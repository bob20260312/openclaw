import fs from "node:fs/promises";
import { request as httpRequest } from "node:http";
import path from "node:path";
import {
  resolveAutoFallbackPrimaryProbe,
  resolveAgentConfig,
  resolveAgentDir,
  resolveAgentWorkspaceDir,
  resolveSessionAgentId,
  resolveAgentSkillsFilter,
} from "../../agents/agent-scope.js";
import { modelKey, resolveModelRefFromString } from "../../agents/model-selection.js";
import { resolveAgentTimeoutMs } from "../../agents/timeout.js";
import { DEFAULT_AGENT_WORKSPACE_DIR, ensureAgentWorkspace } from "../../agents/workspace.js";
import { resolveChannelModelOverride } from "../../channels/model-overrides.js";
import { type OpenClawConfig, getRuntimeConfig } from "../../config/config.js";
import { resolveStorePath } from "../../config/sessions/paths.js";
import { updateSessionStoreEntry } from "../../config/sessions/store.js";
import { logVerbose } from "../../globals.js";
import { measureDiagnosticsTimelineSpan } from "../../infra/diagnostics-timeline.js";
import { formatErrorMessage } from "../../infra/errors.js";
import { normalizeAttachments } from "../../media-understanding/attachments.normalize.js";
import { MEDIA_MAX_BYTES, saveMediaBuffer } from "../../media/store.js";
import { buildAgentHookContextChannelFields } from "../../plugins/hook-agent-context.js";
import { defaultRuntime } from "../../runtime.js";
import { createLazyImportLoader } from "../../shared/lazy-promise.js";
import { normalizeOptionalString } from "../../shared/string-coerce.js";
import { normalizeStringEntries } from "../../shared/string-normalization.js";
import { resolveCommandTurnTargetSessionKey } from "../command-turn-context.js";
import type { GetReplyOptions } from "../get-reply-options.types.js";
import { DEFAULT_HEARTBEAT_ACK_MAX_CHARS, stripHeartbeatToken } from "../heartbeat.js";
import type { ReplyPayload } from "../reply-payload.js";
import type { MsgContext } from "../templating.js";
import { normalizeVerboseLevel } from "../thinking.js";
import { SILENT_REPLY_TOKEN } from "../tokens.js";
import { resolveDefaultModel } from "./directive-handling.defaults.js";
import { clearInlineDirectives } from "./get-reply-directives-utils.js";
import { resolveReplyDirectives } from "./get-reply-directives.js";
import {
  initFastReplySessionState,
  buildFastReplyCommandContext,
  shouldHandleFastReplyTextCommands,
  shouldUseReplyFastDirectiveExecution,
  resolveGetReplyConfig,
  shouldUseReplyFastTestBootstrap,
  shouldUseReplyFastTestRuntime,
} from "./get-reply-fast-path.js";
import { handleInlineActions } from "./get-reply-inline-actions.js";
import { maybeResolveNativeSlashCommandFastReply } from "./get-reply-native-slash-fast-path.js";
import { runPreparedReply } from "./get-reply-run.js";
import { finalizeInboundContext } from "./inbound-context.js";
import { hasInboundMedia } from "./inbound-media.js";
import { emitPreAgentMessageHooks } from "./message-preprocess-hooks.js";
import { createFastTestModelSelectionState, createModelSelectionState } from "./model-selection.js";
import { sanitizePendingFinalDeliveryText } from "./pending-final-delivery.js";
import { initSessionState } from "./session.js";
import {
  isStaleHeartbeatAutoFallbackOverride,
  resolveStoredModelOverride,
} from "./stored-model-override.js";
import { createTypingController } from "./typing.js";

const CONTRACT_ANALYST_ID = "contract-analyst";
const CONTRACT_REVIEW_PATH = `/api/v1/agents/${CONTRACT_ANALYST_ID}/contract-review`;
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const CONTRACT_REVIEW_INTENT_PATTERNS = [
  "审查",
  "审核",
  "审一下",
  "风险",
  "高亮",
  "问题清单",
  "修改建议",
  "红线",
  "批注",
  "review",
  "revise",
  "highlight",
  "check contract",
];

type ContractReviewApiSuccess = {
  summary: string;
  highlighted_filename: string;
  highlighted_download_url: string;
  severity_counts: Record<string, number>;
  findings: Array<{
    severity: string;
    clause_title: string;
    rule_id: string;
    evidence_text: string;
    comment: string;
    suggestion: string;
  }>;
};

function normalizeContractIntentText(ctx: MsgContext): string {
  return (
    normalizeOptionalString(ctx.BodyForAgent) ??
    normalizeOptionalString(ctx.CommandBody) ??
    normalizeOptionalString(ctx.RawBody) ??
    normalizeOptionalString(ctx.Body) ??
    ""
  ).trim();
}

function hasAnyKeyword(text: string, keywords: readonly string[]): boolean {
  const normalized = text.trim().toLowerCase();
  if (!normalized) {
    return false;
  }
  return keywords.some((keyword) => normalized.includes(keyword.toLowerCase()));
}

function isContractReviewIntent(text: string): boolean {
  return hasAnyKeyword(text, CONTRACT_REVIEW_INTENT_PATTERNS);
}

function isEmptyContractFollowupText(text: string): boolean {
  return text.trim().length === 0 || text.trim() === "[用户上传了附件]";
}

async function maybeHandleContractAnalystUploadAck(params: {
  agentId: string;
  ctx: MsgContext;
  sessionKey?: string;
  cfg: OpenClawConfig;
}): Promise<ReplyPayload | undefined> {
  if (params.agentId !== CONTRACT_ANALYST_ID) {
    return undefined;
  }
  const text = normalizeContractIntentText(params.ctx);
  if (!isEmptyContractFollowupText(text)) {
    return undefined;
  }
  const attachments = normalizeAttachments(params.ctx);
  const inboundDocx = attachments.find((attachment) =>
    isDocxPath(attachment.path, attachment.mime),
  );
  const inboundDocxPath = normalizeOptionalString(inboundDocx?.path);
  if (!inboundDocxPath) {
    return undefined;
  }
  const resolvedInboundDocxPath = resolveContractAttachmentPath(
    inboundDocxPath,
    normalizeOptionalString(params.ctx.MediaWorkspaceDir) ?? undefined,
  );
  const storePath = params.sessionKey
    ? resolveStorePath(params.cfg.session?.store, { agentId: params.agentId })
    : undefined;
  await rememberRecentContractAttachment({
    storePath,
    sessionKey: params.sessionKey,
    filePath: resolvedInboundDocxPath,
    fileName: path.basename(resolvedInboundDocxPath),
    mimeType: normalizeOptionalString(inboundDocx?.mime) ?? undefined,
    sourceChannel:
      normalizeOptionalString(params.ctx.OriginatingChannel ?? params.ctx.Provider) ?? undefined,
  });
  return {
    text: "我已收到文件，请描述你的需求吧。",
  };
}

function isDocxPath(filePath?: string, mimeType?: string): boolean {
  const normalizedPath = normalizeOptionalString(filePath)?.toLowerCase();
  const normalizedMime = normalizeOptionalString(mimeType)?.toLowerCase();
  return normalizedPath?.endsWith(".docx") === true || normalizedMime === DOCX_MIME;
}

function resolveContractAttachmentPath(filePath: string, mediaWorkspaceDir?: string): string {
  if (path.isAbsolute(filePath)) {
    return filePath;
  }
  if (mediaWorkspaceDir) {
    return path.resolve(mediaWorkspaceDir, filePath);
  }
  return path.resolve(filePath);
}

async function rememberRecentContractAttachment(params: {
  storePath?: string;
  sessionKey?: string;
  filePath: string;
  fileName: string;
  mimeType?: string;
  sourceChannel?: string;
}): Promise<void> {
  if (!params.storePath || !params.sessionKey) {
    return;
  }
  await updateSessionStoreEntry({
    storePath: params.storePath,
    sessionKey: params.sessionKey,
    update: async () => ({
      recentContractAttachment: {
        path: params.filePath,
        fileName: params.fileName,
        capturedAt: Date.now(),
        ...(params.mimeType ? { mimeType: params.mimeType } : {}),
        ...(params.sourceChannel ? { sourceChannel: params.sourceChannel } : {}),
      },
    }),
  });
}

function buildContractReviewSummaryReply(params: {
  result: ContractReviewApiSuccess;
  highlightedMediaPath: string;
}): ReplyPayload {
  const lines: string[] = [
    "合同审查已完成。",
    params.result.summary,
    "",
    `高风险：${params.result.severity_counts.HIGH ?? 0}`,
    `中风险：${params.result.severity_counts.MEDIUM ?? 0}`,
    `低风险：${params.result.severity_counts.LOW ?? 0}`,
  ];
  if (params.result.findings.length > 0) {
    lines.push("", "问题清单：");
    for (const [index, finding] of params.result.findings.entries()) {
      lines.push(
        `${index + 1}. [${finding.severity}] ${finding.clause_title} / ${finding.rule_id}`,
      );
      lines.push(`证据：${finding.evidence_text}`);
      lines.push(`问题：${finding.comment}`);
      lines.push(`建议：${finding.suggestion}`);
      lines.push("");
    }
  }
  return {
    text: lines.join("\n").trim(),
    mediaUrl: params.highlightedMediaPath,
  };
}

async function postMultipartContractReview(params: {
  fileBuffer: Buffer;
  fileName: string;
  note: string;
}): Promise<ContractReviewApiSuccess> {
  const baseUrl =
    normalizeOptionalString(process.env.AGENT_CHAT_SERVICE_URL) ?? "http://agent-chat-service:8104";
  const endpoint = new URL(CONTRACT_REVIEW_PATH, baseUrl);
  const boundary = `----openclaw-contract-review-${Date.now().toString(16)}`;
  const fileHeader = Buffer.from(
    `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="file"; filename="${params.fileName}"\r\n` +
      `Content-Type: ${DOCX_MIME}\r\n\r\n`,
    "utf8",
  );
  const notePart = Buffer.from(
    `\r\n--${boundary}\r\n` +
      `Content-Disposition: form-data; name="note"\r\n\r\n` +
      `${params.note}\r\n` +
      `--${boundary}--\r\n`,
    "utf8",
  );
  const body = Buffer.concat([fileHeader, params.fileBuffer, notePart]);
  return await new Promise<ContractReviewApiSuccess>((resolve, reject) => {
    const req = httpRequest(
      endpoint,
      {
        method: "POST",
        headers: {
          "Content-Type": `multipart/form-data; boundary=${boundary}`,
          "Content-Length": String(body.byteLength),
        },
      },
      (res) => {
        const chunks: Buffer[] = [];
        res.on("data", (chunk) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });
        res.on("end", () => {
          const raw = Buffer.concat(chunks).toString("utf8");
          if ((res.statusCode ?? 500) >= 400) {
            reject(new Error(`contract-review failed (${res.statusCode ?? 500}): ${raw}`));
            return;
          }
          try {
            resolve(JSON.parse(raw) as ContractReviewApiSuccess);
          } catch (error) {
            reject(error);
          }
        });
      },
    );
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

async function downloadContractArtifact(params: {
  downloadUrl: string;
  fileName: string;
}): Promise<string> {
  const baseUrl =
    normalizeOptionalString(process.env.AGENT_CHAT_SERVICE_URL) ?? "http://agent-chat-service:8104";
  const artifactUrl = new URL(params.downloadUrl, baseUrl);
  const buffer = await new Promise<Buffer>((resolve, reject) => {
    const req = httpRequest(artifactUrl, { method: "GET" }, (res) => {
      const chunks: Buffer[] = [];
      res.on("data", (chunk) => {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      });
      res.on("end", () => {
        if ((res.statusCode ?? 500) >= 400) {
          reject(new Error(`contract artifact download failed (${res.statusCode ?? 500})`));
          return;
        }
        resolve(Buffer.concat(chunks));
      });
    });
    req.on("error", reject);
    req.end();
  });
  const saved = await saveMediaBuffer(
    buffer,
    DOCX_MIME,
    "outbound",
    MEDIA_MAX_BYTES,
    params.fileName,
    params.fileName,
  );
  return saved.path;
}

async function maybeHandleContractAnalystReview(params: {
  agentId: string;
  ctx: MsgContext;
  sessionEntry: {
    recentContractAttachment?: {
      path: string;
      fileName: string;
      mimeType?: string;
    };
  };
  sessionKey?: string;
  storePath?: string;
}): Promise<ReplyPayload | undefined> {
  if (params.agentId !== CONTRACT_ANALYST_ID) {
    return undefined;
  }
  const text = normalizeContractIntentText(params.ctx);
  const attachments = normalizeAttachments(params.ctx);
  const inboundDocx = attachments.find((attachment) =>
    isDocxPath(attachment.path, attachment.mime),
  );
  const inboundDocxPath = normalizeOptionalString(inboundDocx?.path);
  const resolvedInboundDocxPath = inboundDocxPath
    ? resolveContractAttachmentPath(
        inboundDocxPath,
        normalizeOptionalString(params.ctx.MediaWorkspaceDir) ?? undefined,
      )
    : undefined;

  if (resolvedInboundDocxPath) {
    await rememberRecentContractAttachment({
      storePath: params.storePath,
      sessionKey: params.sessionKey,
      filePath: resolvedInboundDocxPath,
      fileName: path.basename(resolvedInboundDocxPath),
      mimeType: normalizeOptionalString(inboundDocx?.mime) ?? undefined,
      sourceChannel:
        normalizeOptionalString(params.ctx.OriginatingChannel ?? params.ctx.Provider) ?? undefined,
    });
  }

  const recentAttachment = resolvedInboundDocxPath
    ? {
        path: resolvedInboundDocxPath,
        fileName: path.basename(resolvedInboundDocxPath),
        mimeType: normalizeOptionalString(inboundDocx?.mime) ?? undefined,
      }
    : params.sessionEntry.recentContractAttachment;

  if (!recentAttachment) {
    if (!text) {
      return {
        text: "请先上传一份 .docx 合同文件，再告诉我你想审查、摘要还是解释其中的条款。",
      };
    }
    if (!isContractReviewIntent(text)) {
      return undefined;
    }
    return {
      text: "请先上传一份 .docx 合同文件，我才能继续做合同摘要、条款解释或风险审查。",
    };
  }

  if (!isContractReviewIntent(text)) {
    return undefined;
  }

  try {
    const fileBuffer = await fs.readFile(recentAttachment.path);
    const reviewResult = await postMultipartContractReview({
      fileBuffer,
      fileName: recentAttachment.fileName,
      note: text,
    });
    const highlightedMediaPath = await downloadContractArtifact({
      downloadUrl: reviewResult.highlighted_download_url,
      fileName: reviewResult.highlighted_filename,
    });
    return buildContractReviewSummaryReply({
      result: reviewResult,
      highlightedMediaPath,
    });
  } catch (error) {
    return {
      text:
        `合同审查暂时失败：${formatErrorMessage(error)}\n` +
        "如果这是之前上传的文件，请重新上传一次 .docx 合同后再试。",
    };
  }
}

type ResetCommandAction = "new" | "reset";

function classifyHeartbeatPendingFinalDelivery(text: string, ackMaxChars: number) {
  const stripped = stripHeartbeatToken(text, {
    mode: "heartbeat",
    maxAckChars: ackMaxChars,
  });
  return {
    shouldClear: stripped.shouldSkip,
    replayText: stripped.didStrip && stripped.text ? stripped.text : text,
  };
}

function resolveHeartbeatAckMaxChars(cfg: OpenClawConfig, agentId: string): number {
  const agentHeartbeat = resolveAgentConfig(cfg, agentId)?.heartbeat;
  return Math.max(
    0,
    agentHeartbeat?.ackMaxChars ??
      cfg.agents?.defaults?.heartbeat?.ackMaxChars ??
      DEFAULT_HEARTBEAT_ACK_MAX_CHARS,
  );
}

const sessionResetModelRuntimeLoader = createLazyImportLoader(
  () => import("./session-reset-model.runtime.js"),
);
const stageSandboxMediaRuntimeLoader = createLazyImportLoader(
  () => import("./stage-sandbox-media.runtime.js"),
);
const mediaUnderstandingApplyRuntimeLoader = createLazyImportLoader(
  () => import("../../media-understanding/apply.runtime.js"),
);
const linkUnderstandingApplyRuntimeLoader = createLazyImportLoader(
  () => import("../../link-understanding/apply.runtime.js"),
);
const commandsCoreRuntimeLoader = createLazyImportLoader(
  () => import("./commands-core.runtime.js"),
);

function loadSessionResetModelRuntime() {
  return sessionResetModelRuntimeLoader.load();
}

function loadStageSandboxMediaRuntime() {
  return stageSandboxMediaRuntimeLoader.load();
}

function loadMediaUnderstandingApplyRuntime() {
  return mediaUnderstandingApplyRuntimeLoader.load();
}

function loadLinkUnderstandingApplyRuntime() {
  return linkUnderstandingApplyRuntimeLoader.load();
}

function loadCommandsCoreRuntime() {
  return commandsCoreRuntimeLoader.load();
}

const hookRunnerGlobalLoader = createLazyImportLoader(
  () => import("../../plugins/hook-runner-global.js"),
);
const originRoutingLoader = createLazyImportLoader(() => import("./origin-routing.js"));

function loadHookRunnerGlobal() {
  return hookRunnerGlobalLoader.load();
}

function loadOriginRouting() {
  return originRoutingLoader.load();
}

function mergeSkillFilters(channelFilter?: string[], agentFilter?: string[]): string[] | undefined {
  const normalize = (list?: string[]) => {
    if (!Array.isArray(list)) {
      return undefined;
    }
    return normalizeStringEntries(list);
  };
  const channel = normalize(channelFilter);
  const agent = normalize(agentFilter);
  if (!channel && !agent) {
    return undefined;
  }
  if (!channel) {
    return agent;
  }
  if (!agent) {
    return channel;
  }
  if (channel.length === 0 || agent.length === 0) {
    return [];
  }
  const agentSet = new Set(agent);
  return channel.filter((name) => agentSet.has(name));
}

function hasLinkCandidate(ctx: MsgContext): boolean {
  const message = ctx.BodyForCommands ?? ctx.CommandBody ?? ctx.RawBody ?? ctx.Body;
  if (!message) {
    return false;
  }
  return /\bhttps?:\/\/\S+/i.test(message);
}

async function applyMediaUnderstandingIfNeeded(params: {
  ctx: MsgContext;
  cfg: OpenClawConfig;
  agentDir?: string;
  workspaceDir?: string;
  activeModel: { provider: string; model: string };
}): Promise<boolean> {
  if (!hasInboundMedia(params.ctx)) {
    return false;
  }
  try {
    const { applyMediaUnderstanding } = await loadMediaUnderstandingApplyRuntime();
    await applyMediaUnderstanding(params);
    return true;
  } catch (err) {
    mediaUnderstandingApplyRuntimeLoader.clear();
    logVerbose(
      `media understanding failed, proceeding with raw content: ${formatErrorMessage(err)}`,
    );
    return false;
  }
}

async function applyLinkUnderstandingIfNeeded(params: {
  ctx: MsgContext;
  cfg: OpenClawConfig;
}): Promise<boolean> {
  if (!hasLinkCandidate(params.ctx)) {
    return false;
  }
  try {
    const { applyLinkUnderstanding } = await loadLinkUnderstandingApplyRuntime();
    await applyLinkUnderstanding(params);
    return true;
  } catch (err) {
    linkUnderstandingApplyRuntimeLoader.clear();
    logVerbose(
      `link understanding failed, proceeding with raw content: ${formatErrorMessage(err)}`,
    );
    return false;
  }
}

export async function getReplyFromConfig(
  ctx: MsgContext,
  opts?: GetReplyOptions,
  configOverride?: OpenClawConfig,
): Promise<ReplyPayload | ReplyPayload[] | undefined> {
  const isFastTestEnv = process.env.OPENCLAW_TEST_FAST === "1";
  const cfg = resolveGetReplyConfig({
    getRuntimeConfig,
    isFastTestEnv,
    configOverride,
  });
  const useFastTestBootstrap = shouldUseReplyFastTestBootstrap({
    isFastTestEnv,
    configOverride,
  });
  const useFastTestRuntime = shouldUseReplyFastTestRuntime({
    cfg,
    isFastTestEnv,
  });
  const finalized = finalizeInboundContext(ctx);
  const targetSessionKey = resolveCommandTurnTargetSessionKey(finalized);
  const agentSessionKey = targetSessionKey || finalized.SessionKey;
  const traceAttributes = {
    surface: normalizeOptionalString(finalized.Surface ?? finalized.Provider) ?? "unknown",
    hasSessionKey: Boolean(agentSessionKey),
    isHeartbeat: opts?.isHeartbeat === true,
    hasMedia: hasInboundMedia(finalized),
  };
  const traceGetReplyPhase = <T>(name: string, run: () => Promise<T> | T): Promise<T> =>
    measureDiagnosticsTimelineSpan(name, run, {
      phase: "agent-turn",
      config: cfg,
      attributes: traceAttributes,
    });
  const agentId = resolveSessionAgentId({
    sessionKey: agentSessionKey,
    config: cfg,
  });
  const mergedSkillFilter = mergeSkillFilters(
    opts?.skillFilter,
    resolveAgentSkillsFilter(cfg, agentId),
  );
  const resolvedOpts =
    mergedSkillFilter !== undefined ? { ...opts, skillFilter: mergedSkillFilter } : opts;
  const agentCfg = cfg.agents?.defaults;
  const sessionCfg = cfg.session;
  const { defaultProvider, defaultModel, aliasIndex } = resolveDefaultModel({
    cfg,
    agentId,
  });
  let provider = defaultProvider;
  let model = defaultModel;
  let hasResolvedHeartbeatModelOverride = false;
  let hasAppliedImageModelOverride = false;
  let imageModelFallbacksOverride: string[] | undefined;
  const modelOverrideRaw = normalizeOptionalString(opts?.modelOverride);
  if (modelOverrideRaw) {
    const modelOverrideRef = resolveModelRefFromString({
      raw: modelOverrideRaw,
      defaultProvider,
      aliasIndex,
    });
    if (modelOverrideRef) {
      provider = modelOverrideRef.ref.provider;
      model = modelOverrideRef.ref.model;
      hasAppliedImageModelOverride = true;
      imageModelFallbacksOverride = opts?.modelOverrideFallbacks?.filter(
        (fallback): fallback is string => normalizeOptionalString(fallback) !== undefined,
      );
    } else {
      defaultRuntime.log?.(
        `[image-model-switch] Failed to resolve image model override ${modelOverrideRaw}; using default model ${modelKey(defaultProvider, defaultModel)}`,
      );
    }
  } else if (opts?.isHeartbeat) {
    // Prefer the resolved per-agent heartbeat model passed from the heartbeat runner,
    // fall back to the global defaults heartbeat model for backward compatibility.
    const heartbeatRaw =
      normalizeOptionalString(opts.heartbeatModelOverride) ??
      normalizeOptionalString(agentCfg?.heartbeat?.model) ??
      "";
    const heartbeatRef = heartbeatRaw
      ? resolveModelRefFromString({
          raw: heartbeatRaw,
          defaultProvider,
          aliasIndex,
        })
      : null;
    if (heartbeatRef) {
      provider = heartbeatRef.ref.provider;
      model = heartbeatRef.ref.model;
      hasResolvedHeartbeatModelOverride = true;
    }
  }

  const workspaceDirRaw = resolveAgentWorkspaceDir(cfg, agentId) ?? DEFAULT_AGENT_WORKSPACE_DIR;
  const workspaceDirForNativeCommand = workspaceDirRaw;
  const agentDir = resolveAgentDir(cfg, agentId);
  const timeoutMs = resolveAgentTimeoutMs({ cfg, overrideSeconds: opts?.timeoutOverrideSeconds });
  const configuredTypingSeconds =
    agentCfg?.typingIntervalSeconds ?? sessionCfg?.typingIntervalSeconds;
  const typingIntervalSeconds =
    typeof configuredTypingSeconds === "number" ? configuredTypingSeconds : 6;
  const typing = createTypingController({
    onReplyStart: opts?.onReplyStart,
    onCleanup: opts?.onTypingCleanup,
    typingIntervalSeconds,
    silentToken: SILENT_REPLY_TOKEN,
    log: defaultRuntime.log,
  });
  opts?.onTypingController?.(typing);

  const nativeSlashCommandFastReply = await traceGetReplyPhase(
    "reply.native_slash_command_fast_path",
    () =>
      maybeResolveNativeSlashCommandFastReply({
        ctx: finalized,
        cfg,
        agentId,
        agentDir,
        agentCfg,
        commandAuthorized: finalized.CommandAuthorized,
        defaultProvider,
        defaultModel,
        aliasIndex,
        provider,
        model,
        workspaceDir: workspaceDirForNativeCommand,
        typing,
        opts: resolvedOpts,
        skillFilter: mergedSkillFilter,
      }),
  );
  if (nativeSlashCommandFastReply.handled) {
    return nativeSlashCommandFastReply.reply;
  }

  const contractUploadAckReply = await traceGetReplyPhase("reply.contract_upload_ack", () =>
    maybeHandleContractAnalystUploadAck({
      agentId,
      ctx: finalized,
      sessionKey: agentSessionKey ?? undefined,
      cfg,
    }),
  );
  if (contractUploadAckReply) {
    return contractUploadAckReply;
  }

  const workspace = await traceGetReplyPhase("reply.ensure_workspace", async () =>
    useFastTestBootstrap
      ? (await fs.mkdir(workspaceDirRaw, { recursive: true }), { dir: workspaceDirRaw })
      : await ensureAgentWorkspace({
          dir: workspaceDirRaw,
          ensureBootstrapFiles: !agentCfg?.skipBootstrap && !isFastTestEnv,
          skipOptionalBootstrapFiles: agentCfg?.skipOptionalBootstrapFiles,
        }),
  );
  const workspaceDir = workspace.dir;

  if (!isFastTestEnv && hasInboundMedia(finalized)) {
    await traceGetReplyPhase("reply.apply_media_understanding", () =>
      applyMediaUnderstandingIfNeeded({
        ctx: finalized,
        cfg,
        agentDir,
        workspaceDir,
        activeModel: { provider, model },
      }),
    );
  }
  if (!isFastTestEnv && hasLinkCandidate(finalized)) {
    await traceGetReplyPhase("reply.apply_link_understanding", () =>
      applyLinkUnderstandingIfNeeded({
        ctx: finalized,
        cfg,
      }),
    );
  }
  emitPreAgentMessageHooks({
    ctx: finalized,
    cfg,
    isFastTestEnv,
  });

  const commandAuthorized = finalized.CommandAuthorized;
  const sessionState = useFastTestBootstrap
    ? initFastReplySessionState({
        ctx: finalized,
        cfg,
        agentId,
        commandAuthorized,
        workspaceDir,
      })
    : await traceGetReplyPhase("reply.init_session_state", () =>
        initSessionState({
          ctx: finalized,
          cfg,
          commandAuthorized,
        }),
      );
  let {
    sessionCtx,
    sessionEntry,
    previousSessionEntry,
    sessionStore,
    sessionKey,
    sessionId,
    isNewSession,
    resetTriggered,
    systemSent,
    abortedLastRun,
    storePath,
    sessionScope,
    groupResolution,
    isGroup,
    triggerBodyNormalized,
    bodyStripped,
  } = sessionState;

  if (sessionEntry?.pendingFinalDelivery && sessionEntry.pendingFinalDeliveryText) {
    const text = sanitizePendingFinalDeliveryText(sessionEntry.pendingFinalDeliveryText);

    // If it's a heartbeat, we definitely want to try delivering the lost reply now.
    // If it's a user message, we deliver the lost reply first, then continue.
    // For now, let's just return the lost reply if it's a heartbeat.
    if (opts?.isHeartbeat) {
      const heartbeatPending = classifyHeartbeatPendingFinalDelivery(
        text,
        resolveHeartbeatAckMaxChars(cfg, agentId),
      );
      if (heartbeatPending.shouldClear) {
        sessionEntry.pendingFinalDelivery = undefined;
        sessionEntry.pendingFinalDeliveryText = undefined;
        sessionEntry.pendingFinalDeliveryCreatedAt = undefined;
        sessionEntry.pendingFinalDeliveryLastAttemptAt = undefined;
        sessionEntry.pendingFinalDeliveryAttemptCount = undefined;
        sessionEntry.pendingFinalDeliveryLastError = undefined;
        sessionEntry.pendingFinalDeliveryContext = undefined;
        if (sessionKey && sessionStore) {
          sessionStore[sessionKey] = sessionEntry;
        }
        if (sessionKey && storePath) {
          const { updateSessionStoreEntry } = await import("../../config/sessions.js");
          await updateSessionStoreEntry({
            storePath,
            sessionKey,
            update: async () => ({
              pendingFinalDelivery: undefined,
              pendingFinalDeliveryText: undefined,
              pendingFinalDeliveryCreatedAt: undefined,
              pendingFinalDeliveryLastAttemptAt: undefined,
              pendingFinalDeliveryAttemptCount: undefined,
              pendingFinalDeliveryLastError: undefined,
              pendingFinalDeliveryContext: undefined,
            }),
          });
        }
      } else {
        const updatedAt = Date.now();
        const attemptCount = (sessionEntry.pendingFinalDeliveryAttemptCount ?? 0) + 1;
        sessionEntry.pendingFinalDeliveryLastAttemptAt = updatedAt;
        sessionEntry.pendingFinalDeliveryAttemptCount = attemptCount;
        sessionEntry.pendingFinalDeliveryLastError = null;
        const replayText = sanitizePendingFinalDeliveryText(heartbeatPending.replayText);
        sessionEntry.pendingFinalDeliveryText = replayText;
        sessionEntry.updatedAt = updatedAt;
        if (sessionKey && sessionStore) {
          sessionStore[sessionKey] = sessionEntry;
        }
        if (sessionKey && storePath) {
          const { updateSessionStoreEntry } = await import("../../config/sessions.js");
          await updateSessionStoreEntry({
            storePath,
            sessionKey,
            update: async () => ({
              pendingFinalDeliveryText: replayText,
              pendingFinalDeliveryLastAttemptAt: updatedAt,
              pendingFinalDeliveryAttemptCount: attemptCount,
              pendingFinalDeliveryLastError: null,
              updatedAt,
            }),
          });
        }
        return { text: replayText };
      }
    }
  }

  if (resetTriggered && normalizeOptionalString(bodyStripped)) {
    const { applyResetModelOverride } = await loadSessionResetModelRuntime();
    await applyResetModelOverride({
      cfg,
      agentId,
      resetTriggered,
      bodyStripped,
      sessionCtx,
      ctx: finalized,
      sessionEntry,
      sessionStore,
      sessionKey,
      storePath,
      defaultProvider,
      defaultModel,
      aliasIndex,
    });
  }

  const contractReviewReply = await traceGetReplyPhase("reply.contract_review_bridge", () =>
    maybeHandleContractAnalystReview({
      agentId,
      ctx: finalized,
      sessionEntry,
      sessionKey,
      storePath,
    }),
  );
  if (contractReviewReply) {
    return contractReviewReply;
  }

  const channelModelOverride = cfg.channels?.modelByChannel
    ? resolveChannelModelOverride({
        cfg,
        channel:
          groupResolution?.channel ??
          sessionEntry.channel ??
          sessionEntry.origin?.provider ??
          (typeof finalized.OriginatingChannel === "string"
            ? finalized.OriginatingChannel
            : undefined) ??
          finalized.Provider,
        groupId: groupResolution?.id ?? sessionEntry.groupId,
        groupChatType: sessionEntry.chatType ?? sessionCtx.ChatType ?? finalized.ChatType,
        groupChannel:
          sessionEntry.groupChannel ?? sessionCtx.GroupChannel ?? finalized.GroupChannel,
        groupSubject: sessionEntry.subject ?? sessionCtx.GroupSubject ?? finalized.GroupSubject,
        parentSessionKey: sessionCtx.ModelParentSessionKey ?? sessionCtx.ParentSessionKey,
      })
    : null;
  const resolvedChannelModelOverride =
    channelModelOverride && !hasResolvedHeartbeatModelOverride
      ? resolveModelRefFromString({
          raw: channelModelOverride.model,
          defaultProvider,
          aliasIndex,
        })
      : null;
  const primaryProvider = resolvedChannelModelOverride?.ref.provider ?? defaultProvider;
  const primaryModel = resolvedChannelModelOverride?.ref.model ?? defaultModel;
  const hasSessionModelOverride = Boolean(
    normalizeOptionalString(sessionEntry.modelOverride) ||
    normalizeOptionalString(sessionEntry.providerOverride),
  );
  const storedModelOverride = resolveStoredModelOverride({
    sessionEntry,
    sessionStore,
    sessionKey,
    parentSessionKey:
      sessionEntry.parentSessionKey ??
      sessionCtx.ModelParentSessionKey ??
      sessionCtx.ParentSessionKey,
    defaultProvider,
  });
  const staleHeartbeatAutoFallbackOverride = isStaleHeartbeatAutoFallbackOverride({
    isHeartbeat: opts?.isHeartbeat === true,
    hasResolvedHeartbeatModelOverride,
    sessionEntry,
    storedOverride: storedModelOverride,
    defaultProvider,
    defaultModel,
    primaryProvider,
    primaryModel,
  });
  if (
    storedModelOverride?.model &&
    !hasResolvedHeartbeatModelOverride &&
    !hasAppliedImageModelOverride &&
    !staleHeartbeatAutoFallbackOverride
  ) {
    provider = storedModelOverride.provider ?? defaultProvider;
    model = storedModelOverride.model;
  }
  const canApplyAutoFallbackPrimaryProbe =
    !hasResolvedHeartbeatModelOverride &&
    !hasAppliedImageModelOverride &&
    !staleHeartbeatAutoFallbackOverride;
  const autoFallbackPrimaryProbe = canApplyAutoFallbackPrimaryProbe
    ? resolveAutoFallbackPrimaryProbe({
        entry: sessionEntry,
        sessionKey,
        primaryProvider,
        primaryModel,
      })
    : undefined;
  const hasEffectiveSessionModelOverride =
    hasSessionModelOverride && !staleHeartbeatAutoFallbackOverride;
  if (
    !hasResolvedHeartbeatModelOverride &&
    !hasEffectiveSessionModelOverride &&
    !hasAppliedImageModelOverride &&
    resolvedChannelModelOverride
  ) {
    provider = resolvedChannelModelOverride.ref.provider;
    model = resolvedChannelModelOverride.ref.model;
  }
  const imageModelOverrideBaseProvider = hasAppliedImageModelOverride
    ? (() => {
        if (
          storedModelOverride?.model &&
          !hasResolvedHeartbeatModelOverride &&
          !staleHeartbeatAutoFallbackOverride
        ) {
          return storedModelOverride.provider ?? defaultProvider;
        }
        if (!hasEffectiveSessionModelOverride && resolvedChannelModelOverride) {
          return resolvedChannelModelOverride.ref.provider;
        }
        const runtimeProvider = normalizeOptionalString(sessionEntry.modelProvider);
        const runtimeModel = normalizeOptionalString(sessionEntry.model);
        if (runtimeProvider && runtimeModel) {
          return runtimeProvider;
        }
        return defaultProvider;
      })()
    : undefined;

  if (
    shouldUseReplyFastDirectiveExecution({
      isFastTestBootstrap: useFastTestRuntime,
      isGroup,
      isHeartbeat: opts?.isHeartbeat === true,
      resetTriggered,
      triggerBodyNormalized,
    })
  ) {
    const fastCommand = buildFastReplyCommandContext({
      ctx,
      cfg,
      agentId,
      sessionKey,
      isGroup,
      triggerBodyNormalized,
      commandAuthorized,
    });
    return await traceGetReplyPhase("reply.run_prepared_reply", () =>
      runPreparedReply({
        ctx,
        sessionCtx,
        cfg,
        agentId,
        agentDir,
        agentCfg,
        sessionCfg,
        commandAuthorized,
        command: fastCommand,
        commandSource:
          finalized.BodyForCommands ?? finalized.CommandBody ?? finalized.RawBody ?? "",
        allowTextCommands: shouldHandleFastReplyTextCommands({
          cfg,
          commandSource: finalized.CommandSource,
        }),
        directives: clearInlineDirectives(
          finalized.BodyForCommands ?? finalized.CommandBody ?? finalized.RawBody ?? "",
        ),
        defaultActivation: "always",
        resolvedThinkLevel: undefined,
        resolvedVerboseLevel: normalizeVerboseLevel(agentCfg?.verboseDefault),
        resolvedReasoningLevel: "off",
        resolvedElevatedLevel: "off",
        execOverrides: undefined,
        elevatedEnabled: false,
        elevatedAllowed: false,
        blockStreamingEnabled: false,
        blockReplyChunking: undefined,
        resolvedBlockStreamingBreak: "text_end",
        modelState: createFastTestModelSelectionState({
          agentCfg,
          provider: autoFallbackPrimaryProbe?.provider ?? provider,
          model: autoFallbackPrimaryProbe?.model ?? model,
        }),
        provider: autoFallbackPrimaryProbe?.provider ?? provider,
        model: autoFallbackPrimaryProbe?.model ?? model,
        perMessageQueueMode: undefined,
        perMessageQueueOptions: undefined,
        typing,
        opts: resolvedOpts,
        defaultProvider,
        defaultModel,
        timeoutMs,
        isNewSession,
        resetTriggered,
        systemSent,
        sessionEntry,
        sessionStore,
        sessionKey,
        sessionId,
        storePath,
        workspaceDir,
        abortedLastRun,
        hasAppliedImageModelOverride,
        imageModelOverrideBaseProvider,
        imageModelFallbacksOverride,
        autoFallbackPrimaryProbe,
      }),
    );
  }

  const directiveResult = await traceGetReplyPhase("reply.resolve_directives", () =>
    resolveReplyDirectives({
      ctx: finalized,
      cfg,
      agentId,
      agentDir,
      workspaceDir,
      agentCfg,
      sessionCtx,
      sessionEntry,
      sessionStore,
      sessionKey,
      storePath,
      sessionScope,
      groupResolution,
      isGroup,
      triggerBodyNormalized,
      resetTriggered,
      commandAuthorized,
      defaultProvider,
      defaultModel,
      primaryProvider,
      primaryModel,
      aliasIndex,
      provider,
      model,
      hasOneTurnModelOverride: hasAppliedImageModelOverride,
      hasResolvedHeartbeatModelOverride,
      typing,
      opts: resolvedOpts,
      skillFilter: mergedSkillFilter,
    }),
  );
  if (directiveResult.kind === "reply") {
    return directiveResult.reply;
  }

  let {
    commandSource,
    command,
    allowTextCommands,
    skillCommands,
    directives,
    cleanedBody,
    elevatedEnabled,
    elevatedAllowed,
    elevatedFailures,
    defaultActivation,
    resolvedThinkLevel,
    resolvedVerboseLevel,
    resolvedReasoningLevel,
    resolvedElevatedLevel,
    execOverrides,
    blockStreamingEnabled,
    blockReplyChunking,
    resolvedBlockStreamingBreak,
    provider: resolvedProvider,
    model: resolvedModel,
    modelState,
    contextTokens,
    inlineStatusRequested,
    directiveAck,
    perMessageQueueMode,
    perMessageQueueOptions,
  } = directiveResult.result;
  provider = resolvedProvider;
  model = resolvedModel;

  const maybeEmitMissingResetHooks = async () => {
    if (!resetTriggered || !command.isAuthorizedSender || command.resetHookTriggered) {
      return;
    }
    const resetMatch = command.commandBodyNormalized.match(/^\/(new|reset)(?:\s|$)/);
    if (!resetMatch) {
      return;
    }
    const { emitResetCommandHooks } = await loadCommandsCoreRuntime();
    const action: ResetCommandAction = resetMatch[1] === "reset" ? "reset" : "new";
    await emitResetCommandHooks({
      action,
      ctx,
      cfg,
      command,
      sessionKey,
      sessionEntry,
      previousSessionEntry,
      workspaceDir,
    });
  };

  const inlineActionResult = await traceGetReplyPhase("reply.handle_inline_actions", () =>
    handleInlineActions({
      ctx,
      sessionCtx,
      cfg,
      agentId,
      agentDir,
      sessionEntry,
      previousSessionEntry,
      sessionStore,
      sessionKey,
      storePath,
      sessionScope,
      workspaceDir,
      isGroup,
      opts: resolvedOpts,
      typing,
      allowTextCommands,
      inlineStatusRequested,
      command,
      skillCommands,
      directives,
      cleanedBody,
      elevatedEnabled,
      elevatedAllowed,
      elevatedFailures,
      defaultActivation: () => defaultActivation,
      resolvedThinkLevel,
      resolvedVerboseLevel,
      resolvedReasoningLevel,
      resolvedElevatedLevel,
      blockReplyChunking,
      resolvedBlockStreamingBreak,
      resolveDefaultThinkingLevel: modelState.resolveDefaultThinkingLevel,
      provider,
      model,
      contextTokens,
      directiveAck,
      abortedLastRun,
      skillFilter: mergedSkillFilter,
    }),
  );
  if (inlineActionResult.kind === "reply") {
    await maybeEmitMissingResetHooks();
    return inlineActionResult.reply;
  }
  await maybeEmitMissingResetHooks();
  directives = inlineActionResult.directives;
  cleanedBody = inlineActionResult.cleanedBody;
  abortedLastRun = inlineActionResult.abortedLastRun ?? abortedLastRun;
  const runAutoFallbackPrimaryProbe = directives.hasModelDirective
    ? undefined
    : autoFallbackPrimaryProbe;
  const runProvider = runAutoFallbackPrimaryProbe?.provider ?? provider;
  const runModel = runAutoFallbackPrimaryProbe?.model ?? model;
  let runModelState = modelState;
  if (runAutoFallbackPrimaryProbe) {
    runModelState = await createModelSelectionState({
      cfg,
      agentId,
      agentCfg,
      sessionEntry,
      sessionStore,
      sessionKey,
      parentSessionKey:
        sessionEntry.parentSessionKey ??
        sessionCtx.ModelParentSessionKey ??
        sessionCtx.ParentSessionKey,
      storePath,
      defaultProvider,
      defaultModel,
      primaryProvider,
      primaryModel,
      provider: runProvider,
      model: runModel,
      hasModelDirective: false,
      hasOneTurnModelOverride: hasAppliedImageModelOverride,
      skipStoredModelOverride: true,
      hasResolvedHeartbeatModelOverride,
      isHeartbeat: opts?.isHeartbeat === true,
    });
    const hasExplicitThinkLevel =
      resolvedOpts?.thinkingLevelOverride !== undefined ||
      directives.thinkLevel !== undefined ||
      (!directives.clearThinkLevel && sessionEntry.thinkingLevel !== undefined) ||
      agentCfg?.thinkingDefault !== undefined;
    if (!hasExplicitThinkLevel) {
      resolvedThinkLevel = await runModelState.resolveDefaultThinkingLevel();
    }
    const agentEntry = resolveAgentConfig(cfg, agentId);
    const rawSessionReasoningLevel = sessionEntry.reasoningLevel;
    const canUseReasoningState =
      command.isAuthorizedSender ||
      command.senderIsOwner ||
      (Array.isArray(ctx.GatewayClientScopes) &&
        ctx.GatewayClientScopes.includes("operator.admin"));
    const hasExplicitReasoningLevel =
      directives.reasoningLevel !== undefined ||
      (rawSessionReasoningLevel != null && canUseReasoningState) ||
      (rawSessionReasoningLevel != null && !canUseReasoningState) ||
      agentEntry?.reasoningDefault != null ||
      agentCfg?.reasoningDefault != null;
    if (!hasExplicitReasoningLevel && resolvedThinkLevel === "off") {
      resolvedReasoningLevel = await runModelState.resolveDefaultReasoningLevel();
    }
  }

  // Allow plugins to intercept and return a synthetic reply before the LLM runs.
  if (!useFastTestBootstrap) {
    const { getGlobalHookRunner } = await loadHookRunnerGlobal();
    const hookRunner = getGlobalHookRunner();
    if (hookRunner?.hasHooks("before_agent_reply")) {
      const { resolveOriginMessageProvider } = await loadOriginRouting();
      const hookMessageProvider = resolveOriginMessageProvider({
        originatingChannel: sessionCtx.OriginatingChannel,
        provider: sessionCtx.Provider,
      });
      const hookResult = await traceGetReplyPhase("reply.before_agent_reply_hooks", () =>
        hookRunner.runBeforeAgentReply(
          { cleanedBody },
          {
            agentId,
            sessionKey: agentSessionKey,
            sessionId,
            workspaceDir,
            trigger: opts?.isHeartbeat ? "heartbeat" : "user",
            ...buildAgentHookContextChannelFields({
              sessionKey: agentSessionKey,
              messageProvider: hookMessageProvider,
              currentChannelId: sessionCtx.OriginatingTo ?? ctx.OriginatingTo ?? ctx.To,
              messageTo: sessionCtx.OriginatingTo ?? ctx.OriginatingTo ?? ctx.To,
            }),
          },
        ),
      );
      if (hookResult?.handled) {
        return hookResult.reply ?? { text: SILENT_REPLY_TOKEN };
      }
    }
  }

  // ctx.MediaStaged=true means the caller (e.g. chat.send RPC) already staged
  // synchronously so it could surface 5xx before respond(). Skipping here keeps
  // staging a single-call contract instead of relying on relative-path no-op
  // semantics in stageSandboxMedia.
  if (!useFastTestBootstrap && sessionKey && !ctx.MediaStaged && hasInboundMedia(ctx)) {
    const { stageSandboxMedia } = await loadStageSandboxMediaRuntime();
    await traceGetReplyPhase("reply.stage_media", () =>
      stageSandboxMedia({
        ctx,
        sessionCtx,
        cfg,
        sessionKey,
        workspaceDir,
      }),
    );
  }

  return await traceGetReplyPhase("reply.run_prepared_reply", () =>
    runPreparedReply({
      ctx,
      sessionCtx,
      cfg,
      agentId,
      agentDir,
      agentCfg,
      sessionCfg,
      commandAuthorized,
      command,
      commandSource,
      allowTextCommands,
      directives,
      defaultActivation,
      resolvedThinkLevel,
      resolvedVerboseLevel,
      resolvedReasoningLevel,
      resolvedElevatedLevel,
      execOverrides,
      elevatedEnabled,
      elevatedAllowed,
      blockStreamingEnabled,
      blockReplyChunking,
      resolvedBlockStreamingBreak,
      modelState: runModelState,
      provider: runProvider,
      model: runModel,
      perMessageQueueMode,
      perMessageQueueOptions,
      typing,
      opts: resolvedOpts,
      defaultProvider,
      defaultModel,
      timeoutMs,
      isNewSession,
      resetTriggered,
      systemSent,
      sessionEntry,
      sessionStore,
      sessionKey,
      sessionId,
      storePath,
      workspaceDir,
      abortedLastRun,
      hasAppliedImageModelOverride,
      imageModelOverrideBaseProvider,
      imageModelFallbacksOverride,
      autoFallbackPrimaryProbe: runAutoFallbackPrimaryProbe,
    }),
  );
}
