import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  buildGetReplyCtx,
  createGetReplySessionState,
  registerGetReplyRuntimeOverrides,
} from "./get-reply.test-fixtures.js";
import { loadGetReplyModuleForTest } from "./get-reply.test-loader.js";
import "./get-reply.test-runtime-mocks.js";

const mocks = vi.hoisted(() => ({
  resolveReplyDirectives: vi.fn(),
  initSessionState: vi.fn(),
}));

registerGetReplyRuntimeOverrides(mocks);

let getReplyFromConfig: typeof import("./get-reply.js").getReplyFromConfig;

async function loadRuntime() {
  ({ getReplyFromConfig } = await loadGetReplyModuleForTest({
    cacheKey: import.meta.url,
    fresh: true,
  }));
}

describe("contract analyst upload ack", () => {
  beforeEach(async () => {
    await loadRuntime();
    vi.stubEnv("OPENCLAW_ALLOW_SLOW_REPLY_TESTS", "1");
    mocks.resolveReplyDirectives.mockReset();
    mocks.initSessionState.mockReset();
    mocks.resolveReplyDirectives.mockResolvedValue({ kind: "reply", reply: { text: "ok" } });
    mocks.initSessionState.mockResolvedValue(
      createGetReplySessionState({
        sessionKey: "agent:contract-analyst:wechat:direct:1",
        storePath: "/tmp/sessions.json",
      }),
    );
  });

  it("returns a fixed acknowledgement immediately when a docx is uploaded without a text request", async () => {
    const reply = await getReplyFromConfig(
      buildGetReplyCtx({
        SessionKey: "agent:contract-analyst:wechat:direct:1",
        Body: "[用户上传了附件]",
        BodyForAgent: "[用户上传了附件]",
        RawBody: "[用户上传了附件]",
        CommandBody: "[用户上传了附件]",
        MediaPaths: ["风险合同---abc.docx"],
        MediaTypes: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
        MediaWorkspaceDir: "/tmp/channel-media/default/inbound",
      }),
      undefined,
      {
        session: { store: "/tmp/sessions.json" },
      } as never,
    );

    expect(reply).toEqual({
      text: "我已收到文件，请描述你的需求吧。",
    });
    expect(mocks.initSessionState).not.toHaveBeenCalled();
    expect(mocks.resolveReplyDirectives).not.toHaveBeenCalled();
  });
});
