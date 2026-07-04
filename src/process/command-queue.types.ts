export type CommandQueueEnqueueOptions = {
  warnAfterMs?: number;
  onWait?: (waitMs: number, queuedAhead: number) => void;
  taskTimeoutMs?: number;
  taskTimeoutProgressAtMs?: () => number | undefined;
  priority?: "foreground" | "normal" | "background";
  /** Maximum allowed queue depth before rejecting with CommandLaneOverflowError. */
  maxQueueDepth?: number;
};

export type CommandQueueEnqueueFn = <T>(
  task: () => Promise<T>,
  opts?: CommandQueueEnqueueOptions,
) => Promise<T>;
