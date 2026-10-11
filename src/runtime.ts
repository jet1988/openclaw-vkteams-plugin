import type { PluginRuntime } from "openclaw/plugin-sdk/core";
import { createPluginRuntimeStore } from "openclaw/plugin-sdk/runtime-store";
import type { CoreConfig } from "./types.js";

const { setRuntime: setVkTeamsRuntime, getRuntime: getVkTeamsRuntime } =
  createPluginRuntimeStore<PluginRuntime>(
    "VK Teams runtime not initialized - plugin not registered",
  );

export function readVkTeamsRuntimeConfig(runtime: PluginRuntime = getVkTeamsRuntime()): CoreConfig {
  if (typeof runtime.config.current !== "function") {
    throw new Error("OpenClaw runtime does not expose config.current()");
  }
  return runtime.config.current() as unknown as CoreConfig;
}

export { getVkTeamsRuntime, setVkTeamsRuntime };