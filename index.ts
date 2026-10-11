import type { ChannelPlugin } from "openclaw/plugin-sdk/core";
import { defineChannelPluginEntry } from "openclaw/plugin-sdk/core";
import { vkTeamsPlugin } from "./src/channel.js";
import { setVkTeamsRuntime } from "./src/runtime.js";

export { vkTeamsPlugin } from "./src/channel.js";
export { setVkTeamsRuntime } from "./src/runtime.js";

export default defineChannelPluginEntry({
  id: "vkteams",
  name: "VK Teams",
  description: "VK Teams (Среда) channel plugin",
  plugin: vkTeamsPlugin as ChannelPlugin,
  setRuntime: setVkTeamsRuntime,
});