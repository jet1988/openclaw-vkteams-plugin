import { defineSetupPluginEntry } from "openclaw/plugin-sdk/core";
import { vkTeamsPlugin } from "./src/channel.js";

export default defineSetupPluginEntry(vkTeamsPlugin);