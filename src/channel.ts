// VK Teams (Среда) channel plugin for OpenClaw.
// Long-polling receive path + sendText/sendFile/sendVoice outbound.

import { buildChannelConfigSchema } from "openclaw/plugin-sdk/channel-config-schema";
import {
  buildComputedAccountStatusSnapshot,
  buildTokenChannelStatusSummary,
} from "openclaw/plugin-sdk/channel-status";
import { createAccountStatusSink } from "openclaw/plugin-sdk/channel-outbound";
import { DEFAULT_ACCOUNT_ID } from "openclaw/plugin-sdk/account-id";
import type { ChannelStatusIssue } from "openclaw/plugin-sdk/channel-contract";
import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { ChannelPlugin } from "openclaw/plugin-sdk/core";
import {
  listVkTeamsAccountIds,
  resolveDefaultVkTeamsAccountId,
  resolveVkTeamsAccount,
  type ResolvedVkTeamsAccount,
} from "./accounts.js";
import { VkTeamsConfigSchema } from "./config-schema.js";
import { monitorVkTeamsProvider } from "./monitor.js";
import { probeVkTeamsBot } from "./probe.js";
import { getVkTeamsRuntime } from "./runtime.js";
import {
  deleteMessageVkTeams,
  editMessageVkTeams,
  sendFormattedMediaVkTeams,
  sendFormattedTextVkTeams,
  sendMessageVkTeams,
  sendPayloadVkTeams,
} from "./send.js";
import type { CoreConfig, VkTeamsProbe } from "./types.js";

const meta = {
  id: "vkteams",
  label: "VK Teams",
  selectionLabel: "VK Teams (Среда)",
  detailLabel: "VK Teams Bot",
  docsPath: "/channels/vkteams",
  docsLabel: "vkteams",
  blurb: "Общение с агентом через VK Teams (Среда).",
  systemImage: "message.fill",
} as const;

const CHANNEL_ID = "vkteams";
const VKTEAMS_DM_APPROVE_HINT = "openclaw pairing approve vkteams <code>";

function normalizeVkTeamsDmAllowEntry(raw: string): string {
  return raw.replace(/^vkteams:/i, "");
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

const vkTeamsSecurityAdapter = {
  resolveDmPolicy: ({
    account,
  }: {
    cfg: OpenClawConfig;
    accountId?: string | null;
    account: ResolvedVkTeamsAccount;
  }) => ({
    policy: account.config.dmPolicy ?? "pairing",
    allowFrom: account.config.allowFrom ?? [],
    allowFromPath: `channels.${CHANNEL_ID}.`,
    approveHint: VKTEAMS_DM_APPROVE_HINT,
    normalizeEntry: normalizeVkTeamsDmAllowEntry,
  }),
};

export const vkTeamsPlugin: ChannelPlugin<ResolvedVkTeamsAccount, VkTeamsProbe> = {
  id: CHANNEL_ID,
  meta: {
    ...meta,
    quickstartAllowFrom: true,
  },
  pairing: {
    idLabel: "vkTeamsUserId",
    normalizeAllowEntry: (entry) => {
      return entry.replace(/^vkteams:/i, "");
    },
    notifyApproval: async ({ cfg, id, accountId }) => {
      await sendMessageVkTeams(
        id,
        "OpenClaw: ваш доступ подтверждён.",
        { cfg: cfg as CoreConfig, accountId },
      );
    },
  },
  allowlist: {
    supportsScope: ({ scope }) => scope === "dm" || scope === "all",
    readConfig: ({ cfg, accountId }) => {
      const account = resolveVkTeamsAccount({
        cfg: cfg as CoreConfig,
        accountId,
      });
      return account.config.allowFrom ?? [];
    },
    applyConfigEdit: ({ cfg, parsedConfig, action, entry }) => {
      const section = (cfg.channels as Record<string, unknown>)?.[CHANNEL_ID] as
        | Record<string, unknown>
        | undefined;
      const current = Array.isArray(section?.allowFrom)
        ? (section.allowFrom as Array<string | number>)
        : [];
      const next = [...current];
      const normalized = String(entry).replace(/^vkteams:/i, "");
      if (action === "add" && !next.map(String).includes(normalized)) {
        next.push(normalized);
      } else if (action === "remove") {
        const idx = next.map(String).indexOf(normalized);
        if (idx !== -1) next.splice(idx, 1);
      }
      return {
        ...cfg,
        channels: {
          ...cfg.channels,
          [CHANNEL_ID]: {
            ...section,
            ...parsedConfig,
            allowFrom: next,
          },
        },
      };
    },
  },
  capabilities: {
    chatTypes: ["direct"],
    reactions: false,
    threads: false,
    media: true,
    nativeCommands: false,
    blockStreaming: true,
  },
  reload: { configPrefixes: ["channels.vkteams"] },
  configSchema: buildChannelConfigSchema(VkTeamsConfigSchema),
  config: {
    listAccountIds: (cfg: OpenClawConfig) => listVkTeamsAccountIds(cfg as CoreConfig),
    resolveAccount: (cfg: OpenClawConfig, accountId?: string | null) =>
      resolveVkTeamsAccount({ cfg: cfg as CoreConfig, accountId: accountId ?? undefined }),
    defaultAccountId: (cfg: OpenClawConfig) =>
      resolveDefaultVkTeamsAccountId(cfg as CoreConfig),
    isConfigured: (account) => Boolean(account.token?.trim()),
    describeAccount: (account) => ({
      accountId: account.accountId,
      name: account.name,
      enabled: account.enabled,
      configured: Boolean(account.token?.trim()),
      tokenSource: account.tokenSource ?? undefined,
    }),
    resolveAllowFrom: ({ cfg, accountId }) =>
      resolveVkTeamsAccount({
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
      }).config.allowFrom,
    formatAllowFrom: ({ allowFrom }) =>
      allowFrom
        .map((entry) => String(entry).trim())
        .filter(Boolean)
        .map((entry) => normalizeVkTeamsDmAllowEntry(entry)),
  },
  security: vkTeamsSecurityAdapter,
  messaging: {
    normalizeTarget: (target) => {
      const trimmed = target.trim();
      if (!trimmed) {
        return undefined;
      }
      return trimmed.replace(/^vkteams:/i, "");
    },
    parseExplicitTarget: ({ raw }) => {
      const normalized = raw.trim().replace(/^vkteams:/i, "");
      if (!normalized) {
        return null;
      }
      return {
        to: normalized,
        chatType: "direct" as const,
      };
    },
    inferTargetChatType: ({ to }) => {
      const normalized = to.trim().replace(/^vkteams:/i, "");
      return normalized ? ("direct" as const) : undefined;
    },
    targetResolver: {
      looksLikeId: (id) => {
        const trimmed = id?.trim();
        if (!trimmed) {
          return false;
        }
        return /^vkteams:/i.test(trimmed) || /@/.test(trimmed) || /^\d+$/.test(trimmed);
      },
      hint: "<userId|sn@domain>",
    },
  },
  outbound: {
    deliveryMode: "direct",
    textChunkLimit: 4096,
    sendPayload: async ({ to, payload, accountId, cfg, replyToId }) => {
      const result = await sendPayloadVkTeams(to, payload as never, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
        replyTo: replyToId ?? undefined,
      });
      return result
        ? { channel: CHANNEL_ID, ...result }
        : { channel: CHANNEL_ID, messageId: "", chatId: to };
    },
    sendFormattedText: async ({ cfg, to, text, accountId, replyToId }) => {
      const results = await sendFormattedTextVkTeams(to, text, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
        replyTo: replyToId ?? undefined,
      });
      return results.map((result) => ({ channel: CHANNEL_ID as const, ...result }));
    },
    sendFormattedMedia: async ({ cfg, to, text, mediaUrl, accountId, replyToId }) => {
      const result = await sendFormattedMediaVkTeams(to, text, mediaUrl, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
        replyTo: replyToId ?? undefined,
      });
      return { channel: CHANNEL_ID, ...result };
    },
    sendText: async ({ cfg, to, text, accountId, replyToId }) => {
      const result = await sendMessageVkTeams(to, text, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
        replyTo: replyToId ?? undefined,
      });
      return { channel: CHANNEL_ID, ...result };
    },
    sendMedia: async ({ cfg, to, text, mediaUrl, accountId, replyToId }) => {
      const result = await sendFormattedMediaVkTeams(to, text, mediaUrl, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
        replyTo: replyToId ?? undefined,
      });
      return { channel: CHANNEL_ID, ...result };
    },
    editMessage: async ({ cfg, to, messageId, text, accountId }) => {
      const result = await editMessageVkTeams(to, messageId, text, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
      });
      return { channel: CHANNEL_ID, ...result };
    },
    deleteMessage: async ({ cfg, to, messageId, accountId }) => {
      const result = await deleteMessageVkTeams(to, messageId, {
        cfg: cfg as CoreConfig,
        accountId: accountId ?? undefined,
      });
      return { channel: CHANNEL_ID, ...result };
    },
  },
  status: {
    defaultRuntime: {
      accountId: DEFAULT_ACCOUNT_ID,
      running: false,
      lastStartAt: null,
      lastStopAt: null,
      lastError: null,
    },
    collectStatusIssues: (accounts) => {
      const issues: ChannelStatusIssue[] = [];
      for (const account of accounts) {
        const accountId = account.accountId ?? DEFAULT_ACCOUNT_ID;
        if (!account.configured) {
          issues.push({
            channel: CHANNEL_ID,
            accountId,
            kind: "config",
            message: "VK Teams bot token not configured",
          });
        }
      }
      return issues;
    },
    buildChannelSummary: ({ snapshot }) => buildTokenChannelStatusSummary(snapshot),
    formatCapabilitiesProbe: ({ probe }) => {
      const lines: Array<{ text: string }> = [];
      if (probe?.ok) {
        if (probe.botName) {
          lines.push({ text: `Bot: ${probe.botName}` });
        }
        if (probe.botId) {
          lines.push({ text: `ID: ${probe.botId}` });
        }
      }
      return lines;
    },
    probeAccount: async ({ account, timeoutMs }) =>
      probeVkTeamsBot(account.token, timeoutMs),
    buildAccountSnapshot: ({ account, runtime, probe }) => {
      const configured = Boolean(account.token?.trim());
      const base = buildComputedAccountStatusSnapshot({
        accountId: account.accountId,
        name: account.name,
        enabled: account.enabled,
        configured,
        runtime,
        probe,
      });
      return {
        ...base,
        tokenSource: account.tokenSource,
        mode: "longpoll",
      };
    },
  },
  gateway: {
    startAccount: async (ctx) => {
      const account = ctx.account;
      const setStatus = createAccountStatusSink({
        accountId: account.accountId,
        setStatus: ctx.setStatus,
      });
      const token = account.token.trim();
      if (!token) {
        throw new Error(
          `VK Teams long poll mode requires a non-empty bot token for account "${account.accountId}".`,
        );
      }

      let botLabel = "";
      try {
        const probe = await probeVkTeamsBot(token, 2500);
        const displayName = probe.ok ? probe.botName?.trim() : null;
        if (displayName) {
          botLabel = ` (${displayName})`;
        }
      } catch (err) {
        if (getVkTeamsRuntime().logging.shouldLogVerbose()) {
          ctx.log?.debug?.(`[${account.accountId}] VK Teams bot probe failed: ${String(err)}`);
        }
      }

      ctx.log?.info(`[${account.accountId}] starting VK Teams provider${botLabel}`);

      const monitor = await monitorVkTeamsProvider({
        token,
        accountId: account.accountId,
        config: ctx.cfg as CoreConfig,
        runtime: ctx.runtime,
        abortSignal: ctx.abortSignal,
        setStatus,
      });

      return monitor;
    },
    logoutAccount: async ({ accountId, cfg }) => {
      const nextCfg = { ...cfg } as OpenClawConfig;
      const vkConfig = ((cfg.channels as Record<string, unknown>)?.vkteams ?? {}) as Record<
        string,
        unknown
      >;
      const nextVk = { ...vkConfig };
      let cleared = false;
      let changed = false;

      if (accountId === DEFAULT_ACCOUNT_ID) {
        if (nextVk.token || nextVk.tokenFile) {
          delete nextVk.token;
          delete nextVk.tokenFile;
          cleared = true;
          changed = true;
        }
      }

      if (accountId !== DEFAULT_ACCOUNT_ID && isObjectRecord(nextVk.accounts?.[accountId])) {
        const nextAccounts = { ...(nextVk.accounts as Record<string, unknown>) };
        delete nextAccounts[accountId];
        if (Object.keys(nextAccounts).length > 0) {
          nextVk.accounts = nextAccounts;
        } else {
          delete nextVk.accounts;
        }
        cleared = true;
        changed = true;
      }

      if (changed) {
        if (Object.keys(nextVk).length > 0) {
          (nextCfg.channels as Record<string, unknown>) = {
            ...nextCfg.channels,
            vkteams: nextVk,
          };
        } else {
          const nextChannels = {
            ...nextCfg.channels,
          } as Record<string, unknown>;
          delete nextChannels.vkteams;
          if (Object.keys(nextChannels).length > 0) {
            nextCfg.channels = nextChannels;
          } else {
            delete nextCfg.channels;
          }
        }
        await getVkTeamsRuntime().config.replaceConfigFile({
          nextConfig: nextCfg,
          afterWrite: { mode: "auto" },
        });
      }

      const resolved = resolveVkTeamsAccount({
        cfg: changed ? (nextCfg as CoreConfig) : (cfg as CoreConfig),
        accountId,
      });
      const loggedOut = resolved.tokenSource === "none";

      return {
        cleared,
        envToken: Boolean(process.env.VKTEAMS_TOKEN?.trim()),
        loggedOut,
      };
    },
  },
};