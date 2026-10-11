import { DEFAULT_ACCOUNT_ID, tryReadSecretFileSync } from "openclaw/plugin-sdk/core";
import { normalizeAccountId } from "openclaw/plugin-sdk/account-id";
import type { ResolvedVkTeamsAccount, VkTeamsAccountConfig, VkTeamsConfig, CoreConfig } from "./types.js";
export type { ResolvedVkTeamsAccount } from "./types.js";

export const DEFAULT_API_BASE_URL = "https://api.icq.net/bot/v1";
export const DEFAULT_POLL_TIME = 30;

function mergeVkTeamsAccountConfig(cfg: CoreConfig, accountId: string): VkTeamsAccountConfig {
  const vkConfig = (cfg.channels?.vkteams ?? {}) as VkTeamsConfig;
  const base: VkTeamsAccountConfig = {
    name: vkConfig.name,
    enabled: vkConfig.enabled,
    token: vkConfig.token,
    tokenFile: vkConfig.tokenFile,
    apiBaseUrl: vkConfig.apiBaseUrl,
    pollTime: vkConfig.pollTime,
    dmPolicy: vkConfig.dmPolicy,
    allowFrom: vkConfig.allowFrom,
    defaultTo: vkConfig.defaultTo,
  };

  if (accountId !== DEFAULT_ACCOUNT_ID && vkConfig.accounts?.[accountId]) {
    return {
      ...base,
      ...vkConfig.accounts[accountId],
    };
  }

  return base;
}

function resolveVkTeamsToken(
  accountId: string,
  config: VkTeamsAccountConfig,
): { token: string; source: "env" | "tokenFile" | "config" | "none" } {
  // 1. Environment variable
  if (accountId === DEFAULT_ACCOUNT_ID) {
    const envToken = process.env.VKTEAMS_TOKEN?.trim();
    if (envToken) {
      return { token: envToken, source: "env" };
    }
  }

  // 2. Token file
  if (config.tokenFile?.trim()) {
    const fileValue = tryReadSecretFileSync(config.tokenFile.trim(), "VK Teams token");
    if (fileValue?.trim()) {
      return { token: fileValue.trim(), source: "tokenFile" };
    }
  }

  // 3. Config value
  if (config.token?.trim()) {
    return { token: config.token.trim(), source: "config" };
  }

  return { token: "", source: "none" };
}

export function resolveVkTeamsAccount(params: {
  cfg: CoreConfig;
  accountId?: string | null;
}): ResolvedVkTeamsAccount {
  const accountId = normalizeAccountId(params.accountId ?? undefined);
  const config = mergeVkTeamsAccountConfig(params.cfg, accountId);
  const { token, source } = resolveVkTeamsToken(accountId, config);

  return {
    accountId,
    enabled: config.enabled !== false,
    name: config.name?.trim() || undefined,
    token,
    tokenSource: source,
    config,
  };
}

export function listVkTeamsAccountIds(cfg: CoreConfig): string[] {
  const vkConfig = (cfg.channels?.vkteams ?? {}) as VkTeamsConfig;
  const ids = Object.keys(vkConfig.accounts ?? {});
  return ids.length > 0 ? ids : [DEFAULT_ACCOUNT_ID];
}

export function resolveDefaultVkTeamsAccountId(cfg: CoreConfig): string {
  const vkConfig = (cfg.channels?.vkteams ?? {}) as VkTeamsConfig;
  const ids = Object.keys(vkConfig.accounts ?? {});
  return ids.length > 0 ? (ids[0] ?? DEFAULT_ACCOUNT_ID) : DEFAULT_ACCOUNT_ID;
}