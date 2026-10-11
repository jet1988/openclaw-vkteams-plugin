export type DmPolicy = "pairing" | "allowlist" | "open" | "disabled";

export type VkTeamsAccountConfig = {
  name?: string;
  enabled?: boolean;
  token?: string;
  tokenFile?: string;
  apiBaseUrl?: string;
  pollTime?: number;
  dmPolicy?: DmPolicy;
  allowFrom?: Array<string | number>;
  defaultTo?: string;
};

export type VkTeamsConfig = VkTeamsAccountConfig & {
  accounts?: Record<string, VkTeamsAccountConfig>;
};

export type ResolvedVkTeamsAccount = {
  accountId: string;
  enabled: boolean;
  name?: string;
  token: string;
  tokenSource: "env" | "tokenFile" | "config" | "none";
  config: VkTeamsAccountConfig;
};

export type VkTeamsInboundMessage = {
  messageId: string;
  chatId: string;
  senderId: string;
  text: string;
  timestamp?: number;
  parts?: Array<{
    type: string;
    payload: Record<string, unknown>;
  }>;
  replyToMsgId?: string;
};

export type VkTeamsInboundAttachment = {
  type: string;
  kind: "image" | "video" | "audio" | "document" | "voice" | "sticker" | "unknown";
  fileId?: string;
  url?: string;
  title?: string;
  mimeType?: string;
};

export type VkTeamsInboundResolvedMedia = {
  path?: string;
  url?: string;
  contentType?: string;
  attachment: VkTeamsInboundAttachment;
};

export type VkTeamsProbe = {
  ok: boolean;
  botName?: string;
  botId?: string;
  error?: string;
};

export type CoreConfig = {
  channels?: {
    vkteams?: VkTeamsConfig;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};