// Отправка сообщений в VK Teams (Среда).
import { resolveVkTeamsAccount } from "./accounts.js";
import { VkTeamsClient } from "./client.js";
import { renderVkTeamsMarkdownChunks, type VkTeamsPreparedFormattedMessage } from "./format.js";
import { getVkTeamsRuntime, readVkTeamsRuntimeConfig } from "./runtime.js";
import type { CoreConfig, ResolvedVkTeamsAccount } from "./types.js";

const VKTEAMS_MESSAGE_TEXT_LIMIT = 4096;
const VKTEAMS_TRANSIENT_RETRY_ATTEMPTS = 3;

export type SendVkTeamsOptions = {
  cfg?: CoreConfig;
  accountId?: string;
  replyTo?: string;
};

export type SendVkTeamsResult = {
  messageId: string;
  chatId: string;
};

export type VkTeamsOutboundPayloadLike = {
  text?: string;
  mediaUrl?: string;
  mediaUrls?: string[];
  replyToId?: string;
  channelData?: Record<string, unknown>;
};

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function readVkTeamsErrorMessage(error: unknown): string {
  if (!error || typeof error !== "object") {
    return "";
  }
  const record = error as Record<string, unknown>;
  return [
    typeof record.message === "string" ? record.message : "",
    typeof record.name === "string" ? record.name : "",
    typeof record.description === "string" ? record.description : "",
  ]
    .filter(Boolean)
    .join(" ");
}

function isRetryableVkTeamsError(error: unknown): boolean {
  const message = readVkTeamsErrorMessage(error).toLowerCase();
  return /time-?out|timed out|temporar|econnreset|socket hang up|too many requests|rate limit|abort|429|5\d\d/.test(
    message,
  );
}

async function withVkTeamsRetry<T>(operation: () => Promise<T>): Promise<T> {
  let attempt = 0;
  while (true) {
    try {
      return await operation();
    } catch (error) {
      attempt += 1;
      if (attempt >= VKTEAMS_TRANSIENT_RETRY_ATTEMPTS || !isRetryableVkTeamsError(error)) {
        throw error;
      }
      await sleep(250 * attempt);
    }
  }
}

function getLastSendResult(results: readonly SendVkTeamsResult[]): SendVkTeamsResult | null {
  return results.length > 0 ? (results[results.length - 1] ?? null) : null;
}

function prepareVkTeamsMessageChunks(text: string): VkTeamsPreparedFormattedMessage[] {
  return renderVkTeamsMarkdownChunks(text, { chunkSize: VKTEAMS_MESSAGE_TEXT_LIMIT }).filter(
    (chunk) => chunk.text.length > 0,
  );
}

function toPreparedVkTeamsMessages(
  message: string | VkTeamsPreparedFormattedMessage | undefined,
): VkTeamsPreparedFormattedMessage[] {
  if (!message) {
    return [{ text: "" }];
  }
  if (typeof message !== "string") {
    return [message];
  }
  const chunks = prepareVkTeamsMessageChunks(message);
  return chunks.length > 0 ? chunks : [{ text: "" }];
}

function resolveSendTarget(params: { cfg?: CoreConfig; accountId?: string; to: string }) {
  const runtime = getVkTeamsRuntime();
  const cfg = params.cfg ?? readVkTeamsRuntimeConfig(runtime);
  const account = resolveVkTeamsAccount({ cfg, accountId: params.accountId });
  if (!account.token) {
    throw new Error("VK Teams token not configured");
  }
  return { account, to: params.to };
}

function resolvePayloadMediaReferences(payload: VkTeamsOutboundPayloadLike): string[] {
  const rawUrls = payload.mediaUrls?.length
    ? payload.mediaUrls
    : payload.mediaUrl
      ? [payload.mediaUrl]
      : [];
  return rawUrls.filter((entry): entry is string => typeof entry === "string").filter(Boolean);
}

function isHttpUrl(value: string): boolean {
  return /^https?:\/\//i.test(value);
}

function isLikelyVkTeamsMediaReference(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }
  if (isHttpUrl(trimmed) || /^data:/i.test(trimmed) || trimmed.startsWith("file://")) {
    return true;
  }
  if (
    trimmed.startsWith("/") ||
    trimmed.startsWith("./") ||
    trimmed.startsWith("../") ||
    /^[A-Za-z]:[\\/]/.test(trimmed)
  ) {
    return true;
  }
  return /\.(?:apng|avif|bmp|gif|heic|heif|jpeg|jpg|png|webp|aac|flac|m4a|mp3|oga|ogg|opus|wav|pdf|mp4|webm)(?:$|[?#])/i.test(
    trimmed,
  );
}

const MARKDOWN_LINK_RE = /(!?)\[([^\]\n]*)\]\(([^)\n]+)\)/g;

function resolveVkTeamsMarkdownAttachmentPayload(text: string): {
  text: string;
  mediaRefs: string[];
} {
  const mediaRefs: string[] = [];
  const stripped = text.replace(
    MARKDOWN_LINK_RE,
    (fullMatch, bang: string, _label: string, rawUrl: string) => {
      const candidate = rawUrl.trim();
      const shouldExtract = bang === "!" ? isLikelyVkTeamsMediaReference(candidate) : false;
      if (!shouldExtract) {
        return fullMatch;
      }
      mediaRefs.push(candidate);
      return "";
    },
  );
  const normalizedText = stripped
    .split("\n")
    .map((line) => line.replace(/[ \t]+$/g, ""))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  return { text: normalizedText, mediaRefs };
}

function resolveVkTeamsPayloadParts(
  payload: VkTeamsOutboundPayloadLike,
  _opts: SendVkTeamsOptions,
): {
  text: string;
  mediaRefs: string[];
  replyTo: string | undefined;
} {
  const parsedText = resolveVkTeamsMarkdownAttachmentPayload(payload.text ?? "");
  const text = parsedText.text;
  const mediaRefs = [...resolvePayloadMediaReferences(payload), ...parsedText.mediaRefs];
  const replyTo = payload.replyToId ?? _opts.replyTo;
  return { text, mediaRefs, replyTo };
}

function recordOutboundActivity(accountId: string): void {
  getVkTeamsRuntime().channel.activity.record({
    channel: "vkteams",
    accountId,
    direction: "outbound",
    at: Date.now(),
  });
}

async function sendVkTeamsApiMessage(params: {
  to: string;
  account: ResolvedVkTeamsAccount;
  formatted: VkTeamsPreparedFormattedMessage;
  replyTo?: string;
}): Promise<SendVkTeamsResult> {
  const client = new VkTeamsClient({
    token: params.account.token,
    apiBaseUrl: params.account.config.apiBaseUrl,
  });
  const result = await client.sendText({
    chatId: params.to,
    text: params.formatted.text,
    ...(params.formatted.formatData
      ? { format: JSON.stringify(params.formatted.formatData) }
      : {}),
    ...(params.replyTo ? { replyMsgId: params.replyTo } : {}),
  });
  return { messageId: result.msgId, chatId: params.to };
}

async function sendPayloadResultsVkTeams(params: {
  to: string;
  text: string;
  mediaRefs: string[];
  opts: SendVkTeamsOptions;
}): Promise<SendVkTeamsResult[]> {
  const { account, to } = resolveSendTarget({
    cfg: params.opts.cfg,
    accountId: params.opts.accountId,
    to: params.to,
  });
  const results: SendVkTeamsResult[] = [];
  const client = new VkTeamsClient({
    token: account.token,
    apiBaseUrl: account.config.apiBaseUrl,
  });

  // Сначала медиа (если есть)
  for (const mediaUrl of params.mediaRefs) {
    let fileData: Buffer;
    let filename = "file";
    if (isHttpUrl(mediaUrl)) {
      fileData = await withVkTeamsRetry(() => client.downloadFile(mediaUrl));
      try {
        filename = decodeURIComponent(mediaUrl.split("/").pop()?.split("?")[0] ?? "file");
      } catch {
        filename = "file";
      }
    } else {
      const fs = await import("node:fs/promises");
      fileData = await fs.readFile(mediaUrl.replace(/^file:\/\//, ""));
      filename = mediaUrl.split("/").pop() ?? "file";
    }
    const result = await withVkTeamsRetry(() =>
      client.sendFile({
        chatId: to,
        file: fileData,
        filename,
        caption: results.length === 0 && params.text.trim() ? params.text : undefined,
        ...(params.opts.replyTo ? { replyMsgId: params.opts.replyTo } : {}),
      }),
    );
    results.push({ messageId: result.msgId, chatId: to });
    recordOutboundActivity(account.accountId);
  }

  // Затем текст (если не ушёл как caption к первому медиа)
  if (params.text.trim() && !(results.length > 0 && params.mediaRefs.length > 0)) {
    const chunks = toPreparedVkTeamsMessages(params.text);
    for (const chunk of chunks) {
      const result = await withVkTeamsRetry(() =>
        sendVkTeamsApiMessage({
          to,
          account,
          formatted: chunk,
          replyTo: params.opts.replyTo,
        }),
      );
      results.push(result);
      recordOutboundActivity(account.accountId);
    }
  }

  return results;
}

export async function sendMessageVkTeams(
  to: string,
  text: string,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult> {
  const parsed = resolveVkTeamsMarkdownAttachmentPayload(text);
  const results = await sendPayloadResultsVkTeams({
    to,
    text: parsed.text,
    mediaRefs: parsed.mediaRefs,
    opts,
  });
  return getLastSendResult(results) ?? { messageId: "", chatId: to };
}

export async function sendFormattedTextVkTeams(
  to: string,
  text: string,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult[]> {
  const parsed = resolveVkTeamsMarkdownAttachmentPayload(text);
  const results = await sendPayloadResultsVkTeams({
    to,
    text: parsed.text,
    mediaRefs: parsed.mediaRefs,
    opts,
  });
  return results;
}

export async function sendFormattedMediaVkTeams(
  to: string,
  text: string,
  mediaUrl: string,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult> {
  const results = await sendPayloadResultsVkTeams({
    to,
    text,
    mediaRefs: [mediaUrl],
    opts,
  });
  return getLastSendResult(results) ?? { messageId: "", chatId: to };
}

export async function sendTypingVkTeams(
  to: string,
  account: ResolvedVkTeamsAccount,
): Promise<void> {
  if (!account.token) {
    return;
  }
  const client = new VkTeamsClient({
    token: account.token,
    apiBaseUrl: account.config.apiBaseUrl,
  });
  await withVkTeamsRetry(async () => {
    await client.sendActions(to, "typing");
  });
}

export async function sendPayloadVkTeams(
  to: string,
  payload: VkTeamsOutboundPayloadLike,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult | null> {
  const { text, mediaRefs, replyTo } = resolveVkTeamsPayloadParts(payload, opts);
  return getLastSendResult(
    await sendPayloadResultsVkTeams({
      to,
      text,
      mediaRefs,
      opts: { ...opts, replyTo },
    }),
  );
}
/** messages/editText — edit an existing message (markdown → format). */
export async function editMessageVkTeams(
  to: string,
  messageId: string,
  text: string,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult> {
  const { account } = resolveSendTarget({
    cfg: opts.cfg,
    accountId: opts.accountId,
    to,
  });
  const client = new VkTeamsClient({
    token: account.token,
    apiBaseUrl: account.config.apiBaseUrl,
  });
  const parsed = resolveVkTeamsMarkdownAttachmentPayload(text);
  const chunks = toPreparedVkTeamsMessages(parsed.text);
  const chunk = chunks[0] ?? { text: "" };
  const result = await withVkTeamsRetry(() =>
    client.editText({
      chatId: to,
      msgId: messageId,
      text: chunk.text,
      ...(chunk.formatData
        ? { format: JSON.stringify(chunk.formatData) }
        : {}),
    }),
  );
  recordOutboundActivity(account.accountId);
  return { messageId: result.msgId, chatId: to };
}

/** messages/deleteMessages — delete a message. */
export async function deleteMessageVkTeams(
  to: string,
  messageId: string,
  opts: SendVkTeamsOptions = {},
): Promise<SendVkTeamsResult> {
  const { account } = resolveSendTarget({
    cfg: opts.cfg,
    accountId: opts.accountId,
    to,
  });
  const client = new VkTeamsClient({
    token: account.token,
    apiBaseUrl: account.config.apiBaseUrl,
  });
  await withVkTeamsRetry(() => client.deleteMessages(to, messageId));
  recordOutboundActivity(account.accountId);
  return { messageId, chatId: to };
}
