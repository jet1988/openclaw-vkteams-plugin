// Обработка входящих сообщений VK Teams (Среда).
import { logInboundDrop, toInboundMediaFacts } from "openclaw/plugin-sdk/channel-inbound";
import { createTypingCallbacks, logTypingFailure } from "openclaw/plugin-sdk/channel-outbound";
import { createChannelPairingController } from "openclaw/plugin-sdk/channel-pairing";
import type { RuntimeEnv } from "openclaw/plugin-sdk/runtime";
import { VkTeamsClient } from "./client.js";
import { resolveVkTeamsInboundResolvedMedia } from "./media.js";
import { getVkTeamsRuntime } from "./runtime.js";
import { sendPayloadVkTeams, sendTypingVkTeams } from "./send.js";
import type { CoreConfig, ResolvedVkTeamsAccount, VkTeamsInboundMessage } from "./types.js";

const CHANNEL_ID = "vkteams" as const;

function normalizeVkTeamsAllowlist(allowFrom: Array<string | number> | undefined): string[] {
  if (!allowFrom) {
    return [];
  }
  return allowFrom.map((entry) => String(entry).trim().toLowerCase()).filter(Boolean);
}

function resolveVkTeamsAllowlistMatch(params: {
  allowFrom: string[];
  senderId: string;
}): { allowed: boolean } {
  const senderStr = params.senderId.toLowerCase();
  if (params.allowFrom.length === 0) {
    return { allowed: false };
  }
  if (params.allowFrom.includes("*")) {
    return { allowed: true };
  }
  return {
    allowed: params.allowFrom.some(
      (entry) => entry === senderStr || entry === `vkteams:${senderStr}`,
    ),
  };
}

function resolveVkTeamsInboundMediaKind(kind: string): "image" | "video" | "audio" | "document" | "sticker" | "voice" | "unknown" {
  switch (kind) {
    case "image":
    case "video":
    case "audio":
    case "document":
    case "sticker":
    case "voice":
      return kind;
    default:
      return "unknown";
  }
}

type VkTeamsDispatchPayload = {
  text?: string;
  mediaUrl?: string;
  mediaUrls?: string[];
  replyToId?: string;
  channelData?: Record<string, unknown>;
};

async function deliverVkTeamsReply(params: {
  payload: VkTeamsDispatchPayload;
  chatId: string;
  accountId: string;
  statusSink?: (patch: { lastOutboundAt?: number }) => void;
}) {
  const result = await sendPayloadVkTeams(String(params.chatId), params.payload, {
    accountId: params.accountId,
  });
  if (!result) {
    return;
  }
  params.statusSink?.({ lastOutboundAt: Date.now() });
}

/**
 * Извлекает вложения из parts сообщения VK Teams.
 * Части: file (payload.fileId), voice (payload.fileId), sticker (payload.fileId),
 * mention, forward, reply.
 */
function extractVkTeamsAttachments(message: VkTeamsInboundMessage): Array<{
  type: string;
  kind: "image" | "video" | "audio" | "document" | "voice" | "sticker" | "unknown";
  fileId?: string;
  url?: string;
  title?: string;
  mimeType?: string;
}> {
  const out: Array<{
    type: string;
    kind: "image" | "video" | "audio" | "document" | "voice" | "sticker" | "unknown";
    fileId?: string;
    url?: string;
    title?: string;
    mimeType?: string;
  }> = [];
  for (const part of message.parts ?? []) {
    if (part.type === "file" || part.type === "voice" || part.type === "sticker") {
      const payload = part.payload ?? {};
      const fileId = typeof payload.fileId === "string" ? payload.fileId : undefined;
      const type = part.type;
      const kind =
        type === "voice"
          ? "voice"
          : type === "sticker"
            ? "sticker"
            : resolveVkTeamsInboundMediaKind(
                fileId,
                typeof payload.mimeType === "string" ? payload.mimeType : undefined,
                typeof payload.filename === "string" ? payload.filename : undefined,
              );
      out.push({
        type,
        kind,
        fileId,
        title: typeof payload.filename === "string" ? payload.filename : undefined,
        mimeType: typeof payload.mimeType === "string" ? payload.mimeType : undefined,
      });
    }
  }
  return out;
}

/**
 * Получает url для fileId через files/getInfo.
 */
async function resolveVkTeamsAttachmentUrls(
  attachments: ReturnType<typeof extractVkTeamsAttachments>,
  account: ResolvedVkTeamsAccount,
): Promise<void> {
  if (attachments.length === 0) {
    return;
  }
  const client = new VkTeamsClient({
    token: account.token,
    apiBaseUrl: account.config.apiBaseUrl,
  });
  for (const attachment of attachments) {
    if (!attachment.fileId || attachment.url) {
      continue;
    }
    try {
      const info = await client.getFileInfo(attachment.fileId);
      attachment.url = info.url;
      attachment.title = attachment.title ?? info.filename;
      attachment.mimeType = attachment.mimeType ?? mimeFromFilename(info.filename);
    } catch (err) {
      getVkTeamsRuntime().logging?.log?.(
        `vkteams: files/getInfo failed for ${attachment.fileId}: ${String(err)}`,
      );
    }
  }
}

function mimeFromFilename(filename: string): string | undefined {
  const ext = filename.split(".").pop()?.toLowerCase();
  if (!ext) return undefined;
  const map: Record<string, string> = {
    jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif",
    webp: "image/webp", bmp: "image/bmp", mp4: "video/mp4", webm: "video/webm",
    mov: "video/quicktime", mp3: "audio/mpeg", ogg: "audio/ogg", opus: "audio/opus",
    wav: "audio/wav", m4a: "audio/mp4", pdf: "application/pdf", txt: "text/plain",
    md: "text/markdown", json: "application/json", csv: "text/csv",
    doc: "application/msword", docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xls: "application/vnd.ms-excel", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    zip: "application/zip", tar: "application/x-tar", gz: "application/gzip",
  };
  return map[ext];
}

export async function handleVkTeamsInbound(params: {
  message: VkTeamsInboundMessage;
  account: ResolvedVkTeamsAccount;
  config: CoreConfig;
  runtime: RuntimeEnv;
  statusSink?: (patch: { lastInboundAt?: number; lastOutboundAt?: number }) => void;
}): Promise<void> {
  const { message, account, config, runtime, statusSink } = params;
  const core = getVkTeamsRuntime();
  const pairing = createChannelPairingController({
    core,
    channel: CHANNEL_ID,
    accountId: account.accountId,
  });

  const rawBody = message.text ?? "";
  if (!rawBody.trim() && (message.parts?.length ?? 0) === 0) {
    return;
  }

  statusSink?.({ lastInboundAt: message.timestamp ?? Date.now() });

  const senderDisplay = String(message.senderId);
  const dmPolicy = account.config.dmPolicy ?? "pairing";
  const configAllowFrom = normalizeVkTeamsAllowlist(account.config.allowFrom);

  // Sender authorization (DM)
  if (dmPolicy === "disabled") {
    runtime.log?.(`vkteams: drop DM sender=${senderDisplay} (dmPolicy=disabled)`);
    return;
  }
  if (dmPolicy !== "open") {
    const dmAllowed = resolveVkTeamsAllowlistMatch({
      allowFrom: configAllowFrom,
      senderId: senderDisplay,
    });
    if (!dmAllowed.allowed) {
      if (dmPolicy === "pairing") {
        await pairing.issueChallenge({
          senderId: senderDisplay,
          senderIdLine: `Your VK Teams user id: ${senderDisplay}`,
          meta: {},
          sendPairingReply: async (text) => {
            await deliverVkTeamsReply({
              payload: { text },
              chatId: message.chatId,
              accountId: account.accountId,
              statusSink,
            });
          },
          onReplyError: (err) => {
            runtime.error?.(`vkteams: pairing reply failed for ${senderDisplay}: ${String(err)}`);
          },
        });
      }
      runtime.log?.(`vkteams: drop DM sender ${senderDisplay} (dmPolicy=${dmPolicy})`);
      return;
    }
  }

  // Command gating
  const allowTextCommands = core.channel.commands.shouldHandleTextCommands({
    cfg: config as never,
    surface: CHANNEL_ID,
  });
  const hasControlCommand = core.channel.text.hasControlCommand(rawBody, config as never);
  if (hasControlCommand && !allowTextCommands) {
    logInboundDrop({
      log: (line) => runtime.log?.(line),
      channel: CHANNEL_ID,
      reason: "control command (disabled)",
      target: senderDisplay,
    });
    return;
  }

  // Build route and dispatch
  const chatId = String(message.chatId);
  const route = core.channel.routing.resolveAgentRoute({
    cfg: config as never,
    channel: CHANNEL_ID,
    accountId: account.accountId,
    peer: {
      kind: "direct",
      id: chatId,
    },
  });

  const fromLabel = `vkteams:${message.senderId}`;
  const storePath = core.channel.session.resolveStorePath(
    (config as Record<string, Record<string, unknown>>).session?.store as string | undefined,
    { agentId: route.agentId },
  );
  const envelopeOptions = core.channel.reply.resolveEnvelopeFormatOptions(config as never);
  const previousTimestamp = core.channel.session.readSessionUpdatedAt({
    storePath,
    sessionKey: route.sessionKey,
  });
  const body = core.channel.reply.formatAgentEnvelope({
    channel: "VK Teams",
    from: fromLabel,
    timestamp: message.timestamp ?? Date.now(),
    previousTimestamp,
    envelope: envelopeOptions,
    body: rawBody,
  });

  // Media
  const attachments = extractVkTeamsAttachments(message);
  await resolveVkTeamsAttachmentUrls(attachments, account);
  const resolvedMedia = await resolveVkTeamsInboundResolvedMedia({
    attachments,
    mediaRuntime: core.channel.media,
    logError: (line) => runtime.log?.(line),
  });
  const media = toInboundMediaFacts(
    resolvedMedia.map((entry) => ({
      path: entry.path,
      url: entry.url,
      contentType: entry.contentType ?? entry.attachment.mimeType,
      fileName: entry.attachment.title,
      kind: resolveVkTeamsInboundMediaKind(entry.attachment.kind),
    })),
    { messageId: message.messageId },
  );

  const ctxPayload = core.channel.reply.finalizeInboundContext({
    Body: body,
    BodyForAgent: rawBody,
    RawBody: rawBody,
    CommandBody: rawBody,
    From: fromLabel,
    To: `vkteams:${chatId}`,
    SessionKey: route.sessionKey,
    AccountId: route.accountId,
    ChatType: "direct",
    ConversationLabel: fromLabel,
    SenderName: undefined,
    SenderId: senderDisplay,
    Provider: CHANNEL_ID,
    Surface: CHANNEL_ID,
    MessageSid: message.messageId,
    Timestamp: message.timestamp,
    OriginatingChannel: CHANNEL_ID,
    OriginatingTo: `vkteams:${chatId}`,
    CommandAuthorized: false,
    media: media.length > 0 ? media : undefined,
    ReplyToId: message.replyToMsgId,
    ReplyToIdFull: message.replyToMsgId,
  });

  const onDispatchError = (err: unknown, info: { kind: string }) => {
    runtime.error?.(`vkteams ${info.kind} reply failed: ${String(err)}`);
  };
  const typingCallbacks = createTypingCallbacks({
    start: async () => {
      await sendTypingVkTeams(String(message.chatId), account);
    },
    onStartError: (err) => {
      logTypingFailure({
        log: (line) => runtime.log?.(line),
        channel: CHANNEL_ID,
        target: String(message.chatId),
        error: err,
      });
    },
  });
  let typingStarted = false;
  const startTypingOnce = async () => {
    if (typingStarted) {
      return;
    }
    typingStarted = true;
    await typingCallbacks.onReplyStart();
  };

  await core.channel.session.recordInboundSession({
    storePath,
    ctx: ctxPayload,
    sessionKey: ctxPayload.SessionKey ?? route.sessionKey,
    onRecordError: (err) => {
      runtime.error?.(`vkteams: failed updating session meta: ${String(err)}`);
    },
  });

  await startTypingOnce();

  try {
    await core.channel.reply.dispatchReplyWithBufferedBlockDispatcher({
      ctx: ctxPayload,
      cfg: config as never,
      dispatcherOptions: {
        onReplyStart: async () => {
          await startTypingOnce();
        },
        typingCallbacks,
        deliver: async (payload: unknown, info?: { kind?: string }) => {
          const normalized =
            payload && typeof payload === "object" && !Array.isArray(payload)
              ? (payload as VkTeamsDispatchPayload)
              : {};
          await deliverVkTeamsReply({
            payload: normalized,
            chatId: message.chatId,
            accountId: account.accountId,
            statusSink,
          });
          void info;
        },
        onError: onDispatchError,
      },
    });
  } finally {
    typingCallbacks.onReplyDone?.();
  }
}