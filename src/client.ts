// VK Teams (Среда) Bot API client.
// Base URL: https://api.icq.net/bot/v1
// Auth: token passed as `token` query parameter (per official libraries).
// Zero-dependency: uses built-in fetch/FormData/Blob (Node 24+).

export const DEFAULT_API_BASE_URL = "https://api.icq.net/bot/v1";

export type VkTeamsEventType =
  | "newMessage"
  | "editedMessage"
  | "deletedMessage"
  | "pinnedMessage"
  | "unpinnedMessage"
  | "newChatMembers"
  | "leftChatMembers"
  | "changedChatInfo"
  | "callbackQuery";

export interface VkTeamsChatInfo {
  chatId: string;
  type: "private" | "group" | "channel";
  title?: string;
}

export interface VkTeamsUserInfo {
  userId: string;
  firstName?: string;
  lastName?: string;
  nick?: string;
}

export interface VkTeamsPart {
  type: "file" | "sticker" | "mention" | "voice" | "forward" | "reply";
  payload: Record<string, unknown>;
}

export interface VkTeamsMessagePayload {
  msgId: string;
  chat: VkTeamsChatInfo;
  from?: VkTeamsUserInfo;
  text?: string;
  parts?: VkTeamsPart[];
  timestamp?: number;
  format?: unknown;
}

export interface VkTeamsEvent {
  eventId: number;
  type: VkTeamsEventType;
  payload: VkTeamsMessagePayload;
}

export interface VkTeamsBotInfo {
  firstName?: string;
  nick?: string;
  userId?: string;
  about?: string;
  photo?: Array<{ url: string }>;
}

export interface VkTeamsFileInfo {
  fileId: string;
  type: string;
  size: number;
  filename: string;
  url: string;
}

export interface VkTeamsApiOptions {
  token: string;
  apiBaseUrl?: string;
  timeoutMs?: number;
  signal?: AbortSignal;
}

export class VkTeamsApiError extends Error {
  constructor(
    message: string,
    public readonly ok: boolean,
    public readonly description?: string,
    public readonly statusCode?: number,
  ) {
    super(message);
    this.name = "VkTeamsApiError";
  }
}

export class VkTeamsClient {
  private readonly token: string;
  private readonly apiBaseUrl: string;
  private readonly timeoutMs: number;
  private readonly externalSignal?: AbortSignal;

  constructor(options: VkTeamsApiOptions) {
    this.token = options.token;
    this.apiBaseUrl = (options.apiBaseUrl ?? DEFAULT_API_BASE_URL).replace(/\/+$/, "");
    this.timeoutMs = options.timeoutMs ?? 60_000;
    this.externalSignal = options.signal;
  }

  private async call<T>(
    method: "GET" | "POST",
    path: string,
    params: Record<string, string | number | undefined>,
    body?: FormData,
  ): Promise<T> {
    const url = new URL(`${this.apiBaseUrl}${path}`);
    url.searchParams.set("token", this.token);
    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined && value !== null && value !== "") {
        url.searchParams.set(key, String(value));
      }
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    const onExternalAbort = () => controller.abort();
    this.externalSignal?.addEventListener("abort", onExternalAbort, { once: true });
    try {
      const response = await fetch(url.toString(), {
        method,
        body: body ?? undefined,
        signal: controller.signal,
      });

      const text = await response.text();
      let json: any;
      try {
        json = JSON.parse(text);
      } catch {
        throw new VkTeamsApiError(
          `VK Teams API returned non-JSON response (HTTP ${response.status})`,
          false,
          text.slice(0, 200),
          response.status,
        );
      }

      if (json && json.ok === false) {
        throw new VkTeamsApiError(
          `VK Teams API error: ${json.description ?? "unknown"}`,
          false,
          json.description,
          response.status,
        );
      }
      return json as T;
    } finally {
      clearTimeout(timeout);
      this.externalSignal?.removeEventListener("abort", onExternalAbort);
    }
  }

  /** self/get — bot info */
  async getSelf(): Promise<VkTeamsBotInfo> {
    return this.call<VkTeamsBotInfo>("GET", "/self/get", {});
  }

  /**
   * events/get — long polling. Blocks up to pollTime seconds.
   * Returns events with eventId > lastEventId (or all if lastEventId is 0).
   */
  async getEvents(pollTime: number, lastEventId: number): Promise<VkTeamsEvent[]> {
    const json = await this.call<{ events?: VkTeamsEvent[] }>("GET", "/events/get", {
      pollTime,
      lastEventId,
    });
    return json.events ?? [];
  }

  /** messages/sendText — send a text message */
  async sendText(params: {
    chatId: string;
    text: string;
    replyMsgId?: string;
    parseMode?: "HTML" | "MarkdownV2";
    format?: string;
  }): Promise<{ msgId: string }> {
    return this.call<{ msgId: string }>("GET", "/messages/sendText", {
      chatId: params.chatId,
      text: params.text,
      replyMsgId: params.replyMsgId,
      parseMode: params.parseMode,
      format: params.format,
    });
  }

  /** messages/sendFile — send a file (multipart upload) */
  async sendFile(params: {
    chatId: string;
    file: Buffer | Uint8Array;
    filename: string;
    caption?: string;
    replyMsgId?: string;
  }): Promise<{ msgId: string }> {
    const form = new FormData();
    form.append(
      "file",
      new Blob([params.file as BlobPart], { type: "application/octet-stream" }),
      params.filename,
    );
    return this.call<{ msgId: string }>(
      "POST",
      "/messages/sendFile",
      {
        chatId: params.chatId,
        caption: params.caption,
        replyMsgId: params.replyMsgId,
      },
      form,
    );
  }

  /** messages/sendVoice — send a voice message (multipart upload) */
  async sendVoice(params: {
    chatId: string;
    file: Buffer | Uint8Array;
    filename: string;
    replyMsgId?: string;
  }): Promise<{ msgId: string }> {
    const form = new FormData();
    form.append(
      "file",
      new Blob([params.file as BlobPart], { type: "audio/ogg" }),
      params.filename,
    );
    return this.call<{ msgId: string }>(
      "POST",
      "/messages/sendVoice",
      {
        chatId: params.chatId,
        replyMsgId: params.replyMsgId,
      },
      form,
    );
  }

  /** messages/editText — edit a message */
  async editText(params: {
    chatId: string;
    msgId: string;
    text: string;
    parseMode?: "HTML" | "MarkdownV2";
    format?: string;
  }): Promise<{ msgId: string }> {
    return this.call<{ msgId: string }>("GET", "/messages/editText", {
      chatId: params.chatId,
      msgId: params.msgId,
      text: params.text,
      parseMode: params.parseMode,
      format: params.format,
    });
  }

  /** messages/deleteMessages — delete a message */
  async deleteMessages(chatId: string, msgId: string): Promise<{ ok: boolean }> {
    return this.call<{ ok: boolean }>("GET", "/messages/deleteMessages", {
      chatId,
      msgId,
    });
  }

  /** chats/sendActions — typing indicator etc. */
  async sendActions(chatId: string, actions: string): Promise<{ ok: boolean }> {
    return this.call<{ ok: boolean }>("GET", "/chats/sendActions", {
      chatId,
      actions,
    });
  }

  /** files/getInfo — file metadata incl. download url */
  async getFileInfo(fileId: string): Promise<VkTeamsFileInfo> {
    return this.call<VkTeamsFileInfo>("GET", "/files/getInfo", { fileId });
  }

  /** Download a file from a url returned by files/getInfo */
  async downloadFile(url: string): Promise<Buffer> {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60_000);
    try {
      const response = await fetch(url, { signal: controller.signal });
      if (response.status !== 200) {
        throw new VkTeamsApiError(
          `File download failed (HTTP ${response.status})`,
          false,
          undefined,
          response.status,
        );
      }
      return Buffer.from(await response.arrayBuffer());
    } finally {
      clearTimeout(timeout);
    }
  }
}