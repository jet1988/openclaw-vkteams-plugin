// Long-polling монитор VK Teams (Среда).
import { VkTeamsClient, VkTeamsApiError } from "./client.js";
import { resolveVkTeamsAccount } from "./accounts.js";
import { handleVkTeamsInbound } from "./inbound.js";
import { readVkTeamsRuntimeConfig } from "./runtime.js";
import type { RuntimeEnv } from "openclaw/plugin-sdk/runtime";
import type { CoreConfig } from "./types.js";

const POLL_RETRY_DELAY_MS = 3_000;
const TYPING_INTERVAL_MS = 5_000;

export type VkTeamsMonitorOptions = {
  token: string;
  accountId: string;
  config: CoreConfig;
  runtime: RuntimeEnv;
  abortSignal?: AbortSignal;
  setStatus?: (patch: Record<string, unknown>) => void;
};

function waitForAbort(signal?: AbortSignal): Promise<void> {
  if (!signal) {
    return new Promise<void>(() => {});
  }
  if (signal.aborted) {
    return Promise.resolve();
  }
  return new Promise<void>((resolve) => {
    signal.addEventListener("abort", () => resolve(), { once: true });
  });
}

const INBOUND_PROCESS_TIMEOUT_MS = 120_000;

function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error(`inbound processing timed out after ${ms}ms`)),
      ms,
    );
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}

/**
 * Запускает long-polling цикл VK Teams: получает события, обрабатывает
 * newMessage из личных чатов, шлёт typing-индикатор во время обработки.
 * Завершается по abortSignal.
 */
export async function monitorVkTeamsProvider(opts: VkTeamsMonitorOptions): Promise<void> {
  const account = resolveVkTeamsAccount({
    cfg: opts.config,
    accountId: opts.accountId,
  });
  const client = new VkTeamsClient({
    token: opts.token,
    apiBaseUrl: account.config.apiBaseUrl,
    signal: opts.abortSignal,
  });
  const pollTime = account.config.pollTime ?? 30;

  let lastEventId = 0;
  let typingTimer: ReturnType<typeof setInterval> | undefined;
  let lastTypingChat: string | null = null;

  const stopTyping = () => {
    if (typingTimer) {
      clearInterval(typingTimer);
      typingTimer = undefined;
    }
    lastTypingChat = null;
  };

  const startTyping = (chatId: string) => {
    if (typingTimer && lastTypingChat === chatId) return;
    stopTyping();
    lastTypingChat = chatId;
    void client.sendActions(chatId, "typing").catch(() => undefined);
    typingTimer = setInterval(() => {
      void client.sendActions(chatId, "typing").catch(() => undefined);
    }, TYPING_INTERVAL_MS);
    typingTimer.unref?.();
  };

  try {
    while (!opts.abortSignal?.aborted) {
      let events;
      try {
        events = await client.getEvents(pollTime, lastEventId);
      } catch (err) {
        if (opts.abortSignal?.aborted) return;
        if (err instanceof VkTeamsApiError && err.description === "Invalid token") {
          throw err;
        }
        // Таймаут long-poll — нормальная ситуация, просто продолжаем
        await new Promise((r) => setTimeout(r, POLL_RETRY_DELAY_MS));
        continue;
      }

      if (events.length > 0) {
        lastEventId = Math.max(...events.map((e) => e.eventId));
        opts.setStatus?.({ lastInboundAt: Date.now() });
      }

      for (const event of events) {
        if (event.type !== "newMessage") continue;
        const payload = event.payload;
        if (!payload?.chat || payload.chat.type !== "private") continue;
        const sender = payload.from?.userId;
        if (!sender) continue;
        // Игнорируем собственные сообщения бота
        if (sender === opts.token.split(":").pop()) continue;

        const text = payload.text ?? "";
        if (text.trim()) {
          startTyping(payload.chat.chatId);
        }

        // ⚠️ Не блокируем polling-цикл обработкой входящего: один зависший
        // агентный run не должен останавливать приём новых сообщений и
        // держать drain при reload/stop канала. Диспатчим асинхронно
        // с таймаутом (проверено 08.10.2026: зависший run держал канал,
        // reload падал с «admitted work did not settle within 60s»).
        void (async () => {
          try {
            await withTimeout(
              handleVkTeamsInbound({
                message: {
                  messageId: payload.msgId,
                  chatId: payload.chat.chatId,
                  senderId: sender,
                  text,
                  timestamp: payload.timestamp,
                  parts: payload.parts,
                },
                account,
                config: opts.config,
                runtime: opts.runtime,
                setStatus: opts.setStatus as never,
              }),
              INBOUND_PROCESS_TIMEOUT_MS,
            );
          } catch (err) {
            opts.runtime.error?.(`[${opts.accountId}] vkteams inbound error: ${String(err)}`);
          } finally {
            stopTyping();
          }
        })();
      }
    }
  } finally {
    stopTyping();
  }
}

export { readVkTeamsRuntimeConfig };