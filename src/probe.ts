import { VkTeamsClient } from "./client.js";
import type { VkTeamsProbe } from "./types.js";

/**
 * Проверяет токен VK Teams через self/get.
 */
export async function probeVkTeamsBot(
  token: string,
  timeoutMs?: number,
): Promise<VkTeamsProbe> {
  if (!token.trim()) {
    return { ok: false, error: "No token provided" };
  }

  const controller = new AbortController();
  const timer = timeoutMs ? setTimeout(() => controller.abort(), timeoutMs) : null;

  try {
    const client = new VkTeamsClient({ token, timeoutMs: timeoutMs ?? 10_000 });
    const self = await client.getSelf();
    return {
      ok: true,
      botName: self.firstName ?? self.nick,
      botId: self.userId,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { ok: false, error: message };
  } finally {
    if (timer) {
      clearTimeout(timer);
    }
  }
}