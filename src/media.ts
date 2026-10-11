// Обработка входящих и исходящих медиа для VK Teams (Среда).
import { readFile } from "node:fs/promises";
import { basename, extname } from "node:path";
import type { PluginRuntime } from "openclaw/plugin-sdk/core";
import type { VkTeamsInboundAttachment, VkTeamsInboundResolvedMedia } from "./types.js";

const IMAGE_EXTENSIONS = new Set([
  ".apng", ".avif", ".bmp", ".heic", ".heif", ".jpeg", ".jpg", ".png", ".webp",
]);
const AUDIO_EXTENSIONS = new Set([
  ".aac", ".flac", ".m4a", ".mp3", ".oga", ".ogg", ".opus", ".wav",
]);
const VIDEO_EXTENSIONS = new Set([".mp4", ".webm", ".mov", ".avi", ".mkv"]);

const DEFAULT_VKTEAMS_INBOUND_MEDIA_MAX_BYTES = 20 * 1024 * 1024;

const MIME_BY_EXTENSION: Record<string, string> = {
  ".aac": "audio/aac",
  ".bmp": "image/bmp",
  ".flac": "audio/flac",
  ".gif": "image/gif",
  ".heic": "image/heic",
  ".heif": "image/heif",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".m4a": "audio/mp4",
  ".mp3": "audio/mpeg",
  ".mp4": "video/mp4",
  ".mov": "video/quicktime",
  ".avi": "video/x-msvideo",
  ".mkv": "video/x-matroska",
  ".oga": "audio/ogg",
  ".ogg": "audio/ogg",
  ".opus": "audio/opus",
  ".pdf": "application/pdf",
  ".txt": "text/plain",
  ".md": "text/markdown",
  ".json": "application/json",
  ".csv": "text/csv",
  ".tsv": "text/tab-separated-values",
  ".log": "text/plain",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".zip": "application/zip",
  ".tar": "application/x-tar",
  ".gz": "application/gzip",
  ".tgz": "application/gzip",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".tif": "image/tiff",
  ".tiff": "image/tiff",
  ".wav": "audio/wav",
  ".webm": "video/webm",
  ".webp": "image/webp",
};

function mimeFromExtension(ext?: string): string | undefined {
  const normalized = ext?.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }
  return MIME_BY_EXTENSION[normalized.startsWith(".") ? normalized : `.${normalized}`];
}

function extensionFromMimeType(value?: string): string | undefined {
  const normalized = value?.trim().toLowerCase().split(";")[0];
  if (!normalized) {
    return undefined;
  }
  const entry = Object.entries(MIME_BY_EXTENSION).find(([, mime]) => mime === normalized);
  return entry?.[0];
}

function kindFromMimeType(value?: string): VkTeamsInboundAttachment["kind"] {
  const normalized = value?.trim().toLowerCase().split(";")[0] ?? "";
  if (normalized.startsWith("image/")) return "image";
  if (normalized.startsWith("video/")) return "video";
  if (normalized.startsWith("audio/")) return "audio";
  if (normalized === "application/ogg" || normalized === "audio/ogg") return "voice";
  return "document";
}

function kindFromExtension(ext?: string): VkTeamsInboundAttachment["kind"] {
  const normalized = ext?.trim().toLowerCase();
  if (!normalized) return "unknown";
  if (IMAGE_EXTENSIONS.has(normalized)) return "image";
  if (VIDEO_EXTENSIONS.has(normalized)) return "video";
  if (AUDIO_EXTENSIONS.has(normalized)) return "audio";
  return "document";
}

function buildFileHint(attachment: VkTeamsInboundAttachment): string {
  const title = attachment.title?.trim();
  if (title) {
    return title;
  }
  const ext = extensionFromMimeType(attachment.mimeType);
  return `attachment${ext ?? ".bin"}`;
}

function shouldMaterialize(attachment: VkTeamsInboundAttachment): boolean {
  // Скачиваем только то, что реально нужно агенту: изображения, видео, аудио, документы
  return attachment.kind !== "sticker" && attachment.kind !== "unknown";
}

/**
 * Разрешает входящие вложения: скачивает файл через files/getInfo url
 * и сохраняет через mediaRuntime, либо возвращает url как есть.
 */
export async function resolveVkTeamsInboundResolvedMedia(params: {
  attachments?: readonly VkTeamsInboundAttachment[];
  mediaRuntime?: Pick<PluginRuntime["channel"]["media"], "fetchRemoteMedia" | "saveMediaBuffer">;
  maxBytes?: number;
  logError?: (line: string) => void;
}): Promise<VkTeamsInboundResolvedMedia[]> {
  const attachments = params.attachments ?? [];
  if (attachments.length === 0) {
    return [];
  }

  const maxBytes = params.maxBytes ?? DEFAULT_VKTEAMS_INBOUND_MEDIA_MAX_BYTES;
  const out: VkTeamsInboundResolvedMedia[] = [];
  for (const attachment of attachments) {
    const url = attachment.url?.trim();
    if (!url) {
      continue;
    }
    if (!params.mediaRuntime || !shouldMaterialize(attachment)) {
      out.push({
        url,
        contentType: attachment.mimeType,
        attachment,
      });
      continue;
    }

    try {
      const fileHint = buildFileHint(attachment);
      const fetched = await params.mediaRuntime.fetchRemoteMedia({
        url,
        filePathHint: fileHint,
        maxBytes,
      });
      const saved = await params.mediaRuntime.saveMediaBuffer(
        Buffer.from(fetched.buffer),
        fetched.contentType ?? attachment.mimeType,
        "inbound",
        maxBytes,
        fileHint,
      );
      out.push({
        path: saved.path,
        url,
        contentType: saved.contentType ?? attachment.mimeType,
        attachment,
      });
    } catch (err) {
      params.logError?.(
        `vkteams: inbound media download failed for ${url}: ${String(err)}`,
      );
      out.push({
        url,
        contentType: attachment.mimeType,
        attachment,
      });
    }
  }
  return out;
}

export function resolveVkTeamsInboundMediaPaths(
  media: readonly VkTeamsInboundResolvedMedia[] | undefined,
): string[] {
  return (media ?? [])
    .map((entry) => entry.path?.trim())
    .filter((entry): entry is string => Boolean(entry));
}

export function resolveVkTeamsInboundMediaUrls(
  media: readonly VkTeamsInboundResolvedMedia[] | undefined,
): string[] {
  return (media ?? [])
    .map((entry) => entry.url?.trim())
    .filter((entry): entry is string => Boolean(entry));
}

export function resolveVkTeamsInboundMediaTypes(
  media: readonly VkTeamsInboundResolvedMedia[] | undefined,
): string[] {
  return (media ?? [])
    .map((entry) => entry.contentType?.trim())
    .filter((entry): entry is string => Boolean(entry));
}

/**
 * Загружает исходящее медиа (локальный путь или URL) в Buffer.
 */
export async function loadVkTeamsOutboundMedia(params: {
  mediaUrl: string;
  mediaLocalRoots?: readonly string[];
  timeoutMs?: number;
}): Promise<{ buffer: Buffer; filename: string; mimeType?: string }> {
  const { mediaUrl } = params;
  const timeoutMs = params.timeoutMs ?? 15_000;

  // Локальный путь
  if (
    mediaUrl.startsWith("/") ||
    mediaUrl.startsWith("./") ||
    mediaUrl.startsWith("../") ||
    /^[A-Za-z]:[\\/]/.test(mediaUrl) ||
    mediaUrl.startsWith("file://")
  ) {
    const filePath = mediaUrl.replace(/^file:\/\//, "");
    const buffer = await readFile(filePath);
    const ext = extname(filePath).toLowerCase();
    return {
      buffer,
      filename: basename(filePath) || "file",
      mimeType: mimeFromExtension(ext),
    };
  }

  // URL
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(mediaUrl, { signal: controller.signal });
    if (!response.ok) {
      throw new Error(`Failed to fetch media: HTTP ${response.status}`);
    }
    const contentType = response.headers.get("content-type")?.split(";")[0]?.trim().toLowerCase();
    const buffer = Buffer.from(await response.arrayBuffer());
    let filename = "file";
    try {
      const pathname = new URL(mediaUrl).pathname;
      const candidate = decodeURIComponent(basename(pathname));
      if (candidate && candidate !== "/") {
        filename = candidate;
      }
    } catch {
      // ignore
    }
    if (filename === "file" || filename === "/") {
      const ext = extensionFromMimeType(contentType);
      filename = `file${ext ?? ".bin"}`;
    }
    return { buffer, filename, mimeType: contentType };
  } finally {
    clearTimeout(timeout);
  }
}

export function resolveVkTeamsInboundMediaKind(
  fileId: string | undefined,
  mimeType: string | undefined,
  filename: string | undefined,
): VkTeamsInboundAttachment["kind"] {
  if (mimeType) {
    const kind = kindFromMimeType(mimeType);
    if (kind !== "document") return kind;
  }
  const ext = filename ? extname(filename).toLowerCase() : undefined;
  if (ext) {
    const kind = kindFromExtension(ext);
    if (kind !== "unknown") return kind;
  }
  // fileId первого символа кодирует тип (см. util.py библиотеки):
  // 0-7 image, 8-F video, G-N audio
  if (fileId) {
    const first = fileId[0];
    if (first >= "0" && first <= "7") return "image";
    if (first >= "8" && first <= "F") return "video";
    if (first >= "G" && first <= "N") return "audio";
  }
  return "document";
}