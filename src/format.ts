// Конвертация текста (markdown/HTML от OpenClaw) в формат VK Teams.
// VK Teams поддерживает parseMode=HTML/MarkdownV2 либо format-стили:
//   {"bold":[{"offset":0,"length":2}],"italic":[...],"link":[{"offset":0,"length":2,"url":"..."}]}
// Поддерживаемые стили: bold, italic, underline, strikethrough, link, mention,
// inline_code, pre, ordered_list, unordered_list, quote.

export type VkTeamsFormatType =
  | "bold"
  | "italic"
  | "underline"
  | "strikethrough"
  | "link"
  | "mention"
  | "inline_code"
  | "pre"
  | "ordered_list"
  | "unordered_list"
  | "quote";

export type VkTeamsFormatItem = {
  type: VkTeamsFormatType;
  offset: number;
  length: number;
  url?: string;
};

export type VkTeamsFormatData = Partial<Record<VkTeamsFormatType, VkTeamsFormatItem[]>>;

export type VkTeamsPreparedFormattedMessage = {
  text: string;
  formatData?: VkTeamsFormatData;
};

export type VkTeamsMarkdownPipelineOptions = {
  chunkSize?: number;
};

type StyleAccumulator = Record<VkTeamsFormatType, VkTeamsFormatItem[]>;

const EMPTY_STYLES: StyleAccumulator = {
  bold: [],
  italic: [],
  underline: [],
  strikethrough: [],
  link: [],
  mention: [],
  inline_code: [],
  pre: [],
  ordered_list: [],
  unordered_list: [],
  quote: [],
};

const VOID_TAGS = new Set([
  "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
  "meta", "param", "source", "track", "wbr",
]);

function hasStyles(styles: StyleAccumulator): boolean {
  return Object.values(styles).some((ranges) => ranges.length > 0);
}

function toFormatData(styles: StyleAccumulator): VkTeamsFormatData | undefined {
  if (!hasStyles(styles)) {
    return undefined;
  }
  // VK Teams API принимает format только в группированном виде:
  //   {"bold":[{"offset":0,"length":4}], "link":[{"offset":0,"length":2,"url":"..."}]}
  // (проверено 08.10.2026: {"version":1,"items":[...]} -> "Format error").
  const data: VkTeamsFormatData = {};
  for (const [type, ranges] of Object.entries(styles) as Array<[VkTeamsFormatType, VkTeamsFormatItem[]]>) {
    const items = ranges.map(({ type: _type, ...rest }) => rest);
    if (items.length > 0) {
      data[type] = items;
    }
  }
  return data;
}

/**
 * Собирает format-объект VK Teams из HTML-строки.
 * Поддерживает: b/strong, i/em, u, s/strike/del, a[href], code, pre,
 * blockquote, ul/ol/li.
 */
export function htmlToVkTeamsFormat(html: string): VkTeamsPreparedFormattedMessage {
  const styles: StyleAccumulator = {
    bold: [],
    italic: [],
    underline: [],
    strikethrough: [],
    link: [],
    mention: [],
    inline_code: [],
    pre: [],
    ordered_list: [],
    unordered_list: [],
    quote: [],
  };
  const raw: string[] = [];
  let rawLen = 0;
  const stack: Array<{ style: VkTeamsFormatType; url?: string; offset: number }> = [];

  const pushRaw = (s: string) => {
    if (!s) return;
    raw.push(s);
    rawLen += s.length;
  };

  const closeStyle = (style: VkTeamsFormatType, url?: string) => {
    const idx = stack
      .map((s) => `${s.style}:${s.url ?? ""}`)
      .lastIndexOf(`${style}:${url ?? ""}`);
    if (idx === -1) return;
    const entry = stack[idx];
    const length = rawLen - entry.offset;
    if (length > 0) {
      const item: VkTeamsFormatItem = { type: style, offset: entry.offset, length };
      if (entry.url) item.url = entry.url;
      styles[style].push(item);
    }
    stack.splice(idx, 1);
  };

  const openStyle = (style: VkTeamsFormatType, url?: string) => {
    stack.push({ style, url, offset: rawLen });
  };

  const tagRe = /<\/?([a-zA-Z][a-zA-Z0-9]*)((?:\s+[a-zA-Z-]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*\/?>/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = tagRe.exec(html)) !== null) {
    if (match.index > lastIndex) {
      pushRaw(html.slice(lastIndex, match.index));
    }
    lastIndex = tagRe.lastIndex;

    const full = match[0];
    const tagName = match[1].toLowerCase();
    const attrsRaw = match[2] ?? "";
    const isClosing = full.startsWith("</");

    let href: string | undefined;
    const hrefMatch = attrsRaw.match(/href\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i);
    if (hrefMatch) {
      href = hrefMatch[1] ?? hrefMatch[2] ?? hrefMatch[3];
    }

    if (VOID_TAGS.has(tagName) || full.endsWith("/>")) {
      continue;
    }

    switch (tagName) {
      case "b":
      case "strong":
        if (isClosing) closeStyle("bold");
        else openStyle("bold");
        break;
      case "i":
      case "em":
        if (isClosing) closeStyle("italic");
        else openStyle("italic");
        break;
      case "u":
        if (isClosing) closeStyle("underline");
        else openStyle("underline");
        break;
      case "s":
      case "strike":
      case "del":
        if (isClosing) closeStyle("strikethrough");
        else openStyle("strikethrough");
        break;
      case "a":
        if (isClosing) closeStyle("link", href);
        else openStyle("link", href);
        break;
      case "code":
        if (isClosing) closeStyle("inline_code");
        else openStyle("inline_code");
        break;
      case "pre":
        if (isClosing) closeStyle("pre");
        else openStyle("pre");
        break;
      case "blockquote":
        if (isClosing) closeStyle("quote");
        else openStyle("quote");
        break;
      case "ul":
        if (isClosing) closeStyle("unordered_list");
        else openStyle("unordered_list");
        break;
      case "ol":
        if (isClosing) closeStyle("ordered_list");
        else openStyle("ordered_list");
        break;
      case "li":
        if (isClosing) closeStyle("unordered_list");
        else openStyle("unordered_list");
        break;
      case "br":
      case "p":
      case "div":
      case "h1":
      case "h2":
      case "h3":
      case "h4":
      case "h5":
      case "h6":
        if (!isClosing && (tagName === "br" || tagName === "p" || tagName === "div")) {
          pushRaw("\n");
        }
        break;
      default:
        break;
    }
  }

  if (lastIndex < html.length) {
    pushRaw(html.slice(lastIndex));
  }

  while (stack.length > 0) {
    const entry = stack.pop()!;
    const length = rawLen - entry.offset;
    if (length > 0) {
      const item: VkTeamsFormatItem = { type: entry.style, offset: entry.offset, length };
      if (entry.url) item.url = entry.url;
      styles[entry.style].push(item);
    }
  }

  const rawText = raw.join("");
  const collapsed = rawText.replace(/\n{3,}/g, "\n\n");
  const leading = collapsed.length - collapsed.trimStart().length;
  const text = collapsed.trim();

  // Карта соответствия индексов: исходный текст -> нормализованный.
  const map = new Array<number>(rawText.length + 1).fill(0);
  {
    let out = 0;
    let i = 0;
    while (i < rawText.length) {
      if (rawText[i] === "\n") {
        let j = i;
        while (j < rawText.length && rawText[j] === "\n") j++;
        const run = j - i;
        const keep = Math.min(run, 2);
        for (let k = 0; k < keep; k++) {
          map[i + k] = out;
          out += 1;
        }
        for (let k = keep; k < run; k++) {
          map[i + k] = out;
        }
        i = j;
      } else {
        map[i] = out;
        out += 1;
        i += 1;
      }
    }
    map[rawText.length] = out;
  }

  const remapped: StyleAccumulator = {
    bold: [],
    italic: [],
    underline: [],
    strikethrough: [],
    link: [],
    mention: [],
    inline_code: [],
    pre: [],
    ordered_list: [],
    unordered_list: [],
    quote: [],
  };
  for (const [type, items] of Object.entries(styles) as Array<[VkTeamsFormatType, VkTeamsFormatItem[]]>) {
    for (const item of items) {
      const start = (map[item.offset] ?? 0) - leading;
      const end = (map[item.offset + item.length] ?? text.length) - leading;
      const s = Math.max(0, start);
      const e = Math.min(text.length, end);
      if (e > s) {
        const next: VkTeamsFormatItem = { type, offset: s, length: e - s };
        if (item.url) next.url = item.url;
        remapped[type].push(next);
      }
    }
  }

  return { text, formatData: toFormatData(remapped) };
}

/**
 * Собирает format-объект из простого markdown (жирный, курсив, код, ссылки).
 * Используется как fallback, если текст пришёл в markdown-разметке.
 */
export function markdownToVkTeamsFormat(md: string): VkTeamsPreparedFormattedMessage {
  const styles: StyleAccumulator = {
    bold: [],
    italic: [],
    underline: [],
    strikethrough: [],
    link: [],
    mention: [],
    inline_code: [],
    pre: [],
    ordered_list: [],
    unordered_list: [],
    quote: [],
  };
  const out: string[] = [];
  let offset = 0;

  const push = (s: string) => {
    out.push(s);
    offset += s.length;
  };

  // Инлайн-обработка: вырезает маркеры, регистрирует стили по чистому тексту.
  // Поддержка: **bold**/__bold__, *italic*/_italic_, `code`, ~~strike~~, [text](url).
  const pushInline = (raw: string) => {
    const mdRe =
      /(\*\*|__)(.+?)\1|(?<!\*)\*([^*\n]+)\*(?!\*)|(?<!_)_([^_\n]+)_(?!_)|`([^`\n]+)`|~~([^~\n]+)~~|\[([^\]]+)\]\(([^)\s]+)\)/g;
    let last = 0;
    let m: RegExpExecArray | null;
    while ((m = mdRe.exec(raw)) !== null) {
      push(raw.slice(last, m.index));
      last = mdRe.lastIndex;
      if (m[2] !== undefined) {
        const inner = m[2];
        const start = offset;
        push(inner);
        styles.bold.push({ type: "bold", offset: start, length: inner.length });
      } else if (m[3] !== undefined || m[4] !== undefined) {
        const inner = m[3] ?? m[4];
        const start = offset;
        push(inner);
        styles.italic.push({ type: "italic", offset: start, length: inner.length });
      } else if (m[5] !== undefined) {
        const inner = m[5];
        const start = offset;
        push(inner);
        styles.inline_code.push({ type: "inline_code", offset: start, length: inner.length });
      } else if (m[6] !== undefined) {
        const inner = m[6];
        const start = offset;
        push(inner);
        styles.strikethrough.push({ type: "strikethrough", offset: start, length: inner.length });
      } else if (m[7] !== undefined && m[8] !== undefined) {
        const text = m[7];
        const url = m[8];
        const start = offset;
        push(text);
        styles.link.push({ type: "link", offset: start, length: text.length, url });
      }
    }
    push(raw.slice(last));
  };

  // Блочная обработка: код-блоки ```...```, цитаты >, списки -/*/1.
  const lines = md.split("\n");
  let inCodeBlock = false;
  let codeBlockBuffer: string[] = [];
  let listType: "ordered_list" | "unordered_list" | null = null;
  let listStart = 0;
  let inQuote = false;
  let quoteStart = 0;

  const flushList = () => {
    if (listType) {
      const length = offset - listStart;
      if (length > 0) {
        styles[listType].push({ type: listType, offset: listStart, length });
      }
      listType = null;
    }
  };

  const flushQuote = () => {
    if (inQuote) {
      const length = offset - quoteStart;
      if (length > 0) {
        styles.quote.push({ type: "quote", offset: quoteStart, length });
      }
      inQuote = false;
    }
  };

  for (const line of lines) {
    if (line.trimStart().startsWith("```")) {
      if (!inCodeBlock) {
        inCodeBlock = true;
        codeBlockBuffer = [];
      } else {
        inCodeBlock = false;
        const content = codeBlockBuffer.join("\n");
        const start = offset;
        push(content);
        if (content.length > 0) {
          styles.pre.push({ type: "pre", offset: start, length: content.length });
        }
        push("\n");
      }
      continue;
    }
    if (inCodeBlock) {
      codeBlockBuffer.push(line);
      continue;
    }

    const quoteMatch = line.match(/^>\s?(.*)$/);
    if (quoteMatch) {
      flushList();
      if (!inQuote) {
        inQuote = true;
        quoteStart = offset;
      }
      pushInline(quoteMatch[1]);
      push("\n");
      continue;
    }
    flushQuote();

    const ulMatch = line.match(/^[-*+]\s+(.*)$/);
    const olMatch = line.match(/^\d+[.)]\s+(.*)$/);
    if (ulMatch || olMatch) {
      const type = olMatch ? "ordered_list" : "unordered_list";
      if (listType !== type) {
        flushList();
        listType = type;
        listStart = offset;
      }
      pushInline((ulMatch ?? olMatch)[1]);
      push("\n");
      continue;
    }
    flushList();

    pushInline(line);
    push("\n");
  }
  flushList();
  flushQuote();
  if (inCodeBlock && codeBlockBuffer.length > 0) {
    const content = codeBlockBuffer.join("\n");
    const start = offset;
    push(content);
    if (content.length > 0) {
      styles.pre.push({ type: "pre", offset: start, length: content.length });
    }
  }

  let text = out.join("");
  if (text.endsWith("\n")) {
    text = text.slice(0, -1);
  }

  return { text, formatData: toFormatData(styles) };
}


function isHtml(text: string): boolean {
  return /<[a-zA-Z][^>]*>/.test(text);
}

/**
 * Разбивает markdown-текст на чанки и возвращает подготовленные сообщения
 * с formatData для VK Teams API.
 */
export function renderVkTeamsMarkdownChunks(
  markdown: string,
  options: VkTeamsMarkdownPipelineOptions = {},
): VkTeamsPreparedFormattedMessage[] {
  const source = markdown ?? "";
  if (!source.trim()) {
    return [{ text: "" }];
  }

  const chunkSize = options.chunkSize ?? 4096;
  const prepared = isHtml(source)
    ? htmlToVkTeamsFormat(source)
    : markdownToVkTeamsFormat(source);

  const chunks: VkTeamsPreparedFormattedMessage[] = [];
  const text = prepared.text;
  if (text.length <= chunkSize) {
    chunks.push(prepared);
    return chunks;
  }

  // Нарезка по границам абзацев/строк с переносом format-офсетов
  const boundaries: number[] = [0];
  const lines = text.split("\n");
  let acc = "";
  for (const line of lines) {
    if (acc.length + line.length + 1 > chunkSize && acc.length > 0) {
      boundaries.push(acc.length);
      acc = line;
    } else {
      acc = acc ? `${acc}\n${line}` : line;
    }
  }
  if (acc.length > 0) {
    boundaries.push(text.length);
  }

  const shiftItems = (data: VkTeamsFormatData | undefined, start: number, end: number): VkTeamsFormatData => {
    if (!data) return {};
    const out: VkTeamsFormatData = {};
    for (const [type, items] of Object.entries(data) as Array<[VkTeamsFormatType, VkTeamsFormatItem[]]>) {
      const shifted: VkTeamsFormatItem[] = [];
      for (const item of items ?? []) {
        const itemStart = item.offset;
        const itemEnd = item.offset + item.length;
        if (itemEnd <= start || itemStart >= end) continue;
        const newStart = Math.max(itemStart, start) - start;
        const newEnd = Math.min(itemEnd, end) - start;
        if (newEnd > newStart) {
          shifted.push({ ...item, offset: newStart, length: newEnd - newStart });
        }
      }
      if (shifted.length > 0) {
        out[type] = shifted;
      }
    }
    return out;
  };

  for (let i = 0; i < boundaries.length - 1; i++) {
    const start = boundaries[i] ?? 0;
    const end = boundaries[i + 1] ?? text.length;
    const chunkText = text.slice(start, end).trim();
    if (!chunkText) continue;
    const formatData = shiftItems(prepared.formatData, start, end);
    chunks.push(
      Object.keys(formatData).length > 0
        ? { text: chunkText, formatData }
        : { text: chunkText },
    );
  }

  return chunks.length > 0 ? chunks : [{ text }];
}