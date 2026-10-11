# OpenClaw VK Teams (Среда) Channel Plugin

Плагин-канал VK Teams (Среда) для [OpenClaw](https://openclaw.ai).

Интеграция бота через long polling: приём сообщений через `events/get`, отправка
текста с format-стилями (жирный, курсив, ссылки, цитаты, списки, код-блоки),
редактирование и удаление сообщений, индикатор набора текста.

---

## Возможности / Features

- **Long polling** — `events/get` с `lastEventId`, автоматический retry с backoff
- **Markdown → VK Teams format-стили** — конвертер с учётом кириллицы
  (оффсеты по символам): жирный, курсив, подчёркнутый, зачёркнутый, ссылки,
  упоминания, инлайн-код, код-блоки, нумерованные/маркированные списки, цитаты
- **editMessage / deleteMessage** — редактирование и удаление сообщений
- **Typing indicator** — `chats/sendActions` (с дебаунсом)
- **Фильтрация своих сообщений** — бот не реагирует на самого себя
- **Медиа** — файлы, голосовые, стикеры через `files/getInfo`
- **DM-политики** — `pairing` / `allowlist` / `open` / `disabled`
- **Разбивка длинных сообщений** — 4096 символов

## Установка / Install

```bash
cd vkteams-plugin
openclaw plugins install . --force --accept-capabilities
```

## Настройка / Configuration

В `openclaw.json`:

```json5
channels: {
  vkteams: {
    enabled: true,
    token: "<bot-token>",                 // или tokenFile
    apiBaseUrl: "https://api.icq.net/bot/v1",
    dmPolicy: "pairing",
    allowFrom: ["user1@domain", "user2@domain"],
    pollTime: 30
  }
}
```

| Ключ / Key | Описание / Description | По умолчанию / Default |
|---|---|---|
| `token` | Токен бота VK Teams | — (обязательно) |
| `tokenFile` | Путь к файлу с токеном | — |
| `apiBaseUrl` | Базовый URL Bot API | `https://api.icq.net/bot/v1` |
| `dmPolicy` | `pairing` / `allowlist` / `open` / `disabled` | `pairing` |
| `allowFrom` | Разрешённые ID пользователей | `[]` |
| `pollTime` | Секунд на запрос `events/get` | `30` |

## Заметки по VK Teams Bot API

- **Авторизация:** параметр `token` в query (URL-кодирование; `:` → `%3A`).
- **Webhook'ов нет** — только long polling:
  `events/get?pollTime=<s>&lastEventId=<id>`.
- `messages/sendText` принимает JSON `format` с группированными стилями:
  `{"bold":[{"offset":0,"length":4}], "link":[{"offset":0,"length":2,"url":"..."}]}`.
  Оффсеты — **по символам**, не по байтам.

## Сборка / Build

```bash
esbuild index.ts --bundle --format=esm --platform=node --target=node24 \
  '--external:openclaw/*' --outfile=dist/index.js
esbuild setup-entry.ts --bundle --format=esm --platform=node --target=node24 \
  '--external:openclaw/*' --outfile=dist/setup-entry.js
```

## Лицензия / License

MIT
