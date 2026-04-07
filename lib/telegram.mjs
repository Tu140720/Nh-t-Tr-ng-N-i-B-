import { createHmac, timingSafeEqual } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const DEFAULT_INIT_DATA_MAX_AGE_SECONDS = 86_400;

function formatNow() {
  return new Date().toISOString();
}

async function ensureTelegramLinkStore(filePath) {
  await mkdir(path.dirname(filePath), { recursive: true });
  try {
    await readFile(filePath, "utf8");
  } catch (error) {
    if (error && error.code !== "ENOENT") {
      throw error;
    }
    await writeFile(filePath, `${JSON.stringify({ links: [] }, null, 2)}\n`, "utf8");
  }
}

function normalizeTelegramLink(record = {}) {
  return {
    internal_user_id: String(record.internal_user_id || "").trim(),
    telegram_user_id: String(record.telegram_user_id || "").trim(),
    telegram_username: String(record.telegram_username || "").trim(),
    first_name: String(record.first_name || "").trim(),
    last_name: String(record.last_name || "").trim(),
    linked_at: String(record.linked_at || formatNow()).trim(),
    updated_at: String(record.updated_at || record.linked_at || formatNow()).trim(),
  };
}

async function writeTelegramLinks(filePath, links = []) {
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify({ links }, null, 2)}\n`, "utf8");
}

export async function readTelegramLinks(filePath) {
  await ensureTelegramLinkStore(filePath);
  const raw = await readFile(filePath, "utf8");
  const parsed = JSON.parse(raw);
  return (Array.isArray(parsed?.links) ? parsed.links : [])
    .map((item) => normalizeTelegramLink(item))
    .filter((item) => item.internal_user_id && item.telegram_user_id);
}

export async function findTelegramLinkByInternalUserId(filePath, internalUserId) {
  const targetUserId = String(internalUserId || "").trim();
  if (!targetUserId) {
    return null;
  }
  const links = await readTelegramLinks(filePath);
  return links.find((item) => item.internal_user_id === targetUserId) || null;
}

export async function findTelegramLinkByTelegramUserId(filePath, telegramUserId) {
  const targetUserId = String(telegramUserId || "").trim();
  if (!targetUserId) {
    return null;
  }
  const links = await readTelegramLinks(filePath);
  return links.find((item) => item.telegram_user_id === targetUserId) || null;
}

export async function upsertTelegramLink(filePath, { internalUserId, telegramUser } = {}) {
  const normalizedInternalUserId = String(internalUserId || "").trim();
  const normalizedTelegramUserId = String(telegramUser?.id || "").trim();
  if (!normalizedInternalUserId || !normalizedTelegramUserId) {
    throw new Error("Khong du thong tin de lien ket Telegram.");
  }

  const links = await readTelegramLinks(filePath);
  const nowIso = formatNow();
  const previous =
    links.find((item) => item.internal_user_id === normalizedInternalUserId || item.telegram_user_id === normalizedTelegramUserId) || null;
  const nextLink = normalizeTelegramLink({
    internal_user_id: normalizedInternalUserId,
    telegram_user_id: normalizedTelegramUserId,
    telegram_username: String(telegramUser?.username || "").trim(),
    first_name: String(telegramUser?.first_name || "").trim(),
    last_name: String(telegramUser?.last_name || "").trim(),
    linked_at: previous?.linked_at || nowIso,
    updated_at: nowIso,
  });
  const nextLinks = [
    nextLink,
    ...links.filter(
      (item) =>
        item.internal_user_id !== normalizedInternalUserId &&
        item.telegram_user_id !== normalizedTelegramUserId,
    ),
  ];
  await writeTelegramLinks(filePath, nextLinks);
  return nextLink;
}

export function validateTelegramInitData(initData, botToken, options = {}) {
  const rawInitData = String(initData || "").trim();
  const normalizedBotToken = String(botToken || "").trim();
  if (!rawInitData) {
    throw new Error("Thieu init data tu Telegram.");
  }
  if (!normalizedBotToken) {
    throw new Error("He thong chua cau hinh bot Telegram.");
  }

  const params = new URLSearchParams(rawInitData);
  const hash = String(params.get("hash") || "").trim();
  if (!hash) {
    throw new Error("Init data Telegram khong hop le.");
  }

  const dataCheckString = [...params.entries()]
    .filter(([key]) => key !== "hash")
    .sort(([leftKey], [rightKey]) => leftKey.localeCompare(rightKey))
    .map(([key, value]) => `${key}=${value}`)
    .join("\n");

  const secretKey = createHmac("sha256", "WebAppData").update(normalizedBotToken).digest();
  const computedHash = createHmac("sha256", secretKey).update(dataCheckString).digest();
  const providedHash = Buffer.from(hash, "hex");
  if (providedHash.length !== computedHash.length || !timingSafeEqual(providedHash, computedHash)) {
    throw new Error("Khong the xac minh du lieu Telegram.");
  }

  const maxAgeSeconds = Number(options.maxAgeSeconds || DEFAULT_INIT_DATA_MAX_AGE_SECONDS);
  const authDate = Number(params.get("auth_date") || 0);
  if (Number.isFinite(maxAgeSeconds) && maxAgeSeconds > 0 && Number.isFinite(authDate) && authDate > 0) {
    const ageSeconds = Math.floor(Date.now() / 1000) - authDate;
    if (ageSeconds > maxAgeSeconds) {
      throw new Error("Du lieu Telegram da het han.");
    }
  }

  let telegramUser = null;
  const userRaw = String(params.get("user") || "").trim();
  if (userRaw) {
    try {
      telegramUser = JSON.parse(userRaw);
    } catch {
      throw new Error("Khong doc duoc thong tin nguoi dung Telegram.");
    }
  }

  return {
    authDate,
    queryId: String(params.get("query_id") || "").trim(),
    user: telegramUser,
  };
}

export async function sendTelegramTextMessage(botToken, chatId, text, options = {}) {
  const normalizedBotToken = String(botToken || "").trim();
  const normalizedChatId = String(chatId || "").trim();
  const normalizedText = String(text || "").trim();
  if (!normalizedBotToken || !normalizedChatId || !normalizedText) {
    throw new Error("Khong du thong tin gui tin nhan Telegram.");
  }

  const response = await fetch(`https://api.telegram.org/bot${normalizedBotToken}/sendMessage`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify({
      chat_id: normalizedChatId,
      text: normalizedText.slice(0, 4000),
      disable_notification: options.silent === true,
      disable_web_page_preview: true,
    }),
  });

  let payload = {};
  try {
    payload = await response.json();
  } catch {
    payload = {};
  }

  if (!response.ok || payload?.ok === false) {
    throw new Error(payload?.description || `Telegram API ${response.status}`);
  }

  return payload?.result || null;
}
