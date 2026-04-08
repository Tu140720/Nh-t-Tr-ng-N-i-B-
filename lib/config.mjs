import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

export function loadConfig(cwd = process.cwd()) {
  loadDotEnv(cwd);

  const port = Number(process.env.PORT || 5000);
  const host = String(process.env.HOST || "0.0.0.0").trim() || "0.0.0.0";
  const sheetSyncIntervalMs = Number(
    process.env.SHEET_SYNC_INTERVAL_MS || process.env.SHEET_SYNC_INTERVAL_MINUTES * 60_000 || 120_000,
  );
  const internalDocsDir = path.resolve(
    cwd,
    process.env.INTERNAL_DOCS_DIR || "./data/internal",
  );
  const sourcesDir = path.resolve(
    cwd,
    process.env.SOURCES_DIR || "./data/sources",
  );
  const uploadsDir = path.resolve(
    cwd,
    process.env.UPLOADS_DIR || "./data/uploads",
  );
  const usersFile = path.resolve(
    cwd,
    process.env.USERS_FILE || "./data/access/users.json",
  );
  const auditLogFile = path.resolve(
    cwd,
    process.env.AUDIT_LOG_FILE || "./data/audit/log.jsonl",
  );
  const telegramLinksFile = path.resolve(
    cwd,
    process.env.TELEGRAM_LINKS_FILE || "./data/telegram/links.json",
  );
  const usersSheetUrl = String(process.env.USERS_SHEET_URL || "").trim();
  const usersSyncIntervalMs = Number(
    process.env.USERS_SYNC_INTERVAL_MS || sheetSyncIntervalMs || 120_000,
  );
  const publicDir = path.resolve(cwd, "./public");

  return {
    cwd,
    host,
    port: Number.isFinite(port) ? port : 5000,
    sheetSyncIntervalMs:
      Number.isFinite(sheetSyncIntervalMs) && sheetSyncIntervalMs >= 0
        ? sheetSyncIntervalMs
        : 120_000,
    publicDir,
    internalDocsDir,
    sourcesDir,
    uploadsDir,
    usersFile,
    auditLogFile,
    usersSheetUrl,
    usersSyncIntervalMs:
      Number.isFinite(usersSyncIntervalMs) && usersSyncIntervalMs >= 0
        ? usersSyncIntervalMs
        : 120_000,
    auth: {
      adminOtpTtlMs: Number(process.env.ADMIN_OTP_TTL_MS || 300000),
      sessionTtlMs: Number(process.env.SESSION_TTL_MS || 30 * 24 * 60 * 60 * 1000),
    },
    mail: {
      host: String(process.env.SMTP_HOST || "").trim(),
      port: Number(process.env.SMTP_PORT || 465),
      secure: String(process.env.SMTP_SECURE || "true").trim().toLowerCase() !== "false",
      user: String(process.env.SMTP_USER || "").trim(),
      pass: String(process.env.SMTP_PASS || "").trim(),
      from: String(process.env.SMTP_FROM || "").trim(),
    },
    gemini: {
      apiKey: process.env.GEMINI_API_KEY || "",
      baseUrl: (
        process.env.GEMINI_BASE_URL || "https://generativelanguage.googleapis.com/v1beta"
      ).replace(/\/$/, ""),
      model: process.env.GEMINI_MODEL || "gemini-3-flash-preview",
      fallbackModels: parseModelList(process.env.GEMINI_FALLBACK_MODELS),
      webModel: process.env.GEMINI_WEB_MODEL || "",
      webFallbackModel: process.env.GEMINI_WEB_FALLBACK_MODEL || "gemini-2.5-flash",
      webFallbackModels: parseModelList(process.env.GEMINI_WEB_FALLBACK_MODELS),
      thinkingLevel: process.env.GEMINI_THINKING_LEVEL || "minimal",
    },
    telegram: {
      botToken: String(process.env.TELEGRAM_BOT_TOKEN || "").trim(),
      linksFile: telegramLinksFile,
      initDataMaxAgeSeconds: Number(process.env.TELEGRAM_INIT_DATA_MAX_AGE_SECONDS || 86_400),
    },
  };
}

function parseModelList(value) {
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function loadDotEnv(cwd) {
  const envPath = path.resolve(cwd, ".env");
  if (!existsSync(envPath)) {
    return;
  }

  const content = readFileSync(envPath, "utf8");
  const lines = content.split(/\r?\n/);

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex === -1) {
      continue;
    }

    const key = trimmed.slice(0, separatorIndex).trim();
    let value = trimmed.slice(separatorIndex + 1).trim();

    if (!key || process.env[key] !== undefined) {
      continue;
    }

    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    process.env[key] = value;
  }
}
