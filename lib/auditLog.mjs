import { appendFile, mkdir, readFile } from "node:fs/promises";
import path from "node:path";

export async function appendAuditLog(logFile, entry) {
  await mkdir(path.dirname(logFile), { recursive: true });
  const record = {
    id: buildAuditId(),
    created_at: new Date().toISOString(),
    ...entry,
  };
  await appendFile(logFile, `${JSON.stringify(record)}\n`, "utf8");
  return record;
}

export async function readAuditLogs(logFile, limit = 100) {
  try {
    const raw = await readFile(logFile, "utf8");
    return raw
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .slice(-Math.max(1, limit))
      .reverse();
  } catch {
    return [];
  }
}

function buildAuditId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}
