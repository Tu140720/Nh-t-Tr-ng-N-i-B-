import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { appendAuditLog, readAuditLogs } from "../lib/auditLog.mjs";

test("ghi va doc audit log theo thu tu moi nhat", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "audit-log-"));

  try {
    const logFile = path.join(cwd, "data", "audit", "log.jsonl");
    await appendAuditLog(logFile, {
      action: "folder.create",
      actor: { id: "u1", username: "nt002" },
      detail: { folder_name: "Kho A" },
    });
    await appendAuditLog(logFile, {
      action: "document.delete",
      actor: { id: "u2", username: "nt003" },
      detail: { title: "Tai lieu X" },
    });

    const logs = await readAuditLogs(logFile, 10);
    assert.equal(logs.length, 2);
    assert.equal(logs[0].action, "document.delete");
    assert.equal(logs[1].action, "folder.create");
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});
