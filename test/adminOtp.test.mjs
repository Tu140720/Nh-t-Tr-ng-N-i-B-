import test from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { loginWithPassword, verifyLoginOtp } from "../lib/accessControl.mjs";

test("admin dang nhap qua OTP email", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "admin-otp-"));
  const usersFile = path.join(cwd, "data", "access", "users.json");
  await mkdir(path.dirname(usersFile), { recursive: true });
  await writeFile(
    usersFile,
    `${JSON.stringify({
      users: [
        {
          id: "u-admin-code",
          name: "Admin Code",
          username: "nt002",
          role: "admin",
          department: "production",
          password_hash: "8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92",
          email: "admin@example.com",
          policy_override: {},
        },
      ],
    }, null, 2)}\n`,
    "utf8",
  );

  let deliveredCode = "";

  try {
    const firstStep = await loginWithPassword(usersFile, "nt002", "123456", {
      requireAdminOtp: true,
      sendAdminOtp: async ({ code }) => {
        deliveredCode = code;
      },
      adminOtpTtlMs: 60000,
    });

    assert.equal(firstStep.otp_required, true);
    assert.ok(firstStep.challengeToken);
    assert.match(firstStep.delivery, /admin@example.com/);
    assert.equal(deliveredCode.length, 6);

    const session = await verifyLoginOtp(usersFile, firstStep.challengeToken, deliveredCode);
    assert.ok(session.token);
    assert.equal(session.user.username, "nt002");
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});
