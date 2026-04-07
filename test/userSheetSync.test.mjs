import test from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { loadUsers, loginWithPassword } from "../lib/accessControl.mjs";

test("tu dong dong bo user tu Google Sheet khi tai danh sach user", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "user-sheet-sync-"));
  const usersFile = path.join(cwd, "data", "access", "users.json");
  await mkdir(path.dirname(usersFile), { recursive: true });
  await writeFile(usersFile, `${JSON.stringify({ users: [] }, null, 2)}\n`, "utf8");

  const csv = [
    "STT,Mã NV,Họ Tên,Chức Vụ,Bộ Phận,Tài Khoản,Quyền Hạn",
    "1,NT001,Đào Thế Anh,Quản Lí,Sản Xuất,NT001,QL",
    "2,NT002,Nguyễn Phương Tú,Nhân Viên/Admin Code,Sản xuất,NT002,NV",
    "3,NT006,Đào Hoàng Sơn,Giám Đốc,Admin,NT006,GĐ",
  ].join("\n");

  const fetchImpl = async () =>
    new Response(csv, {
      status: 200,
      headers: {
        "content-type": "text/csv; charset=utf-8",
      },
    });

  try {
    const users = await loadUsers(usersFile, {
      usersSheetUrl: "https://docs.google.com/spreadsheets/d/test-sheet/edit?gid=0#gid=0",
      usersSyncIntervalMs: 0,
      fetchImpl,
      throwOnUserSyncError: true,
    });
    const raw = JSON.parse(await readFile(usersFile, "utf8"));

    assert.equal(users.length, 3);
    assert.equal(users.find((user) => user.username === "nt001")?.role, "manager");
    assert.equal(users.find((user) => user.username === "nt002")?.role, "admin");
    assert.equal(users.find((user) => user.username === "nt006")?.role, "director");
    assert.equal(users.find((user) => user.username === "nt002")?.department, "production");
    assert.equal(raw.users.every((user) => user.managed_by === "sheet"), true);

    const session = await loginWithPassword(usersFile, "NT002", "123456", {
      usersSheetUrl: "https://docs.google.com/spreadsheets/d/test-sheet/edit?gid=0#gid=0",
      usersSyncIntervalMs: 0,
      fetchImpl,
      throwOnUserSyncError: true,
    });

    assert.equal(session.user.username, "nt002");
    assert.equal(session.user.role, "admin");
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("giu nguyen mat khau da doi khi sheet khong con cot mat khau", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "user-sheet-password-"));
  const usersFile = path.join(cwd, "data", "access", "users.json");
  await mkdir(path.dirname(usersFile), { recursive: true });
  await writeFile(
    usersFile,
    `${JSON.stringify({
      users: [
        {
          id: "sheet-nt002",
          name: "Nguyen Phuong Tu",
          username: "nt002",
          role: "admin",
          department: "production",
          password_hash: "481f6cc0511143ccdd7e2d1b1b94faf0a700a8b49cd13922a70b5ae28acaa8c5",
          managed_by: "sheet",
          employee_code: "NT002",
          title: "Nhan Vien/Admin Code",
          source_permission: "NV",
          policy_override: {},
        },
      ],
    }, null, 2)}\n`,
    "utf8",
  );

  const csv = [
    "STT,Mã NV,Họ Tên,Chức Vụ,Bộ Phận,Tài Khoản,Quyền Hạn",
    "1,NT002,Nguyễn Phương Tú,Nhân Viên/Admin Code,Sản xuất,NT002,NV",
  ].join("\n");

  const fetchImpl = async () =>
    new Response(csv, {
      status: 200,
      headers: {
        "content-type": "text/csv; charset=utf-8",
      },
    });

  try {
    await loadUsers(usersFile, {
      usersSheetUrl: "https://docs.google.com/spreadsheets/d/test-sheet/edit?gid=0#gid=0",
      usersSyncIntervalMs: 0,
      fetchImpl,
      throwOnUserSyncError: true,
    });

    const session = await loginWithPassword(usersFile, "NT002", "654321", {
      usersSheetUrl: "https://docs.google.com/spreadsheets/d/test-sheet/edit?gid=0#gid=0",
      usersSyncIntervalMs: 0,
      fetchImpl,
      throwOnUserSyncError: true,
    });

    assert.equal(session.user.username, "nt002");
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});
