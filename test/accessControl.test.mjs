import test from "node:test";
import assert from "node:assert/strict";

import { canViewDocument, getEffectivePolicy } from "../lib/accessControl.mjs";

test("admin co quyen ngang director", () => {
  const adminPolicy = getEffectivePolicy({
    role: "admin",
    department: "production",
    policy_override: {},
  });
  const directorPolicy = getEffectivePolicy({
    role: "director",
    department: "executive",
    policy_override: {},
  });

  assert.deepEqual(adminPolicy, directorPolicy);
});

test("tai lieu rieng tu chi nguoi tao va nguoi duoc chia se moi xem duoc", () => {
  const metadata = {
    access_level: "sensitive",
    owner_user_id: "uploader-1",
    shared_with_users: "user-2,user-3",
  };

  assert.equal(canViewDocument({ id: "uploader-1", role: "employee", department: "sales" }, metadata), true);
  assert.equal(canViewDocument({ id: "user-2", role: "employee", department: "finance" }, metadata), true);
  assert.equal(canViewDocument({ id: "admin-1", role: "admin", department: "executive" }, metadata), false);
  assert.equal(canViewDocument({ id: "user-9", role: "manager", department: "sales" }, metadata), false);
});

test("nhan vien xem duoc moi tai lieu muc co ban", () => {
  const metadata = {
    access_level: "basic",
    owner_department: "operations",
    allowed_departments: "operations",
  };

  assert.equal(canViewDocument({ id: "emp-1", role: "employee", department: "operations" }, metadata), true);
  assert.equal(canViewDocument({ id: "emp-2", role: "employee", department: "sales" }, metadata), false);
});

test("quan ly xem duoc moi tai lieu muc nang cao", () => {
  const metadata = {
    access_level: "advanced",
    owner_department: "finance",
    allowed_departments: "finance",
  };

  assert.equal(canViewDocument({ id: "mgr-1", role: "manager", department: "finance" }, metadata), true);
  assert.equal(canViewDocument({ id: "mgr-2", role: "manager", department: "production" }, metadata), false);
  assert.equal(canViewDocument({ id: "dir-1", role: "director", department: "production" }, metadata), true);
  assert.equal(canViewDocument({ id: "emp-1", role: "employee", department: "finance" }, metadata), false);
});

test("policy override cho phep xem them phong ban duoc cap", () => {
  const metadata = {
    access_level: "basic",
    owner_department: "finance",
    allowed_departments: "finance",
  };

  assert.equal(
    canViewDocument(
      {
        id: "mgr-3",
        role: "manager",
        department: "sales",
        policy_override: {
          allowed_departments: ["finance"],
        },
      },
      metadata,
    ),
    true,
  );
});
