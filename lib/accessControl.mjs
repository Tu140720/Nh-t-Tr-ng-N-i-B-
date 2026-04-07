import { createHash, randomBytes } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const ACCESS_LEVELS = ["basic", "advanced", "sensitive"];

const ROLE_POLICIES = {
  employee: {
    max_access_level: "basic",
    department_scope: "own",
    allowed_departments: [],
    can_manage_permissions: false,
  },
  manager: {
    max_access_level: "advanced",
    department_scope: "own",
    allowed_departments: [],
    can_manage_permissions: false,
  },
  director: {
    max_access_level: "sensitive",
    department_scope: "all",
    allowed_departments: [],
    can_manage_permissions: true,
  },
  admin: {
    max_access_level: "sensitive",
    department_scope: "all",
    allowed_departments: [],
    can_manage_permissions: true,
  },
};

const DEFAULT_USERS = [];

const sessions = new Map();
const userSheetSyncState = new Map();
const loginChallenges = new Map();

async function loadSessions(usersFile) {
  void usersFile;
}

async function saveSessions(usersFile) {
  void usersFile;
}

export async function ensureUserStore(usersFile) {
  await mkdir(path.dirname(usersFile), { recursive: true });
  try {
    await readFile(usersFile, "utf8");
  } catch {
    await writeFile(
      usersFile,
      `${JSON.stringify({ users: DEFAULT_USERS }, null, 2)}\n`,
      "utf8",
    );
  }
}

export async function loadUsers(usersFile, options = {}) {
  await ensureUserStore(usersFile);
  await maybeSyncUsersFromSheet(usersFile, options);
  const raw = await readFile(usersFile, "utf8");
  const parsed = JSON.parse(raw);
  const users = Array.isArray(parsed.users) ? parsed.users : DEFAULT_USERS;
  const normalizedUsers = users.map((user) => mergeWithDefaultUser(normalizeUser(user)));
  await writeUsers(usersFile, normalizedUsers);
  return normalizedUsers;
}

export async function resolveCurrentUser(usersFile, request, options = {}) {
  const users = await loadUsers(usersFile, options);
  await loadSessions(usersFile);
  const token = extractBearerToken(request.headers.authorization);
  const session = token ? sessions.get(token) : null;
  const currentUser = session ? users.find((user) => user.id === session.user_id) || null : null;
  return {
    currentUser,
    users,
  };
}

export async function resolveCurrentUserFromToken(usersFile, token, options = {}) {
  const users = await loadUsers(usersFile, options);
  await loadSessions(usersFile);
  const normalizedToken = String(token || "").trim();
  const session = normalizedToken ? sessions.get(normalizedToken) : null;
  const currentUser = session ? users.find((user) => user.id === session.user_id) || null : null;
  return {
    currentUser,
    users,
  };
}

export function sanitizeUser(user) {
  const policy = getEffectivePolicy(user);
  return {
    id: user.id,
    name: user.name,
    username: user.username,
    role: user.role,
    department: user.department,
    employee_code: user.employee_code || "",
    title: user.title || "",
    email: user.email || "",
    policy,
  };
}

export async function loginWithPassword(usersFile, username, password, options = {}) {
  const users = await loadUsers(usersFile, options);
  const normalizedUsername = String(username || "").trim().toLowerCase();
  const user = users.find((item) => item.username === normalizedUsername);
  if (!user || user.password_hash !== hashPassword(password)) {
    throw new Error("Tai khoan hoac mat khau khong dung.");
  }

  if (options.requireAdminOtp === true && user.role === "admin") {
    const email = String(user.email || "").trim();
    if (!email) {
      throw new Error("Tai khoan admin code chua co email de nhan ma xac thuc.");
    }
    if (typeof options.sendAdminOtp !== "function") {
      throw new Error("He thong chua duoc cau hinh de gui ma xac thuc qua email.");
    }

    const code = String(Math.floor(100000 + Math.random() * 900000));
    const challengeToken = randomBytes(24).toString("hex");
    loginChallenges.set(challengeToken, {
      user_id: user.id,
      code_hash: hashPassword(code),
      expires_at: Date.now() + (options.adminOtpTtlMs || 5 * 60_000),
    });

    await options.sendAdminOtp({
      user: normalizeUser(user),
      email,
      code,
      ttlMs: options.adminOtpTtlMs || 5 * 60_000,
    });

    return {
      otp_required: true,
      challengeToken,
      currentUser: sanitizeUser(user),
      delivery: `Ma xac thuc da duoc gui ve ${email}.`,
    };
  }

  const token = randomBytes(24).toString("hex");
  sessions.set(token, {
    user_id: user.id,
    created_at: Date.now(),
  });
  await saveSessions(usersFile);

  return {
    token,
    user: normalizeUser(user),
  };
}

export async function verifyLoginOtp(usersFile, challengeToken, code, options = {}) {
  const token = String(challengeToken || "").trim();
  const rawCode = String(code || "").trim();
  const challenge = loginChallenges.get(token);
  if (!challenge || !rawCode) {
    throw new Error("Ma xac thuc khong hop le.");
  }
  if (challenge.expires_at < Date.now()) {
    loginChallenges.delete(token);
    throw new Error("Ma xac thuc da het han.");
  }
  if (challenge.code_hash !== hashPassword(rawCode)) {
    throw new Error("Ma xac thuc khong dung.");
  }

  loginChallenges.delete(token);
  const users = await loadUsers(usersFile, options);
  const user = users.find((item) => item.id === challenge.user_id);
  if (!user) {
    throw new Error("Khong tim thay tai khoan dang nhap.");
  }

  const sessionToken = randomBytes(24).toString("hex");
  sessions.set(sessionToken, {
    user_id: user.id,
    created_at: Date.now(),
  });
  await saveSessions(usersFile);

  return {
    token: sessionToken,
    user: normalizeUser(user),
  };
}

export async function logoutSession(usersFile, request) {
  await loadSessions(usersFile);
  const token = extractBearerToken(request.headers.authorization);
  if (token) {
    sessions.delete(token);
    await saveSessions(usersFile);
  }
}

export function canManagePermissions(user) {
  return Boolean(getEffectivePolicy(user).can_manage_permissions);
}

export function getEffectivePolicy(user) {
  const rolePolicy = ROLE_POLICIES[user.role] || ROLE_POLICIES.employee;
  const override = user.policy_override || {};
  return {
    max_access_level: normalizeAccessLevel(
      override.max_access_level || rolePolicy.max_access_level,
    ),
    department_scope:
      override.department_scope === "all" || rolePolicy.department_scope === "all"
        ? "all"
        : "own",
    allowed_departments: dedupeStrings([
      ...rolePolicy.allowed_departments,
      ...(Array.isArray(override.allowed_departments) ? override.allowed_departments : []),
    ]),
    can_manage_permissions:
      override.can_manage_permissions === true || rolePolicy.can_manage_permissions === true,
  };
}

export function filterDocumentsForUser(documents, user) {
  return documents.filter((document) => canViewDocument(user, document.metadata || {}));
}

export function filterChunksForUser(chunks, user) {
  return chunks.filter((chunk) => canViewDocument(user, chunk.metadata || {}));
}

export function canViewDocument(user, metadata = {}) {
  const statusAccess = canViewDocumentByStatus(user, metadata);
  if (statusAccess !== null) {
    return statusAccess;
  }

  const privateAccess = canViewPrivateDocument(user, metadata);
  if (privateAccess !== null) {
    return privateAccess;
  }

  const policy = getEffectivePolicy(user);
  const docLevel = normalizeAccessLevel(metadata.access_level || "basic");
  if (ACCESS_LEVELS.indexOf(docLevel) > ACCESS_LEVELS.indexOf(policy.max_access_level)) {
    return false;
  }

  if (!canViewDocumentByDepartment(user, metadata, policy)) {
    return false;
  }

  return true;
}

function canViewDocumentByStatus(user, metadata = {}) {
  const status = String(metadata.status || "").trim().toLowerCase();
  if (!["draft", "rejected"].includes(status)) {
    return null;
  }

  const currentUserId = String(user?.id || "").trim();
  const ownerUserId = String(metadata.owner_user_id || "").trim();
  if (canManagePermissions(user)) {
    return true;
  }
  if (currentUserId && ownerUserId && currentUserId === ownerUserId) {
    return true;
  }
  return false;
}

function canViewPrivateDocument(user, metadata = {}) {
  const docLevel = normalizeAccessLevel(metadata.access_level || "basic");
  if (docLevel !== "sensitive") {
    return null;
  }

  const ownerUserId = String(metadata.owner_user_id || "").trim();
  const sharedUsers = getSharedUserIds(metadata);

  if (!ownerUserId && sharedUsers.length === 0) {
    return null;
  }

  const currentUserId = String(user?.id || "").trim();
  if (!currentUserId) {
    return false;
  }

  return currentUserId === ownerUserId || sharedUsers.includes(currentUserId);
}

function canViewDocumentByDepartment(user, metadata = {}, policy = getEffectivePolicy(user || {})) {
  if (policy.department_scope === "all") {
    return true;
  }

  const documentDepartments = getDocumentDepartments(metadata);
  if (documentDepartments.length === 0) {
    return true;
  }

  const viewerDepartments = dedupeStrings([
    normalizeDepartment(user?.department || ""),
    ...(Array.isArray(policy.allowed_departments) ? policy.allowed_departments : []),
  ]);

  if (viewerDepartments.length === 0) {
    return false;
  }

  return documentDepartments.some((department) => viewerDepartments.includes(department));
}

export async function updateUserAccess(usersFile, actor, payload, options = {}) {
  if (String(actor?.role || "").trim().toLowerCase() !== "admin") {
    throw new Error("Chi admin moi duoc phan quyen nguoi dung.");
  }

  const users = await loadUsers(usersFile, options);
  const targetId = String(payload.user_id || "").trim();
  const target = users.find((user) => user.id === targetId);
  if (!target) {
    throw new Error("Khong tim thay nguoi dung can cap quyen.");
  }

  const nextOverride = {
    max_access_level: normalizeAccessLevel(payload.max_access_level || target.policy_override?.max_access_level),
    department_scope:
      String(payload.department_scope || target.policy_override?.department_scope || "own") === "all"
        ? "all"
        : "own",
    allowed_departments: dedupeStrings(
      String(payload.allowed_departments || "")
        .split(",")
        .map((item) => normalizeDepartment(item))
        .filter(Boolean),
    ),
    can_manage_permissions: Boolean(payload.can_manage_permissions),
  };

  target.policy_override = cleanupOverride(nextOverride, target.role);
  await writeUsers(usersFile, users);
  return normalizeUser(target);
}

export async function updateUserProfile(usersFile, actor, payload, options = {}) {
  if (String(actor?.role || "").trim().toLowerCase() !== "admin") {
    throw new Error("Chi admin moi duoc sua thong tin nguoi dung.");
  }

  const users = await loadUsers(usersFile, options);
  const targetId = String(payload.user_id || "").trim();
  const target = users.find((user) => user.id === targetId);
  if (!target) {
    throw new Error("Khong tim thay nguoi dung can cap nhat.");
  }

  const username = String(payload.username || target.username || "").trim().toLowerCase();
  const name = String(payload.name || target.name || "").trim();
  const department = normalizeDepartment(payload.department || target.department || "general");
  const employeeCode = String(payload.employee_code || target.employee_code || username).trim();
  const title = String(payload.title || target.title || "").trim();
  const email = String(payload.email || target.email || "").trim().toLowerCase();

  if (!username || !name) {
    throw new Error("Can nhap tai khoan va ho ten.");
  }

  if (
    users.some(
      (user) =>
        String(user.id || "").trim() !== targetId &&
        String(user.username || "").trim().toLowerCase() === username,
    )
  ) {
    throw new Error("Ten dang nhap da ton tai.");
  }

  target.username = username;
  target.name = name;
  target.department = department;
  target.employee_code = employeeCode;
  target.title = title;
  target.email = email;

  await writeUsers(usersFile, users);
  return normalizeUser(target);
}

export async function updateUserPassword(usersFile, actor, payload, options = {}) {
  const users = await loadUsers(usersFile, options);
  const targetId = String(payload.user_id || actor.id || "").trim();
  const currentPassword = String(payload.current_password || "");
  const nextPassword = String(payload.new_password || "");

  if (!nextPassword || nextPassword.length < 6) {
    throw new Error("Mat khau moi phai co it nhat 6 ky tu.");
  }

  const target = users.find((user) => user.id === targetId);
  if (!target) {
    throw new Error("Khong tim thay nguoi dung can doi mat khau.");
  }

  const isSelf = target.id === actor.id;
  const isAdminReset = String(actor?.role || "").trim().toLowerCase() === "admin" && !isSelf;

  if (!isSelf && !isAdminReset) {
    throw new Error("Ban khong co quyen doi mat khau cua nguoi dung nay.");
  }

  if (!isAdminReset && target.password_hash !== hashPassword(currentPassword)) {
    throw new Error("Mat khau hien tai khong dung.");
  }

  target.password_hash = hashPassword(nextPassword);
  await writeUsers(usersFile, users);
  return normalizeUser(target);
}

export async function createUser(usersFile, actor, payload, options = {}) {
  if (String(actor?.role || "").trim().toLowerCase() !== "admin") {
    throw new Error("Chi admin moi duoc tao nguoi dung.");
  }

  const users = await loadUsers(usersFile, options);
  const username = String(payload.username || "").trim().toLowerCase();
  const password = String(payload.password || "");
  const name = String(payload.name || "").trim();
  const role = normalizeRole(payload.role || "employee");
  const department = normalizeDepartment(payload.department || "general");
  const employeeCode = String(payload.employee_code || username).trim();
  const email = String(payload.email || "").trim().toLowerCase();
  const title = String(payload.title || "").trim();

  if (!username || !name) {
    throw new Error("Can nhap ten dang nhap va ho ten.");
  }
  if (!password || password.length < 6) {
    throw new Error("Mat khau phai co it nhat 6 ky tu.");
  }
  if (users.some((user) => String(user.username || "").trim().toLowerCase() === username)) {
    throw new Error("Ten dang nhap da ton tai.");
  }

  const nextUser = normalizeUser({
    id: `local-${employeeCode || username}`,
    name,
    username,
    role,
    department,
    password_hash: hashPassword(password),
    email,
    managed_by: "local",
    employee_code: employeeCode,
    title,
    source_permission: "",
    policy_override: {},
  });

  users.push(nextUser);
  await writeUsers(usersFile, dedupeUsers(users));
  return nextUser;
}

export async function deleteUser(usersFile, actor, payload, options = {}) {
  if (String(actor?.role || "").trim().toLowerCase() !== "admin") {
    throw new Error("Chi admin moi duoc xoa nguoi dung.");
  }

  const users = await loadUsers(usersFile, options);
  const targetId = String(payload.user_id || "").trim();
  if (!targetId) {
    throw new Error("Khong tim thay nguoi dung can xoa.");
  }
  if (String(actor?.id || "").trim() === targetId) {
    throw new Error("Khong the tu xoa tai khoan dang dang nhap.");
  }

  const target = users.find((user) => String(user.id || "").trim() === targetId);
  if (!target) {
    throw new Error("Khong tim thay nguoi dung can xoa.");
  }

  const adminCount = users.filter((user) => String(user.role || "").trim().toLowerCase() === "admin").length;
  if (String(target.role || "").trim().toLowerCase() === "admin" && adminCount <= 1) {
    throw new Error("Khong the xoa admin cuoi cung.");
  }

  const nextUsers = users.filter((user) => String(user.id || "").trim() !== targetId);
  await writeUsers(usersFile, nextUsers);
  return normalizeUser(target);
}

export function requireAuthenticatedUser(currentUser) {
  if (!currentUser) {
    throw new Error("Ban can dang nhap de su dung he thong.");
  }
}

function cleanupOverride(override, role) {
  const rolePolicy = ROLE_POLICIES[role] || ROLE_POLICIES.employee;
  const normalized = {
    max_access_level: normalizeAccessLevel(override.max_access_level),
    department_scope: override.department_scope === "all" ? "all" : "own",
    allowed_departments: dedupeStrings(override.allowed_departments || []),
    can_manage_permissions: Boolean(override.can_manage_permissions),
  };

  const result = {};
  if (normalized.max_access_level !== rolePolicy.max_access_level) {
    result.max_access_level = normalized.max_access_level;
  }
  if (normalized.department_scope !== rolePolicy.department_scope) {
    result.department_scope = normalized.department_scope;
  }
  if (normalized.allowed_departments.length > 0) {
    result.allowed_departments = normalized.allowed_departments;
  }
  if (normalized.can_manage_permissions !== rolePolicy.can_manage_permissions) {
    result.can_manage_permissions = normalized.can_manage_permissions;
  }
  return result;
}

async function writeUsers(usersFile, users) {
  await writeFile(usersFile, `${JSON.stringify({ users }, null, 2)}\n`, "utf8");
}

async function maybeSyncUsersFromSheet(usersFile, options = {}) {
  const sheetUrl = String(options.usersSheetUrl || options.sheetUrl || "").trim();
  if (!sheetUrl) {
    return;
  }

  const syncState = userSheetSyncState.get(usersFile) || {
    inFlight: null,
    lastStartedAt: 0,
    lastFinishedAt: 0,
    lastSheetUrl: "",
  };
  const intervalMs = Math.max(0, Number(options.usersSyncIntervalMs || options.syncIntervalMs || 120_000) || 0);
  const now = Date.now();
  const shouldStart =
    syncState.lastSheetUrl !== sheetUrl ||
    !syncState.lastStartedAt ||
    intervalMs === 0 ||
    now - syncState.lastStartedAt >= intervalMs;

  if (!shouldStart) {
    if (syncState.inFlight) {
      await syncState.inFlight;
    }
    return;
  }

  syncState.lastStartedAt = now;
  syncState.lastSheetUrl = sheetUrl;
  syncState.inFlight = (async () => {
    try {
      await syncUsersFromSheetWithOptions(usersFile, {
        ...options,
        sheetUrl,
      });
      syncState.lastFinishedAt = Date.now();
    } catch (error) {
      syncState.lastFinishedAt = Date.now();
      if (options.throwOnUserSyncError) {
        throw error;
      }
    } finally {
      syncState.inFlight = null;
    }
  })();

  userSheetSyncState.set(usersFile, syncState);
  await syncState.inFlight;
}

async function syncUsersFromSheetWithOptions(usersFile, options = {}) {
  const csv = await fetchSheetCsv(options.sheetUrl, options.fetchImpl);
  const rows = parseCsv(csv);
  if (rows.length < 2) {
    return;
  }

  const headers = rows[0].map(normalizeHeader);
  const existing = JSON.parse(await readFile(usersFile, "utf8"));
  const existingUsers = Array.isArray(existing.users) ? existing.users : [];
  const nextUsers = [];

  for (const row of rows.slice(1)) {
    const record = Object.fromEntries(
      headers.map((header, index) => [header, String(row[index] || "").trim()]),
    );
    const syncedUser = mapSheetRowToUser(record, existingUsers);
    if (!syncedUser) {
      continue;
    }
    nextUsers.push(syncedUser);
  }

  const sheetUsernames = new Set(
    nextUsers.map((user) => String(user.username || "").trim().toLowerCase()).filter(Boolean),
  );
  const preservedUsers = existingUsers.filter((user) => {
    const username = String(user.username || "").trim().toLowerCase();
    return (
      String(user.managed_by || "").toLowerCase() !== "sheet" &&
      !sheetUsernames.has(username)
    );
  });

  const deduped = dedupeUsers([...preservedUsers, ...nextUsers]);
  await writeUsers(usersFile, deduped);
}

async function fetchSheetCsv(sheetUrl, fetchImpl = fetch) {
  const downloadUrl = resolveSheetCsvUrl(sheetUrl);
  const response = await fetchImpl(downloadUrl, {
    cache: "no-store",
    headers: {
      "cache-control": "no-cache, no-store, max-age=0",
      pragma: "no-cache",
    },
  });

  if (!response.ok) {
    throw new Error(`Khong the dong bo user tu sheet (${response.status}).`);
  }

  return response.text();
}

function resolveSheetCsvUrl(value) {
  const url = new URL(String(value || "").trim());
  if (url.hostname === "docs.google.com" && url.pathname.includes("/spreadsheets/d/")) {
    const match = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    const gidFromHash = url.hash.match(/gid=(\d+)/i)?.[1] || "";
    const gid = url.searchParams.get("gid") || gidFromHash || "0";
    if (match?.[1]) {
      return `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv&gid=${gid}`;
    }
  }

  return value;
}

function parseCsv(raw) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    const next = raw[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(field);
      field = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(field);
      field = "";
      if (row.some((item) => String(item || "").trim())) {
        rows.push(row);
      }
      row = [];
      continue;
    }

    field += char;
  }

  row.push(field);
  if (row.some((item) => String(item || "").trim())) {
    rows.push(row);
  }

  return rows;
}

function normalizeHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function mapSheetRowToUser(record, existingUsers = []) {
  const username = String(record.tai_khoan || record.username || "").trim().toLowerCase();
  const password = String(record.mat_khau || record.password || "").trim();
  const name = String(record.ho_ten || record.name || username).trim();

  if (!username) {
    return null;
  }

  const existing = existingUsers.find(
    (user) => String(user.username || "").trim().toLowerCase() === username,
  );
  const employeeCode = String(record.ma_nv || record.employee_code || username)
    .trim()
    .toLowerCase();
  const passwordHash = password
    ? hashPassword(password)
    : existing?.password_hash || hashPassword("123456");

  return {
    id: existing?.id || `sheet-${employeeCode || username}`,
    name,
    username,
    role: mapSheetRole(record),
    department: normalizeDepartment(record.bo_phan || record.department || "general"),
    password_hash: passwordHash,
    email: String(record.mail || record.email || existing?.email || "").trim().toLowerCase(),
    policy_override: {
      ...(existing?.policy_override || {}),
      allowed_departments: Array.isArray(existing?.policy_override?.allowed_departments)
        ? dedupeStrings(existing.policy_override.allowed_departments)
        : [],
    },
    managed_by: "sheet",
    employee_code: String(record.ma_nv || "").trim(),
    title: String(record.chuc_vu || "").trim(),
    source_permission: String(record.quyen_han || "").trim(),
  };
}

function mapSheetRole(record) {
  const permission = String(record.quyen_han || "").trim().toLowerCase();
  const title = String(record.chuc_vu || "").trim().toLowerCase();

  if (title.includes("admin code") || title.includes("admin")) {
    return "admin";
  }
  if (permission === "gđ" || permission === "gd") {
    return "director";
  }
  if (permission === "ql") {
    return "manager";
  }
  if (permission === "kt") {
    return "manager";
  }
  return "employee";
}

function dedupeUsers(users) {
  const seen = new Set();
  const result = [];

  for (const user of users) {
    const key = String(user.username || "").trim().toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(user);
  }

  return result;
}

function normalizeUser(user) {
  return {
    id: String(user.id || "").trim(),
    name: String(user.name || "").trim(),
    username: String(user.username || "").trim().toLowerCase(),
    role: normalizeRole(user.role),
    department: normalizeDepartment(user.department || "general"),
    password_hash: String(user.password_hash || "").trim(),
    email: String(user.email || "").trim().toLowerCase(),
    managed_by: String(user.managed_by || "").trim().toLowerCase(),
    employee_code: String(user.employee_code || "").trim(),
    title: String(user.title || "").trim(),
    source_permission: String(user.source_permission || "").trim(),
    policy_override: {
      ...(user.policy_override || {}),
      allowed_departments: Array.isArray(user.policy_override?.allowed_departments)
        ? dedupeStrings(user.policy_override.allowed_departments)
        : [],
    },
  };
}

function mergeWithDefaultUser(user) {
  const fallback = DEFAULT_USERS.find((item) => item.id === user.id);
  if (!fallback) {
    return user;
  }

  return {
    ...fallback,
    ...user,
    username: user.username || fallback.username,
    password_hash: user.password_hash || fallback.password_hash,
    email: user.email || fallback.email || "",
    policy_override: {
      ...(fallback.policy_override || {}),
      ...(user.policy_override || {}),
    },
  };
}

function normalizeRole(role) {
  const normalized = String(role || "").trim().toLowerCase();
  return ROLE_POLICIES[normalized] ? normalized : "employee";
}

function normalizeAccessLevel(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ACCESS_LEVELS.includes(normalized) ? normalized : "basic";
}

function normalizeDepartment(value) {
  const normalized = String(value || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  const aliases = {
    "kinh-doanh": "sales",
    sales: "sales",
    "nhan-su": "hr",
    hr: "hr",
    "tai-chinh": "finance",
    "ke-toan": "finance",
    finance: "finance",
    "van-hanh": "operations",
    "van-chuyen": "operations",
    operations: "operations",
    "san-xuat": "production",
    production: "production",
    admin: "executive",
    "ban-giam-doc": "executive",
    executive: "executive",
    unknown: "",
  };

  return aliases[normalized] ?? normalized;
}

function getDocumentDepartments(metadata = {}) {
  const values = [];
  const allowed = String(metadata.allowed_departments || "").trim();
  const ownerDepartment = String(metadata.owner_department || metadata.owner || "").trim();
  if (allowed) {
    values.push(
      ...allowed
        .split(",")
        .map((item) => normalizeDepartment(item))
        .filter(Boolean),
    );
  } else if (ownerDepartment && ownerDepartment.toLowerCase() !== "unknown") {
    values.push(normalizeDepartment(ownerDepartment));
  }
  return dedupeStrings(values);
}

function getSharedUserIds(metadata = {}) {
  return dedupeStrings(
    String(metadata.shared_with_users || "")
      .split(",")
      .map((item) => String(item || "").trim()),
  );
}

function dedupeStrings(values) {
  return [...new Set(values.map((item) => String(item || "").trim()).filter(Boolean))];
}

function hashPassword(value) {
  return createHash("sha256").update(String(value || "")).digest("hex");
}

function extractBearerToken(authorization) {
  const value = String(authorization || "");
  const match = value.match(/^Bearer\s+(.+)$/i);
  return match?.[1]?.trim() || "";
}
