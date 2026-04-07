import { createServer } from "node:http";
import { access, mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import XLSX from "xlsx";

import { loadConfig } from "./lib/config.mjs";
import { appendAuditLog, readAuditLogs } from "./lib/auditLog.mjs";
import {
  canViewDocument,
  canManagePermissions,
  filterChunksForUser,
  filterDocumentsForUser,
  getEffectivePolicy,
  loginWithPassword,
  logoutSession,
  requireAuthenticatedUser,
  resolveCurrentUser,
  resolveCurrentUserFromToken,
  sanitizeUser,
  verifyLoginOtp,
  createUser,
  deleteUser,
  updateUserProfile,
  updateUserPassword,
  updateUserAccess,
} from "./lib/accessControl.mjs";
import { sendMail } from "./lib/mailer.mjs";
import {
  buildInternalFallbackAnswer,
  buildNoSearchAnswer,
  buildSuggestions,
  buildWebSearchErrorAnswer,
  buildWorkflow,
  pickSnippet,
} from "./lib/fallback.mjs";
import {
  createStorageFolder,
  deleteStorageFolder,
  deleteInternalDocument,
  deleteSyncSource,
  importSheetDocument,
  importUploadedDocument,
  listCustomStorageFolders,
  listSyncSources,
  listInternalDocuments,
  loadInternalRepository,
  readInternalDocument,
  reviewInternalDocument,
  reviewSyncSource,
  renameStorageFolder,
  saveSyncSource,
  saveInternalDocument,
  setSourceSyncable,
  syncSourceById,
  syncRegisteredSources,
  updateInternalDocumentMetadata,
  validateSheetDocument,
} from "./lib/internalRepo.mjs";
import {
  answerWithInternalContext,
  answerWithWebSearch,
  extractTextFromImage,
  isGeminiEnabled,
} from "./lib/geminiClient.mjs";
import {
  detectConflictingMatches,
  hasConfidentInternalMatch,
  isConversationalQuery,
  searchDocuments,
  selectTopDocumentMatches,
} from "./lib/retrieval.mjs";
import { createOrderStore } from "./lib/orderStore.mjs";

const config = loadConfig(process.cwd());
const orderStore = await createOrderStore({
  dbFilePath: path.resolve(config.cwd, "./data/orders/orders.sqlite"),
  legacyOrdersFile: path.resolve(config.cwd, "./data/orders/orders.json"),
  legacyCatalogFile: path.resolve(config.cwd, "./data/orders/catalog.json"),
  legacyCompletionsFile: path.resolve(config.cwd, "./data/delivery/completions.json"),
  legacyNotificationsFile: path.resolve(config.cwd, "./data/notifications/inbox.json"),
  defaultProducts: defaultOrderProducts,
});
const server = createServer(handleRequest);
const notificationStreams = new Map();
const TAPE_CALCULATOR_SHEET_URL = "https://docs.google.com/spreadsheets/d/1AkD44QMnjm0OsN39LcrGwIX2TdRrvKXChyWXcTaLprs/export?format=xlsx";
const TAPE_CALCULATOR_PRODUCT_CODES = [
  "BOPP45",
  "BOPP48",
  "BOPP50",
  "BOPP57",
  "BOPP60",
  "HDV51",
  "IN1T50",
  "IN2T50",
  "IN1TS50",
  "IN2TS50",
  "IN1M50",
  "IN2M50",
  "INPTS50",
  "XX80",
  "XV2",
  "2M38",
  "GT",
  "SML",
  "VXD260",
  "MV50",
  "MD50",
  "CT09",
  "CT11",
  "CT13",
  "MXD50",
  "PDN17",
  "PDN18",
  "PDN33",
  "PTEFMN13",
  "PTEFMN17",
  "PPETPI",
  "PPETSMX38",
  "PPETSMX25",
  "PCTD",
  "PSTT1",
  "PSTT2",
  "P2MKV1",
  "P2MKV2",
  "PVXD10",
  "PVXD15",
  "PVXD20",
  "IN1T55",
];
const DEFAULT_TAPE_CALCULATOR_PRODUCTS = [
  { code: "BOPP45", jumbo_height: 1260 },
  { code: "BOPP48", jumbo_height: 1260 },
  { code: "BOPP50", jumbo_height: 1260 },
  { code: "BOPP57", jumbo_height: 1260 },
  { code: "BOPP60", jumbo_height: 1260 },
  { code: "HDV51", jumbo_height: 1260 },
  { code: "IN1T50", jumbo_height: 980 },
  { code: "IN2T50", jumbo_height: 980 },
  { code: "IN1TS50", jumbo_height: 980 },
  { code: "IN2TS50", jumbo_height: 980 },
];
const DEFAULT_TAPE_CALCULATOR_CORES = [
  { code: "G10", width_mm: 10 },
  { code: "G12", width_mm: 12 },
  { code: "G15", width_mm: 15 },
  { code: "G18", width_mm: 18 },
  { code: "G20", width_mm: 20 },
  { code: "G25", width_mm: 25 },
  { code: "G46", width_mm: 46 },
  { code: "G70", width_mm: 70 },
  { code: "N36", width_mm: 36 },
];
const DEFAULT_TAPE_CALCULATOR_FORM = {
  tape_type: "BOPP60",
  order_quantity: 212,
  core_type: "G15",
  packaging: 20,
  finished_quantity: 4240,
  jumbo_height: 1260,
  core_width: 15,
};
let tapeCalculatorConfigCache = { expiresAt: 0, payload: null };
const TAPE_CALCULATOR_FETCH_TIMEOUT_MS = 20_000;
const TAPE_CALCULATOR_FETCH_RETRY_COUNT = 3;
const userSyncOptions = {
  usersSheetUrl: config.usersSheetUrl,
  usersSyncIntervalMs: config.usersSyncIntervalMs,
};

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
};

const sheetSyncState = {
  inFlight: null,
  lastStartedAt: 0,
  lastFinishedAt: "",
  lastResult: {
    enabled: config.sheetSyncIntervalMs > 0,
    intervalMs: config.sheetSyncIntervalMs,
    checked: 0,
    updated: 0,
    unchanged: 0,
    errors: [],
  },
};

const STORAGE_FOLDER_OPTIONS = {
  internal: {
    level: "basic",
    storage_folder: "internal",
    owner: "internal",
    owner_department: "",
    allowed_departments: "",
    isPrivate: false,
  },
  "van-chuyen": {
    level: "basic",
    storage_folder: "Vận Chuyển",
    owner: "operations",
    owner_department: "operations",
    allowed_departments: "operations",
    isPrivate: false,
  },
  "ke-toan": {
    level: "advanced",
    storage_folder: "Kế Toán",
    owner: "finance",
    owner_department: "finance",
    allowed_departments: "finance",
    isPrivate: false,
  },
  "phong-kinh-doanh": {
    level: "advanced",
    storage_folder: "phòng kinh doanh",
    owner: "sales",
    owner_department: "sales",
    allowed_departments: "sales",
    isPrivate: false,
  },
  "phong-san-xuat": {
    level: "advanced",
    storage_folder: "phòng Sản Xuất",
    owner: "production",
    owner_department: "production",
    allowed_departments: "production",
    isPrivate: false,
  },
  "du-lieu-ca-nhan": {
    level: "sensitive",
    storage_folder: "dữ liệu cá nhân",
    owner: "private",
    owner_department: "",
    allowed_departments: "",
    isPrivate: true,
  },
};

server.listen(config.port, config.host, () => {
  console.log(`Server dang lang nghe tai http://${config.host}:${config.port}`);
});

async function handleRequest(request, response) {
  try {
    const requestUrl = new URL(request.url, `http://${request.headers.host}`);
    if (request.method === "GET" && requestUrl.pathname === "/api/notifications/stream") {
      const token = String(requestUrl.searchParams.get("token") || "").trim();
      const auth = await resolveCurrentUserFromToken(config.usersFile, token, userSyncOptions);
      if (!auth.currentUser) {
        return sendJson(response, 401, { error: "Unauthorized" });
      }
      return openNotificationStream(response, auth.currentUser);
    }
    const { currentUser, users } = await resolveCurrentUser(
      config.usersFile,
      request,
      userSyncOptions,
    );

    if (request.method === "POST" && requestUrl.pathname === "/api/login") {
      const payload = await readJsonBody(request);
      const session = await loginWithPassword(
        config.usersFile,
        payload.username,
        payload.password,
        {
          ...userSyncOptions,
          requireAdminOtp: true,
          adminOtpTtlMs: config.auth.adminOtpTtlMs,
          sendAdminOtp: ({ email, code, user, ttlMs }) =>
            sendMail(config.mail, {
              to: email,
              subject: "Ma xac thuc dang nhap Admin Code",
              text: [
                `Xin chao ${user.name || user.username},`,
                "",
                `Ma xac thuc dang nhap cua ban la: ${code}`,
                `Ma co hieu luc trong ${Math.max(1, Math.round(ttlMs / 60000))} phut.`,
                "",
                "Neu ban khong yeu cau dang nhap, hay bo qua email nay.",
              ].join("\n"),
            }),
        },
      );
      if (session.otp_required) {
        return sendJson(response, 200, session);
      }
      return sendJson(response, 200, {
        ok: true,
        token: session.token,
        currentUser: sanitizeUser(session.user),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/login/verify") {
      const payload = await readJsonBody(request);
      const session = await verifyLoginOtp(
        config.usersFile,
        payload.challengeToken,
        payload.code,
        userSyncOptions,
      );
      return sendJson(response, 200, {
        ok: true,
        token: session.token,
        currentUser: sanitizeUser(session.user),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/logout") {
      await logoutSession(config.usersFile, request);
      return sendJson(response, 200, { ok: true });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/session") {
      return sendJson(response, 200, {
        authenticated: Boolean(currentUser),
        currentUser: currentUser ? sanitizeUser(currentUser) : null,
      });
    }

    if (request.method === "GET" && !requestUrl.pathname.startsWith("/api/")) {
      return serveStaticFile(requestUrl.pathname, response);
    }

    requireAuthenticatedUser(currentUser);

    if (request.method === "GET" && requestUrl.pathname === "/api/health") {
      const syncInfo = await maybeSyncSheetSources();
      const repo = await loadInternalRepository(config.internalDocsDir, config.cwd);
      const allSources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const visibleSources = allSources.filter((source) => canViewDocument(currentUser, source));
      const visibleDocuments = filterDocumentsForUser(repo.documents, currentUser);
      const visibleChunks = filterChunksForUser(repo.chunks, currentUser);
      return sendJson(response, 200, {
        status: "ok",
        aiEnabled: isGeminiEnabled(config),
        model: config.gemini.model,
        documentCount: visibleDocuments.length,
        chunkCount: visibleChunks.length,
        sourceCount: visibleSources.length,
        currentUser: sanitizeUser(currentUser),
        canManagePermissions: canManagePermissions(currentUser),
        sheetSync: syncInfo,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/users") {
      return sendJson(response, 200, {
        currentUser: sanitizeUser(currentUser),
        users: users.map(sanitizeUser),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/users/create") {
      const payload = await readJsonBody(request);
      await createUser(config.usersFile, currentUser, payload, userSyncOptions);
      const nextUsers = await resolveCurrentUser(config.usersFile, request, userSyncOptions);
      return sendJson(response, 201, {
        ok: true,
        currentUser: nextUsers.currentUser ? sanitizeUser(nextUsers.currentUser) : null,
        users: nextUsers.users.map(sanitizeUser),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/users/delete") {
      const payload = await readJsonBody(request);
      await deleteUser(config.usersFile, currentUser, payload, userSyncOptions);
      const nextUsers = await resolveCurrentUser(config.usersFile, request, userSyncOptions);
      return sendJson(response, 200, {
        ok: true,
        currentUser: nextUsers.currentUser ? sanitizeUser(nextUsers.currentUser) : null,
        users: nextUsers.users.map(sanitizeUser),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/users/profile") {
      const payload = await readJsonBody(request);
      const updatedUser = await updateUserProfile(
        config.usersFile,
        currentUser,
        payload,
        userSyncOptions,
      );
      const nextUsers = await resolveCurrentUser(config.usersFile, request, userSyncOptions);
      return sendJson(response, 200, {
        ok: true,
        updatedUser: sanitizeUser(updatedUser),
        currentUser: nextUsers.currentUser ? sanitizeUser(nextUsers.currentUser) : null,
        users: nextUsers.users.map(sanitizeUser),
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/notifications") {
      const notifications = await readNotifications(resolveNotificationsFile(config.cwd), currentUser);
      return sendJson(response, 200, {
        notifications,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/orders") {
      if (!canViewOrders(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen xem don hang." });
      }
      const orders = await readOrders(resolveOrdersFile(config.cwd));
      return sendJson(response, 200, {
        orders,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/sales-inventory") {
      if (!canViewOrders(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen xem kho kinh doanh." });
      }
      const receipts = await readSalesInventoryReceipts(resolveSalesInventoryFile(config.cwd));
      const transfers = await readSalesInventoryTransfers(resolveSalesInventoryTransfersFile(config.cwd));
      return sendJson(response, 200, {
        receipts,
        transfers,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sales-inventory/receive") {
      if (!canCreateOrder(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen nhap kho kinh doanh." });
      }

      const payload = await readJsonBody(request);
      const name = typeof payload.name === "string" ? payload.name.trim() : "";
      const code = typeof payload.code === "string" ? payload.code.trim() : "";
      const unit = typeof payload.unit === "string" ? payload.unit.trim() : "";
      const supplier = typeof payload.supplier === "string" ? payload.supplier.trim() : "";
      const note = typeof payload.note === "string" ? payload.note.trim() : "";
      const quantity = normalizeNonNegativeNumber(payload.quantity);
      const receivedAt = normalizeOptionalDateTime(payload.received_at) || formatNow();

      if (!name || !unit || quantity <= 0) {
        return sendJson(response, 400, { error: "Can nhap ten hang, don vi tinh va so luong hop le." });
      }

      const receipt = buildSalesInventoryReceiptRecord({
        currentUser,
        code,
        name,
        unit,
        quantity,
        supplier,
        note,
        receivedAt,
      });
      const receipts = await appendSalesInventoryReceipt(resolveSalesInventoryFile(config.cwd), receipt);

      await logAudit(config.auditLogFile, currentUser, "sales-inventory.receive", {
        receipt_id: receipt.id,
        code,
        name,
        unit,
        quantity: String(quantity),
        supplier,
      });

      return sendJson(response, 201, {
        ok: true,
        receipt,
        receipts,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sales-inventory/transfer-from-production") {
      if (!canCreateOrder(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen chuyen hang tu kho san xuat." });
      }

      const payload = await readJsonBody(request);
      const transferItems = normalizeSalesInventoryTransferItems(payload.items);
      const transferredAt = normalizeOptionalDateTime(payload.transferred_at) || formatNow();
      const note = typeof payload.note === "string" ? payload.note.trim() : "";

      if (!transferItems.length) {
        return sendJson(response, 400, { error: "Can chon it nhat mot mat hang de chuyen kho." });
      }

      const orders = await readOrders(resolveOrdersFile(config.cwd));
      const existingTransfers = await readSalesInventoryTransfers(resolveSalesInventoryTransfersFile(config.cwd));
      const availabilityMap = buildProductionAvailabilityMap(orders, existingTransfers);

      for (const item of transferItems) {
        const itemKey = buildInventoryItemKey(item);
        const available = availabilityMap.get(itemKey) || { quantity: 0 };
        if (item.quantity <= 0) {
          return sendJson(response, 400, { error: `So luong chuyen cua mat hang "${item.name || item.code}" phai lon hon 0.` });
        }
        if (item.quantity > available.quantity) {
          return sendJson(response, 400, { error: `Mat hang "${item.name || item.code}" chi con ${formatInventoryNumber(available.quantity)} trong kho san xuat.` });
        }
      }

      const transferBatch = buildSalesInventoryTransferBatch({
        currentUser,
        transferredAt,
        note,
        items: transferItems,
      });
      const transfers = await appendSalesInventoryTransfers(resolveSalesInventoryTransfersFile(config.cwd), transferBatch.items);
      const receipts = await readSalesInventoryReceipts(resolveSalesInventoryFile(config.cwd));

      await logAudit(config.auditLogFile, currentUser, "sales-inventory.transfer", {
        transfer_id: transferBatch.transfer_id,
        item_count: String(transferBatch.items.length),
      });

      return sendJson(response, 201, {
        ok: true,
        transfer_id: transferBatch.transfer_id,
        transfers,
        receipts,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/tape-calculator/config") {
      if (!currentUser) {
        return sendJson(response, 401, { error: "Ban can dang nhap de dung bang tinh." });
      }
      const configPayload = await loadTapeCalculatorConfig();
      return sendJson(response, 200, configPayload);
    }

    if (requestUrl.pathname === "/api/order-products") {
      if (request.method === "GET") {
        if (!currentUser) {
          return sendJson(response, 401, { error: "Ban can dang nhap de xem danh muc hang hoa." });
        }
        const products = await readOrderProducts(resolveOrderProductsFile(config.cwd));
        return sendJson(response, 200, { products });
      }

      if (request.method === "POST") {
        if (!canManageOrders(currentUser)) {
          return sendJson(response, 403, { error: "Chi admin moi duoc quan ly hang hoa." });
        }

        const payload = await readJsonBody(request);
        const products = await saveOrderProducts(resolveOrderProductsFile(config.cwd), payload.products);
        await logAudit(config.auditLogFile, currentUser, "order-products.save", {
          count: products.length,
        });

        return sendJson(response, 200, {
          ok: true,
          products,
        });
      }
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/update") {
      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      const progressUpdates = normalizeProductionProgressUpdates(payload.progress_updates);
      const markCompleted = Boolean(payload.mark_completed);
      const customerName = typeof payload.customer_name === "string" ? payload.customer_name.trim() : "";
      const salesUserId = typeof payload.sales_user_id === "string" ? payload.sales_user_id.trim() : "";
      const deliveryUserId = typeof payload.delivery_user_id === "string" ? payload.delivery_user_id.trim() : "";
      const plannedDeliveryAt = normalizeOptionalDateTime(payload.planned_delivery_at);
      const deliveryAddress = typeof payload.delivery_address === "string" ? payload.delivery_address.trim() : "";
      const orderItems = typeof payload.order_items === "string" ? payload.order_items.trim() : "";
      const orderValue = normalizeOrderValue(payload.order_value);
      const paymentStatus = normalizePaymentStatus(payload.payment_status);
      const paymentMethod = normalizePaymentMethod(payload.payment_method, paymentStatus);
      const note = typeof payload.note === "string" ? payload.note.trim() : "";
      const salesUser = users.find((user) => user.id === salesUserId);
      const deliveryUser = users.find((user) => user.id === deliveryUserId);

      if (!orderId) {
        return sendJson(response, 400, { error: "Can nhap ma don hang." });
      }

      const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
      const existingOrder = existingOrders.find((item) => String(item?.order_id || "").trim() === orderId);

      if (!existingOrder) {
        return sendJson(response, 404, { error: "Khong tim thay don hang." });
      }
      const isProductionProgressOnly = isProductionOrderRecord(existingOrder) && progressUpdates.length > 0 && !customerName;
      if (isProductionProgressOnly && !canUpdateProductionProgressForLines(currentUser, existingOrder, progressUpdates)) {
        return sendJson(response, 403, { error: "Ban chi duoc cap nhat dong san xuat minh da nhan." });
      }
      if (!isProductionProgressOnly && !canEditExistingOrder(currentUser, existingOrder)) {
        return sendJson(response, 403, { error: "Chi admin hoac nguoi tao don moi duoc sua don hang." });
      }
      if (!isProductionProgressOnly && !customerName) {
        return sendJson(response, 400, { error: "Can nhap ma don hang va ten khach hang." });
      }

      if (isProductionProgressOnly) {
        const orders = await orderStore.updateOrder(orderId, (item) => {
          if (orderItems) {
            return {
              ...item,
              order_items: orderItems,
              production_completed_by_json: markCompleted ? appendProductionCompletionRecord(item, currentUser) : item?.production_completed_by_json || "[]",
              production_packaged_by_json: item?.production_packaged_by_json || "",
              updated_at: formatNow(),
            };
          }
          const summaryItems = parseProductionOrderItemsSummary(item?.order_items || "");
          progressUpdates.forEach((update) => {
            const targetIndex = Math.trunc(update.line_number) - 1;
            const currentLine = summaryItems[targetIndex];
            if (!currentLine) {
              return;
            }
            const doneValue = String(update.done || "").trim() || "0";
            const quantityValue = Number(String(currentLine.quantity || "").replace(",", "."));
            const doneNumber = Math.max(0, Number(String(doneValue || "0").replace(",", ".")));
            const missingNumber = Number.isFinite(quantityValue) ? Math.max(0, quantityValue - doneNumber) : 0;
            const extraNumber = Number.isFinite(quantityValue) ? Math.max(0, doneNumber - quantityValue) : 0;
            currentLine.done = doneValue;
            currentLine.team = String(update.team || "").trim() || "-";
            currentLine.missing = String(Number.isFinite(missingNumber) ? missingNumber : currentLine.missing || "0");
            currentLine.extra = String(Number.isFinite(extraNumber) ? extraNumber : currentLine.extra || "0");
          });
          return {
            ...item,
            order_items: buildProductionOrderItemsSummary(summaryItems),
            production_completed_by_json: markCompleted ? appendProductionCompletionRecord(item, currentUser) : item?.production_completed_by_json || "[]",
            production_packaged_by_json: item?.production_packaged_by_json || "",
            updated_at: formatNow(),
          };
        });

        await logAudit(config.auditLogFile, currentUser, "production.progress.update", {
          order_id: orderId,
          updated_lines: progressUpdates.map((item) => String(item.line_number)).join(","),
          via: "order.update",
        });

        return sendJson(response, 200, {
          ok: true,
          orders,
        });
      }

      const orders = await updateOrderRecord(resolveOrdersFile(config.cwd), {
        orderId,
        customerName,
        salesUserId,
        salesUserName: salesUser?.name || "",
        deliveryUserId,
        deliveryUserName: deliveryUser?.name || "",
        plannedDeliveryAt,
        deliveryAddress,
        orderItems,
        orderValue,
        paymentStatus,
        paymentMethod,
        note,
      });

      await logAudit(config.auditLogFile, currentUser, "order.update", {
        order_id: orderId,
        payment_status: paymentStatus,
        payment_method: paymentMethod,
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/order-products/save") {
      if (!canManageOrders(currentUser)) {
        return sendJson(response, 403, { error: "Chi admin moi duoc quan ly hang hoa." });
      }

      const payload = await readJsonBody(request);
      const products = await saveOrderProducts(resolveOrderProductsFile(config.cwd), payload.products);
      await logAudit(config.auditLogFile, currentUser, "order-products.save", {
        count: products.length,
      });

      return sendJson(response, 200, {
        ok: true,
        products,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/delete") {
      if (!canManageOrders(currentUser)) {
        return sendJson(response, 403, { error: "Chi admin moi duoc xoa don hang." });
      }

      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      if (!orderId) {
        return sendJson(response, 400, { error: "Can nhap ma don hang." });
      }

      const orders = await deleteOrderRecord(resolveOrdersFile(config.cwd), orderId);
      await logAudit(config.auditLogFile, currentUser, "order.delete", {
        order_id: orderId,
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/production-claim") {
      if (!canClaimProductionOrder(currentUser)) {
        return sendJson(response, 403, { error: "Bo phan kinh doanh khong duoc nhan phieu san xuat." });
      }
      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      const mode = String(payload.mode || "").trim().toLowerCase() === "partial" ? "partial" : "all";
      const claimedItems = normalizeProductionClaimItems(payload.claimed_items);
      if (!orderId) {
        return sendJson(response, 400, { error: "Thieu ma phieu san xuat." });
      }
      if (mode === "partial" && !claimedItems.length) {
        return sendJson(response, 400, { error: "Can chon it nhat mot dong hang hoa de nhan san xuat." });
      }

      const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
      const targetOrder = existingOrders.find((item) => String(item?.order_id || "").trim() === orderId);
      if (!targetOrder || !isProductionOrderRecord(targetOrder)) {
        return sendJson(response, 404, { error: "Khong tim thay phieu san xuat can nhan." });
      }

      const requestedClaimItems =
        mode === "all" && claimedItems.length ? claimedItems : mode === "all" ? normalizeProductionClaimItemsFromSummary(targetOrder.order_items) : claimedItems;
      const claimItemsForStorage = filterUnclaimedProductionItems(targetOrder, requestedClaimItems);
      if (!claimItemsForStorage.length) {
        return sendJson(response, 400, { error: "Phiếu này không còn dòng nào chưa được nhận sản xuất." });
      }
      const orders = await appendProductionClaimToOrder(resolveOrdersFile(config.cwd), {
        orderId,
        actor: currentUser,
        mode,
        claimedItems: claimItemsForStorage,
      });

      const notificationMessage = buildProductionClaimNotificationMessageV2(currentUser, orderId, mode, claimItemsForStorage, targetOrder);
      const notifyUserId = String(targetOrder.created_by_user_id || targetOrder.sales_user_id || "").trim();
      if (notifyUserId) {
        await appendNotification(resolveNotificationsFile(config.cwd), {
          id: `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
          user_id: notifyUserId,
          type: "production.order.claimed",
          title: `Phiếu ${orderId} đã được nhận sản xuất`,
          message: notificationMessage,
          created_at: formatNow(),
          meta: {
            order_id: orderId,
            claim_mode: mode,
            actor_user_id: currentUser.id,
            actor_user_name: currentUser.name || currentUser.username || currentUser.id,
            claimed_items: claimItemsForStorage,
          },
          read_at: "",
        });
      }

      await logAudit(config.auditLogFile, currentUser, "production.claim", {
        order_id: orderId,
        claim_mode: mode,
        claimed_items: String(claimItemsForStorage.length || 0),
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
        message: notificationMessage,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/production-progress") {
      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      const progressUpdates = normalizeProductionProgressUpdates(payload.progress_updates);
      const orderItems = typeof payload.order_items === "string" ? payload.order_items.trim() : "";
      const markCompleted = Boolean(payload.mark_completed);
      if (!orderId) {
        return sendJson(response, 400, { error: "Thieu ma phieu san xuat." });
      }
      if (!progressUpdates.length) {
        return sendJson(response, 400, { error: "Khong co dong tien do nao de cap nhat." });
      }

      const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
      const targetOrder = existingOrders.find((item) => String(item?.order_id || "").trim() === orderId);
      if (!targetOrder || !isProductionOrderRecord(targetOrder)) {
        return sendJson(response, 404, { error: "Khong tim thay phieu san xuat can cap nhat." });
      }

      const claimUserIdsByLine = getProductionClaimUserIdsByLine(targetOrder);
      const currentUserId = String(currentUser?.id || "").trim();
      const invalidUpdate = progressUpdates.find((item) => {
        const owners = claimUserIdsByLine.get(Math.trunc(item.line_number)) || [];
        return !currentUserId || !owners.includes(currentUserId);
      });
      if (invalidUpdate) {
        return sendJson(response, 403, { error: "Ban chi duoc cap nhat dong san xuat minh da nhan." });
      }

      const orders = await orderStore.updateOrder(orderId, (item) => {
        if (orderItems) {
          return {
            ...item,
            order_items: orderItems,
            production_completed_by_json: markCompleted ? appendProductionCompletionRecord(item, currentUser) : item?.production_completed_by_json || "[]",
            production_packaged_by_json: item?.production_packaged_by_json || "",
            updated_at: formatNow(),
          };
        }
        const summaryItems = parseProductionOrderItemsSummary(item?.order_items || "");
        progressUpdates.forEach((update) => {
          const targetIndex = Math.trunc(update.line_number) - 1;
          const currentLine = summaryItems[targetIndex];
          if (!currentLine) {
            return;
          }
          const doneValue = String(update.done || "").trim() || "0";
          const quantityValue = Number(String(currentLine.quantity || "").replace(",", "."));
          const doneNumber = Math.max(0, Number(String(doneValue || "0").replace(",", ".")));
          const extraValue = String(currentLine.extra || "").trim() || "0";
          const extraNumber = Math.max(0, Number(String(extraValue || "0").replace(",", ".")));
          const missingNumber = Number.isFinite(quantityValue) ? Math.max(0, quantityValue - doneNumber + extraNumber) : 0;
          currentLine.done = doneValue;
          currentLine.team = String(update.team || "").trim() || "-";
          currentLine.missing = String(Number.isFinite(missingNumber) ? missingNumber : currentLine.missing || "0");
        });
        return {
          ...item,
          order_items: buildProductionOrderItemsSummary(summaryItems),
          production_completed_by_json: markCompleted ? appendProductionCompletionRecord(item, currentUser) : item?.production_completed_by_json || "[]",
          production_packaged_by_json: item?.production_packaged_by_json || "",
          updated_at: formatNow(),
        };
      });

      await logAudit(config.auditLogFile, currentUser, "production.progress.update", {
        order_id: orderId,
        updated_lines: progressUpdates.map((item) => String(item.line_number)).join(","),
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/production-packaging-complete") {
      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      if (!orderId) {
        return sendJson(response, 400, { error: "Thiếu mã phiếu sản xuất." });
      }

      const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
      const targetOrder = existingOrders.find((item) => String(item?.order_id || "").trim() === orderId);
      if (!targetOrder || !isProductionOrderRecord(targetOrder)) {
        return sendJson(response, 404, { error: "Không tìm thấy phiếu sản xuất cần đóng gói." });
      }
      if (hasCompletedProductionPackaging(targetOrder)) {
        return sendJson(response, 400, { error: "Phiếu này đã hoàn tất đóng gói." });
      }
      if (!canCompleteProductionPackaging(targetOrder)) {
        return sendJson(response, 400, { error: "Chỉ hoàn tất đóng gói khi tất cả giá trị cột Thiếu đã về 0." });
      }

      const orders = await orderStore.updateOrder(orderId, (item) => ({
        ...item,
        production_packaged_by_json: buildProductionPackagingRecord(currentUser),
        updated_at: formatNow(),
      }));

      const notifyUserId = String(targetOrder.created_by_user_id || "").trim();
      if (notifyUserId) {
        await appendNotification(resolveNotificationsFile(config.cwd), {
          id: `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
          user_id: notifyUserId,
          type: "production.packaging.complete",
          title: `Phiếu ${orderId} đã đóng gói xong`,
          message: buildProductionPackagingNotificationMessage(currentUser, orderId, targetOrder),
          created_at: formatNow(),
          meta: {
            order_id: orderId,
            actor_user_id: currentUser.id,
            actor_user_name: currentUser.name || currentUser.username || currentUser.id,
          },
        });
      }

      await logAudit(config.auditLogFile, currentUser, "production.packaging.complete", {
        order_id: orderId,
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/production-receipt-complete") {
      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      if (!orderId) {
        return sendJson(response, 400, { error: "Thiếu mã phiếu sản xuất." });
      }

      const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
      const targetOrder = existingOrders.find((item) => String(item?.order_id || "").trim() === orderId);
      if (!targetOrder || !isProductionOrderRecord(targetOrder)) {
        return sendJson(response, 404, { error: "Không tìm thấy phiếu sản xuất cần xác nhận nhận hàng." });
      }
      if (!canConfirmProductionReceipt(currentUser, targetOrder)) {
        return sendJson(response, 403, { error: "Chỉ người tạo phiếu mới được xác nhận nhận đủ hàng." });
      }
      if (!hasCompletedProductionPackaging(targetOrder)) {
        return sendJson(response, 400, { error: "Phiếu này chưa hoàn tất đóng gói." });
      }
      if (hasCompletedProductionReceipt(targetOrder)) {
        return sendJson(response, 400, { error: "Phiếu này đã được xác nhận nhận đủ hàng." });
      }

      const orders = await orderStore.updateOrder(orderId, (item) => ({
        ...item,
        production_received_by_json: buildProductionReceiptRecord(currentUser),
        updated_at: formatNow(),
      }));

      await logAudit(config.auditLogFile, currentUser, "production.receipt.complete", {
        order_id: orderId,
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/backfill-details") {
      if (!canViewOrders(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen cap nhat du lieu don hang." });
      }

      const payload = await readJsonBody(request);
      const requestedOrderId = normalizeOrderId(payload.order_id);
      const deliveryAddress = typeof payload.delivery_address === "string" ? payload.delivery_address.trim() : "";
      const orderItems = typeof payload.order_items === "string" ? payload.order_items.trim() : "";
      const orderValue = normalizeOrderValue(payload.order_value);
      const note = typeof payload.note === "string" ? payload.note.trim() : "";

      if (!orderId) {
        return sendJson(response, 400, { error: "Can nhap ma don hang." });
      }

      const orders = await backfillOrderDetails(resolveOrdersFile(config.cwd), {
        orderId,
        deliveryAddress,
        orderItems,
        orderValue,
        note,
        orderKind,
      });

      await logAudit(config.auditLogFile, currentUser, "order.backfill", {
        order_id: orderId,
        fields: [
          deliveryAddress ? "delivery_address" : "",
          orderItems ? "order_items" : "",
          orderValue ? "order_value" : "",
          note ? "note" : "",
        ].filter(Boolean),
      });

      return sendJson(response, 200, {
        ok: true,
        orders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/notifications/read") {
      const payload = await readJsonBody(request);
      const notificationId = typeof payload.notification_id === "string" ? payload.notification_id.trim() : "";
      if (!notificationId) {
        return sendJson(response, 400, { error: "Thiếu mã thông báo cần cập nhật." });
      }
      const notifications = await markNotificationRead(
        resolveNotificationsFile(config.cwd),
        currentUser,
        notificationId,
      );
      return sendJson(response, 200, {
        ok: true,
        notifications,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/notifications/read-all") {
      const notifications = await markAllNotificationsRead(
        resolveNotificationsFile(config.cwd),
        currentUser,
      );
      return sendJson(response, 200, {
        ok: true,
        notifications,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/notifications/delete") {
      const payload = await readJsonBody(request);
      const notificationId = typeof payload.notification_id === "string" ? payload.notification_id.trim() : "";
      if (!notificationId) {
        return sendJson(response, 400, { error: "Thiếu mã thông báo cần xóa." });
      }
      const notifications = await deleteNotification(
        resolveNotificationsFile(config.cwd),
        currentUser,
        notificationId,
      );
      return sendJson(response, 200, {
        ok: true,
        notifications,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/users/access") {
      const payload = await readJsonBody(request);
      const updatedUser = await updateUserAccess(
        config.usersFile,
        currentUser,
        payload,
        userSyncOptions,
      );
      const nextUsers = await resolveCurrentUser(config.usersFile, request, userSyncOptions);
      return sendJson(response, 200, {
        ok: true,
        updatedUser: sanitizeUser(updatedUser),
        currentUser: sanitizeUser(nextUsers.currentUser),
        users: nextUsers.users.map(sanitizeUser),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/users/password") {
      const payload = await readJsonBody(request);
      const updatedUser = await updateUserPassword(
        config.usersFile,
        currentUser,
        payload,
        userSyncOptions,
      );
      const nextUsers = await resolveCurrentUser(config.usersFile, request, userSyncOptions);
      return sendJson(response, 200, {
        ok: true,
        updatedUser: sanitizeUser(updatedUser),
        currentUser: sanitizeUser(nextUsers.currentUser),
        users: nextUsers.users.map(sanitizeUser),
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/sources") {
      await maybeSyncSheetSources();
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      return sendJson(response, 200, {
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sources/sync-one") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen dong bo nguon nay." });
      }
      const payload = await readJsonBody(request);
      const sourceId = typeof payload.source_id === "string" ? payload.source_id.trim() : "";
      const sync = await syncSourceById(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        sourceId,
      );
      await logAudit(config.auditLogFile, currentUser, "source.sync", {
        source_id: sourceId,
        title: sync.source?.title || "",
      });
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sources/syncable") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen cap nhat nguon nay." });
      }
      const payload = await readJsonBody(request);
      const sourceId = typeof payload.source_id === "string" ? payload.source_id.trim() : "";
      const syncable = payload.syncable !== false;
      const updated = await setSourceSyncable(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        sourceId,
        syncable,
      );
      await logAudit(config.auditLogFile, currentUser, "source.syncable", {
        source_id: sourceId,
        syncable: updated.syncable ? "true" : "false",
        title: updated.title,
      });
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sources/delete") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen xoa nguon nay." });
      }
      const payload = await readJsonBody(request);
      const sourceId = typeof payload.source_id === "string" ? payload.source_id.trim() : "";
      const deleted = await deleteSyncSource(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        sourceId,
      );
      await logAudit(config.auditLogFile, currentUser, "source.delete", {
        source_id: sourceId,
        title: deleted.title,
      });
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sources/review") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen duyet nguon nay." });
      }
      const payload = await readJsonBody(request);
      const sourceId = typeof payload.source_id === "string" ? payload.source_id.trim() : "";
      const nextStatus = resolveReviewStatus(payload.decision);
      const reviewed = await reviewSyncSource(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        sourceId,
        {
          status: nextStatus,
          reviewed_at: formatNow(),
          reviewed_by_user_id: currentUser.id,
        },
      );
      await logAudit(config.auditLogFile, currentUser, `source.${nextStatus === "active" ? "approve" : "reject"}`, {
        source_id: sourceId,
        title: reviewed.title,
      });
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/internal-docs") {
      await maybeSyncSheetSources();
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/internal-docs/content") {
      const documentPath = String(requestUrl.searchParams.get("path") || "").trim();
      const document = await readInternalDocument(config.internalDocsDir, config.cwd, documentPath);
      if (!canViewDocument(currentUser, document.metadata || {})) {
        return sendJson(response, 403, { error: "Ban khong co quyen xem tai lieu nay." });
      }
      return sendJson(response, 200, { document });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/internal-docs/download") {
      const documentPath = String(requestUrl.searchParams.get("path") || "").trim();
      const document = await readInternalDocument(config.internalDocsDir, config.cwd, documentPath);
      if (!canViewDocument(currentUser, document.metadata || {})) {
        return sendJson(response, 403, { error: "Ban khong co quyen tai tai lieu nay." });
      }

      const absolutePath = path.resolve(config.cwd, document.path);
      const content = await readFile(absolutePath);
      response.writeHead(200, {
        "Content-Type": MIME_TYPES[path.extname(absolutePath)] ?? "application/octet-stream",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(document.file_name || path.basename(absolutePath))}"`,
        "Cache-Control": "no-store",
      });
      response.end(content);
      return;
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs/delete") {
      const payload = await readJsonBody(request);
      const documentPath = typeof payload.path === "string" ? payload.path.trim() : "";
      const document = await readInternalDocument(config.internalDocsDir, config.cwd, documentPath);
      enforceDocumentManagement(currentUser, document, "delete");
      const deleted = await deleteInternalDocument(
        config.internalDocsDir,
        config.sourcesDir,
        config.cwd,
        documentPath,
      );
      await logAudit(config.auditLogFile, currentUser, "document.delete", {
        title: deleted.title,
        path: deleted.path,
        access_level: deleted.metadata?.access_level || "",
      });
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      return sendJson(response, 200, {
        ok: true,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs/access") {
      const payload = await readJsonBody(request);
      const documentPath = typeof payload.path === "string" ? payload.path.trim() : "";
      const ownerUserId = typeof payload.owner_user_id === "string" ? payload.owner_user_id.trim() : "";
      const sharedWithUsers =
        typeof payload.shared_with_users === "string" ? payload.shared_with_users.trim() : "";
      const currentDocument = await readInternalDocument(config.internalDocsDir, config.cwd, documentPath);
      enforcePrivateDocumentOwnership(currentUser, currentDocument);

      const nextOwner = users.find((user) => user.id === ownerUserId) || null;
      if (ownerUserId && !nextOwner) {
        return sendJson(response, 400, { error: "Khong tim thay nguoi dung moi de chuyen chu so huu." });
      }
      const nextStorageFolder = ownerUserId
        ? `${STORAGE_FOLDER_OPTIONS["du-lieu-ca-nhan"].storage_folder}/${buildPrivateFolderName(nextOwner || currentUser)}`
        : currentDocument.metadata.storage_folder;
      const updated = await updateInternalDocumentMetadata(
        config.internalDocsDir,
        config.cwd,
        documentPath,
        {
          owner_user_id: ownerUserId || currentDocument.metadata.owner_user_id || currentUser.id,
          shared_with_users: normalizeSharedUsers(sharedWithUsers),
          owner: nextOwner?.name || currentDocument.metadata.owner || currentUser.name || currentUser.username || "",
          owner_department: nextOwner?.department || currentDocument.metadata.owner_department || "",
          storage_folder: nextStorageFolder,
        },
      );

      await logAudit(config.auditLogFile, currentUser, "document.access.update", {
        title: updated.title,
        path: updated.path,
        owner_user_id: updated.metadata.owner_user_id || "",
        shared_with_users: updated.metadata.shared_with_users || "",
      });

      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        updated,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs/review") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen duyet tai lieu nay." });
      }
      const payload = await readJsonBody(request);
      const documentPath = typeof payload.path === "string" ? payload.path.trim() : "";
      const nextStatus = resolveReviewStatus(payload.decision);
      const reviewed = await reviewInternalDocument(
        config.internalDocsDir,
        config.cwd,
        documentPath,
        {
          status: nextStatus,
          reviewed_at: formatNow(),
          reviewed_by_user_id: currentUser.id,
        },
      );
      await logAudit(
        config.auditLogFile,
        currentUser,
        `document.${nextStatus === "active" ? "approve" : "reject"}`,
        {
          title: reviewed.title,
          path: reviewed.path,
        },
      );
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      return sendJson(response, 200, {
        ok: true,
        updated: reviewed,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
      });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/audit-logs") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen xem nhat ky he thong." });
      }
      const logs = await readAuditLogs(config.auditLogFile, 200);
      return sendJson(response, 200, { logs });
    }

    if (request.method === "GET" && requestUrl.pathname === "/api/dashboard") {
      if (!canManagePermissions(currentUser)) {
        return sendJson(response, 403, { error: "Ban khong co quyen xem dashboard." });
      }

      await maybeSyncSheetSources();
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const logs = await readAuditLogs(config.auditLogFile, 100);
      return sendJson(response, 200, {
        stats: buildDashboardStats(documents, sources, logs),
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/orders/create") {
      if (!canCreateOrder(currentUser)) {
        return sendJson(response, 403, {
          error: "Chi kinh doanh, quan ly hoac quan tri moi duoc tao don hang.",
        });
      }

      const payload = await readJsonBody(request);
      const requestedOrderId = normalizeOrderId(payload.order_id);
      const customerName = typeof payload.customer_name === "string" ? payload.customer_name.trim() : "";
      const requestedSalesUserId = typeof payload.sales_user_id === "string" ? payload.sales_user_id.trim() : "";
      const orderKind = String(payload.order_kind || "").trim().toLowerCase() === "production" ? "production" : "transport";
      const salesUserId = canChooseSalesAssignee(currentUser)
        ? requestedSalesUserId
        : String(currentUser?.id || "").trim();
      const deliveryUserId = typeof payload.delivery_user_id === "string" ? payload.delivery_user_id.trim() : "";
      const plannedDeliveryAt = normalizeOptionalDateTime(payload.planned_delivery_at);
      const deliveryAddress = typeof payload.delivery_address === "string" ? payload.delivery_address.trim() : "";
      const orderItems = typeof payload.order_items === "string" ? payload.order_items.trim() : "";
      const orderValue = normalizeOrderValue(payload.order_value);
      const note = typeof payload.note === "string" ? payload.note.trim() : "";

      if ((!requestedOrderId && orderKind !== "production") || !customerName || !salesUserId || (orderKind !== "production" && !deliveryUserId)) {
        return sendJson(response, 400, {
          error:
            orderKind === "production"
              ? "Can nhap khach hang va nguoi yeu cau."
              : "Can nhap ma don hang, khach hang, NVKD phu trach va nhan vien giao hang.",
        });
      }

      const salesUser = users.find((user) => user.id === salesUserId);
      if (!salesUser || String(salesUser.department || "").trim().toLowerCase() !== "sales") {
        return sendJson(response, 400, { error: "Nhan vien kinh doanh phu trach khong hop le." });
      }

      const deliveryUser = orderKind === "production" ? null : users.find((user) => user.id === deliveryUserId);
      const deliveryDepartment = String(deliveryUser?.department || "").trim().toLowerCase();
      if (
        orderKind !== "production" &&
        (!deliveryUser || !["operations", "transport", "logistics", "delivery"].includes(deliveryDepartment))
      ) {
        return sendJson(response, 400, { error: "Nhan vien giao hang duoc chon khong hop le." });
      }

      let orderId = await allocateOrderId(requestedOrderId, { orderKind });
      if (!orderId) {
        return sendJson(response, 400, { error: "Can nhap ma don hang hop le." });
      }

      if (orderKind !== "production") {
        const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
        if (existingOrders.some((item) => String(item?.order_id || "").trim().toLowerCase() === orderId.toLowerCase())) {
          return sendJson(response, 409, { error: "Ma don hang nay da ton tai." });
        }
      }

      const order = buildOrderRecord({
        currentUser,
        salesUser,
        deliveryUser,
        orderId,
        customerName,
        plannedDeliveryAt,
        deliveryAddress,
        orderItems,
        orderValue,
        note,
        orderKind,
      });

      if (orderKind !== "production") {
        const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
        const receipts = await readSalesInventoryReceipts(resolveSalesInventoryFile(config.cwd));
        const availabilityMap = buildCombinedInventoryAvailabilityMap(existingOrders, receipts);
        const requestedItems = parseOrderItemsSummary(orderItems).filter((item) => item?.name);
        const shortages = requestedItems
          .map((item) => {
            const key = buildInventoryItemKey(item);
            const available = availabilityMap.get(key)?.quantity || 0;
            const requestedQuantity = Math.max(0, normalizeNonNegativeNumber(item?.quantity || 0));
            return {
              name: String(item?.name || "").trim(),
              unit: String(item?.unit || "").trim() || "-",
              available,
              requestedQuantity,
              missing: Math.max(0, requestedQuantity - available),
            };
          })
          .filter((item) => item.requestedQuantity > 0 && item.available < item.requestedQuantity);

        if (shortages.length) {
          const summary = shortages
            .map((item) => `${item.name} (${formatInventoryNumber(item.available)}/${formatInventoryNumber(item.requestedQuantity)} ${item.unit})`)
            .join("; ");
          return sendJson(response, 400, {
            error: `Khong du ton kho de tao don giao hang: ${summary}. Hay tao phieu san xuat hoac nhap them hang.`,
          });
        }
      }

      if (orderKind === "production") {
        order.document_title = `Phieu de nghi san xuat ${orderId}`;
      }

      if (orderKind === "production") {
        const productionRecipients = users.filter(
          (user) => String(user?.department || "").trim().toLowerCase() === "production",
        );
        while (true) {
          try {
            await createOrder(resolveOrdersFile(config.cwd), order);
            break;
          } catch (error) {
            if (!isDuplicateOrderConstraintError(error)) {
              throw error;
            }
            orderId = formatSequentialOrderId(extractNumericOrderSuffix(orderId) + 1);
            order.order_id = orderId;
            order.document_title = `Phieu de nghi san xuat ${orderId}`;
          }
        }
        const saved = await saveInternalDocument(config.internalDocsDir, config.cwd, {
          title: order.document_title,
          content: buildProductionOrderDocument(order),
          effective_date: order.created_at.slice(0, 10),
          owner: "phòng Sản Xuất",
          owner_department: "production",
          access_level: "advanced",
          allowed_departments: "sales,production",
          owner_user_id: currentUser.id,
          shared_with_users: normalizeSharedUsers(
            [salesUser.id, ...productionRecipients.map((user) => String(user.id || "").trim())].join(","),
          ),
          storage_folder: "phòng Sản Xuất",
          status: "active",
          topic_key: `production-order-${slugifyValue(orderId)}`,
        });

        order.document_path = saved.path;
        await updateOrderRecord(resolveOrdersFile(config.cwd), {
          orderId,
          customerName,
          salesUserId: salesUser.id,
          salesUserName: salesUser.name || "",
          deliveryUserId: "",
          deliveryUserName: "",
          plannedDeliveryAt: "",
          deliveryAddress: "",
          orderItems,
          orderValue: "",
          paymentStatus: "unpaid",
          paymentMethod: "",
          note,
        });

        const notification = await notifyProductionDepartmentAboutOrder({
          users,
          salesUser,
          orderId,
          customerName,
          orderKind: "production",
          documentPath: saved.path,
        });

        await logAudit(config.auditLogFile, currentUser, "order.create", {
          order_id: orderId,
          order_kind: "production",
          customer_name: customerName,
          sales_user_id: salesUser.id,
          delivery_user_id: "",
          planned_delivery_at: "",
          document_path: saved.path,
          notification_sent: notification.sent ? "true" : "false",
          notification_recipients: String(notification.recipients || 0),
        });

        const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
        const folders = await listCustomStorageFolders(config.internalDocsDir);
        const orders = await readOrders(resolveOrdersFile(config.cwd));
        return sendJson(response, 201, {
          ok: true,
          order,
          salesUser: sanitizeUser(salesUser),
          deliveryUser: null,
          notification,
          productionRecipients: productionRecipients.map((user) => sanitizeUser(user)),
          orders,
          documents: filterDocumentsForUser(documents, currentUser),
          folders,
        });
      }

      const saved = await saveInternalDocument(config.internalDocsDir, config.cwd, {
        title: order.document_title,
        content: buildOrderAssignmentDocument(order),
        effective_date: (order.planned_delivery_at || order.created_at).slice(0, 10),
        owner: "Vận chuyển",
        owner_department: "operations",
        access_level: "basic",
        allowed_departments: "sales,operations",
        owner_user_id: currentUser.id,
        shared_with_users: normalizeSharedUsers([salesUser.id, deliveryUser.id].join(",")),
        storage_folder: "Vận Chuyển",
        status: "active",
        topic_key: `order-assigned-${slugifyValue(orderId)}`,
      });

      order.document_path = saved.path;
      await appendOrder(resolveOrdersFile(config.cwd), order);

      await appendNotification(resolveNotificationsFile(config.cwd), {
        id: `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
        user_id: deliveryUser.id,
        type: "order.assigned",
        title: `Co don giao moi ${orderId}`,
        message: `${salesUser.name || salesUser.username || "NVKD"} da giao don ${orderId} cho ban xu ly.`,
        created_at: formatNow(),
        meta: {
          order_id: orderId,
          customer_name: customerName,
          sales_user_id: salesUser.id,
          sales_user_name: salesUser.name || salesUser.username || salesUser.id,
          document_path: saved.path,
        },
        read_at: "",
      });

      let notification = {
        sent: false,
        reason: "",
      };

      if (deliveryUser.email) {
        try {
          await sendMail(config.mail, {
            to: deliveryUser.email,
            subject: `Don giao moi ${orderId}`,
            text: buildOrderAssignmentMail(order, deliveryUser, salesUser),
          });
          notification = {
            sent: true,
            email: deliveryUser.email,
          };
        } catch (error) {
          notification = {
            sent: false,
            reason: error instanceof Error ? error.message : "Khong gui duoc email dieu phoi don hang.",
          };
        }
      } else {
        notification = {
          sent: false,
          reason: "Nhan vien giao hang chua co email de nhan thong bao.",
        };
      }

      await logAudit(config.auditLogFile, currentUser, "order.create", {
        order_id: orderId,
        customer_name: customerName,
        sales_user_id: salesUser.id,
        delivery_user_id: deliveryUser.id,
        planned_delivery_at: plannedDeliveryAt || "",
        document_path: saved.path,
        notification_sent: notification.sent ? "true" : "false",
      });

      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      const orders = await readOrders(resolveOrdersFile(config.cwd));
      return sendJson(response, 201, {
        ok: true,
        order,
        salesUser: sanitizeUser(salesUser),
        deliveryUser: sanitizeUser(deliveryUser),
        notification,
        orders,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/delivery/complete") {
      if (!canCompleteDelivery(currentUser)) {
        return sendJson(response, 403, {
          error: "Chỉ bộ phận vận chuyển hoặc quản trị mới được xác nhận giao hàng.",
        });
      }

      const payload = await readJsonBody(request);
      const orderId = normalizeOrderId(payload.order_id);
      const customerName = typeof payload.customer_name === "string" ? payload.customer_name.trim() : "";
      const salesUserId = typeof payload.sales_user_id === "string" ? payload.sales_user_id.trim() : "";
      const resultStatus = normalizeDeliveryResultStatus(payload.result_status);
      const completedAt = normalizeDeliveryCompletedAt(payload.completed_at);
      const paymentStatus = normalizePaymentStatus(payload.payment_status);
      const paymentMethod = normalizePaymentMethod(payload.payment_method, paymentStatus);
      const note = typeof payload.note === "string" ? payload.note.trim() : "";

      if (!orderId || !customerName || !salesUserId) {
        return sendJson(response, 400, {
          error: "Cần nhập mã đơn hàng, khách hàng và nhân viên kinh doanh phụ trách.",
        });
      }

      const salesUser = users.find((user) => user.id === salesUserId);
      if (!salesUser) {
        return sendJson(response, 400, { error: "Không tìm thấy nhân viên kinh doanh phụ trách đơn hàng." });
      }
      if (String(salesUser.department || "").trim().toLowerCase() !== "sales") {
        return sendJson(response, 400, { error: "Người được chọn phải thuộc bộ phận kinh doanh." });
      }

      const completion = buildDeliveryCompletionRecord({
        currentUser,
        salesUser,
        orderId,
        customerName,
        resultStatus,
        completedAt,
        paymentStatus,
        paymentMethod,
        note,
      });

      await appendDeliveryCompletion(resolveDeliveryCompletionsFile(config.cwd), completion);
      let saved = { path: "" };
      await appendNotification(resolveNotificationsFile(config.cwd), {
        id: `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
        user_id: salesUser.id,
        type: "delivery.complete",
        title: `Đơn ${orderId} đã giao xong`,
        message: `${currentUser.name || currentUser.username || "Nhân viên giao hàng"} đã cập nhật hoàn thành giao hàng cho khách ${customerName}.`,
        created_at: formatNow(),
        meta: {
          order_id: orderId,
          customer_name: customerName,
          result_status: resultStatus,
          document_path: saved.path,
        },
        read_at: "",
      });

      saved = await saveInternalDocument(config.internalDocsDir, config.cwd, {
        title: completion.document_title,
        content: buildDeliveryCompletionDocument(completion),
        effective_date: completion.completed_at.slice(0, 10),
        owner: "Vận chuyển",
        owner_department: "operations",
        access_level: "basic",
        allowed_departments: "sales,operations",
        owner_user_id: currentUser.id,
        shared_with_users: salesUser.id,
        storage_folder: "Vận Chuyển",
        status: "active",
        topic_key: `delivery-complete-${slugifyValue(orderId)}`,
      });

      await updateOrderCompletion(resolveOrdersFile(config.cwd), {
        orderId,
        completion,
        documentPath: saved.path,
      });

      await patchNotificationDocumentPath(resolveNotificationsFile(config.cwd), {
        userId: salesUser.id,
        type: "delivery.complete",
        orderId,
        documentPath: saved.path,
      });

      let notification = {
        sent: false,
        reason: "",
      };

      if (salesUser.email) {
        try {
          await sendMail(config.mail, {
            to: salesUser.email,
            subject: `Đơn ${orderId} đã hoàn thành giao hàng`,
            text: buildDeliveryNotificationMail(completion, salesUser, currentUser),
          });
          notification = {
            sent: true,
            email: salesUser.email,
          };
        } catch (error) {
          notification = {
            sent: false,
            reason: error instanceof Error ? error.message : "Không gửi được email thông báo.",
          };
        }
      } else {
        notification = {
          sent: false,
          reason: "Nhân viên kinh doanh chưa có email để nhận thông báo.",
        };
      }

      await logAudit(config.auditLogFile, currentUser, "delivery.complete", {
        order_id: orderId,
        customer_name: customerName,
        sales_user_id: salesUser.id,
        sales_user_name: salesUser.name || salesUser.username || salesUser.id,
        result_status: resultStatus,
        payment_status: paymentStatus,
        payment_method: paymentMethod,
        completed_at: completedAt,
        document_path: saved.path,
        notification_sent: notification.sent ? "true" : "false",
      });

      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      const orders = await readOrders(resolveOrdersFile(config.cwd));
      return sendJson(response, 201, {
        ok: true,
        completion,
        saved,
        salesUser: sanitizeUser(salesUser),
        notification,
        orders,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-folders") {
      if (!canCreateManagedContent(currentUser)) {
        throw new Error("Chỉ quản trị viên và giám đốc mới được tạo thư mục.");
      }

      const payload = await readJsonBody(request);
      const folderName =
        typeof payload.folder_name === "string" ? payload.folder_name.trim() : "";
      const level =
        typeof payload.level === "string" ? payload.level.trim().toLowerCase() : "";
      const nextFolder = resolveMutableCustomFolder(folderName);

      await createStorageFolder(config.internalDocsDir, nextFolder, level);
      await logAudit(config.auditLogFile, currentUser, "folder.create", {
        folder_name: nextFolder,
        level,
      });

      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 201, {
        ok: true,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-folders/rename") {
      if (!canManagePermissions(currentUser)) {
        throw new Error("Ban khong co quyen doi ten thu muc.");
      }

      const payload = await readJsonBody(request);
      const folderName =
        typeof payload.folder_name === "string" ? payload.folder_name.trim() : "";
      const nextFolderName =
        typeof payload.next_folder_name === "string" ? payload.next_folder_name.trim() : "";
      const currentFolder = resolveMutableCustomFolder(folderName);
      const nextFolder = sanitizeCustomFolderName(nextFolderName);

      if (!nextFolder) {
        return sendJson(response, 400, {
          error: "Hay nhap ten thu muc moi.",
        });
      }

      if (isProtectedStorageFolder(nextFolder)) {
        return sendJson(response, 400, {
          error: "Khong the doi thanh ten thu muc he thong.",
        });
      }

      await renameStorageFolder(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        currentFolder,
        nextFolder,
      );
      await logAudit(config.auditLogFile, currentUser, "folder.rename", {
        folder_name: currentFolder,
        next_folder_name: nextFolder,
      });

      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-folders/delete") {
      if (!canManagePermissions(currentUser)) {
        throw new Error("Ban khong co quyen xoa thu muc.");
      }

      const payload = await readJsonBody(request);
      const folderName =
        typeof payload.folder_name === "string" ? payload.folder_name.trim() : "";
      const currentFolder = resolveMutableCustomFolder(folderName);

      await deleteStorageFolder(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        currentFolder,
      );
      await logAudit(config.auditLogFile, currentUser, "folder.delete", {
        folder_name: currentFolder,
      });

      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs") {
      const payload = await readJsonBody(request);
      const title = typeof payload.title === "string" ? payload.title.trim() : "";
      const content = typeof payload.content === "string" ? payload.content.trim() : "";
      const sheetUrl = typeof payload.sheetUrl === "string" ? payload.sheetUrl.trim() : "";
      const effective_date =
        typeof payload.effective_date === "string" ? payload.effective_date.trim() : "";
      const requestedStatus = typeof payload.status === "string" ? payload.status.trim() : "";
      const topic_key = typeof payload.topic_key === "string" ? payload.topic_key.trim() : "";
      const storageLevel =
        typeof payload.storage_level === "string" ? payload.storage_level.trim() : "";
      const folderKey =
        typeof payload.storage_folder_key === "string" ? payload.storage_folder_key.trim() : "";
      const customFolderName =
        typeof payload.custom_folder_name === "string" ? payload.custom_folder_name.trim() : "";
      const sharedWithUsers =
        typeof payload.shared_with_users === "string" ? payload.shared_with_users.trim() : "";
      const fileName = typeof payload.fileName === "string" ? payload.fileName.trim() : "";
      const contentBase64 =
        typeof payload.contentBase64 === "string" ? payload.contentBase64.trim() : "";
      const sourceType =
        typeof payload.sourceType === "string" ? payload.sourceType.trim() : "text";

      if (!title) {
        return sendJson(response, 400, {
          error: "Tieu de tai lieu la bat buoc.",
        });
      }

      const storage = resolveStorageSelection(
        currentUser,
        storageLevel,
        folderKey,
        sharedWithUsers,
        customFolderName,
      );
      if (!canCreateManagedContent(currentUser)) {
        if (sourceType !== "file" || storage.isPrivate !== true || storage.access_level !== "sensitive") {
          return sendJson(response, 403, {
            error: "Nhân viên chỉ được tải file ở mục cá nhân của chính mình.",
          });
        }
      }
      enforceRequestedAccess(currentUser, {
        access_level: storage.access_level,
        is_private: storage.isPrivate,
      });

      let saved;

      if (sourceType === "sheet") {
        return sendJson(response, 400, {
          error: "Hay dung API /api/sources cho nguon dong bo.",
        });
      } else if (sourceType === "file") {
        if (!fileName || !contentBase64) {
          return sendJson(response, 400, {
            error: "Can cung cap file hop le.",
          });
        }

        saved = await importUploadedDocument(config.internalDocsDir, config.cwd, {
          title,
          fileName,
          contentBase64,
          effective_date,
          owner: storage.owner,
          owner_department: storage.owner_department,
          access_level: storage.access_level,
          allowed_departments: storage.allowed_departments,
          owner_user_id: currentUser.id,
          shared_with_users: storage.shared_with_users,
          storage_folder: storage.storage_folder,
          status: resolveRequestedDocumentStatus(currentUser, requestedStatus),
          topic_key,
        });
      } else {
        if (!content) {
          return sendJson(response, 400, {
            error: "Noi dung tai lieu la bat buoc.",
          });
        }

        saved = await saveInternalDocument(config.internalDocsDir, config.cwd, {
          title,
          content,
          effective_date,
          owner: storage.owner,
          owner_department: storage.owner_department,
          access_level: storage.access_level,
          allowed_departments: storage.allowed_departments,
          owner_user_id: currentUser.id,
          shared_with_users: storage.shared_with_users,
          storage_folder: storage.storage_folder,
          status: resolveRequestedDocumentStatus(currentUser, requestedStatus),
          topic_key,
        });
      }

      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      await logAudit(config.auditLogFile, currentUser, "document.create", {
        title: saved.title,
        path: saved.path,
        source_type: sourceType,
        access_level: storage.access_level,
      });
      return sendJson(response, 201, {
        saved,
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/sources") {
      if (!canCreateManagedContent(currentUser)) {
        return sendJson(response, 403, {
          error: "Chỉ quản trị viên và giám đốc mới được tạo nguồn đồng bộ.",
        });
      }
      const payload = await readJsonBody(request);
      const title = typeof payload.title === "string" ? payload.title.trim() : "";
      const sheetUrl = typeof payload.sheetUrl === "string" ? payload.sheetUrl.trim() : "";
      const effective_date =
        typeof payload.effective_date === "string" ? payload.effective_date.trim() : "";
      const requestedStatus = typeof payload.status === "string" ? payload.status.trim() : "";
      const topic_key = typeof payload.topic_key === "string" ? payload.topic_key.trim() : "";
      const storageLevel =
        typeof payload.storage_level === "string" ? payload.storage_level.trim() : "";
      const folderKey =
        typeof payload.storage_folder_key === "string" ? payload.storage_folder_key.trim() : "";
      const customFolderName =
        typeof payload.custom_folder_name === "string" ? payload.custom_folder_name.trim() : "";
      const sharedWithUsers =
        typeof payload.shared_with_users === "string" ? payload.shared_with_users.trim() : "";

      if (!title || !sheetUrl) {
        return sendJson(response, 400, {
          error: "Can cung cap tieu de va link nguon dong bo.",
        });
      }

      const storage = resolveStorageSelection(
        currentUser,
        storageLevel,
        folderKey,
        sharedWithUsers,
        customFolderName,
      );
      enforceRequestedAccess(currentUser, {
        access_level: storage.access_level,
        is_private: storage.isPrivate,
      });

      const saved = await saveSyncSource(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
        {
          title,
          sheetUrl,
          effective_date,
          owner: storage.owner,
          owner_department: storage.owner_department,
          access_level: storage.access_level,
          allowed_departments: storage.allowed_departments,
          owner_user_id: currentUser.id,
          shared_with_users: storage.shared_with_users,
          storage_folder: storage.storage_folder,
          status: resolveRequestedDocumentStatus(currentUser, requestedStatus),
          topic_key,
        },
      );
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      await logAudit(config.auditLogFile, currentUser, "source.create", {
        title,
        source_url: sheetUrl,
        access_level: storage.access_level,
        storage_folder: storage.storage_folder,
      });
      return sendJson(response, 201, {
        saved,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs/check") {
      const payload = await readJsonBody(request);
      const title = typeof payload.title === "string" ? payload.title.trim() : "";
      const sheetUrl = typeof payload.sheetUrl === "string" ? payload.sheetUrl.trim() : "";
      const effective_date =
        typeof payload.effective_date === "string" ? payload.effective_date.trim() : "";
      const requestedStatus = typeof payload.status === "string" ? payload.status.trim() : "";
      const topic_key = typeof payload.topic_key === "string" ? payload.topic_key.trim() : "";
      const storageLevel =
        typeof payload.storage_level === "string" ? payload.storage_level.trim() : "";
      const folderKey =
        typeof payload.storage_folder_key === "string" ? payload.storage_folder_key.trim() : "";
      const customFolderName =
        typeof payload.custom_folder_name === "string" ? payload.custom_folder_name.trim() : "";
      const sharedWithUsers =
        typeof payload.shared_with_users === "string" ? payload.shared_with_users.trim() : "";

      if (!title) {
        return sendJson(response, 400, {
          error: "Tieu de tai lieu la bat buoc.",
        });
      }

      if (!sheetUrl) {
        return sendJson(response, 400, {
          error: "Can cung cap link Google Sheet.",
        });
      }

      const storage = resolveStorageSelection(
        currentUser,
        storageLevel,
        folderKey,
        sharedWithUsers,
        customFolderName,
      );
      const result = await validateSheetDocument(config.internalDocsDir, config.cwd, {
        title,
        sheetUrl,
        effective_date,
        owner: storage.owner,
        owner_department: storage.owner_department,
        access_level: storage.access_level,
        allowed_departments: storage.allowed_departments,
        owner_user_id: currentUser.id,
        shared_with_users: storage.shared_with_users,
        storage_folder: storage.storage_folder,
        status: resolveRequestedDocumentStatus(currentUser, requestedStatus),
        topic_key,
      });
      return sendJson(response, 200, { ok: true, result });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/internal-docs/sync") {
      const sync = await runSheetSync(true);
      const sources = await listSyncSources(config.sourcesDir, config.internalDocsDir, config.cwd);
      const documents = await listInternalDocuments(config.internalDocsDir, config.cwd);
      const folders = await listCustomStorageFolders(config.internalDocsDir);
      return sendJson(response, 200, {
        ok: true,
        sync,
        sources: sources.filter((source) => canViewDocument(currentUser, source)),
        documents: filterDocumentsForUser(documents, currentUser),
        folders,
      });
    }

    if (request.method === "POST" && requestUrl.pathname === "/api/ask") {
      await maybeSyncSheetSources();

      const payload = await readJsonBody(request);
      const rawQuestion = typeof payload.question === "string" ? payload.question.trim() : "";
      const image = isImagePayload(payload.image) ? payload.image : null;
      const previewOnly = Boolean(payload.previewOnly);
      const extractedText = image ? await extractImageText(image) : "";
      const filters = normalizeAskFilters(currentUser, payload.filters || {});

      if (previewOnly) {
        return sendJson(response, 200, {
          extractedText: extractedText || "Khong trich duoc noi dung tu anh.",
        });
      }

      const question = buildQuestionFromPayload(rawQuestion, extractedText);
      if (!question) {
        return sendJson(response, 400, { error: "Cau hoi khong duoc de trong." });
      }

      if (isConversationalQuery(question)) {
        return sendJson(response, 200, buildConversationalResponse(question));
      }

      const repo = await loadInternalRepository(config.internalDocsDir, config.cwd);
      const accessibleDocuments = filterDocumentsForUser(repo.documents, currentUser).filter(
        (document) => isSearchableDocumentStatus(document.metadata || {}) && matchesAskFilters(document.metadata || {}, filters),
      );
      const accessibleChunks = filterChunksForUser(repo.chunks, currentUser).filter(
        (chunk) => isSearchableDocumentStatus(chunk.metadata || {}) && matchesAskFilters(chunk.metadata || {}, filters),
      );
      const rawMatches = searchDocuments(question, accessibleChunks, { limit: 8 });
      const matches = selectTopDocumentMatches(rawMatches, 4);
      const conflictingMatches = detectConflictingMatches(matches);
      const hasInternalMatch = hasConfidentInternalMatch(question, matches);

      if (conflictingMatches.length > 0) {
        return sendJson(response, 200, buildConflictResponse(conflictingMatches));
      }

      if (hasInternalMatch) {
        return sendJson(
          response,
          200,
          await handleInternalRoute(question, matches, accessibleDocuments),
        );
      }

      return sendJson(response, 200, await handleWebRoute(question, accessibleDocuments.length));
    }

    return sendJson(response, 405, { error: "Method khong duoc ho tro." });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Da xay ra loi khong xac dinh.";
    const statusCode =
      message === "Ban can dang nhap de su dung he thong."
        ? 401
        : message === "Tai khoan hoac mat khau khong dung."
          ? 401
        : message === "Ban khong co quyen phan quyen nguoi dung."
          ? 403
          : 500;
    return sendJson(response, statusCode, {
      error: message,
    });
  }
}

async function maybeSyncSheetSources() {
  return runSheetSync(false);
}

async function runSheetSync(force = false) {
  if (config.sheetSyncIntervalMs <= 0) {
    return {
      ...sheetSyncState.lastResult,
      enabled: false,
      intervalMs: 0,
      lastFinishedAt: sheetSyncState.lastFinishedAt,
    };
  }

  const now = Date.now();
  const shouldStart =
    force ||
    !sheetSyncState.lastStartedAt ||
    now - sheetSyncState.lastStartedAt >= config.sheetSyncIntervalMs;

  if (!shouldStart) {
    if (sheetSyncState.inFlight) {
      await sheetSyncState.inFlight;
    }

    return {
      ...sheetSyncState.lastResult,
      lastFinishedAt: sheetSyncState.lastFinishedAt,
    };
  }

  sheetSyncState.lastStartedAt = now;
  sheetSyncState.inFlight = (async () => {
    try {
      const result = await syncRegisteredSources(
        config.sourcesDir,
        config.internalDocsDir,
        config.cwd,
      );
      sheetSyncState.lastResult = {
        enabled: true,
        intervalMs: config.sheetSyncIntervalMs,
        ...result,
      };
      sheetSyncState.lastFinishedAt = formatSyncTimestamp(new Date());
    } catch (error) {
      sheetSyncState.lastResult = {
        enabled: true,
        intervalMs: config.sheetSyncIntervalMs,
        checked: 0,
        updated: 0,
        unchanged: 0,
        errors: [
          {
            message:
              error instanceof Error ? error.message : "Khong dong bo duoc Google Sheet.",
          },
        ],
      };
      sheetSyncState.lastFinishedAt = formatSyncTimestamp(new Date());
    } finally {
      sheetSyncState.inFlight = null;
    }
  })();

  await sheetSyncState.inFlight;
  return {
    ...sheetSyncState.lastResult,
    lastFinishedAt: sheetSyncState.lastFinishedAt,
  };
}

async function extractImageText(image) {
  if (!image) {
    return "";
  }

  if (typeof image.extractedText === "string" && image.extractedText.trim()) {
    return image.extractedText.trim();
  }

  if (!isGeminiEnabled(config)) {
    throw new Error("Can cau hinh GEMINI_API_KEY de doc du lieu tu anh.");
  }

  return extractTextFromImage(config, {
    mimeType: image.mimeType,
    dataBase64: image.dataBase64,
    prompt:
      "Hay OCR anh nay va tra ve toan bo noi dung van ban co trong anh. Neu co bang du lieu, giu thu tu de doc de hieu.",
  });
}

function buildQuestionFromPayload(question, extractedText) {
  if (!extractedText) {
    return question;
  }

  if (!question) {
    return `Hay doc va tra loi dua tren noi dung OCR sau:\n\n${extractedText}`;
  }

  return `${question}\n\nVan ban trich tu anh:\n${extractedText}`;
}

function isImagePayload(value) {
  return Boolean(
    value &&
      typeof value === "object" &&
      typeof value.mimeType === "string" &&
      typeof value.dataBase64 === "string" &&
      value.mimeType.trim() &&
      value.dataBase64.trim(),
  );
}

async function handleInternalRoute(question, matches, documents) {
  const focusedMatches = focusMatches(matches);
  const selectedDocs = focusedMatches
    .map((match) => documents.find((document) => document.id === match.documentId))
    .filter(Boolean);

  const uniqueDocs = dedupeByKey(selectedDocs, "id");
  const aiEnabled = isGeminiEnabled(config);
  let answer = "";
  let usedAi = false;

  if (aiEnabled) {
    try {
      answer = await answerWithInternalContext(config, {
        question,
        matches: focusedMatches,
        documents: uniqueDocs,
      });
      usedAi = true;
    } catch {
      answer = buildInternalFallbackAnswer(question, focusedMatches);
    }
  } else {
    answer = buildInternalFallbackAnswer(question, focusedMatches);
  }

  const sources = focusedMatches.map((match) => ({
    type: "internal",
    title: match.title,
    path: match.relativePath,
    score: match.score,
    snippet: pickSnippet(match.text, question),
    effectiveDate: match.metadata?.effective_date || "",
    topicKey: match.metadata?.topic_key || "",
    version: match.metadata?.version || "",
    owner: match.metadata?.owner || "",
  }));

  return {
    route: "internal",
    answer,
    suggestions: buildSuggestions("internal", question, sources),
    sources,
    workflow: buildWorkflow({
      route: "internal",
      aiEnabled,
      usedAi,
      topMatchTitle: focusedMatches[0]?.title ?? "Khong ro",
      documentCount: documents.length,
    }),
  };
}

async function handleWebRoute(question, documentCount) {
  const aiEnabled = isGeminiEnabled(config);

  if (!aiEnabled) {
    return {
      route: "web",
      answer: buildNoSearchAnswer(),
      suggestions: buildSuggestions("web-disabled", question, []),
      sources: [],
      workflow: buildWorkflow({
        route: "web",
        aiEnabled,
        usedAi: false,
        usedWebSearch: false,
        documentCount,
      }),
    };
  }

  try {
    const result = await answerWithWebSearch(config, { question });
    return {
      route: "web",
      answer: result.answer,
      suggestions: buildSuggestions("web", question, result.sources),
      sources: result.sources,
      workflow: buildWorkflow({
        route: "web",
        aiEnabled,
        usedAi: true,
        usedWebSearch: true,
        documentCount,
      }),
    };
  } catch (error) {
    return {
      route: "web",
      answer: buildWebSearchErrorAnswer(
        error instanceof Error ? error.message : "Khong the tim kiem Internet.",
      ),
      suggestions: buildSuggestions("web", question, []),
      sources: [],
      workflow: buildWorkflow({
        route: "web",
        aiEnabled,
        usedAi: false,
        usedWebSearch: false,
        documentCount,
      }),
    };
  }
}

function buildConflictResponse(matches) {
  return {
    route: "internal",
    answer: [
      "Co nhieu tai lieu cung chu de nhung so lieu khong khop.",
      "He thong tam thoi khong tra loi thang de tranh sai lech.",
      "Ban hay chon tai lieu theo ngay hieu luc moi nhat hoac xac nhan nguon can dung.",
    ].join(" "),
    suggestions: [],
    sources: matches.map((match) => ({
      type: "internal",
      title: match.title,
      path: match.relativePath,
      score: match.score,
      snippet: pickSnippet(match.text, ""),
      effectiveDate: match.metadata?.effective_date || "",
      topicKey: match.metadata?.topic_key || "",
      owner: match.metadata?.owner || "",
      version: match.metadata?.version || "",
    })),
    workflow: [],
  };
}

function buildConversationalResponse(question) {
  const normalized = String(question || "").trim().toLowerCase();
  const answer = normalized.includes("cam on")
    ? "Không có gì. Bạn cứ đặt câu hỏi, tôi sẽ kiểm tra tài liệu nội bộ hoặc tìm web khi cần."
    : "Chào bạn, tôi sẵn sàng hỗ trợ. Bạn hãy hỏi cụ thể, tôi sẽ kiểm tra dữ liệu nội bộ trước, nếu không đủ sẽ chuyển sang tìm web.";

  return {
    route: "chat",
    answer,
    suggestions: [
      "Hỏi một câu nội bộ cụ thể để tra trong kho tài liệu.",
      "Hỏi một vấn đề hiện tại để hệ thống cân nhắc tìm web.",
      "Dán ảnh hoặc file nếu bạn muốn OCR và phân tích nội dung.",
    ],
    sources: [],
    workflow: [],
  };
}

async function serveStaticFile(pathname, response) {
  const filePath = resolvePublicPath(pathname);

  if (!filePath) {
    return sendJson(response, 404, { error: "Khong tim thay tai nguyen." });
  }

  try {
    await access(filePath);
    const fileStat = await stat(filePath);

    if (!fileStat.isFile()) {
      return sendJson(response, 404, { error: "Khong tim thay tai nguyen." });
    }

    const content = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type": MIME_TYPES[path.extname(filePath)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
    });
    response.end(content);
  } catch {
    return sendJson(response, 404, { error: "Khong tim thay tai nguyen." });
  }
}

function resolvePublicPath(pathname) {
  let cleanPath = pathname === "/" ? "/index.html" : pathname;
  if (!path.extname(cleanPath)) {
    cleanPath = cleanPath.endsWith("/") ? `${cleanPath}index.html` : `${cleanPath}/index.html`;
  }
  const safePath = path.resolve(config.publicDir, `.${cleanPath}`);

  if (!safePath.startsWith(config.publicDir)) {
    return null;
  }

  return safePath;
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(payload, null, 2));
}

function readJsonBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;

    request.on("data", (chunk) => {
      size += chunk.length;
      if (size > 1_000_000) {
        reject(new Error("Payload qua lon."));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });

    request.on("end", () => {
      if (chunks.length === 0) {
        resolve({});
        return;
      }

      try {
        const raw = Buffer.concat(chunks).toString("utf8");
        resolve(JSON.parse(raw));
      } catch {
        reject(new Error("JSON khong hop le."));
      }
    });

    request.on("error", reject);
  });
}

function dedupeByKey(items, key) {
  const seen = new Set();
  const result = [];

  for (const item of items) {
    if (!item || seen.has(item[key])) {
      continue;
    }

    seen.add(item[key]);
    result.push(item);
  }

  return result;
}

function focusMatches(matches) {
  if (matches.length <= 1) {
    return matches;
  }

  const topMatch = matches[0];
  const topScore = topMatch.score || 0;
  let focused = matches
    .filter(
      (match, index) =>
        index === 0 ||
        (match.score >= topScore * 0.35 &&
          match.overlapCount >= 2 &&
          match.coverage >= 0.22),
    )
    .slice(0, 3);

  const topSourceType = String(topMatch.metadata?.source_type || "").toLowerCase();
  if (topSourceType === "sheet") {
    const sheetFirst = focused.filter(
      (match) => String(match.metadata?.source_type || "").toLowerCase() === "sheet",
    );
    if (sheetFirst.length > 0) {
      focused = sheetFirst;
    }
  }

  return focused;
}

function formatSyncTimestamp(value) {
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "medium",
    timeStyle: "medium",
  }).format(value);
}

function enforceRequestedAccess(currentUser, payload) {
  const policy = getEffectivePolicy(currentUser);
  const requestedLevel = String(payload.access_level || "basic").trim().toLowerCase();
  if (requestedLevel === "sensitive" && payload.is_private === true) {
    return;
  }
  const levels = ["basic", "advanced", "sensitive"];
  if (levels.indexOf(requestedLevel) > levels.indexOf(policy.max_access_level)) {
    throw new Error("Ban khong duoc tao tai lieu co muc truy cap cao hon quyen hien tai.");
  }
}

function resolveStorageSelection(
  currentUser,
  storageLevel,
  folderKey,
  sharedWithUsers,
  customFolderName = "",
) {
  const normalizedKey = String(folderKey || "").trim();
  const customConfig = resolveCustomStorageSelection(
    currentUser,
    storageLevel,
    normalizedKey,
    customFolderName,
  );
  const config = customConfig || STORAGE_FOLDER_OPTIONS[normalizedKey];
  if (!config) {
    const existingCustomConfig = resolveExistingCustomStorageSelection(storageLevel, normalizedKey);
    if (!existingCustomConfig) {
      throw new Error("Hay chon muc luu tru va thu muc hop le.");
    }
    return buildResolvedStorageConfig(existingCustomConfig, currentUser, sharedWithUsers);
  }
  return buildResolvedStorageConfig(config, currentUser, sharedWithUsers);
}

function resolveCustomStorageSelection(currentUser, storageLevel, folderKey, customFolderName) {
  if (folderKey !== "__new__") {
    return null;
  }

  if (!canCreateManagedContent(currentUser)) {
    throw new Error("Chỉ quản trị viên và giám đốc mới được tạo thư mục mới ở mức cơ bản hoặc nâng cao.");
  }

  const level = normalizeStorageLevel(storageLevel);
  if (!["basic", "advanced"].includes(level)) {
    throw new Error("Chi co the tao thu muc moi o muc co ban hoac nang cao.");
  }

  const folderName = sanitizeCustomFolderName(customFolderName);
  if (!folderName) {
    throw new Error("Hay nhap ten thu muc moi.");
  }

  return {
    level,
    storage_folder: folderName,
    owner: currentUser.name || currentUser.username || folderName,
    owner_department: currentUser.department || "",
    allowed_departments: "",
    isPrivate: false,
  };
}

function resolveExistingCustomStorageSelection(storageLevel, folderKey) {
  const level = normalizeStorageLevel(storageLevel);
  if (!["basic", "advanced"].includes(level)) {
    return null;
  }

  const folderName = sanitizeCustomFolderName(folderKey);
  if (!folderName) {
    return null;
  }

  return {
    level,
    storage_folder: folderName,
    owner: folderName,
    owner_department: "",
    allowed_departments: "",
    isPrivate: false,
  };
}

function buildResolvedStorageConfig(config, currentUser, sharedWithUsers) {
  const privateFolderName = buildPrivateFolderName(currentUser);

  return {
    ...config,
    owner: config.isPrivate ? currentUser.name || currentUser.username || "private" : config.owner,
    owner_department: config.isPrivate ? currentUser.department || "" : config.owner_department,
    access_level: config.level,
    shared_with_users: normalizeSharedUsers(sharedWithUsers),
    storage_folder: config.isPrivate
      ? `${config.storage_folder}/${privateFolderName}`
      : config.storage_folder,
  };
}

function normalizeSharedUsers(value) {
  return [...new Set(
    String(value || "")
      .split(",")
      .map((item) => String(item || "").trim())
      .filter(Boolean),
  )].join(",");
}

function normalizeStorageLevel(value) {
  const level = String(value || "").trim().toLowerCase();
  return ["basic", "advanced", "sensitive"].includes(level) ? level : "basic";
}

function sanitizeCustomFolderName(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeStorageFolderPath(value) {
  return String(value || "")
    .split(/[\\/]+/)
    .map((segment) => String(segment || "").trim())
    .filter((segment) => segment && segment !== "." && segment !== "..")
    .join("/");
}

function isProtectedStorageFolder(value) {
  const normalized = normalizeStorageFolderPath(value);
  return Object.values(STORAGE_FOLDER_OPTIONS)
    .map((option) => normalizeStorageFolderPath(option.storage_folder))
    .includes(normalized);
}

function resolveMutableCustomFolder(value) {
  const folderName = sanitizeCustomFolderName(value);
  const normalized = normalizeStorageFolderPath(folderName);

  if (!normalized || normalized.includes("/")) {
    throw new Error("Chi co the thao tac voi thu muc custom cap goc.");
  }

  if (isProtectedStorageFolder(normalized)) {
    throw new Error("Khong the chinh sua thu muc he thong.");
  }

  return normalized;
}

function buildPrivateFolderName(user) {
  const raw =
    String(user?.employee_code || "").trim() ||
    String(user?.username || "").trim() ||
    String(user?.id || "").trim() ||
    "private-user";

  return raw.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, "-");
}

function enforceDocumentManagement(currentUser, document, action = "cap nhat") {
  const ownerUserId = String(document?.metadata?.owner_user_id || "").trim();
  const currentUserId = String(currentUser?.id || "").trim();
  const isOwner = ownerUserId && ownerUserId === currentUserId;
  if (canManagePermissions(currentUser) || isOwner) {
    return;
  }

  throw new Error(`Ban khong co quyen ${action} tai lieu nay.`);
}

function enforcePrivateDocumentOwnership(currentUser, document) {
  const level = String(document?.metadata?.access_level || "").trim().toLowerCase();
  const ownerUserId = String(document?.metadata?.owner_user_id || "").trim();
  if (level !== "sensitive") {
    throw new Error("Chỉ tài liệu cá nhân mới có thể cập nhật chia sẻ.");
  }
  if (!ownerUserId || ownerUserId !== String(currentUser?.id || "").trim()) {
    throw new Error("Chỉ người sở hữu mới được cập nhật quyền xem tài liệu cá nhân.");
  }
}

async function logAudit(logFile, user, action, detail = {}) {
  await appendAuditLog(logFile, {
    action,
    actor: {
      id: String(user?.id || "").trim(),
      username: String(user?.username || "").trim(),
      name: String(user?.name || "").trim(),
      role: String(user?.role || "").trim(),
    },
    detail,
  });
}

function buildDashboardStats(documents = [], sources = [], logs = []) {
  const folderCounts = new Map();
  const uploaderCounts = new Map();
  let privateCount = 0;
  let pendingApprovalCount = 0;
  let sourcePausedCount = 0;
  let sourceErrorCount = 0;

  for (const document of documents) {
    const metadata = document.metadata || {};
    const folder = normalizeStorageFolderPath(metadata.storage_folder || "") || "internal";
    folderCounts.set(folder, (folderCounts.get(folder) || 0) + 1);
    const uploader = String(metadata.owner || metadata.owner_user_id || "").trim() || "Khong ro";
    uploaderCounts.set(uploader, (uploaderCounts.get(uploader) || 0) + 1);
    if (String(metadata.access_level || "").trim().toLowerCase() === "sensitive") {
      privateCount += 1;
    }
    if (String(metadata.status || "").trim().toLowerCase() === "draft") {
      pendingApprovalCount += 1;
    }
  }

  for (const source of sources) {
    if (source.syncable === false) {
      sourcePausedCount += 1;
    }
    if (String(source.last_error || "").trim()) {
      sourceErrorCount += 1;
    }
  }

  return {
    documentCount: documents.length,
    sourceCount: sources.length,
    privateCount,
    pendingApprovalCount,
    sourcePausedCount,
    sourceErrorCount,
    topFolders: toSortedSummary(folderCounts),
    topUploaders: toSortedSummary(uploaderCounts),
    recentActions: logs.slice(0, 8),
  };
}

function toSortedSummary(map) {
  return [...map.entries()]
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0], "vi"))
    .slice(0, 5)
    .map(([label, count]) => ({ label, count }));
}

function resolveRequestedDocumentStatus(currentUser, requestedStatus) {
  if (!canManagePermissions(currentUser)) {
    return "draft";
  }

  const normalized = String(requestedStatus || "").trim().toLowerCase();
  return ["active", "draft", "rejected", "superseded"].includes(normalized) ? normalized : "active";
}

function canCreateManagedContent(currentUser) {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin" || role === "director";
}

function canCreateOrder(currentUser) {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager" || department === "sales";
}

function canChooseSalesAssignee(currentUser) {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager";
}

function canManageOrders(currentUser) {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin";
}

function canEditExistingOrder(currentUser, order) {
  if (!currentUser || !order) {
    return false;
  }
  if (canManageOrders(currentUser)) {
    return true;
  }
  return String(order?.created_by_user_id || "").trim() === String(currentUser?.id || "").trim();
}

function canUpdateProductionProgressForLines(currentUser, order, updates = []) {
  if (!currentUser || !order || !isProductionOrderRecord(order)) {
    return false;
  }
  if (canManageOrders(currentUser)) {
    return true;
  }
  const currentUserId = String(currentUser?.id || "").trim();
  if (!currentUserId) {
    return false;
  }
  const claimUserIdsByLine = getProductionClaimUserIdsByLine(order);
  return normalizeProductionProgressUpdates(updates).every((item) => {
    const owners = claimUserIdsByLine.get(Math.trunc(item.line_number)) || [];
    return owners.includes(currentUserId);
  });
}

function canViewOrders(currentUser) {
  if (!currentUser) {
    return false;
  }
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    role === "manager" ||
    department === "sales" ||
    department === "production" ||
    ["operations", "transport", "logistics", "delivery"].includes(department)
  );
}

function canClaimProductionOrder(currentUser) {
  if (!currentUser) {
    return false;
  }
  const department = String(currentUser?.department || "").trim().toLowerCase();
  if (department === "sales") {
    return false;
  }
  return true;
}

function canCompleteDelivery(currentUser) {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    ["operations", "transport", "logistics", "delivery"].includes(department)
  );
}

function isProductionOrderRecord(order) {
  const explicitKind = String(order?.order_kind || "").trim().toLowerCase();
  if (explicitKind === "production") {
    return true;
  }
  if (explicitKind === "transport") {
    return false;
  }
  return String(order?.document_title || "").trim().toLowerCase().includes("phieu de nghi san xuat");
}

function canConfirmProductionReceipt(currentUser, order) {
  if (!currentUser || !order || !isProductionOrderRecord(order)) {
    return false;
  }
  if (canManageOrders(currentUser)) {
    return true;
  }
  return String(order?.created_by_user_id || "").trim() === String(currentUser?.id || "").trim();
}

function normalizeProductionClaimItems(items) {
  return (Array.isArray(items) ? items : [])
    .map((item, index) => ({
      line_number: Math.max(1, Number(item?.line_number || index + 1)),
      code: String(item?.code || "").trim(),
      name: String(item?.name || "").trim(),
      quantity: String(item?.quantity || "").trim(),
      unit: String(item?.unit || "").trim(),
    }))
    .filter((item) => item.name || item.code);
}

function collectClaimedProductionLineNumbers(order) {
  let existingClaims = [];
  try {
    existingClaims = JSON.parse(order?.production_claims_json || "[]");
  } catch {
    existingClaims = [];
  }
  const claimedLineNumbers = new Set();
  (Array.isArray(existingClaims) ? existingClaims : []).forEach((claim) => {
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Math.max(1, Number(item?.line_number || index + 1));
      if (Number.isFinite(lineNumber)) {
        claimedLineNumbers.add(Math.trunc(lineNumber));
      }
    });
  });
  return claimedLineNumbers;
}

function filterUnclaimedProductionItems(order, items = []) {
  const claimedLineNumbers = collectClaimedProductionLineNumbers(order);
  return normalizeProductionClaimItems(items).filter((item) => !claimedLineNumbers.has(Math.trunc(Number(item?.line_number || 0))));
}

function appendProductionCompletionRecord(order, actor) {
  let existing = [];
  try {
    existing = JSON.parse(order?.production_completed_by_json || "[]");
  } catch {
    existing = [];
  }
  const nextEntries = Array.isArray(existing) ? [...existing] : [];
  const actorUserId = String(actor?.id || "").trim();
  if (actorUserId && nextEntries.some((entry) => String(entry?.user_id || "").trim() === actorUserId)) {
    return JSON.stringify(nextEntries);
  }
  nextEntries.push({
    user_id: actorUserId,
    user_name: actor?.name || actor?.username || actor?.id || "Nhân viên sản xuất",
    completed_at: formatNow(),
  });
  return JSON.stringify(nextEntries);
}

function buildProductionPackagingRecord(actor) {
  return JSON.stringify({
    user_id: String(actor?.id || "").trim(),
    user_name: actor?.name || actor?.username || actor?.id || "Nhân viên đóng gói",
    completed_at: formatNow(),
  });
}

function buildProductionReceiptRecord(actor) {
  return JSON.stringify({
    user_id: String(actor?.id || "").trim(),
    user_name: actor?.name || actor?.username || actor?.id || "Người nhận trả hàng",
    completed_at: formatNow(),
  });
}

function hasCompletedProductionPackaging(order) {
  try {
    const parsed = JSON.parse(order?.production_packaged_by_json || "{}");
    return Boolean(String(parsed?.user_name || parsed?.user_id || "").trim());
  } catch {
    return false;
  }
}

function hasCompletedProductionReceipt(order) {
  try {
    const parsed = JSON.parse(order?.production_received_by_json || "{}");
    return Boolean(String(parsed?.user_name || parsed?.user_id || "").trim());
  } catch {
    return false;
  }
}

function canCompleteProductionPackaging(order) {
  const summaryItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  if (!summaryItems.length) {
    return false;
  }
  return summaryItems.every((item) => {
    const missing = Math.max(0, Number(String(item?.missing || "0").replace(",", ".")));
    return missing === 0;
  });
}

function getProductionClaimUserIdsByLine(order) {
  let existingClaims = [];
  try {
    existingClaims = JSON.parse(order?.production_claims_json || "[]");
  } catch {
    existingClaims = [];
  }
  const userIdsByLine = new Map();
  (Array.isArray(existingClaims) ? existingClaims : []).forEach((claim) => {
    const userId = String(claim?.user_id || "").trim();
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Math.max(1, Number(item?.line_number || index + 1));
      if (!Number.isFinite(lineNumber)) {
        return;
      }
      const normalizedLine = Math.trunc(lineNumber);
      const currentUserIds = userIdsByLine.get(normalizedLine) || [];
      if (userId && !currentUserIds.includes(userId)) {
        currentUserIds.push(userId);
      }
      userIdsByLine.set(normalizedLine, currentUserIds);
    });
  });
  return userIdsByLine;
}

function normalizeProductionClaimItemsFromSummary(summary = "") {
  return String(summary || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const match = normalized.match(
        /^(.*?)\s*\|\s*(.*?)\s*\|\s*ĐM\s*(.*?)\s*\|\s*(.*?)\s*\|\s*SL\s*(.*?)\s*\|\s*Đã SX\s*(.*?)\s*\|\s*Thiếu\s*(.*?)\s*\|\s*Dư\s*(.*?)\s*\|\s*Tổ\s*(.*)$/i,
      );
      if (!match) {
        return {
          line_number: index + 1,
          code: "",
          name: normalized,
          quantity: "",
          unit: "",
        };
      }
      return {
        line_number: index + 1,
        code: match[1].trim(),
        name: match[2].trim(),
        quantity: match[5].trim(),
        unit: match[4].trim(),
      };
    })
    .filter((item) => item.name || item.code);
}

function parseProductionOrderItemsSummary(summary = "") {
  return String(summary || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const parts = normalized.split("|").map((part) => String(part || "").trim());
      const readValue = (part, patterns = []) => {
        let nextValue = String(part || "").trim();
        patterns.forEach((pattern) => {
          nextValue = nextValue.replace(pattern, "").trim();
        });
        return nextValue;
      };

      if (parts.length >= 17 && parts[0] === "-") {
        return {
          index: index + 1,
          code: parts[1] || "",
          name: parts[2] || "",
          norm: readValue(parts[3], [/^(?:ĐM|DM|ÄM|\?M)\s*/i]),
          unit: parts[4] || "",
          quantity: readValue(parts[5], [/^SL\s*/i]),
          done: readValue(parts[13], [/^(?:Đã SX|Da SX|ÄÃ£ SX|\?+ SX(?: SX)?)\s*/i]),
          missing: readValue(parts[14], [/^Thiếu\s*/i, /^Thieu\s*/i, /^Thiáº¿u\s*/i, /^Thi\?u\s*/i]),
          extra: readValue(parts[15], [/^Dư\s*/i, /^Du\s*/i, /^DÆ°\s*/i, /^D\?\s*/i]),
          team: readValue(parts[16], [/^Tổ\s*/i, /^To\s*/i, /^Tá»•\s*/i, /^T\?\s*/i]),
        };
      }

      if (parts.length < 9) {
        return {
          index: index + 1,
          code: "",
          name: normalized,
          norm: "",
          unit: "",
          quantity: "",
          done: "",
          missing: "",
          extra: "",
          team: "",
        };
      }
      return {
        index: index + 1,
        code: parts[0] || "",
        name: parts[1] || "",
        norm: readValue(parts[2], [/^(?:ĐM|DM|ÄM|\?M)\s*/i]),
        unit: parts[3] || "",
        quantity: readValue(parts[4], [/^SL\s*/i]),
        done: readValue(parts[5], [/^(?:Đã SX|Da SX|ÄÃ£ SX|\?+ SX(?: SX)?)\s*/i]),
        missing: readValue(parts[6], [/^Thiếu\s*/i, /^Thieu\s*/i, /^Thiáº¿u\s*/i, /^Thi\?u\s*/i]),
        extra: readValue(parts[7], [/^Dư\s*/i, /^Du\s*/i, /^DÆ°\s*/i, /^D\?\s*/i]),
        team: readValue(parts[8], [/^Tổ\s*/i, /^To\s*/i, /^Tá»•\s*/i, /^T\?\s*/i]),
      };
    });
}

function parseOrderItemsSummary(value = "") {
  return String(value || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const match = normalized.match(/^(.*?)\s+x\s+(\d+(?:[.,]\d+)?)\s+(\S+)/i);
      if (!match) {
        return {
          index: index + 1,
          name: normalized,
          quantity: "",
          unit: "",
        };
      }
      return {
        index: index + 1,
        name: match[1].trim(),
        quantity: match[2].trim(),
        unit: match[3].trim(),
      };
    });
}

function buildProductionOrderItemsSummary(items = []) {
  return (Array.isArray(items) ? items : [])
    .map((item, index) => {
      const code = String(item?.code || "").trim() || "-";
      const name = String(item?.name || "").trim();
      const norm = String(item?.norm || "").trim() || "-";
      const unit = String(item?.unit || "").trim() || "-";
      const quantity = String(item?.quantity || "").trim() || "0";
      const done = String(item?.done || "").trim() || "0";
      const missing = String(item?.missing || "").trim() || "0";
      const extra = String(item?.extra || "").trim() || "0";
      const team = String(item?.team || "").trim() || "-";
      return `${index + 1}. ${code} | ${name} | ĐM ${norm} | ${unit} | SL ${quantity} | Đã SX ${done} | Thiếu ${missing} | Dư ${extra} | Tổ ${team}`;
    })
    .filter(Boolean)
    .join("\n");
}

function normalizeProductionProgressUpdates(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => ({
      line_number: Math.max(1, Number(item?.line_number || 0)),
      done: String(item?.done || "").trim(),
      team: String(item?.team || "").trim(),
    }))
    .filter((item) => Number.isFinite(item.line_number) && item.line_number > 0);
}

function formatProductionClaimItem(item) {
  return [String(item?.quantity || "").trim(), String(item?.unit || "").trim(), String(item?.name || item?.code || "").trim()]
    .filter(Boolean)
    .join(" ");
}

function buildProductionClaimNotificationMessage(actor, orderId, mode, claimedItems, order) {
  const actorName = actor?.name || actor?.username || "Nhân viên sản xuất";
  if (mode === "all") {
    return `${actorName} đã nhận sản xuất toàn bộ phiếu ${orderId} cho ${order?.customer_name || "khách hàng"}.`;
  }
  const itemsText = claimedItems.map(formatProductionClaimItem).filter(Boolean).join(", ");
  return `${actorName} đã nhận sản xuất ${itemsText || `một phần phiếu ${orderId}`}.`;
}

function buildProductionClaimNotificationMessageV2(actor, orderId, mode, claimedItems, order) {
  const actorName = actor?.name || actor?.username || "NhÃ¢n viÃªn sáº£n xuáº¥t";
  const normalizedClaimItems = normalizeProductionClaimItems(claimedItems);
  const totalItems = normalizeProductionClaimItemsFromSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const claimedLineNumbers = collectClaimedProductionLineNumbers(order);
  const remainingBeforeCount = Math.max(0, totalItems.length - claimedLineNumbers.size);
  const completesRemaining = remainingBeforeCount > 0 && normalizedClaimItems.length >= remainingBeforeCount;

  let existingClaims = [];
  try {
    existingClaims = JSON.parse(order?.production_claims_json || "[]");
  } catch {
    existingClaims = [];
  }

  const actorUserId = String(actor?.id || "").trim();
  const actorUserName = String(actorName || "").trim().toLowerCase();
  const priorClaimOwners = new Set();
  (Array.isArray(existingClaims) ? existingClaims : []).forEach((claim) => {
    const userId = String(claim?.user_id || "").trim();
    const userName = String(claim?.user_name || "").trim().toLowerCase();
    priorClaimOwners.add(userId || userName);
  });
  const onlyActorHadClaimedBefore =
    priorClaimOwners.size === 0 ||
    (priorClaimOwners.size === 1 && priorClaimOwners.has(actorUserId || actorUserName));

  if (completesRemaining) {
    if (priorClaimOwners.size === 0 || onlyActorHadClaimedBefore) {
      return `${actorName} Ä‘Ã£ nháº­n sáº£n xuáº¥t toÃ n bá»™ phiáº¿u ${orderId} cho ${order?.customer_name || "khÃ¡ch hÃ ng"}.`;
    }
    return `${actorName} Ä‘Ã£ nháº­n sáº£n xuáº¥t pháº§n cÃ²n láº¡i cá»§a phiáº¿u ${orderId} cho ${order?.customer_name || "khÃ¡ch hÃ ng"}.`;
  }

  if (mode === "all") {
    return `${actorName} Ä‘Ã£ nháº­n sáº£n xuáº¥t toÃ n bá»™ phiáº¿u ${orderId} cho ${order?.customer_name || "khÃ¡ch hÃ ng"}.`;
  }
  const itemsText = normalizedClaimItems.map(formatProductionClaimItem).filter(Boolean).join(", ");
  return `${actorName} Ä‘Ã£ nháº­n sáº£n xuáº¥t ${itemsText || `má»™t pháº§n phiáº¿u ${orderId}`}.`;
}

function buildProductionPackagingNotificationMessage(actor, orderId, order) {
  const actorName = actor?.name || actor?.username || "Nhân viên đóng gói";
  return `${actorName} đã hoàn tất đóng gói phiếu ${orderId} cho ${order?.customer_name || "khách hàng"}. Bạn có thể xác nhận nhận đủ hàng.`;
}

function resolveReviewStatus(decision) {
  const normalized = String(decision || "").trim().toLowerCase();
  if (normalized === "approve") {
    return "active";
  }
  if (normalized === "reject") {
    return "rejected";
  }
  throw new Error("Quyet dinh duyet khong hop le.");
}

function isSearchableDocumentStatus(metadata = {}) {
  return String(metadata.status || "").trim().toLowerCase() === "active";
}

function formatNow() {
  return new Date().toISOString();
}

function getUsersByDepartment(users, department) {
  const normalizedDepartment = String(department || "").trim().toLowerCase();
  return (Array.isArray(users) ? users : []).filter(
    (user) => String(user?.department || "").trim().toLowerCase() === normalizedDepartment,
  );
}

async function notifyProductionDepartmentAboutOrder({
  users,
  salesUser,
  orderId,
  customerName,
  orderKind = "transport",
  documentPath = "",
}) {
  const productionRecipients = getUsersByDepartment(users, "production");
  const safeDocumentPath = String(documentPath || "").trim();
  const title =
    orderKind === "production" ? `Co phieu san xuat moi ${orderId}` : `Co don hang moi ${orderId}`;
  const message =
    orderKind === "production"
      ? `${salesUser.name || salesUser.username || "NVKD"} vua tao phieu san xuat ${orderId} cho ${customerName}.`
      : `${salesUser.name || salesUser.username || "NVKD"} vua tao don hang ${orderId} cho ${customerName}. Bo phan san xuat vui long tiep nhan thong tin.`;

  for (const recipient of productionRecipients) {
    await appendNotification(resolveNotificationsFile(config.cwd), {
      id: `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      user_id: recipient.id,
      type: "production.order.created",
      title,
      message,
      created_at: formatNow(),
      meta: {
        order_id: orderId,
        customer_name: customerName,
        sales_user_id: salesUser.id,
        sales_user_name: salesUser.name || salesUser.username || salesUser.id,
        document_path: safeDocumentPath,
        order_kind: orderKind,
      },
      read_at: "",
    });
  }

  return {
    sent: productionRecipients.length > 0,
    reason: productionRecipients.length > 0 ? "" : "Chua co nhan vien san xuat nao de nhan thong bao.",
    recipients: productionRecipients.length,
  };
}

function resolveDeliveryCompletionsFile(cwd) {
  return path.resolve(cwd, "./data/delivery/completions.json");
}

function resolveOrdersFile(cwd) {
  return path.resolve(cwd, "./data/orders/orders.json");
}

function resolveOrderProductsFile(cwd) {
  return path.resolve(cwd, "./data/orders/catalog.json");
}

function resolveNotificationsFile(cwd) {
  return path.resolve(cwd, "./data/notifications/inbox.json");
}

function resolveSalesInventoryFile(cwd) {
  return path.resolve(cwd, "./data/inventory/sales-receipts.json");
}

function resolveSalesInventoryTransfersFile(cwd) {
  return path.resolve(cwd, "./data/inventory/sales-transfers.json");
}

function defaultOrderProducts() {
  return [
    { id: "foam-10", name: "Xốp nổ 10cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-20", name: "Xốp nổ 20cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-30", name: "Xốp nổ 30cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-40", name: "Xốp nổ 40cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-50", name: "Xốp nổ 50cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-60", name: "Xốp nổ 60cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "tape-1kg", name: "Băng dính vàng 1kg", units: ["cuộn"], default_unit: "cuộn" },
    { id: "tape-500g", name: "Băng dính vàng 500g", units: ["cuộn"], default_unit: "cuộn" },
    { id: "core-n3", name: "Lõi nhựa N3", units: ["cây"], default_unit: "cây" },
    { id: "core-n5", name: "Lõi nhựa N5", units: ["cây"], default_unit: "cây" },
    { id: "core-n6", name: "Lõi nhựa N6", units: ["cây"], default_unit: "cây" },
  ];
}

function normalizeOrderProducts(items) {
  const seen = new Set();
  return (Array.isArray(items) ? items : [])
    .map((item, index) => {
      const name = String(item?.name || "").trim();
      const units = (Array.isArray(item?.units) ? item.units : [])
        .map((unit) => String(unit || "").trim())
        .filter(Boolean);
      const defaultUnit = String(item?.default_unit || units[0] || "").trim();
      if (!name) {
        return null;
      }
      const key = name.toLowerCase();
      if (seen.has(key)) {
        return null;
      }
      seen.add(key);
      return {
        id: String(item?.id || `product-${index + 1}`).trim(),
        name,
        units: units.length ? units : [defaultUnit || "cây"],
        default_unit: defaultUnit || units[0] || "cây",
      };
    })
    .filter(Boolean);
}

async function readOrderProducts(filePath) {
  return orderStore.readProducts();
}

async function saveOrderProducts(filePath, items) {
  return orderStore.saveProducts(items);
}

async function appendOrder(filePath, record) {
  await orderStore.appendOrder(record);
}

async function createOrder(filePath, record) {
  return orderStore.appendOrder(record, { allowReplace: false });
}

async function readOrders(filePath) {
  return orderStore.readOrders();
}

async function appendDeliveryCompletion(filePath, record) {
  await orderStore.appendDeliveryCompletion(record);
}

async function readDeliveryCompletions(filePath) {
  return orderStore.readDeliveryCompletions();
}

async function appendNotification(filePath, record) {
  await orderStore.appendNotification(record);
  broadcastNotification(record);
}

async function readNotifications(filePath, currentUser = null, options = {}) {
  return orderStore.readNotifications(currentUser, options);
}

function openNotificationStream(response, currentUser) {
  const userId = String(currentUser?.id || "").trim();
  if (!userId) {
    return sendJson(response, 401, { error: "Unauthorized" });
  }

  response.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-cache, no-transform",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });
  response.write(": connected\n\n");

  const client = {
    response,
    heartbeat: setInterval(() => {
      try {
        response.write(": ping\n\n");
      } catch {}
    }, 15000),
  };

  const bucket = notificationStreams.get(userId) || new Set();
  bucket.add(client);
  notificationStreams.set(userId, bucket);

  const cleanup = () => {
    clearInterval(client.heartbeat);
    const current = notificationStreams.get(userId);
    if (!current) {
      return;
    }
    current.delete(client);
    if (current.size === 0) {
      notificationStreams.delete(userId);
    }
  };

  response.on("close", cleanup);
  response.on("error", cleanup);
}

function broadcastNotification(record) {
  const userId = String(record?.user_id || "").trim();
  if (!userId) {
    return;
  }
  const clients = notificationStreams.get(userId);
  if (!clients || clients.size === 0) {
    return;
  }

  const payload = `event: notification\ndata: ${JSON.stringify(record)}\n\n`;
  for (const client of [...clients]) {
    try {
      client.response.write(payload);
    } catch {
      clearInterval(client.heartbeat);
      clients.delete(client);
    }
  }

  if (clients.size === 0) {
    notificationStreams.delete(userId);
  }
}

async function markNotificationRead(filePath, currentUser, notificationId) {
  return orderStore.markNotificationRead(currentUser, notificationId, formatNow());
}

async function markAllNotificationsRead(filePath, currentUser) {
  return orderStore.markAllNotificationsRead(currentUser, formatNow());
}

async function deleteNotification(filePath, currentUser, notificationId) {
  return orderStore.deleteNotification(currentUser, notificationId);
}

async function patchNotificationDocumentPath(filePath, { userId, type, orderId, documentPath }) {
  await orderStore.patchNotificationDocumentPath({ userId, type, orderId, documentPath });
}

function normalizeNonNegativeNumber(value) {
  const numeric = Number.parseFloat(String(value ?? "").replace(/,/g, "."));
  if (!Number.isFinite(numeric) || numeric < 0) {
    return 0;
  }
  return Math.round(numeric * 1000) / 1000;
}

function buildSalesInventoryReceiptRecord({
  currentUser,
  code = "",
  name = "",
  unit = "",
  quantity = 0,
  supplier = "",
  note = "",
  receivedAt = formatNow(),
}) {
  const normalizedAt = normalizeOptionalDateTime(receivedAt) || formatNow();
  return {
    id: `sales-stock-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
    code: String(code || "").trim(),
    name: String(name || "").trim(),
    unit: String(unit || "").trim(),
    quantity: normalizeNonNegativeNumber(quantity),
    supplier: String(supplier || "").trim(),
    note: String(note || "").trim(),
    received_at: normalizedAt,
    created_at: formatNow(),
    created_by_user_id: String(currentUser?.id || "").trim(),
    created_by_user_name: String(currentUser?.name || currentUser?.username || "").trim(),
  };
}

async function readSalesInventoryReceipts(filePath) {
  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
}

async function appendSalesInventoryReceipt(filePath, record) {
  const items = await readSalesInventoryReceipts(filePath);
  const nextItems = [record, ...items];
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  return nextItems;
}

function buildInventoryItemKey(item) {
  return `${String(item?.code || "").trim().toLowerCase()}|${String(item?.name || "").trim().toLowerCase()}|${String(item?.unit || "").trim().toLowerCase()}`;
}

function formatInventoryNumber(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return "0";
  }
  return Number.isInteger(numeric) ? String(numeric) : String(numeric.toFixed(2)).replace(/\.?0+$/, "");
}

function normalizeSalesInventoryTransferItems(items) {
  return (Array.isArray(items) ? items : [])
    .map((item) => ({
      code: String(item?.code || "").trim(),
      name: String(item?.name || "").trim(),
      unit: String(item?.unit || "").trim(),
      quantity: normalizeNonNegativeNumber(item?.quantity),
    }))
    .filter((item) => (item.name || item.code) && item.unit && item.quantity > 0);
}

function buildSalesInventoryTransferBatch({ currentUser, transferredAt = formatNow(), note = "", items = [] }) {
  const transferId = `sales-transfer-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  const normalizedAt = normalizeOptionalDateTime(transferredAt) || formatNow();
  return {
    transfer_id: transferId,
    items: normalizeSalesInventoryTransferItems(items).map((item) => ({
      id: `${transferId}-${Math.random().toString(16).slice(2, 6)}`,
      transfer_id: transferId,
      source: "production",
      code: item.code,
      name: item.name,
      unit: item.unit,
      quantity: item.quantity,
      note: String(note || "").trim(),
      transferred_at: normalizedAt,
      created_at: formatNow(),
      created_by_user_id: String(currentUser?.id || "").trim(),
      created_by_user_name: String(currentUser?.name || currentUser?.username || "").trim(),
    })),
  };
}

async function readSalesInventoryTransfers(filePath) {
  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
}

async function appendSalesInventoryTransfers(filePath, items = []) {
  const existingItems = await readSalesInventoryTransfers(filePath);
  const nextItems = [...(Array.isArray(items) ? items : []), ...existingItems];
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  return nextItems;
}

function buildProductionAvailabilityMap(orders = [], transfers = []) {
  const itemMap = new Map();

  (Array.isArray(orders) ? orders : []).forEach((order) => {
    if (isProductionOrderRecord(order)) {
      if (!hasCompletedProductionReceipt(order)) {
        return;
      }
      const parsedItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
      parsedItems.forEach((item) => {
        const key = buildInventoryItemKey(item);
        const current = itemMap.get(key) || {
          code: String(item?.code || "").trim(),
          name: String(item?.name || "").trim(),
          unit: String(item?.unit || "").trim() || "-",
          quantity: 0,
        };
        current.quantity += Math.max(0, normalizeNonNegativeNumber(item?.done || 0));
        itemMap.set(key, current);
      });
      return;
    }

    const parsedItems = parseOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
    parsedItems.forEach((item) => {
      const key = buildInventoryItemKey(item);
      const current = itemMap.get(key) || {
        code: String(item?.code || "").trim(),
        name: String(item?.name || "").trim(),
        unit: String(item?.unit || "").trim() || "-",
        quantity: 0,
      };
      current.quantity -= Math.max(0, normalizeNonNegativeNumber(item?.quantity || 0));
      itemMap.set(key, current);
    });
  });

  (Array.isArray(transfers) ? transfers : []).forEach((item) => {
    const key = buildInventoryItemKey(item);
    const current = itemMap.get(key) || {
      code: String(item?.code || "").trim(),
      name: String(item?.name || "").trim(),
      unit: String(item?.unit || "").trim() || "-",
      quantity: 0,
    };
    current.quantity -= Math.max(0, normalizeNonNegativeNumber(item?.quantity || 0));
    itemMap.set(key, current);
  });

  return itemMap;
}

function buildCombinedInventoryAvailabilityMap(orders = [], receipts = []) {
  const itemMap = new Map();

  (Array.isArray(orders) ? orders : []).forEach((order) => {
    if (isProductionOrderRecord(order)) {
      if (!hasCompletedProductionReceipt(order)) {
        return;
      }
      const parsedItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
      parsedItems.forEach((item) => {
        const key = buildInventoryItemKey(item);
        const current = itemMap.get(key) || {
          code: String(item?.code || "").trim(),
          name: String(item?.name || "").trim(),
          unit: String(item?.unit || "").trim() || "-",
          quantity: 0,
        };
        current.quantity += Math.max(0, normalizeNonNegativeNumber(item?.done || 0));
        itemMap.set(key, current);
      });
      return;
    }

    const parsedItems = parseOrderItemsSummary(order?.order_items || "").filter((item) => item?.name);
    parsedItems.forEach((item) => {
      const key = buildInventoryItemKey(item);
      const current = itemMap.get(key) || {
        code: "",
        name: String(item?.name || "").trim(),
        unit: String(item?.unit || "").trim() || "-",
        quantity: 0,
      };
      current.quantity -= Math.max(0, normalizeNonNegativeNumber(item?.quantity || 0));
      itemMap.set(key, current);
    });
  });

  (Array.isArray(receipts) ? receipts : []).forEach((receipt) => {
    const key = buildInventoryItemKey(receipt);
    const current = itemMap.get(key) || {
      code: String(receipt?.code || "").trim(),
      name: String(receipt?.name || "").trim(),
      unit: String(receipt?.unit || "").trim() || "-",
      quantity: 0,
    };
    current.quantity += Math.max(0, normalizeNonNegativeNumber(receipt?.quantity || 0));
    itemMap.set(key, current);
  });

  return itemMap;
}

async function legacyUpdateOrderCompletion(filePath, { orderId, completion, documentPath = "" }) {
  const normalizedOrderId = String(orderId || "").trim();
  if (!normalizedOrderId) {
    return;
  }

  const items = await readOrders(filePath);
  let changed = false;
  const nextItems = items.map((item) => {
    if (String(item?.order_id || "").trim().toLowerCase() !== normalizedOrderId.toLowerCase()) {
      return item;
    }
    changed = true;
    const normalizedStatus = String(item?.status || "").trim().toLowerCase();
    return {
      ...item,
      status: "completed",
      completed_at: completion?.completed_at || formatNow(),
      completion_result_status: completion?.result_status || "delivered",
      completion_result_label: completion?.result_label || labelDeliveryResultStatus(completion?.result_status),
      payment_status: completion?.payment_status || "unpaid",
      payment_label: completion?.payment_label || (completion?.payment_status === "paid" ? "Đã thanh toán" : "Chưa thanh toán"),
      payment_method: completion?.payment_method || "",
      payment_method_label: completion?.payment_method_label || labelPaymentMethod(completion?.payment_method),
      completion_note: completion?.note || "",
      completed_by_user_id: completion?.driver_user_id || "",
      completed_by_user_name: completion?.driver_user_name || "",
      completion_document_path: documentPath || item.completion_document_path || "",
      updated_at: formatNow(),
    };
  });

  if (changed) {
    await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  }
}

async function legacyUpdateOrderRecord(
  filePath,
  { orderId, customerName, salesUserId, deliveryUserId, plannedDeliveryAt, deliveryAddress, orderItems, orderValue, paymentStatus, paymentMethod, note },
) {
  const normalizedOrderId = normalizeOrderId(orderId);
  const items = await readOrders(filePath);
  let changed = false;
  const nextItems = items.map((item) => {
    if (String(item?.order_id || "").trim().toUpperCase() !== normalizedOrderId) {
      return item;
    }
    changed = true;
    const normalizedStatus = String(item?.status || "").trim().toLowerCase();
    const nextSalesUser = salesUserId ? findUserById(salesUserId) : null;
    const nextDeliveryUser = deliveryUserId ? findUserById(deliveryUserId) : null;
    return {
      ...item,
      customer_name: customerName,
      sales_user_id: salesUserId || item.sales_user_id || "",
      sales_user_name: nextSalesUser?.name || item.sales_user_name || "",
      delivery_user_id: deliveryUserId || item.delivery_user_id || "",
      delivery_user_name: nextDeliveryUser?.name || item.delivery_user_name || "",
      planned_delivery_at: plannedDeliveryAt || item.planned_delivery_at || "",
      delivery_address: deliveryAddress,
      order_items: orderItems,
      order_value: normalizeOrderValue(orderValue),
      payment_status: normalizedStatus === "completed" ? paymentStatus : item.payment_status || "unpaid",
      payment_label: normalizedStatus === "completed" ? (paymentStatus === "paid" ? "Đã thanh toán" : "Chưa thanh toán") : item.payment_label || "",
      payment_label: paymentStatus === "paid" ? "Đã thanh toán" : "Chưa thanh toán",
      payment_method: normalizedStatus === "completed" ? paymentMethod : item.payment_method || "",
      payment_method_label: normalizedStatus === "completed" ? labelPaymentMethod(paymentMethod) : item.payment_method_label || "",
      completion_note: normalizedStatus === "completed" ? note : item.completion_note || "",
      note: normalizedStatus === "completed" ? item.note || "" : note,
      updated_at: formatNow(),
    };
  });

  if (!changed) {
    throw new Error("Khong tim thay don hang de cap nhat.");
  }

  await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  return nextItems;
}

async function legacyDeleteOrderRecord(filePath, orderId) {
  const normalizedOrderId = normalizeOrderId(orderId);
  const items = await readOrders(filePath);
  const nextItems = items.filter((item) => String(item?.order_id || "").trim().toUpperCase() !== normalizedOrderId);

  if (nextItems.length === items.length) {
    throw new Error("Khong tim thay don hang de xoa.");
  }

  await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  return nextItems;
}

async function legacyBackfillOrderDetails(filePath, { orderId, deliveryAddress, orderItems, orderValue, note }) {
  const normalizedOrderId = normalizeOrderId(orderId);
  const items = await readOrders(filePath);
  let changed = false;
  let found = false;

  const nextItems = items.map((item) => {
    if (String(item?.order_id || "").trim().toUpperCase() !== normalizedOrderId) {
      return item;
    }

    found = true;
    const nextItem = { ...item };

    if (!String(nextItem.delivery_address || "").trim() && deliveryAddress) {
      nextItem.delivery_address = deliveryAddress;
      changed = true;
    }
    if (!String(nextItem.order_items || "").trim() && orderItems) {
      nextItem.order_items = orderItems;
      changed = true;
    }
    if (!String(nextItem.order_value || "").trim() && orderValue) {
      nextItem.order_value = normalizeOrderValue(orderValue);
      changed = true;
    }
    if (!String(nextItem.note || nextItem.completion_note || "").trim() && note) {
      if (String(nextItem.status || "").trim().toLowerCase() === "completed") {
        nextItem.completion_note = note;
      } else {
        nextItem.note = note;
      }
      changed = true;
    }

    if (changed) {
      nextItem.updated_at = formatNow();
    }

    return nextItem;
  });

  if (!found) {
    throw new Error("Khong tim thay don hang de bo sung.");
  }

  if (changed) {
    await writeFile(filePath, `${JSON.stringify(nextItems, null, 2)}\n`, "utf8");
  }

  return nextItems;
}

async function updateOrderCompletion(filePath, { orderId, completion, documentPath = "" }) {
  const normalizedOrderId = String(orderId || "").trim();
  if (!normalizedOrderId) {
    return;
  }

  await orderStore.updateOrder(normalizedOrderId, (item) => ({
    ...item,
    status: "completed",
    completed_at: completion?.completed_at || formatNow(),
    completion_result_status: completion?.result_status || "delivered",
    completion_result_label: completion?.result_label || labelDeliveryResultStatus(completion?.result_status),
    payment_status: completion?.payment_status || "unpaid",
    payment_label: completion?.payment_label || (completion?.payment_status === "paid" ? "Đã thanh toán" : "Chưa thanh toán"),
    payment_method: completion?.payment_method || "",
    payment_method_label: completion?.payment_method_label || labelPaymentMethod(completion?.payment_method),
    completion_note: completion?.note || "",
    completed_by_user_id: completion?.driver_user_id || "",
    completed_by_user_name: completion?.driver_user_name || "",
    completion_document_path: documentPath || item.completion_document_path || "",
    updated_at: formatNow(),
  }));
}

async function updateOrderRecord(
  filePath,
  {
    orderId,
    customerName,
    salesUserId,
    salesUserName,
    deliveryUserId,
    deliveryUserName,
    plannedDeliveryAt,
    deliveryAddress,
    orderItems,
    orderValue,
    paymentStatus,
    paymentMethod,
    note,
  },
) {
  const normalizedOrderId = normalizeOrderId(orderId);
  return orderStore.updateOrder(normalizedOrderId, (item) => {
    const normalizedStatus = String(item?.status || "").trim().toLowerCase();
    return {
      ...item,
      customer_name: customerName,
      sales_user_id: salesUserId || item.sales_user_id || "",
      sales_user_name: salesUserName || item.sales_user_name || "",
      delivery_user_id: deliveryUserId || item.delivery_user_id || "",
      delivery_user_name: deliveryUserName || item.delivery_user_name || "",
      planned_delivery_at: plannedDeliveryAt || item.planned_delivery_at || "",
      delivery_address: deliveryAddress,
      order_items: orderItems,
      order_value: normalizeOrderValue(orderValue),
      payment_status: normalizedStatus === "completed" ? paymentStatus : item.payment_status || "unpaid",
      payment_label:
        normalizedStatus === "completed"
          ? paymentStatus === "paid"
            ? "Đã thanh toán"
            : "Chưa thanh toán"
          : item.payment_label || "",
      payment_method: normalizedStatus === "completed" ? paymentMethod : item.payment_method || "",
      payment_method_label: normalizedStatus === "completed" ? labelPaymentMethod(paymentMethod) : item.payment_method_label || "",
      completion_note: normalizedStatus === "completed" ? note : item.completion_note || "",
      note: normalizedStatus === "completed" ? item.note || "" : note,
      updated_at: formatNow(),
    };
  });
}

async function appendProductionClaimToOrder(filePath, { orderId, actor, mode = "all", claimedItems = [] }) {
  const normalizedOrderId = normalizeOrderId(orderId);
  return orderStore.updateOrder(normalizedOrderId, (item) => {
    let existingClaims = [];
    try {
      existingClaims = JSON.parse(item?.production_claims_json || "[]");
    } catch {
      existingClaims = [];
    }
    const nextClaim = {
      id: `production-claim-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      mode,
      claimed_at: formatNow(),
      user_id: String(actor?.id || "").trim(),
      user_name: actor?.name || actor?.username || actor?.id || "Nhân viên sản xuất",
      items: normalizeProductionClaimItems(claimedItems),
    };
    const nextClaims = [...(Array.isArray(existingClaims) ? existingClaims : []), nextClaim];
    const claimedByNames = [...new Set(nextClaims.map((entry) => String(entry?.user_name || "").trim()).filter(Boolean))];
    return {
      ...item,
      production_claims_json: JSON.stringify(nextClaims),
      production_claimed_by_names: claimedByNames.join(", "),
      updated_at: formatNow(),
    };
  });
}

async function deleteOrderRecord(filePath, orderId) {
  const normalizedOrderId = normalizeOrderId(orderId);
  return orderStore.deleteOrder(normalizedOrderId);
}

async function backfillOrderDetails(filePath, { orderId, deliveryAddress, orderItems, orderValue, note }) {
  const normalizedOrderId = normalizeOrderId(orderId);
  return orderStore.updateOrder(normalizedOrderId, (item) => {
    const nextItem = { ...item };
    let changed = false;

    if (!String(nextItem.delivery_address || "").trim() && deliveryAddress) {
      nextItem.delivery_address = deliveryAddress;
      changed = true;
    }
    if (!String(nextItem.order_items || "").trim() && orderItems) {
      nextItem.order_items = orderItems;
      changed = true;
    }
    if (!String(nextItem.order_value || "").trim() && orderValue) {
      nextItem.order_value = normalizeOrderValue(orderValue);
      changed = true;
    }
    if (!String(nextItem.note || nextItem.completion_note || "").trim() && note) {
      if (String(nextItem.status || "").trim().toLowerCase() === "completed") {
        nextItem.completion_note = note;
      } else {
        nextItem.note = note;
      }
      changed = true;
    }

    if (changed) {
      nextItem.updated_at = formatNow();
    }

    return nextItem;
  });
}

function normalizeDeliveryResultStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["delivered", "partial", "rescheduled", "failed"].includes(normalized)) {
    return normalized;
  }
  return "delivered";
}

function resolveTapeCalculatorConfigCacheFile(cwd) {
  return path.resolve(cwd, "./data/cache/tape-calculator-config.json");
}

async function readTapeCalculatorConfigCacheFile(filePath) {
  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || !parsed.payload || typeof parsed.payload !== "object") {
      return null;
    }
    return parsed.payload;
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return null;
    }
    throw error;
  }
}

async function writeTapeCalculatorConfigCacheFile(filePath, payload) {
  await mkdir(path.dirname(filePath), { recursive: true });
  await writeFile(filePath, `${JSON.stringify({ saved_at: formatNow(), payload }, null, 2)}\n`, "utf8");
}

async function fetchWithTimeout(url, timeoutMs) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      signal: controller.signal,
      headers: {
        "user-agent": "AI-Beta-TapeCalculator/1.0",
        accept: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
      },
    });
  } finally {
    clearTimeout(timeoutId);
  }
}

async function fetchTapeCalculatorWorkbookBuffer() {
  let lastError = null;
  for (let attempt = 1; attempt <= TAPE_CALCULATOR_FETCH_RETRY_COUNT; attempt += 1) {
    try {
      const response = await fetchWithTimeout(TAPE_CALCULATOR_SHEET_URL, TAPE_CALCULATOR_FETCH_TIMEOUT_MS);
      if (!response.ok) {
        throw new Error(`Khong the tai du lieu bang tinh tu Google Sheet (HTTP ${response.status}).`);
      }
      const arrayBuffer = await response.arrayBuffer();
      return Buffer.from(arrayBuffer);
    } catch (error) {
      lastError = error;
      if (attempt < TAPE_CALCULATOR_FETCH_RETRY_COUNT) {
        await new Promise((resolve) => setTimeout(resolve, 600 * attempt));
      }
    }
  }
  throw lastError || new Error("Khong the tai du lieu bang tinh tu Google Sheet.");
}

async function loadTapeCalculatorConfig() {
  if (tapeCalculatorConfigCache.payload && tapeCalculatorConfigCache.expiresAt > Date.now()) {
    return tapeCalculatorConfigCache.payload;
  }

  const cacheFilePath = resolveTapeCalculatorConfigCacheFile(config.cwd);
  const fallbackPayload = {
    products: DEFAULT_TAPE_CALCULATOR_PRODUCTS,
    cores: DEFAULT_TAPE_CALCULATOR_CORES,
    defaults: DEFAULT_TAPE_CALCULATOR_FORM,
    source: TAPE_CALCULATOR_SHEET_URL,
    fallback: true,
  };

  try {
    const workbookBuffer = await fetchTapeCalculatorWorkbookBuffer();
    const workbook = XLSX.read(workbookBuffer, { type: "buffer" });
    const getRows = (sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false }) : [];
    };

    const dlRows = getRows("DL");
    const machineRows = getRows("SX-MAYCAT");
    const productMap = new Map();
    const dlHeaderIndex = dlRows.findIndex((row) =>
      row.some((cell) => String(cell || "").trim().toLowerCase() === "mã hàng hóa" || String(cell || "").trim().toLowerCase() === "ma hang hoa"),
    );
    const dlHeaderRow = dlHeaderIndex >= 0 ? dlRows[dlHeaderIndex] : [];
    const symbolColumnIndex = dlHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "kí hiệu" || String(cell || "").trim().toLowerCase() === "ki hieu",
    );
    const nameColumnIndex = dlHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "tên hàng" || String(cell || "").trim().toLowerCase() === "ten hang",
    );
    const jumboHeightColumnIndex = dlHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "chiều cao (mm)" || String(cell || "").trim().toLowerCase() === "chieu cao (mm)",
    );

    const allowedProductCodes = new Set(TAPE_CALCULATOR_PRODUCT_CODES);
    for (const row of dlRows.slice(dlHeaderIndex + 1)) {
      const code = String(row?.[symbolColumnIndex] || "").trim().toUpperCase();
      const productName = String(row?.[nameColumnIndex] || "").trim();
      const jumboHeight = Number(row?.[jumboHeightColumnIndex] || 0);
      if (!code || !productName) {
        continue;
      }

      if (!allowedProductCodes.has(code)) {
        continue;
      }

      productMap.set(code, {
        code,
        jumbo_height: Number.isFinite(jumboHeight) && jumboHeight > 0 ? jumboHeight : 1260,
      });
    }

    const coreCodes = new Set();
    const coreMap = new Map();
    const addCoreCode = (value, widthValue = 0) => {
      const normalizedValue = String(value || "").trim().toUpperCase();
      const widthMm = Number(widthValue || 0);
      if (!/^(?:G|N|ND)\d{1,3}$/.test(normalizedValue)) {
        return;
      }

      coreCodes.add(normalizedValue);
      if (!coreMap.has(normalizedValue)) {
        coreMap.set(normalizedValue, {
          code: normalizedValue,
          width_mm: Number.isFinite(widthMm) && widthMm > 0 ? widthMm : Number(normalizedValue.replace(/^[A-Z]+/, "")) || 0,
        });
      }
    };

    const coreHeaderIndex = dlRows.findIndex((row, index) => index > dlHeaderIndex && row.some((cell) => String(cell || "").trim().toLowerCase() === "khổ (mm)"));
    const coreHeaderRow = coreHeaderIndex >= 0 ? dlRows[coreHeaderIndex] : [];
    const coreCodeColumnIndex = coreHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "kí hiệu" || String(cell || "").trim().toLowerCase() === "ki hieu",
    );
    const coreWidthColumnIndex = coreHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "khổ (mm)" || String(cell || "").trim().toLowerCase() === "kho (mm)",
    );
    const coreNameColumnIndex = coreHeaderRow.findIndex(
      (cell) => String(cell || "").trim().toLowerCase() === "tên hàng" || String(cell || "").trim().toLowerCase() === "ten hang",
    );

    for (const row of dlRows.slice(coreHeaderIndex + 1)) {
      const code = String(row?.[coreCodeColumnIndex] || "").trim().toUpperCase();
      const name = String(row?.[coreNameColumnIndex] || "").trim().toLowerCase();
      if (!code || !name.includes("lõi") && !name.includes("loi")) {
        continue;
      }
      addCoreCode(code, row?.[coreWidthColumnIndex]);
    }

    const resolvedProducts = productMap.size ? [...productMap.values()] : DEFAULT_TAPE_CALCULATOR_PRODUCTS;
    const productOrderMap = new Map(TAPE_CALCULATOR_PRODUCT_CODES.map((code, index) => [code, index]));
    const products = [...resolvedProducts].sort(
      (left, right) => (productOrderMap.get(left.code) ?? Number.MAX_SAFE_INTEGER) - (productOrderMap.get(right.code) ?? Number.MAX_SAFE_INTEGER),
    );
    const defaultCoreMap = new Map(DEFAULT_TAPE_CALCULATOR_CORES.map((item) => [item.code, item]));
    const cores = [...(coreMap.size ? coreMap.values() : DEFAULT_TAPE_CALCULATOR_CORES)].sort((left, right) => {
      const leftOrder = defaultCoreMap.has(left.code) ? DEFAULT_TAPE_CALCULATOR_CORES.findIndex((item) => item.code === left.code) : Number.MAX_SAFE_INTEGER;
      const rightOrder = defaultCoreMap.has(right.code) ? DEFAULT_TAPE_CALCULATOR_CORES.findIndex((item) => item.code === right.code) : Number.MAX_SAFE_INTEGER;
      if (leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
      }
      return left.code.localeCompare(right.code, "vi", { numeric: true, sensitivity: "base" });
    });

    const defaultsRow =
      machineRows.find((row) => String(row?.[0] || "").trim().toUpperCase() === "BOPP60") ||
      machineRows.find((row) => String(row?.[0] || "").trim() && /^[A-Z0-9]+$/i.test(String(row?.[0] || "").trim()));
    const defaultProduct = products.find((item) => item.code === String(defaultsRow?.[0] || "").trim().toUpperCase()) || products[0];
    const defaultCore = cores.find((item) => item.code === String(defaultsRow?.[2] || "").trim().toUpperCase()) || cores[0];
    const payload = {
      products,
      cores,
      defaults: {
        tape_type: defaultProduct?.code || DEFAULT_TAPE_CALCULATOR_FORM.tape_type,
        order_quantity: Number(defaultsRow?.[1] || DEFAULT_TAPE_CALCULATOR_FORM.order_quantity) || DEFAULT_TAPE_CALCULATOR_FORM.order_quantity,
        core_type: defaultCore?.code || DEFAULT_TAPE_CALCULATOR_FORM.core_type,
        packaging: Number(defaultsRow?.[3] || DEFAULT_TAPE_CALCULATOR_FORM.packaging) || DEFAULT_TAPE_CALCULATOR_FORM.packaging,
        finished_quantity:
          Number(defaultsRow?.[4] || DEFAULT_TAPE_CALCULATOR_FORM.finished_quantity) || DEFAULT_TAPE_CALCULATOR_FORM.finished_quantity,
        jumbo_height:
          Number(defaultsRow?.[5] || defaultProduct?.jumbo_height || DEFAULT_TAPE_CALCULATOR_FORM.jumbo_height) ||
          DEFAULT_TAPE_CALCULATOR_FORM.jumbo_height,
        core_width:
          Number(defaultsRow?.[6] || defaultCore?.width_mm || DEFAULT_TAPE_CALCULATOR_FORM.core_width) ||
          DEFAULT_TAPE_CALCULATOR_FORM.core_width,
      },
      source: TAPE_CALCULATOR_SHEET_URL,
      fallback: false,
    };

    await writeTapeCalculatorConfigCacheFile(cacheFilePath, payload);

    tapeCalculatorConfigCache = {
      payload,
      expiresAt: Date.now() + 10 * 60_000,
    };
    return payload;
  } catch (error) {
    console.error("Khong the nap config bang tinh bang dinh tu Google Sheet:", error);
    try {
      const cachedPayload = await readTapeCalculatorConfigCacheFile(cacheFilePath);
      if (cachedPayload) {
        const nextPayload = {
          ...cachedPayload,
          fallback: true,
          source: `${TAPE_CALCULATOR_SHEET_URL} (cache local)`,
        };
        tapeCalculatorConfigCache = {
          payload: nextPayload,
          expiresAt: Date.now() + 5 * 60_000,
        };
        return nextPayload;
      }
    } catch (cacheError) {
      console.error("Khong the doc cache local bang tinh bang dinh:", cacheError);
    }
    if (tapeCalculatorConfigCache.payload) {
      return {
        ...tapeCalculatorConfigCache.payload,
        fallback: true,
      };
    }

    tapeCalculatorConfigCache = {
      payload: fallbackPayload,
      expiresAt: Date.now() + 60_000,
    };
    return fallbackPayload;
  }
}

function normalizeOrderId(value) {
  const core = String(value || "")
    .trim()
    .replace(/^dh[-\s]*/i, "")
    .replace(/[^0-9a-z-]/gi, "")
    .trim()
    .toUpperCase();
  return core ? `DH-${core}` : "";
}

function extractNumericOrderSuffix(orderId) {
  const normalized = normalizeOrderId(orderId);
  const match = normalized.match(/^DH-(\d+)$/i);
  return match ? Number.parseInt(match[1], 10) : Number.NaN;
}

function formatSequentialOrderId(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return "DH-003000";
  }
  return `DH-${String(Math.trunc(numeric)).padStart(6, "0")}`;
}

function isDuplicateOrderConstraintError(error) {
  return (
    error instanceof Error &&
    typeof error.message === "string" &&
    error.message.toLowerCase().includes("unique") &&
    error.message.toLowerCase().includes("orders.order_id")
  );
}

async function allocateOrderId(preferredOrderId, options = {}) {
  const orderKind = String(options.orderKind || "").trim().toLowerCase();
  if (orderKind !== "production") {
    return normalizeOrderId(preferredOrderId);
  }

  const existingOrders = await readOrders(resolveOrdersFile(config.cwd));
  const existingIds = new Set(
    existingOrders.map((item) => normalizeOrderId(item?.order_id || "")).filter(Boolean),
  );
  const numericSuffixes = existingOrders
    .map((item) => extractNumericOrderSuffix(item?.order_id || ""))
    .filter((value) => Number.isFinite(value));

  let nextValue = extractNumericOrderSuffix(preferredOrderId);
  if (!Number.isFinite(nextValue)) {
    nextValue = Math.max(3000, ...(numericSuffixes.length ? numericSuffixes : [3000]));
  }

  let candidate = formatSequentialOrderId(nextValue);
  while (existingIds.has(candidate)) {
    nextValue += 1;
    candidate = formatSequentialOrderId(nextValue);
  }

  return candidate;
}

function normalizeDeliveryCompletedAt(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return formatNow();
  }
  const parsed = new Date(raw);
  return Number.isFinite(parsed.getTime()) ? parsed.toISOString() : formatNow();
}

function normalizeOptionalDateTime(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  const slashMatch = raw.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year = new Date().getFullYear();
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return new Date(Date.UTC(year, month - 1, day)).toISOString();
    }
  }
  const parsed = new Date(raw);
  return Number.isFinite(parsed.getTime()) ? parsed.toISOString() : "";
}

function normalizePaymentStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "paid" ? "paid" : "unpaid";
}

function normalizePaymentMethod(value, paymentStatus = "unpaid") {
  if (String(paymentStatus || "").trim().toLowerCase() !== "paid") {
    return "";
  }
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "cash" || normalized === "bank_transfer") {
    return normalized;
  }
  return "";
}

function normalizeOrderValue(value) {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return "";
  }
  const numeric = Number(raw.replace(/[^\d.-]/g, ""));
  if (!Number.isFinite(numeric) || numeric < 0) {
    return "";
  }
  return String(Math.round(numeric));
}

function buildOrderRecord({
  currentUser,
  salesUser,
  deliveryUser,
  orderId,
  customerName,
  plannedDeliveryAt,
  deliveryAddress,
  orderItems,
  orderValue,
  note,
  orderKind = "transport",
}) {
  return {
    id: `order-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
    order_kind: orderKind,
    order_id: orderId,
    customer_name: customerName,
    sales_user_id: salesUser.id,
    sales_user_name: salesUser.name || salesUser.username || salesUser.id,
    delivery_user_id: deliveryUser?.id || "",
    delivery_user_name: deliveryUser?.name || deliveryUser?.username || deliveryUser?.id || "",
    created_by_user_id: String(currentUser?.id || "").trim(),
    created_by_user_name: currentUser?.name || currentUser?.username || "Hệ thống",
    planned_delivery_at: plannedDeliveryAt,
    delivery_address: deliveryAddress,
    order_items: orderItems,
    order_value: normalizeOrderValue(orderValue),
    note,
    status: "assigned",
    payment_status: "unpaid",
    payment_method: "",
    created_at: formatNow(),
    updated_at: formatNow(),
    document_title: `Phiếu điều phối đơn ${orderId}`,
    document_path: "",
  };
}

function buildOrderAssignmentDocument(record) {
  const lines = [
    `# ${record.document_title}`,
    "",
    `- Mã đơn hàng: ${record.order_id}`,
    `- Khách hàng: ${record.customer_name}`,
    `- NVKD phụ trách: ${record.sales_user_name}`,
    `- Nhân viên giao hàng: ${record.delivery_user_name}`,
    `- Trạng thái: Đã giao việc`,
    `- Thời điểm tạo: ${formatDeliveryTimestamp(record.created_at)}`,
  ];

  if (record.planned_delivery_at) {
    lines.push(`- Thời gian cần giao: ${formatDeliveryTimestamp(record.planned_delivery_at)}`);
  }
  if (record.delivery_address) {
    lines.push(`- Địa chỉ giao hàng: ${record.delivery_address}`);
  }
  if (record.order_items) {
    lines.push(`- Hàng giao: ${record.order_items}`);
  }
  if (record.order_value) {
    lines.push(`- Giá trị đơn hàng: ${record.order_value} VND`);
  }
  if (record.note) {
    lines.push(`- Ghi chú: ${record.note}`);
  }

  lines.push("", "## Điều phối", "", `Đơn ${record.order_id} đã được giao cho ${record.delivery_user_name} xử lý.`);
  return lines.join("\n");
}

function buildProductionOrderDocument(record) {
  const lines = [
    `# ${record.document_title}`,
    "",
    `- Ma don hang: ${record.order_id}`,
    `- Khach hang: ${record.customer_name}`,
    `- Nguoi yeu cau: ${record.sales_user_name}`,
    `- Trang thai: Cho san xuat`,
    `- Thoi diem tao: ${formatDeliveryTimestamp(record.created_at)}`,
  ];

  if (record.order_items) {
    lines.push(`- Chi tiet san xuat: ${record.order_items}`);
  }
  if (record.note) {
    lines.push(`- Ghi chu: ${record.note}`);
  }

  lines.push("", "## Theo doi", "", `Phieu san xuat ${record.order_id} da duoc gui toi bo phan san xuat de xu ly.`);
  return lines.join("\n");
}

function buildOrderAssignmentMail(record, deliveryUser, salesUser) {
  return [
    `Xin chào ${deliveryUser.name || deliveryUser.username || "bạn"},`,
    "",
    `Bạn vừa được giao xử lý đơn hàng ${record.order_id}.`,
    `Khách hàng: ${record.customer_name}`,
    `NVKD phụ trách: ${salesUser.name || salesUser.username || salesUser.id}`,
    record.planned_delivery_at ? `Thời gian cần giao: ${formatDeliveryTimestamp(record.planned_delivery_at)}` : "",
    record.delivery_address ? `Địa chỉ giao hàng: ${record.delivery_address}` : "",
    record.note ? `Ghi chú: ${record.note}` : "",
    "",
    "Biên bản điều phối đã được ghi vào kho dữ liệu nội bộ để tra cứu.",
  ]
    .filter(Boolean)
    .join("\n");
}

function buildDeliveryCompletionRecord({
  currentUser,
  salesUser,
  orderId,
  customerName,
  resultStatus,
  completedAt,
  paymentStatus,
  paymentMethod,
  note,
}) {
  return {
    id: `delivery-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
    order_id: orderId,
    customer_name: customerName,
    sales_user_id: salesUser.id,
    sales_user_name: salesUser.name || salesUser.username || salesUser.id,
    sales_department: salesUser.department || "",
    driver_user_id: String(currentUser?.id || "").trim(),
    driver_user_name: currentUser?.name || currentUser?.username || "Nhân viên giao hàng",
    driver_department: currentUser?.department || "",
    result_status: resultStatus,
    result_label: labelDeliveryResultStatus(resultStatus),
    payment_status: paymentStatus,
    payment_label: paymentStatus === "paid" ? "Đã thanh toán" : "Chưa thanh toán",
    payment_method: paymentMethod,
    payment_method_label: labelPaymentMethod(paymentMethod),
    completed_at: completedAt,
    note,
    created_at: formatNow(),
    document_title: `Biên bản giao đơn ${orderId}`,
  };
}

function buildDeliveryCompletionDocument(record) {
  const lines = [
    `# ${record.document_title}`,
    "",
    `- Mã đơn hàng: ${record.order_id}`,
    `- Khách hàng: ${record.customer_name}`,
    `- Kết quả giao: ${record.result_label}`,
    `- Thời gian hoàn thành: ${formatDeliveryTimestamp(record.completed_at)}`,
    `- Nhân viên giao hàng: ${record.driver_user_name}`,
    `- Nhân viên kinh doanh phụ trách: ${record.sales_user_name}`,
  ];

  if (record.note) {
    lines.push(`- Ghi chú: ${record.note}`);
  }

  lines.push("", "## Tóm tắt", "", `${record.result_label} cho đơn ${record.order_id} của khách hàng ${record.customer_name}.`);
  lines.push("", "## Thanh toán", "", `- Trạng thái: ${record.payment_label}`);
  if (record.payment_method) {
    lines.push(`- Hình thức: ${record.payment_method_label}`);
  }
  if (record.note) {
    lines.push("", "## Ghi chú giao hàng", "", record.note);
  }
  return lines.join("\n");
}

function buildDeliveryNotificationMail(record, salesUser, actor) {
  return [
    `Xin chào ${salesUser.name || salesUser.username || "bạn"},`,
    "",
    `Đơn hàng ${record.order_id} đã được xác nhận hoàn thành giao hàng.`,
    `Khách hàng: ${record.customer_name}`,
    `Kết quả: ${record.result_label}`,
    `Thời gian hoàn thành: ${formatDeliveryTimestamp(record.completed_at)}`,
    `Người cập nhật: ${actor?.name || actor?.username || "Hệ thống"}`,
    record.note ? `Ghi chú: ${record.note}` : "",
    "",
    "Thông tin này đã được ghi vào kho dữ liệu nội bộ để tra cứu lại khi cần.",
  ]
    .filter(Boolean)
    .join("\n");
}

function formatDeliveryTimestamp(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return value || "-";
  }
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

function labelDeliveryResultStatus(value) {
  if (value === "partial") return "Giao một phần";
  if (value === "rescheduled") return "Khách hẹn lại";
  if (value === "failed") return "Không giao được";
  return "Giao thành công";
}

function labelPaymentMethod(value) {
  if (value === "cash") return "Tiền mặt";
  if (value === "bank_transfer") return "Chuyển khoản";
  return "Chưa rõ";
}

function slugifyValue(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "delivery";
}

function normalizeAskFilters(currentUser, value = {}) {
  const policy = getEffectivePolicy(currentUser);
  const rawAccessLevel = String(value.access_level || "").trim().toLowerCase();
  const levels = ["basic", "advanced", "sensitive"];
  const requestedLevel = levels.includes(rawAccessLevel) ? rawAccessLevel : "";
  const maxIndex = levels.indexOf(String(policy?.max_access_level || "basic").trim().toLowerCase());
  const requestedIndex = levels.indexOf(requestedLevel);
  const allowPersonal = requestedLevel === "sensitive" && Boolean(currentUser);
  return {
    access_level: requestedIndex >= 0 && (requestedIndex <= maxIndex || allowPersonal) ? requestedLevel : "",
    folder: normalizeStorageFolderPath(value.folder || ""),
  };
}

function matchesAskFilters(metadata = {}, filters = {}) {
  const accessLevel = String(filters.access_level || "").trim().toLowerCase();
  const folder = String(filters.folder || "").trim();
  if (accessLevel && String(metadata.access_level || "").trim().toLowerCase() !== accessLevel) {
    return false;
  }
  if (folder) {
    const storageFolder = normalizeStorageFolderPath(metadata.storage_folder || "");
    if (storageFolder !== folder && !storageFolder.startsWith(`${folder}/`)) {
      return false;
    }
  }
  return true;
}
