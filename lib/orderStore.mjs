import { mkdir, readFile } from "node:fs/promises";
import path from "node:path";
import Database from "better-sqlite3";

function parseJsonArray(raw) {
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
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

function normalizeProductItems(items, defaultProducts) {
  const seen = new Set();
  const normalized = (Array.isArray(items) ? items : [])
    .map((item, index) => {
      const name = String(item?.name || "").trim();
      const code = String(item?.code || "").trim();
      const units = (Array.isArray(item?.units) ? item.units : [])
        .map((unit) => String(unit || "").trim())
        .filter(Boolean);
      const defaultUnit = String(item?.default_unit || units[0] || "").trim();
      if (!name) {
        return null;
      }

      const key = name.toLowerCase();
      if (seen.has(key)) {
        throw new Error("Ten hang hoa khong duoc trung nhau. Ma hang co the trung neu ten khac.");
      }
      seen.add(key);

      return {
        id: String(item?.id || `product-${index + 1}`).trim(),
        code,
        name,
        units: units.length ? units : [defaultUnit || "cây"],
        default_unit: defaultUnit || units[0] || "cây",
      };
    })
    .filter(Boolean);

  return normalized.length ? normalized : defaultProducts();
}

function parseOrderItemsSummary(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const match = normalized.match(/^(.*?)\s+x\s+(\d+)\s+(\S+)\s+•\s+([\d.,\s]+)\s*[đdĐ]?\s*=\s*([\d.,\s]+)\s*[đdĐ]?$/i);
      if (!match) {
        return {
          line_number: index + 1,
          product_name: normalized,
          quantity: "",
          unit: "",
          unit_price: "",
          line_total: "",
          raw_line: line,
        };
      }

      return {
        line_number: index + 1,
        product_name: match[1].trim(),
        quantity: match[2].trim(),
        unit: match[3].trim(),
        unit_price: normalizeOrderValue(match[4]),
        line_total: normalizeOrderValue(match[5]),
        raw_line: line,
      };
    });
}

function normalizeOrderRow(row) {
  return {
    ...row,
    planned_delivery_at: row?.planned_delivery_at || "",
    delivery_address: row?.delivery_address || "",
    order_items: row?.order_items || "",
    order_value: row?.order_value || "",
    note: row?.note || "",
    payment_status: row?.payment_status || "unpaid",
    payment_method: row?.payment_method || "",
    payment_label: row?.payment_label || "",
    payment_method_label: row?.payment_method_label || "",
    completion_note: row?.completion_note || "",
    completion_result_status: row?.completion_result_status || "",
    completion_result_label: row?.completion_result_label || "",
    completed_at: row?.completed_at || "",
    completed_by_user_id: row?.completed_by_user_id || "",
    completed_by_user_name: row?.completed_by_user_name || "",
    production_claims_json: row?.production_claims_json || "[]",
    production_claimed_by_names: row?.production_claimed_by_names || "",
    production_completed_by_json: row?.production_completed_by_json || "[]",
    production_packaged_by_json: row?.production_packaged_by_json || "",
    production_received_by_json: row?.production_received_by_json || "",
    document_title: row?.document_title || "",
    document_path: row?.document_path || "",
    completion_document_path: row?.completion_document_path || "",
  };
}

export async function createOrderStore({
  dbFilePath,
  legacyOrdersFile,
  legacyCatalogFile,
  legacyCompletionsFile,
  legacyNotificationsFile,
  defaultProducts,
}) {
  await mkdir(path.dirname(dbFilePath), { recursive: true });

  const db = new Database(dbFilePath);
  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS products (
      id TEXT PRIMARY KEY,
      code TEXT NOT NULL DEFAULT '',
      name TEXT NOT NULL UNIQUE,
      units_json TEXT NOT NULL,
      default_unit TEXT NOT NULL,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS orders (
      id TEXT PRIMARY KEY,
      order_id TEXT NOT NULL UNIQUE,
      customer_name TEXT NOT NULL DEFAULT '',
      sales_user_id TEXT NOT NULL DEFAULT '',
      sales_user_name TEXT NOT NULL DEFAULT '',
      delivery_user_id TEXT NOT NULL DEFAULT '',
      delivery_user_name TEXT NOT NULL DEFAULT '',
      created_by_user_id TEXT NOT NULL DEFAULT '',
      created_by_user_name TEXT NOT NULL DEFAULT '',
      planned_delivery_at TEXT NOT NULL DEFAULT '',
      delivery_address TEXT NOT NULL DEFAULT '',
      order_items TEXT NOT NULL DEFAULT '',
      order_value TEXT NOT NULL DEFAULT '',
      note TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'assigned',
      payment_status TEXT NOT NULL DEFAULT 'unpaid',
      payment_label TEXT NOT NULL DEFAULT '',
      payment_method TEXT NOT NULL DEFAULT '',
      payment_method_label TEXT NOT NULL DEFAULT '',
      completion_note TEXT NOT NULL DEFAULT '',
      completion_result_status TEXT NOT NULL DEFAULT '',
      completion_result_label TEXT NOT NULL DEFAULT '',
      completed_at TEXT NOT NULL DEFAULT '',
      completed_by_user_id TEXT NOT NULL DEFAULT '',
      completed_by_user_name TEXT NOT NULL DEFAULT '',
      production_claims_json TEXT NOT NULL DEFAULT '[]',
      production_claimed_by_names TEXT NOT NULL DEFAULT '',
      production_completed_by_json TEXT NOT NULL DEFAULT '[]',
      production_packaged_by_json TEXT NOT NULL DEFAULT '',
      production_received_by_json TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT '',
      document_title TEXT NOT NULL DEFAULT '',
      document_path TEXT NOT NULL DEFAULT '',
      completion_document_path TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS order_items (
      item_id INTEGER PRIMARY KEY AUTOINCREMENT,
      order_id TEXT NOT NULL,
      line_number INTEGER NOT NULL DEFAULT 1,
      product_name TEXT NOT NULL DEFAULT '',
      quantity TEXT NOT NULL DEFAULT '',
      unit TEXT NOT NULL DEFAULT '',
      unit_price TEXT NOT NULL DEFAULT '',
      line_total TEXT NOT NULL DEFAULT '',
      raw_line TEXT NOT NULL DEFAULT '',
      FOREIGN KEY (order_id) REFERENCES orders(order_id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS delivery_completions (
      id TEXT PRIMARY KEY,
      order_id TEXT NOT NULL DEFAULT '',
      customer_name TEXT NOT NULL DEFAULT '',
      sales_user_id TEXT NOT NULL DEFAULT '',
      sales_user_name TEXT NOT NULL DEFAULT '',
      sales_department TEXT NOT NULL DEFAULT '',
      driver_user_id TEXT NOT NULL DEFAULT '',
      driver_user_name TEXT NOT NULL DEFAULT '',
      driver_department TEXT NOT NULL DEFAULT '',
      result_status TEXT NOT NULL DEFAULT '',
      result_label TEXT NOT NULL DEFAULT '',
      payment_status TEXT NOT NULL DEFAULT 'unpaid',
      payment_label TEXT NOT NULL DEFAULT '',
      payment_method TEXT NOT NULL DEFAULT '',
      payment_method_label TEXT NOT NULL DEFAULT '',
      completed_at TEXT NOT NULL DEFAULT '',
      note TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      document_title TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS notifications (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL DEFAULT '',
      type TEXT NOT NULL DEFAULT '',
      title TEXT NOT NULL DEFAULT '',
      message TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT '',
      read_at TEXT NOT NULL DEFAULT '',
      meta_json TEXT NOT NULL DEFAULT '{}'
    );
  `);
  try {
    db.exec(`
      ALTER TABLE products ADD COLUMN code TEXT NOT NULL DEFAULT '';
    `);
  } catch {}
  try {
    db.exec(`
      ALTER TABLE orders ADD COLUMN production_claims_json TEXT NOT NULL DEFAULT '[]';
    `);
  } catch {}
  try {
    db.exec(`
      ALTER TABLE orders ADD COLUMN production_claimed_by_names TEXT NOT NULL DEFAULT '';
    `);
  } catch {}
  try {
    db.exec(`
      ALTER TABLE orders ADD COLUMN production_completed_by_json TEXT NOT NULL DEFAULT '[]';
    `);
  } catch {}
  try {
    db.exec(`
      ALTER TABLE orders ADD COLUMN production_packaged_by_json TEXT NOT NULL DEFAULT '';
    `);
  } catch {}
  try {
    db.exec(`
      ALTER TABLE orders ADD COLUMN production_received_by_json TEXT NOT NULL DEFAULT '';
    `);
  } catch {}

  const replaceProductStmt = db.prepare(`
    INSERT INTO products (id, code, name, units_json, default_unit, sort_order, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      code = excluded.code,
      name = excluded.name,
      units_json = excluded.units_json,
      default_unit = excluded.default_unit,
      sort_order = excluded.sort_order,
      updated_at = excluded.updated_at
  `);

  const replaceOrderStmt = db.prepare(`
    INSERT INTO orders (
      id, order_id, customer_name, sales_user_id, sales_user_name, delivery_user_id, delivery_user_name,
      created_by_user_id, created_by_user_name, planned_delivery_at, delivery_address, order_items, order_value,
      note, status, payment_status, payment_label, payment_method, payment_method_label, completion_note,
      completion_result_status, completion_result_label, completed_at, completed_by_user_id, completed_by_user_name,
      production_claims_json, production_claimed_by_names, production_completed_by_json, production_packaged_by_json,
      production_received_by_json, created_at, updated_at, document_title, document_path, completion_document_path
    ) VALUES (
      ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
    )
    ON CONFLICT(order_id) DO UPDATE SET
      id = excluded.id,
      customer_name = excluded.customer_name,
      sales_user_id = excluded.sales_user_id,
      sales_user_name = excluded.sales_user_name,
      delivery_user_id = excluded.delivery_user_id,
      delivery_user_name = excluded.delivery_user_name,
      created_by_user_id = excluded.created_by_user_id,
      created_by_user_name = excluded.created_by_user_name,
      planned_delivery_at = excluded.planned_delivery_at,
      delivery_address = excluded.delivery_address,
      order_items = excluded.order_items,
      order_value = excluded.order_value,
      note = excluded.note,
      status = excluded.status,
      payment_status = excluded.payment_status,
      payment_label = excluded.payment_label,
      payment_method = excluded.payment_method,
      payment_method_label = excluded.payment_method_label,
      completion_note = excluded.completion_note,
      completion_result_status = excluded.completion_result_status,
      completion_result_label = excluded.completion_result_label,
      completed_at = excluded.completed_at,
      completed_by_user_id = excluded.completed_by_user_id,
      completed_by_user_name = excluded.completed_by_user_name,
      production_claims_json = excluded.production_claims_json,
      production_claimed_by_names = excluded.production_claimed_by_names,
      production_completed_by_json = excluded.production_completed_by_json,
      production_packaged_by_json = excluded.production_packaged_by_json,
      production_received_by_json = excluded.production_received_by_json,
      created_at = excluded.created_at,
      updated_at = excluded.updated_at,
      document_title = excluded.document_title,
      document_path = excluded.document_path,
      completion_document_path = excluded.completion_document_path
  `);
  const insertOrderStmt = db.prepare(`
    INSERT INTO orders (
      id, order_id, customer_name, sales_user_id, sales_user_name, delivery_user_id, delivery_user_name,
      created_by_user_id, created_by_user_name, planned_delivery_at, delivery_address, order_items, order_value,
      note, status, payment_status, payment_label, payment_method, payment_method_label, completion_note,
      completion_result_status, completion_result_label, completed_at, completed_by_user_id, completed_by_user_name,
      production_claims_json, production_claimed_by_names, production_completed_by_json, production_packaged_by_json,
      production_received_by_json, created_at, updated_at, document_title, document_path, completion_document_path
    ) VALUES (
      ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
    )
  `);

  const deleteOrderItemsStmt = db.prepare("DELETE FROM order_items WHERE order_id = ?");
  const insertOrderItemStmt = db.prepare(`
    INSERT INTO order_items (order_id, line_number, product_name, quantity, unit, unit_price, line_total, raw_line)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const replaceDeliveryCompletionStmt = db.prepare(`
    INSERT INTO delivery_completions (
      id, order_id, customer_name, sales_user_id, sales_user_name, sales_department,
      driver_user_id, driver_user_name, driver_department, result_status, result_label,
      payment_status, payment_label, payment_method, payment_method_label, completed_at,
      note, created_at, document_title
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      order_id = excluded.order_id,
      customer_name = excluded.customer_name,
      sales_user_id = excluded.sales_user_id,
      sales_user_name = excluded.sales_user_name,
      sales_department = excluded.sales_department,
      driver_user_id = excluded.driver_user_id,
      driver_user_name = excluded.driver_user_name,
      driver_department = excluded.driver_department,
      result_status = excluded.result_status,
      result_label = excluded.result_label,
      payment_status = excluded.payment_status,
      payment_label = excluded.payment_label,
      payment_method = excluded.payment_method,
      payment_method_label = excluded.payment_method_label,
      completed_at = excluded.completed_at,
      note = excluded.note,
      created_at = excluded.created_at,
      document_title = excluded.document_title
  `);
  const replaceNotificationStmt = db.prepare(`
    INSERT INTO notifications (id, user_id, type, title, message, created_at, read_at, meta_json)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      user_id = excluded.user_id,
      type = excluded.type,
      title = excluded.title,
      message = excluded.message,
      created_at = excluded.created_at,
      read_at = excluded.read_at,
      meta_json = excluded.meta_json
  `);

  function replaceOrderItems(orderId, summaryValue = "") {
    deleteOrderItemsStmt.run(orderId);
    const items = parseOrderItemsSummary(summaryValue);
    for (const item of items) {
      insertOrderItemStmt.run(
        orderId,
        item.line_number,
        item.product_name,
        item.quantity,
        item.unit,
        item.unit_price,
        item.line_total,
        item.raw_line,
      );
    }
  }

  function storeOrderRecord(record) {
    const row = normalizeOrderRow(record);
    replaceOrderStmt.run(
      row.id || `order-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      row.order_id,
      row.customer_name,
      row.sales_user_id,
      row.sales_user_name,
      row.delivery_user_id,
      row.delivery_user_name,
      row.created_by_user_id,
      row.created_by_user_name,
      row.planned_delivery_at,
      row.delivery_address,
      row.order_items,
      normalizeOrderValue(row.order_value),
      row.note,
      row.status || "assigned",
      row.payment_status || "unpaid",
      row.payment_label || "",
      row.payment_method || "",
      row.payment_method_label || "",
      row.completion_note || "",
      row.completion_result_status || "",
      row.completion_result_label || "",
      row.completed_at || "",
      row.completed_by_user_id || "",
      row.completed_by_user_name || "",
      row.production_claims_json || "[]",
      row.production_claimed_by_names || "",
      row.production_completed_by_json || "[]",
      row.production_packaged_by_json || "",
      row.production_received_by_json || "",
      row.created_at || "",
      row.updated_at || row.created_at || "",
      row.document_title || "",
      row.document_path || "",
      row.completion_document_path || "",
    );
    replaceOrderItems(row.order_id, row.order_items);
  }

  function createOrderRecord(record) {
    const row = normalizeOrderRow(record);
    insertOrderStmt.run(
      row.id || `order-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      row.order_id,
      row.customer_name,
      row.sales_user_id,
      row.sales_user_name,
      row.delivery_user_id,
      row.delivery_user_name,
      row.created_by_user_id,
      row.created_by_user_name,
      row.planned_delivery_at,
      row.delivery_address,
      row.order_items,
      normalizeOrderValue(row.order_value),
      row.note,
      row.status || "assigned",
      row.payment_status || "unpaid",
      row.payment_label || "",
      row.payment_method || "",
      row.payment_method_label || "",
      row.completion_note || "",
      row.completion_result_status || "",
      row.completion_result_label || "",
      row.completed_at || "",
      row.completed_by_user_id || "",
      row.completed_by_user_name || "",
      row.production_claims_json || "[]",
      row.production_claimed_by_names || "",
      row.production_completed_by_json || "[]",
      row.production_packaged_by_json || "",
      row.production_received_by_json || "",
      row.created_at || "",
      row.updated_at || row.created_at || "",
      row.document_title || "",
      row.document_path || "",
      row.completion_document_path || "",
    );
    replaceOrderItems(row.order_id, row.order_items);
    return row;
  }

  function storeDeliveryCompletion(record) {
    replaceDeliveryCompletionStmt.run(
      String(record?.id || `delivery-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`).trim(),
      String(record?.order_id || "").trim(),
      String(record?.customer_name || "").trim(),
      String(record?.sales_user_id || "").trim(),
      String(record?.sales_user_name || "").trim(),
      String(record?.sales_department || "").trim(),
      String(record?.driver_user_id || "").trim(),
      String(record?.driver_user_name || "").trim(),
      String(record?.driver_department || "").trim(),
      String(record?.result_status || "").trim(),
      String(record?.result_label || "").trim(),
      String(record?.payment_status || "unpaid").trim(),
      String(record?.payment_label || "").trim(),
      String(record?.payment_method || "").trim(),
      String(record?.payment_method_label || "").trim(),
      String(record?.completed_at || "").trim(),
      String(record?.note || "").trim(),
      String(record?.created_at || "").trim(),
      String(record?.document_title || "").trim(),
    );
  }

  function normalizeNotificationRow(row) {
    return {
      id: String(row?.id || "").trim(),
      user_id: String(row?.user_id || "").trim(),
      type: String(row?.type || "").trim(),
      title: String(row?.title || "").trim(),
      message: String(row?.message || "").trim(),
      created_at: String(row?.created_at || "").trim(),
      read_at: String(row?.read_at || "").trim(),
      meta: (() => {
        try {
          return typeof row?.meta_json === "string" ? JSON.parse(row.meta_json || "{}") : row?.meta || {};
        } catch {
          return {};
        }
      })(),
    };
  }

  function storeNotification(record) {
    replaceNotificationStmt.run(
      String(record?.id || `notification-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`).trim(),
      String(record?.user_id || "").trim(),
      String(record?.type || "").trim(),
      String(record?.title || "").trim(),
      String(record?.message || "").trim(),
      String(record?.created_at || "").trim(),
      String(record?.read_at || "").trim(),
      JSON.stringify(record?.meta || {}),
    );
  }

  async function bootstrap() {
    const productCount = db.prepare("SELECT COUNT(*) AS count FROM products").get().count || 0;
    if (!productCount) {
      let products = [];
      try {
        products = normalizeProductItems(await readFile(legacyCatalogFile, "utf8").then(parseJsonArray), defaultProducts);
      } catch {
        products = defaultProducts();
      }

      const now = new Date().toISOString();
      db.exec("BEGIN");
      try {
        db.exec("DELETE FROM products");
        products.forEach((item, index) => {
          replaceProductStmt.run(
            item.id,
            item.code || "",
            item.name,
            JSON.stringify(item.units || []),
            item.default_unit,
            index,
            now,
          );
        });
        db.exec("COMMIT");
      } catch (error) {
        db.exec("ROLLBACK");
        throw error;
      }
    }

    const orderCount = db.prepare("SELECT COUNT(*) AS count FROM orders").get().count || 0;
    if (!orderCount) {
      let legacyOrders = [];
      try {
        legacyOrders = await readFile(legacyOrdersFile, "utf8").then(parseJsonArray);
      } catch {
        legacyOrders = [];
      }

      db.exec("BEGIN");
      try {
        legacyOrders.forEach((item) => {
          if (!String(item?.order_id || "").trim()) {
            return;
          }
          storeOrderRecord(item);
        });
        db.exec("COMMIT");
      } catch (error) {
        db.exec("ROLLBACK");
        throw error;
      }
    }

    const completionCount = db.prepare("SELECT COUNT(*) AS count FROM delivery_completions").get().count || 0;
    if (!completionCount && legacyCompletionsFile) {
      let legacyCompletions = [];
      try {
        legacyCompletions = await readFile(legacyCompletionsFile, "utf8").then(parseJsonArray);
      } catch {
        legacyCompletions = [];
      }

      db.exec("BEGIN");
      try {
        legacyCompletions.forEach((item) => {
          if (!String(item?.id || item?.order_id || "").trim()) {
            return;
          }
          storeDeliveryCompletion(item);
        });
        db.exec("COMMIT");
      } catch (error) {
        db.exec("ROLLBACK");
        throw error;
      }
    }

    const notificationCount = db.prepare("SELECT COUNT(*) AS count FROM notifications").get().count || 0;
    if (!notificationCount && legacyNotificationsFile) {
      let legacyNotifications = [];
      try {
        legacyNotifications = await readFile(legacyNotificationsFile, "utf8").then(parseJsonArray);
      } catch {
        legacyNotifications = [];
      }

      db.exec("BEGIN");
      try {
        legacyNotifications.forEach((item) => {
          if (!String(item?.id || item?.user_id || "").trim()) {
            return;
          }
          storeNotification(item);
        });
        db.exec("COMMIT");
      } catch (error) {
        db.exec("ROLLBACK");
        throw error;
      }
    }
  }

  await bootstrap();

  return {
    async readProducts() {
      const rows = db.prepare("SELECT id, code, name, units_json, default_unit FROM products ORDER BY sort_order ASC, name COLLATE NOCASE ASC").all();
      return rows.map((row) => ({
        id: row.id,
        code: row.code || "",
        name: row.name,
        units: parseJsonArray(row.units_json),
        default_unit: row.default_unit,
      }));
    },

    async saveProducts(items) {
      const products = normalizeProductItems(items, defaultProducts);
      const now = new Date().toISOString();
      db.exec("BEGIN");
      try {
        db.exec("DELETE FROM products");
        products.forEach((item, index) => {
          replaceProductStmt.run(
            item.id,
            item.code || "",
            item.name,
            JSON.stringify(item.units || []),
            item.default_unit,
            index,
            now,
          );
        });
        db.exec("COMMIT");
      } catch (error) {
        db.exec("ROLLBACK");
        throw error;
      }
      return products;
    },

    async readOrders() {
      const rows = db
        .prepare("SELECT * FROM orders ORDER BY COALESCE(updated_at, created_at) DESC, order_id DESC")
        .all();
      return rows.map(normalizeOrderRow);
    },

    async appendOrder(record, options = {}) {
      if (options.allowReplace === false) {
        return createOrderRecord(record);
      }
      storeOrderRecord(record);
    },

    async updateOrder(orderId, updater) {
      const existing = db.prepare("SELECT * FROM orders WHERE order_id = ?").get(orderId);
      if (!existing) {
        throw new Error("Khong tim thay don hang de cap nhat.");
      }

      const nextRecord = normalizeOrderRow(typeof updater === "function" ? updater(normalizeOrderRow(existing)) : { ...existing, ...updater });
      storeOrderRecord(nextRecord);
      return this.readOrders();
    },

    async deleteOrder(orderId) {
      const result = db.prepare("DELETE FROM orders WHERE order_id = ?").run(orderId);
      if (!result.changes) {
        throw new Error("Khong tim thay don hang de xoa.");
      }
      return this.readOrders();
    },

    async appendDeliveryCompletion(record) {
      storeDeliveryCompletion(record);
    },

    async readDeliveryCompletions() {
      return db
        .prepare("SELECT * FROM delivery_completions ORDER BY COALESCE(completed_at, created_at) DESC, id DESC")
        .all();
    },

    async appendNotification(record) {
      storeNotification(record);
    },

    async readNotifications(currentUser = null, options = {}) {
      const rows =
        options.includeAll === true
          ? db.prepare("SELECT * FROM notifications ORDER BY created_at DESC, id DESC").all()
          : db
              .prepare("SELECT * FROM notifications WHERE user_id = ? ORDER BY created_at DESC, id DESC")
              .all(String(currentUser?.id || "").trim());
      return rows.map(normalizeNotificationRow);
    },

    async markNotificationRead(currentUser, notificationId, nowIso) {
      const currentUserId = String(currentUser?.id || "").trim();
      db.prepare(
        "UPDATE notifications SET read_at = ? WHERE id = ? AND user_id = ? AND COALESCE(read_at, '') = ''",
      ).run(nowIso, notificationId, currentUserId);
      return this.readNotifications(currentUser);
    },

    async markAllNotificationsRead(currentUser, nowIso) {
      const currentUserId = String(currentUser?.id || "").trim();
      db.prepare(
        "UPDATE notifications SET read_at = ? WHERE user_id = ? AND COALESCE(read_at, '') = ''",
      ).run(nowIso, currentUserId);
      return this.readNotifications(currentUser);
    },

    async deleteNotification(currentUser, notificationId) {
      const currentUserId = String(currentUser?.id || "").trim();
      db.prepare("DELETE FROM notifications WHERE id = ? AND user_id = ?").run(notificationId, currentUserId);
      return this.readNotifications(currentUser);
    },

    async patchNotificationDocumentPath({ userId, type, orderId, documentPath }) {
      const targetUserId = String(userId || "").trim();
      const targetType = String(type || "").trim();
      const targetOrderId = String(orderId || "").trim().toLowerCase();
      const nextDocumentPath = String(documentPath || "").trim();
      if (!targetUserId || !targetType || !targetOrderId || !nextDocumentPath) {
        return;
      }

      const rows = db
        .prepare("SELECT * FROM notifications WHERE user_id = ? AND type = ? ORDER BY created_at DESC")
        .all(targetUserId, targetType);
      for (const row of rows) {
        let meta = {};
        try {
          meta = JSON.parse(row.meta_json || "{}");
        } catch {
          meta = {};
        }

        if (
          String(meta?.order_id || "").trim().toLowerCase() === targetOrderId &&
          !String(meta?.document_path || "").trim()
        ) {
          meta.document_path = nextDocumentPath;
          replaceNotificationStmt.run(
            row.id,
            row.user_id,
            row.type,
            row.title,
            row.message,
            row.created_at,
            row.read_at,
            JSON.stringify(meta),
          );
          break;
        }
      }
    },
  };
}
