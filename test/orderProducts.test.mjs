import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { createOrderStore } from "../lib/orderStore.mjs";

test("giu du cac san pham cung ma neu ten khac nhau", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "order-products-dup-code-"));

  try {
    const store = await createOrderStore({
      dbFilePath: path.join(cwd, "orders.sqlite"),
      legacyOrdersFile: path.join(cwd, "orders.json"),
      legacyCatalogFile: path.join(cwd, "catalog.json"),
      legacyCompletionsFile: path.join(cwd, "completions.json"),
      legacyNotificationsFile: path.join(cwd, "notifications.json"),
      defaultProducts: () => [],
    });

    await store.saveProducts([
      {
        id: "product-a",
        code: "BOPP3",
        name: "Bang dinh 1kg",
        units: ["Cuon"],
        default_unit: "Cuon",
      },
      {
        id: "product-b",
        code: "BOPP3",
        name: "Bang dinh 500g",
        units: ["Cuon"],
        default_unit: "Cuon",
      },
    ]);

    const products = await store.readProducts();
    assert.equal(products.length, 2);
    assert.deepEqual(
      products.map((item) => item.code),
      ["BOPP3", "BOPP3"],
    );
  } finally {
    void cwd;
  }
});

test("bao loi neu ten hang hoa bi trung", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "order-products-dup-name-"));

  try {
    const store = await createOrderStore({
      dbFilePath: path.join(cwd, "orders.sqlite"),
      legacyOrdersFile: path.join(cwd, "orders.json"),
      legacyCatalogFile: path.join(cwd, "catalog.json"),
      legacyCompletionsFile: path.join(cwd, "completions.json"),
      legacyNotificationsFile: path.join(cwd, "notifications.json"),
      defaultProducts: () => [],
    });

    await assert.rejects(
      () =>
        store.saveProducts([
          {
            id: "product-a",
            code: "BOPP3",
            name: "Bang dinh",
            units: ["Cuon"],
            default_unit: "Cuon",
          },
          {
            id: "product-b",
            code: "BOPP4",
            name: "Bang dinh",
            units: ["Cuon"],
            default_unit: "Cuon",
          },
        ]),
      /Ten hang hoa khong duoc trung nhau/i,
    );
  } finally {
    void cwd;
  }
});
