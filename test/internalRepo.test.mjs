import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  createStorageFolder,
  deleteStorageFolder,
  deleteInternalDocument,
  deleteSyncSource,
  importUploadedDocument,
  importSheetDocumentWithOptions,
  listCustomStorageFolders,
  loadInternalRepository,
  readInternalDocument,
  reviewInternalDocument,
  reviewSyncSource,
  renameStorageFolder,
  saveInternalDocument,
  saveSyncSource,
  setSourceSyncable,
  listSyncSources,
  syncSourceById,
  syncSheetDocuments,
  updateInternalDocumentMetadata,
} from "../lib/internalRepo.mjs";

test("tu dong dong bo lai noi dung khi Google Sheet thay doi", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "sheet-sync-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sheetUrl = "https://docs.google.com/spreadsheets/d/test-sheet/edit#gid=0";
    let currentCsv = "Muc,Gia\nSan pham A,100";

    const fetchImpl = async () =>
      new Response(currentCsv, {
        status: 200,
        headers: {
          "content-type": "text/csv; charset=utf-8",
        },
      });

    const saved = await importSheetDocumentWithOptions(
      internalDir,
      cwd,
      {
        title: "Bang gia NPP",
        sheetUrl,
        effective_date: "2026-04-03",
        version: "1.0",
        owner: "Kinh doanh",
        status: "active",
        topic_key: "bang-gia-npp",
      },
      { fetchImpl },
    );

    currentCsv = "Muc,Gia\nSan pham A,150";

    const syncResult = await syncSheetDocuments(internalDir, cwd, { fetchImpl });
    const repo = await loadInternalRepository(internalDir, cwd);
    const savedPath = path.join(cwd, saved.path);
    const raw = await readFile(savedPath, "utf8");

    assert.equal(syncResult.checked, 1);
    assert.equal(syncResult.updated, 1);
    assert.equal(syncResult.errors.length, 0);
    assert.match(raw, /San pham A \| 150/);
    assert.match(raw, /synced_at:/);
    assert.match(repo.documents[0].content, /San pham A \| 150/);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("luu tai lieu vao dung thu muc da chon", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "manual-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const saved = await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu van chuyen test",
      content: "Noi dung thu nghiem luu theo thu muc.",
      effective_date: "2026-04-04",
      owner: "operations",
      owner_department: "operations",
      access_level: "basic",
      allowed_departments: "operations",
      status: "active",
      topic_key: "tai-lieu-van-chuyen-test",
      storage_folder: "Vận Chuyển",
      owner_user_id: "sheet-nt005",
      shared_with_users: "",
    });

    assert.match(saved.path, /data\/internal\/Vận Chuyển\//);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("tai lieu rieng tu duoc luu vao thu muc con theo user", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "private-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const saved = await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu rieng tu test",
      content: "Noi dung rieng tu.",
      effective_date: "2026-04-04",
      owner: "Nguyen Van A",
      owner_department: "",
      access_level: "sensitive",
      allowed_departments: "",
      status: "active",
      topic_key: "tai-lieu-rieng-tu-test",
      storage_folder: "dữ liệu cá nhân/NT003",
      owner_user_id: "sheet-nt003",
      shared_with_users: "",
    });

    assert.match(saved.path, /data\/internal\/dữ liệu cá nhân\/NT003\//);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("luu tai lieu vao thu muc custom", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "custom-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const saved = await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu kho chung",
      content: "Noi dung thu muc custom.",
      effective_date: "2026-04-05",
      owner: "Admin",
      owner_department: "",
      access_level: "advanced",
      allowed_departments: "",
      status: "active",
      topic_key: "tai-lieu-kho-chung",
      storage_folder: "Kho chung mien Bac",
      owner_user_id: "sheet-nt002",
      shared_with_users: "",
    });

    assert.match(saved.path, /data\/internal\/Kho chung mien Bac\//);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("upload file se luu ca file goc va ban md da trich xuat", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "upload-original-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const uploadsDir = path.join(cwd, "data", "uploads");
    const fileName = "Bao gia thang 4.txt";
    const fileContent = "Dong 1\nDong 2";
    const saved = await importUploadedDocument(internalDir, uploadsDir, cwd, {
      title: "Bao gia thang 4",
      fileName,
      contentBase64: Buffer.from(fileContent, "utf8").toString("base64"),
      effective_date: "2026-04-08",
      owner: "sales",
      owner_department: "sales",
      access_level: "advanced",
      allowed_departments: "sales",
      status: "active",
      topic_key: "bao-gia-thang-4",
      storage_folder: "phong kinh doanh",
      owner_user_id: "sheet-nt003",
      shared_with_users: "",
    });

    const document = await readInternalDocument(internalDir, cwd, saved.path);
    const originalFilePath = path.join(cwd, document.metadata.original_file_path);
    const originalFileRaw = await readFile(originalFilePath, "utf8");

    assert.equal(originalFileRaw, fileContent);
    assert.equal(document.metadata.original_file_name, fileName);
    assert.match(document.metadata.original_file_path, /data\/uploads\/phong kinh doanh\/\d{4}-\d{2}\//);
    assert.match(document.metadata.original_file_path, /Bao gia thang 4\.txt$/);
    assert.match(document.content, /Nguon file: Bao gia thang 4\.txt/);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("tao thu muc custom rong van duoc liet ke", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "empty-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    await createStorageFolder(internalDir, "Kho nhanh", "basic");

    const folders = await listCustomStorageFolders(internalDir);
    const repo = await loadInternalRepository(internalDir, cwd);

    assert.deepEqual(folders, [{ name: "Kho nhanh", level: "basic" }]);
    assert.equal(repo.documents.length, 0);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("doi ten thu muc custom se cap nhat duong dan va metadata", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "rename-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sourceDir = path.join(cwd, "data", "sources");
    const sheetUrl = "https://docs.google.com/spreadsheets/d/custom-folder/edit#gid=0";
    const fetchImpl = async () =>
      new Response("Muc,Gia\nA,100", {
        status: 200,
        headers: { "content-type": "text/csv; charset=utf-8" },
      });

    await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu can doi ten thu muc",
      content: "Noi dung thu muc cu.",
      effective_date: "2026-04-05",
      owner: "Admin",
      owner_department: "",
      access_level: "advanced",
      allowed_departments: "",
      status: "active",
      topic_key: "tai-lieu-doi-ten-thu-muc",
      storage_folder: "Kho chung",
      owner_user_id: "sheet-nt002",
      shared_with_users: "",
    });

    await saveSyncSource(sourceDir, internalDir, cwd, {
      title: "Nguon dong bo kho chung",
      sheetUrl,
      effective_date: "2026-04-05",
      owner: "Admin",
      owner_department: "",
      access_level: "advanced",
      allowed_departments: "",
      status: "active",
      topic_key: "nguon-dong-bo-kho-chung",
      storage_folder: "Kho chung",
      owner_user_id: "sheet-nt002",
      shared_with_users: "",
    }, { fetchImpl });

    await renameStorageFolder(sourceDir, internalDir, cwd, "Kho chung", "Kho mien Nam");

    const repo = await loadInternalRepository(internalDir, cwd);
    const sources = await listSyncSources(sourceDir, internalDir, cwd);

    assert.ok(repo.documents.every((document) => String(document.relativePath).includes("Kho mien Nam")));
    assert.ok(
      repo.documents.every(
        (document) => String(document.metadata.storage_folder || "") === "Kho mien Nam",
      ),
    );
    assert.ok(sources.some((source) => source.storage_folder === "Kho mien Nam"));
    assert.ok(
      sources.some((source) => String(source.document_path || "").includes("Kho mien Nam")),
    );
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("xoa thu muc custom se xoa tai lieu va source record ben trong", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "delete-folder-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sourceDir = path.join(cwd, "data", "sources");

    await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu se bi xoa",
      content: "Noi dung trong thu muc xoa.",
      effective_date: "2026-04-05",
      owner: "Admin",
      owner_department: "",
      access_level: "basic",
      allowed_departments: "",
      status: "active",
      topic_key: "tai-lieu-se-bi-xoa",
      storage_folder: "Kho tam",
      owner_user_id: "sheet-nt002",
      shared_with_users: "",
    });

    await deleteStorageFolder(sourceDir, internalDir, cwd, "Kho tam");

    const repo = await loadInternalRepository(internalDir, cwd);
    const sources = await listSyncSources(sourceDir, internalDir, cwd);

    assert.equal(
      repo.documents.filter((document) => String(document.relativePath).includes("Kho tam")).length,
      0,
    );
    assert.equal(sources.filter((source) => source.storage_folder === "Kho tam").length, 0);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("doc cap nhat va xoa tai lieu theo duong dan display", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "manage-doc-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sourceDir = path.join(cwd, "data", "sources");
    const saved = await saveInternalDocument(internalDir, cwd, {
      title: "Tai lieu quan ly",
      content: "Noi dung thao tac tai lieu.",
      effective_date: "2026-04-05",
      owner: "Nguyen Van A",
      owner_department: "",
      access_level: "sensitive",
      allowed_departments: "",
      status: "active",
      topic_key: "tai-lieu-quan-ly",
      storage_folder: "dá»¯ liá»‡u cÃ¡ nhÃ¢n/NT003",
      owner_user_id: "sheet-nt003",
      shared_with_users: "sheet-nt004",
    });

    const loaded = await readInternalDocument(internalDir, cwd, saved.path);
    assert.equal(loaded.title, "Tai lieu quan ly");
    assert.equal(loaded.metadata.owner_user_id, "sheet-nt003");

    const updated = await updateInternalDocumentMetadata(internalDir, cwd, saved.path, {
      owner_user_id: "sheet-nt005",
      shared_with_users: "sheet-nt004,sheet-nt006",
      storage_folder: "dá»¯ liá»‡u cÃ¡ nhÃ¢n/NT005",
    });
    assert.equal(updated.metadata.owner_user_id, "sheet-nt005");
    assert.equal(updated.metadata.shared_with_users, "sheet-nt004,sheet-nt006");
    assert.match(updated.path, /NT005/);

    await deleteInternalDocument(internalDir, sourceDir, cwd, updated.path);
    const repo = await loadInternalRepository(internalDir, cwd);
    assert.equal(repo.documents.length, 0);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("quan ly nguon dong bo: tam dung, dong bo lai va xoa", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "manage-source-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sourceDir = path.join(cwd, "data", "sources");
    let csv = "Muc,Gia\nA,100";
    const fetchImpl = async () =>
      new Response(csv, {
        status: 200,
        headers: { "content-type": "text/csv; charset=utf-8" },
      });

    const saved = await saveSyncSource(
      sourceDir,
      internalDir,
      cwd,
      {
        title: "Nguon test thao tac",
        sheetUrl: "https://docs.google.com/spreadsheets/d/manage-source/edit#gid=0",
        effective_date: "2026-04-05",
        owner: "Admin",
        owner_department: "",
        access_level: "advanced",
        allowed_departments: "",
        status: "active",
        topic_key: "nguon-test-thao-tac",
        storage_folder: "Kho chung",
        owner_user_id: "sheet-nt002",
        shared_with_users: "",
      },
      { fetchImpl },
    );

    const paused = await setSourceSyncable(
      sourceDir,
      internalDir,
      cwd,
      saved.source.id,
      false,
    );
    assert.equal(paused.syncable, false);

    csv = "Muc,Gia\nA,150";
    await setSourceSyncable(sourceDir, internalDir, cwd, saved.source.id, true);
    const synced = await syncSourceById(sourceDir, internalDir, cwd, saved.source.id, { fetchImpl });
    assert.equal(synced.changed, true);

    await deleteSyncSource(sourceDir, internalDir, cwd, saved.source.id);
    const sources = await listSyncSources(sourceDir, internalDir, cwd);
    assert.equal(sources.length, 0);
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("ban nhap moi khong de ban dang ap dung theo cung chu de", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "draft-review-doc-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const active = await saveInternalDocument(internalDir, cwd, {
      title: "Quy trinh giao nhan",
      content: "Ban dang ap dung.",
      effective_date: "2026-04-05",
      owner: "Admin",
      owner_department: "",
      access_level: "basic",
      allowed_departments: "",
      status: "active",
      topic_key: "quy-trinh-giao-nhan",
      storage_folder: "internal",
      owner_user_id: "sheet-nt002",
      shared_with_users: "",
    });

    const draft = await saveInternalDocument(internalDir, cwd, {
      title: "Quy trinh giao nhan",
      content: "Ban cho duyet.",
      effective_date: "2026-04-05",
      owner: "Nhan vien",
      owner_department: "",
      access_level: "basic",
      allowed_departments: "",
      status: "draft",
      topic_key: "quy-trinh-giao-nhan",
      storage_folder: "internal",
      owner_user_id: "sheet-nt003",
      shared_with_users: "",
    });

    const beforeReview = await loadInternalRepository(internalDir, cwd);
    assert.equal(
      beforeReview.documents.find((document) => document.relativePath === active.path)?.metadata?.status,
      "active",
    );
    assert.equal(
      beforeReview.documents.find((document) => document.relativePath === draft.path)?.metadata?.status,
      "draft",
    );

    await reviewInternalDocument(internalDir, cwd, draft.path, {
      status: "active",
      reviewed_by_user_id: "sheet-nt002",
    });

    const afterReview = await loadInternalRepository(internalDir, cwd);
    assert.equal(
      afterReview.documents.find((document) => document.relativePath === active.path)?.metadata?.status,
      "superseded",
    );
    assert.equal(
      afterReview.documents.find((document) => document.relativePath === draft.path)?.metadata?.status,
      "active",
    );
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});

test("duyet nguon dong bo se doi trang thai cho source va tai lieu canon", async () => {
  const cwd = await mkdtemp(path.join(tmpdir(), "draft-review-source-"));

  try {
    const internalDir = path.join(cwd, "data", "internal");
    const sourceDir = path.join(cwd, "data", "sources");
    const fetchImpl = async () =>
      new Response("Muc,Gia\nA,100", {
        status: 200,
        headers: { "content-type": "text/csv; charset=utf-8" },
      });

    const saved = await saveSyncSource(
      sourceDir,
      internalDir,
      cwd,
      {
        title: "Bang gia test duyet",
        sheetUrl: "https://docs.google.com/spreadsheets/d/review-source/edit#gid=0",
        effective_date: "2026-04-05",
        owner: "Nhan vien",
        owner_department: "",
        access_level: "basic",
        allowed_departments: "",
        status: "draft",
        topic_key: "bang-gia-test-duyet",
        storage_folder: "internal",
        owner_user_id: "sheet-nt003",
        shared_with_users: "",
      },
      { fetchImpl },
    );

    const reviewed = await reviewSyncSource(
      sourceDir,
      internalDir,
      cwd,
      saved.source.id,
      {
        status: "active",
        reviewed_by_user_id: "sheet-nt002",
      },
    );
    const document = await readInternalDocument(internalDir, cwd, saved.document.path);

    assert.equal(reviewed.status, "active");
    assert.equal(document.metadata.status, "active");
    assert.equal(document.metadata.reviewed_by_user_id, "sheet-nt002");
  } finally {
    await rm(cwd, { recursive: true, force: true });
  }
});
