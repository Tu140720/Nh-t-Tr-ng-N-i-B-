import { createHash } from "node:crypto";
import { mkdir, readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import mammoth from "mammoth";
import xlsx from "xlsx";
import JSZip from "jszip";

const SUPPORTED_EXTENSIONS = new Set([".csv", ".json", ".md", ".txt"]);
const SUPPORTED_UPLOAD_EXTENSIONS = new Set([
  ".csv",
  ".docx",
  ".json",
  ".md",
  ".pptx",
  ".txt",
  ".xlsx",
]);
const CANONICAL_DIR = "canonical";
const SOURCE_FILE_EXTENSION = ".json";
const WEB_SOURCE_DIR = "dá»¯ liá»‡u ná»™i bá»™ tá»« Web";
const CUSTOM_FOLDER_META_FILE = ".folder-meta";
const SYSTEM_STORAGE_FOLDERS = new Map([
  ["internal", "basic"],
  ["Váº­n Chuyá»ƒn", "basic"],
  ["Káº¿ ToÃ¡n", "advanced"],
  ["phÃ²ng kinh doanh", "advanced"],
  ["phÃ²ng Sáº£n Xuáº¥t", "advanced"],
  ["dá»¯ liá»‡u cÃ¡ nhÃ¢n", "sensitive"],
]);

export async function loadInternalRepository(dir, cwd = process.cwd()) {
  await mkdir(dir, { recursive: true });
  const files = await walkDirectory(dir);
  const documents = [];
  const chunks = [];

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    const relativePath = toDisplayPath(cwd, filePath);

    const document = {
      id: relativePath,
      title: parsed.title,
      relativePath,
      content: parsed.text,
      metadata: parsed.metadata,
    };

    const documentChunks = splitIntoChunks(parsed.text).map((text, index) => ({
      id: `${relativePath}#${index + 1}`,
      documentId: document.id,
      title: document.title,
      relativePath,
      chunkIndex: index,
      text,
      metadata: document.metadata,
    }));

    documents.push(document);
    chunks.push(...documentChunks);
  }

  documents.sort((left, right) => left.title.localeCompare(right.title, "vi"));

  return {
    documents,
    chunks,
  };
}

export async function listInternalDocuments(dir, cwd = process.cwd()) {
  const repo = await loadInternalRepository(dir, cwd);

  return repo.documents.map((document) => ({
    id: document.id,
    title: document.title,
    path: document.relativePath,
    preview: document.content.slice(0, 140).replace(/\s+/g, " ").trim(),
    metadata: document.metadata,
  }));
}

export async function listCustomStorageFolders(dir) {
  await mkdir(dir, { recursive: true });
  const entries = await readdir(dir, { withFileTypes: true });
  const folders = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }

    const folderName = entry.name;
    if (SYSTEM_STORAGE_FOLDERS.has(folderName) || folderName === CANONICAL_DIR) {
      continue;
    }

    const folderPath = path.join(dir, folderName);
    const level = await resolveCustomFolderLevel(folderPath);
    if (!["basic", "advanced"].includes(level)) {
      continue;
    }

    folders.push({
      name: folderName,
      level,
    });
  }

  return folders.sort((left, right) => left.name.localeCompare(right.name, "vi"));
}

export async function listSyncSources(sourceDir, docsDir, cwd = process.cwd()) {
  await bootstrapSourceRegistryFromDocuments(sourceDir, docsDir, cwd);
  const files = await walkDirectorySafe(sourceDir);
  const sourcesById = new Map();

  for (const filePath of files) {
    if (path.extname(filePath).toLowerCase() !== SOURCE_FILE_EXTENSION) {
      continue;
    }

    try {
      const raw = await readFile(filePath, "utf8");
      const parsed = normalizeSourceRecord(JSON.parse(raw), filePath, cwd);
      const current = sourcesById.get(parsed.id);
      if (!current || shouldReplaceSourceRecord(current, parsed)) {
        sourcesById.set(parsed.id, parsed);
      }
    } catch {
      continue;
    }
  }

  return [...sourcesById.values()].sort(compareSources);
}

export async function saveSyncSource(sourceDir, docsDir, cwd, source, options = {}) {
  await mkdir(sourceDir, { recursive: true });
  await bootstrapSourceRegistryFromDocuments(sourceDir, docsDir, cwd);

  const prepared = prepareSourceInput(source, "sheet");
  const existing = await findSourceRecord(sourceDir, docsDir, cwd, prepared);
  const sourceRecord = {
    ...(existing || {}),
    id: existing?.id || buildSourceId(prepared),
    title: prepared.title,
    type: prepared.type,
    source_url: prepared.source_url,
    effective_date: prepared.effective_date,
    version: prepared.version,
    owner: prepared.owner,
    owner_department: prepared.owner_department,
    access_level: prepared.access_level,
    allowed_departments: prepared.allowed_departments,
    owner_user_id: prepared.owner_user_id,
    shared_with_users: prepared.shared_with_users,
    storage_folder: prepared.storage_folder,
    status: prepared.status,
    topic_key: prepared.topic_key,
    reviewed_at: existing?.reviewed_at || prepared.reviewed_at || "",
    reviewed_by_user_id: existing?.reviewed_by_user_id || prepared.reviewed_by_user_id || "",
    syncable: true,
    added_at: existing?.added_at || formatDate(new Date()),
    updated_at: formatDate(new Date()),
    last_synced_at: existing?.last_synced_at || "",
    last_error: "",
    document_path: existing?.document_path || "",
    content_fingerprint: existing?.content_fingerprint || "",
  };

  await writeSourceRecord(sourceDir, sourceRecord);
  const sync = await syncSingleSourceRecord(sourceDir, docsDir, cwd, sourceRecord, options);

  return {
    source: sync.source,
    document: sync.document,
  };
}

export async function saveInternalDocument(dir, cwd, document) {
  await mkdir(dir, { recursive: true });

  const prepared = prepareDocumentInput(document, "manual");
  prepared.metadata.content_fingerprint = fingerprintContent(prepared.content);
  await ensureNoDuplicateDocument(dir, cwd, prepared);
  if (prepared.metadata.status === "active") {
    await supersedeActiveByTopic(dir, prepared.metadata.topic_key);
  }

  return writeCanonicalDocument(dir, cwd, prepared);
}

export async function importSheetDocument(dir, cwd, document) {
  return importSheetDocumentWithOptions(dir, cwd, document);
}

export async function importSheetDocumentWithOptions(dir, cwd, document, options = {}) {
  const prepared = prepareDocumentInput(document, "sheet");
  const text = await extractSheetText(
    prepared.source.source_url,
    prepared.title,
    options.fetchImpl,
  );
  prepared.content = buildSheetContent(text, prepared.source.source_url);
  prepared.metadata.content_fingerprint = fingerprintContent(prepared.content);
  prepared.metadata.synced_at = formatDate(new Date());

  await ensureNoDuplicateDocument(dir, cwd, prepared);
  if (prepared.metadata.status === "active") {
    await supersedeActiveByTopic(dir, prepared.metadata.topic_key);
  }
  return writeCanonicalDocument(dir, cwd, prepared);
}

export async function validateSheetDocument(dir, cwd, document) {
  return validateSheetDocumentWithOptions(dir, cwd, document);
}

export async function validateSheetDocumentWithOptions(dir, cwd, document, options = {}) {
  const prepared = prepareDocumentInput(document, "sheet");
  const duplicate = await findDuplicateDocument(dir, cwd, prepared);
  if (duplicate) {
    throw buildDuplicateDocumentError(duplicate);
  }

  const text = await extractSheetText(
    prepared.source.source_url,
    prepared.title,
    options.fetchImpl,
  );
  return {
    title: prepared.title,
    preview: text.replace(/\s+/g, " ").trim().slice(0, 220),
    lineCount: text.split(/\r?\n/).filter(Boolean).length,
  };
}

export async function syncRegisteredSources(sourceDir, docsDir, cwd = process.cwd(), options = {}) {
  await mkdir(sourceDir, { recursive: true });
  await bootstrapSourceRegistryFromDocuments(sourceDir, docsDir, cwd);

  const sources = await listSyncSources(sourceDir, docsDir, cwd);
  const result = {
    checked: 0,
    updated: 0,
    unchanged: 0,
    errors: [],
  };

  for (const source of sources) {
    if (!shouldSyncSource(source)) {
      continue;
    }

    result.checked += 1;

    try {
      const sync = await syncSingleSourceRecord(sourceDir, docsDir, cwd, source, options);
      if (sync.changed) {
        result.updated += 1;
      } else {
        result.unchanged += 1;
      }
    } catch (error) {
      result.errors.push({
        title: source.title,
        path: source.document_path || source.path || source.source_url,
        message: error instanceof Error ? error.message : "Khong dong bo duoc nguon du lieu.",
      });
      await markSourceSyncError(sourceDir, source, error);
    }
  }

  return result;
}

export async function syncSheetDocuments(dir, cwd = process.cwd(), options = {}) {
  await mkdir(dir, { recursive: true });
  const files = await walkDirectory(dir);
  const result = {
    checked: 0,
    updated: 0,
    unchanged: 0,
    errors: [],
  };

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    if (!shouldAutoSyncSheet(parsed.metadata)) {
      continue;
    }

    result.checked += 1;

    try {
      const text = await extractSheetText(
        parsed.metadata.source_url,
        parsed.title,
        options.fetchImpl,
      );
      const nextContent = buildSheetContent(text, parsed.metadata.source_url);
      const nextFingerprint = fingerprintContent(nextContent);
      const currentFingerprint =
        parsed.metadata.content_fingerprint || fingerprintContent(parsed.text);
      const needsMetadataRefresh =
        parsed.metadata.content_fingerprint !== nextFingerprint || !parsed.metadata.synced_at;

      if (nextFingerprint === currentFingerprint && !needsMetadataRefresh) {
        result.unchanged += 1;
        continue;
      }

      const nextMetadata = {
        ...parsed.metadata,
        content_fingerprint: nextFingerprint,
        synced_at: formatDate(new Date()),
      };

      await writeFile(
        filePath,
        serializeDocument(parsed.title, nextMetadata, nextContent),
        "utf8",
      );

      if (nextFingerprint === currentFingerprint) {
        result.unchanged += 1;
      } else {
        result.updated += 1;
      }
    } catch (error) {
      result.errors.push({
        title: parsed.title,
        path: toDisplayPath(cwd, filePath),
        message: error instanceof Error ? error.message : "Khong dong bo duoc Google Sheet.",
      });
    }
  }

  return result;
}

export async function importUploadedDocument(dir, cwd, document) {
  const prepared = prepareDocumentInput(document, "file_upload");
  const fileName = String(document.fileName || "").trim();
  const extension = path.extname(fileName).toLowerCase();

  if (!SUPPORTED_UPLOAD_EXTENSIONS.has(extension)) {
    throw new Error("Dinh dang file chua duoc ho tro.");
  }

  const buffer = Buffer.from(String(document.contentBase64 || ""), "base64");
  if (!buffer.length) {
    throw new Error("File tai len khong hop le hoac rong.");
  }

  const extractedText = await extractUploadedText(extension, buffer, fileName);
  prepared.content = `${extractedText}\n\nNguon file: ${fileName}`;
  prepared.source.file_name = fileName;
  prepared.metadata.content_fingerprint = fingerprintContent(prepared.content);

  await ensureNoDuplicateDocument(dir, cwd, prepared);
  if (prepared.metadata.status === "active") {
    await supersedeActiveByTopic(dir, prepared.metadata.topic_key);
  }
  return writeCanonicalDocument(dir, cwd, prepared);
}

export async function renameStorageFolder(sourceDir, docsDir, cwd, currentFolder, nextFolder) {
  const currentNormalized = normalizeStorageFolder(currentFolder);
  const nextNormalized = normalizeStorageFolder(nextFolder);

  if (!currentNormalized || !nextNormalized) {
    throw new Error("Thu muc khong hop le.");
  }

  if (currentNormalized === nextNormalized) {
    throw new Error("Ten thu muc moi phai khac ten hien tai.");
  }

  const currentDir = resolveStorageFolderPath(docsDir, currentNormalized);
  const nextDir = resolveStorageFolderPath(docsDir, nextNormalized);
  await ensureDirectoryExists(currentDir);
  await ensureDirectoryMissing(nextDir);
  await mkdir(path.dirname(nextDir), { recursive: true });
  await rename(currentDir, nextDir);
  await rewriteDocumentsForFolder(nextDir, currentNormalized, nextNormalized);
  await rewriteSourcesForFolderRename(sourceDir, docsDir, cwd, currentNormalized, nextNormalized);
}

export async function deleteStorageFolder(sourceDir, docsDir, cwd, folderName) {
  const normalizedFolder = normalizeStorageFolder(folderName);
  if (!normalizedFolder) {
    throw new Error("Thu muc khong hop le.");
  }

  const targetDir = resolveStorageFolderPath(docsDir, normalizedFolder);
  await ensureDirectoryExists(targetDir);
  await rm(targetDir, { recursive: true, force: true });
  await deleteSourcesForFolder(sourceDir, docsDir, cwd, normalizedFolder);
}

export async function createStorageFolder(dir, folderName, level) {
  const normalizedFolder = normalizeStorageFolder(folderName);
  if (!normalizedFolder || normalizedFolder.includes("/")) {
    throw new Error("Thu muc khong hop le.");
  }

  const normalizedLevel = String(level || "").trim().toLowerCase();
  if (!["basic", "advanced"].includes(normalizedLevel)) {
    throw new Error("Chi duoc tao thu muc o muc co ban hoac nang cao.");
  }

  const targetDir = resolveStorageFolderPath(dir, normalizedFolder);
  await ensureDirectoryMissing(targetDir);
  await mkdir(targetDir, { recursive: true });
  await writeCustomFolderMeta(targetDir, normalizedLevel);
}

export async function readInternalDocument(dir, cwd, documentPath) {
  const filePath = resolveInternalDocumentPath(dir, cwd, documentPath);
  const raw = await readFile(filePath, "utf8");
  const parsed = parseDocument(filePath, raw);
  return {
    id: toDisplayPath(cwd, filePath),
    title: parsed.title,
    path: toDisplayPath(cwd, filePath),
    file_name: path.basename(filePath),
    content: parsed.text,
    metadata: parsed.metadata,
  };
}

export async function deleteInternalDocument(dir, sourceDir, cwd, documentPath) {
  const current = await readInternalDocument(dir, cwd, documentPath);
  const filePath = resolveInternalDocumentPath(dir, cwd, documentPath);
  await rm(filePath, { force: true });
  await deleteSourcesForDocument(sourceDir, dir, cwd, current.path);
  return current;
}

export async function updateInternalDocumentMetadata(dir, cwd, documentPath, updates = {}) {
  const filePath = resolveInternalDocumentPath(dir, cwd, documentPath);
  const raw = await readFile(filePath, "utf8");
  const parsed = parseDocument(filePath, raw);
  const nextMetadata = {
    ...parsed.metadata,
    ...updates,
  };

  const nextStorageFolder = normalizeStorageFolder(nextMetadata.storage_folder || "");
  let nextFilePath = filePath;
  if (nextStorageFolder && nextStorageFolder !== normalizeStorageFolder(parsed.metadata.storage_folder || "")) {
    const nextDir = resolveDocumentDirectory(dir, nextStorageFolder);
    await mkdir(nextDir, { recursive: true });
    nextFilePath = path.join(nextDir, path.basename(filePath));
    await rename(filePath, nextFilePath);
  }

  await writeFile(nextFilePath, serializeDocument(parsed.title, nextMetadata, parsed.text), "utf8");
  return {
    id: toDisplayPath(cwd, nextFilePath),
    title: parsed.title,
    path: toDisplayPath(cwd, nextFilePath),
    metadata: nextMetadata,
    content: parsed.text,
  };
}

export async function reviewInternalDocument(dir, cwd, documentPath, review = {}) {
  const nextStatus = normalizeStatus(review.status);
  if (!nextStatus || !["active", "rejected"].includes(nextStatus)) {
    throw new Error("Trang thai duyet tai lieu khong hop le.");
  }

  const current = await readInternalDocument(dir, cwd, documentPath);
  if (nextStatus === "active" && current.metadata?.status !== "active") {
    await supersedeActiveByTopic(dir, current.metadata?.topic_key || "");
  }

  return updateInternalDocumentMetadata(dir, cwd, documentPath, {
    status: nextStatus,
    reviewed_at: String(review.reviewed_at || "").trim() || formatDate(new Date()),
    reviewed_by_user_id: String(review.reviewed_by_user_id || "").trim(),
  });
}

export async function syncSourceById(sourceDir, docsDir, cwd, sourceId, options = {}) {
  const source = await findSourceById(sourceDir, docsDir, cwd, sourceId);
  if (!source) {
    throw new Error("Khong tim thay nguon dong bo.");
  }
  if (source.syncable === false) {
    throw new Error("Nguon dong bo dang tam dung.");
  }
  return syncSingleSourceRecord(sourceDir, docsDir, cwd, source, options);
}

export async function setSourceSyncable(sourceDir, docsDir, cwd, sourceId, syncable) {
  const source = await findSourceById(sourceDir, docsDir, cwd, sourceId);
  if (!source) {
    throw new Error("Khong tim thay nguon dong bo.");
  }
  const nextSource = {
    ...source,
    syncable: syncable !== false,
    updated_at: formatDate(new Date()),
  };
  await writeSourceRecord(sourceDir, nextSource);
  return nextSource;
}

export async function reviewSyncSource(sourceDir, docsDir, cwd, sourceId, review = {}) {
  const nextStatus = normalizeStatus(review.status);
  if (!nextStatus || !["active", "rejected"].includes(nextStatus)) {
    throw new Error("Trang thai duyet nguon khong hop le.");
  }

  const source = await findSourceById(sourceDir, docsDir, cwd, sourceId);
  if (!source) {
    throw new Error("Khong tim thay nguon dong bo.");
  }

  if (nextStatus === "active" && source.status !== "active") {
    await supersedeActiveByTopic(docsDir, source.topic_key || "");
  }

  const reviewedAt = String(review.reviewed_at || "").trim() || formatDate(new Date());
  const reviewedByUserId = String(review.reviewed_by_user_id || "").trim();
  const nextSource = {
    ...source,
    status: nextStatus,
    updated_at: formatDate(new Date()),
    reviewed_at: reviewedAt,
    reviewed_by_user_id: reviewedByUserId,
  };
  await writeSourceRecord(sourceDir, nextSource);

  if (source.document_path) {
    await updateInternalDocumentMetadata(docsDir, cwd, source.document_path, {
      status: nextStatus,
      reviewed_at: reviewedAt,
      reviewed_by_user_id: reviewedByUserId,
    });
  }

  return nextSource;
}

export async function deleteSyncSource(sourceDir, docsDir, cwd, sourceId) {
  const source = await findSourceById(sourceDir, docsDir, cwd, sourceId);
  if (!source) {
    throw new Error("Khong tim thay nguon dong bo.");
  }

  const recordPath = path.resolve(cwd, source.path);
  await rm(recordPath, { force: true });
  if (source.document_path) {
    await rm(resolveInternalDocumentPath(docsDir, cwd, source.document_path), { force: true }).catch(() => {});
  }
  return source;
}

async function writeCanonicalDocument(dir, cwd, prepared) {
  const targetDir = resolveDocumentDirectory(dir, prepared.metadata.storage_folder || "");
  await mkdir(targetDir, { recursive: true });

  const slug = slugify(prepared.metadata.topic_key || prepared.title) || "tai-lieu-noi-bo";
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filePath = path.join(targetDir, `${slug}-${timestamp}.md`);
  const body = serializeDocument(prepared.title, prepared.metadata, prepared.content);

  await writeFile(filePath, body, "utf8");

  return {
    title: prepared.title,
    path: toDisplayPath(cwd, filePath),
    metadata: prepared.metadata,
  };
}

async function upsertCanonicalDocumentForSource(dir, cwd, source, title, metadata, content) {
  await mkdir(dir, { recursive: true });
  const files = await walkDirectory(dir);
  const sourceUrl = normalizeSourceValue(source.source_url);

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    const sameTopic = parsed.metadata.topic_key === source.topic_key;
    const sameSource =
      sourceUrl &&
      normalizeSourceValue(parsed.metadata.source_url) === sourceUrl;

    if (!sameTopic && !sameSource) {
      continue;
    }

    await writeFile(filePath, serializeDocument(title, metadata, content), "utf8");
    return toDisplayPath(cwd, filePath);
  }

  const created = await writeCanonicalDocument(dir, cwd, {
    title,
    metadata,
    content,
  });
  return created.path;
}

async function supersedeActiveByTopic(dir, topicKey) {
  const files = await walkDirectory(dir);

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    if (
      parsed.metadata.topic_key === topicKey &&
      parsed.metadata.status === "active"
    ) {
      parsed.metadata.status = "superseded";
      const updated = serializeDocument(parsed.title, parsed.metadata, parsed.text);
      await writeFile(filePath, updated, "utf8");
    }
  }
}

async function ensureNoDuplicateDocument(dir, cwd, prepared) {
  const duplicate = await findDuplicateDocument(dir, cwd, prepared);
  if (duplicate) {
    throw buildDuplicateDocumentError(duplicate);
  }
}

async function findDuplicateDocument(dir, cwd, prepared) {
  await mkdir(dir, { recursive: true });
  const files = await walkDirectory(dir);

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    const metadata = parsed.metadata;

    const sameSource =
      prepared.source.source_url &&
      metadata.source_url &&
      normalizeSourceValue(prepared.source.source_url) ===
        normalizeSourceValue(metadata.source_url);

    const sameFingerprint =
      prepared.metadata.content_fingerprint &&
      metadata.content_fingerprint &&
      prepared.metadata.content_fingerprint === metadata.content_fingerprint;

    if (!sameSource && !sameFingerprint) {
      continue;
    }

    return {
      title: parsed.title,
      path: toDisplayPath(cwd, filePath),
      sourceUrl: metadata.source_url || prepared.source.source_url || "",
      addedAt: metadata.added_at || (await readFileStatTime(filePath)),
      reason: sameSource ? "duplicate_source" : "duplicate_content",
    };
  }

  return null;
}

async function readFileStatTime(filePath) {
  const details = await stat(filePath);
  return formatDate(details.birthtime);
}

function prepareDocumentInput(document, sourceType) {
  const title = String(document.title || "").trim();
  const effectiveDate = String(document.effective_date || "").trim() || currentDateValue();
  const metadata = {
    effective_date: effectiveDate,
    source_type: sourceType,
    owner: String(document.owner || "").trim(),
    owner_department: normalizeDepartment(document.owner_department || document.owner || ""),
    access_level: normalizeAccessLevel(document.access_level || "basic"),
    allowed_departments: normalizeAllowedDepartments(
      document.allowed_departments || document.owner_department || document.owner || "",
    ),
    status: normalizeStatus(document.status) || "active",
    topic_key: normalizeTopicKey(document.topic_key || title),
    owner_user_id: String(document.owner_user_id || "").trim(),
    shared_with_users: normalizeSharedUsers(document.shared_with_users || ""),
    storage_folder: normalizeStorageFolder(document.storage_folder || ""),
    canonical: true,
    added_at: formatDate(new Date()),
    synced_at: "",
    source_url: String(document.sheetUrl || "").trim(),
    source_label: sourceType,
    content_fingerprint: "",
    reviewed_at: String(document.reviewed_at || "").trim(),
    reviewed_by_user_id: String(document.reviewed_by_user_id || "").trim(),
  };

  if (
    !title ||
    !metadata.effective_date ||
    !metadata.owner ||
    !metadata.status ||
    !metadata.topic_key
  ) {
    throw new Error(
      "Thieu metadata bat buoc: title, effective_date, owner, status, topic_key.",
    );
  }

  return {
    title,
    content: String(document.content || "").trim(),
    metadata,
    source: {
      source_url: metadata.source_url,
    },
  };
}

function prepareSourceInput(source, sourceType) {
  const title = String(source.title || "").trim();
  const effectiveDate = String(source.effective_date || "").trim() || currentDateValue();
  const prepared = {
    title,
    type: sourceType,
    source_url: String(source.sourceUrl || source.sheetUrl || "").trim(),
    effective_date: effectiveDate,
    owner: String(source.owner || "").trim(),
    owner_department: normalizeDepartment(source.owner_department || source.owner || ""),
    access_level: normalizeAccessLevel(source.access_level || "basic"),
    allowed_departments: normalizeAllowedDepartments(
      source.allowed_departments || source.owner_department || source.owner || "",
    ),
    status: normalizeStatus(source.status),
    topic_key: normalizeTopicKey(source.topic_key || title),
    owner_user_id: String(source.owner_user_id || "").trim(),
    shared_with_users: normalizeSharedUsers(source.shared_with_users || ""),
    storage_folder: normalizeStorageFolder(source.storage_folder || ""),
    reviewed_at: String(source.reviewed_at || "").trim(),
    reviewed_by_user_id: String(source.reviewed_by_user_id || "").trim(),
  };

  if (
    !prepared.title ||
    !prepared.source_url ||
    !prepared.effective_date ||
    !prepared.owner ||
    !prepared.status ||
    !prepared.topic_key
  ) {
    throw new Error(
      "Thieu metadata bat buoc: title, source_url, effective_date, owner, status, topic_key.",
    );
  }

  return prepared;
}

async function syncSingleSourceRecord(sourceDir, docsDir, cwd, source, options = {}) {
  const text = await extractSheetText(source.source_url, source.title, options.fetchImpl);
  const content = buildSheetContent(text, source.source_url);
  const nextFingerprint = fingerprintContent(content);
  const currentFingerprint = source.content_fingerprint || "";
  const changed = nextFingerprint !== currentFingerprint;

  const metadata = {
    effective_date: source.effective_date,
    version: source.version,
    source_type: source.type,
    owner: source.owner,
    owner_department: source.owner_department,
    access_level: source.access_level,
    allowed_departments: source.allowed_departments,
    status: source.status,
    topic_key: source.topic_key,
    owner_user_id: source.owner_user_id,
    shared_with_users: source.shared_with_users,
    storage_folder: source.storage_folder,
    canonical: true,
    added_at: source.added_at || formatDate(new Date()),
    synced_at: formatDate(new Date()),
    source_url: source.source_url,
    source_label: source.type,
    content_fingerprint: nextFingerprint,
  };

  const documentPath = await upsertCanonicalDocumentForSource(
    docsDir,
    cwd,
    source,
    source.title,
    metadata,
    content,
  );

  const nextSource = {
    ...source,
    updated_at: formatDate(new Date()),
    last_synced_at: metadata.synced_at,
    last_error: "",
    document_path: documentPath,
    content_fingerprint: nextFingerprint,
  };

  await writeSourceRecord(sourceDir, nextSource);

  return {
    changed,
    source: nextSource,
    document: {
      title: source.title,
      path: documentPath,
      metadata,
    },
  };
}

async function walkDirectory(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await walkDirectory(fullPath)));
    } else if (entry.isFile()) {
      files.push(fullPath);
    } else {
      const details = await stat(fullPath);
      if (details.isFile()) {
        files.push(fullPath);
      }
    }
  }

  return files;
}

async function walkDirectorySafe(dir) {
  try {
    return await walkDirectory(dir);
  } catch {
    return [];
  }
}

async function bootstrapSourceRegistryFromDocuments(sourceDir, docsDir, cwd) {
  await mkdir(sourceDir, { recursive: true });
  const existingSources = await walkDirectorySafe(sourceDir);
  if (existingSources.some((filePath) => path.extname(filePath).toLowerCase() === SOURCE_FILE_EXTENSION)) {
    return;
  }

  const documentFiles = await walkDirectorySafe(docsDir);
  for (const filePath of documentFiles) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    if (!shouldAutoSyncSheet(parsed.metadata)) {
      continue;
    }

    const sourceRecord = {
      id: buildSourceId({
        topic_key: parsed.metadata.topic_key,
        source_url: parsed.metadata.source_url,
      }),
      title: parsed.title,
      type: "sheet",
      source_url: parsed.metadata.source_url,
      effective_date: parsed.metadata.effective_date || "",
      version: parsed.metadata.version || "",
      owner: parsed.metadata.owner || "",
      owner_department: parsed.metadata.owner_department || normalizeDepartment(parsed.metadata.owner || ""),
      access_level: normalizeAccessLevel(parsed.metadata.access_level || "basic"),
      allowed_departments: normalizeAllowedDepartments(
        parsed.metadata.allowed_departments || parsed.metadata.owner_department || parsed.metadata.owner || "",
      ),
      status: parsed.metadata.status || "active",
      topic_key: parsed.metadata.topic_key || normalizeTopicKey(parsed.title),
      owner_user_id: String(parsed.metadata.owner_user_id || "").trim(),
      shared_with_users: normalizeSharedUsers(parsed.metadata.shared_with_users || ""),
      storage_folder: normalizeStorageFolder(parsed.metadata.storage_folder || ""),
      syncable: true,
      added_at: parsed.metadata.added_at || formatDate(new Date()),
      updated_at: formatDate(new Date()),
      last_synced_at: parsed.metadata.synced_at || "",
      last_error: "",
      document_path: toDisplayPath(cwd, filePath),
      content_fingerprint:
        parsed.metadata.content_fingerprint || fingerprintContent(parsed.text),
    };

    await writeSourceRecord(sourceDir, sourceRecord);
  }
}

function normalizeSourceRecord(value, filePath, cwd) {
  return {
    id: String(value.id || path.basename(filePath, SOURCE_FILE_EXTENSION)).trim(),
    path: toDisplayPath(cwd, filePath),
    title: String(value.title || "").trim(),
    type: String(value.type || "sheet").trim().toLowerCase(),
    source_url: String(value.source_url || "").trim(),
    effective_date: String(value.effective_date || "").trim(),
    version: String(value.version || "").trim(),
    owner: String(value.owner || "").trim(),
    owner_department: normalizeDepartment(value.owner_department || value.owner || ""),
    access_level: normalizeAccessLevel(value.access_level || "basic"),
    allowed_departments: normalizeAllowedDepartments(
      value.allowed_departments || value.owner_department || value.owner || "",
    ),
    status: normalizeStatus(value.status) || "active",
    topic_key: normalizeTopicKey(value.topic_key || value.title),
    owner_user_id: String(value.owner_user_id || "").trim(),
    shared_with_users: normalizeSharedUsers(value.shared_with_users || ""),
    storage_folder: normalizeStorageFolder(value.storage_folder || ""),
    reviewed_at: String(value.reviewed_at || "").trim(),
    reviewed_by_user_id: String(value.reviewed_by_user_id || "").trim(),
    syncable: value.syncable !== false,
    added_at: String(value.added_at || "").trim(),
    updated_at: String(value.updated_at || "").trim(),
    last_synced_at: String(value.last_synced_at || "").trim(),
    last_error: String(value.last_error || "").trim(),
    document_path: String(value.document_path || "").trim(),
    content_fingerprint: String(value.content_fingerprint || "").trim(),
  };
}

async function findSourceRecord(sourceDir, docsDir, cwd, prepared) {
  const sources = await listSyncSources(sourceDir, docsDir, cwd);
  const sourceUrl = normalizeSourceValue(prepared.source_url);
  return (
    sources.find(
      (item) =>
        normalizeSourceValue(item.source_url) === sourceUrl ||
        item.topic_key === prepared.topic_key,
    ) || null
  );
}

async function writeSourceRecord(sourceDir, source) {
  const targetDir = path.join(sourceDir, WEB_SOURCE_DIR);
  await mkdir(targetDir, { recursive: true });
  const legacyPath = path.join(
    sourceDir,
    `${slugify(source.id || source.topic_key)}${SOURCE_FILE_EXTENSION}`,
  );
  if (legacyPath !== targetDir) {
    await rm(legacyPath, { force: true }).catch(() => {});
  }
  const filePath = path.join(
    targetDir,
    `${slugify(source.id || source.topic_key)}${SOURCE_FILE_EXTENSION}`,
  );
  const body = JSON.stringify(source, null, 2);
  await writeFile(filePath, `${body}\n`, "utf8");
}

async function findSourceById(sourceDir, docsDir, cwd, sourceId) {
  const normalizedId = String(sourceId || "").trim();
  if (!normalizedId) {
    return null;
  }
  const sources = await listSyncSources(sourceDir, docsDir, cwd);
  return sources.find((source) => String(source.id || "").trim() === normalizedId) || null;
}

async function markSourceSyncError(sourceDir, source, error) {
  const nextSource = {
    ...source,
    updated_at: formatDate(new Date()),
    last_error: error instanceof Error ? error.message : "Khong dong bo duoc nguon du lieu.",
  };
  await writeSourceRecord(sourceDir, nextSource);
}

function shouldSyncSource(source = {}) {
  return (
    source.syncable !== false &&
    String(source.type || "").toLowerCase() === "sheet" &&
    String(source.status || "").toLowerCase() === "active" &&
    Boolean(String(source.source_url || "").trim())
  );
}

function buildSourceId(source) {
  const topic = normalizeTopicKey(source.topic_key || "");
  const hash = fingerprintContent(normalizeSourceValue(source.source_url || "")).slice(0, 10);
  return `${topic || "source"}-${hash}`;
}

function compareSources(left, right) {
  return (
    String(right.updated_at || "").localeCompare(String(left.updated_at || ""), "vi") ||
    left.title.localeCompare(right.title, "vi")
  );
}

function shouldReplaceSourceRecord(current, candidate) {
  const currentScore = sourceRecordPriority(current);
  const candidateScore = sourceRecordPriority(candidate);
  if (candidateScore !== currentScore) {
    return candidateScore > currentScore;
  }

  return compareSources(candidate, current) < 0;
}

function sourceRecordPriority(source) {
  const normalizedPath = String(source.path || "").toLowerCase();
  let score = 0;
  if (normalizedPath.includes(WEB_SOURCE_DIR.toLowerCase())) {
    score += 10;
  }
  if (source.last_synced_at) {
    score += 2;
  }
  return score;
}

function parseDocument(filePath, raw) {
  const extension = path.extname(filePath).toLowerCase();
  const fallbackTitle = path.basename(filePath, extension);

  if (extension === ".json") {
    try {
      const value = JSON.parse(raw);
      return {
        title: value.title || fallbackTitle,
        text: JSON.stringify(value, null, 2),
        metadata: defaultMetadata(value.title || fallbackTitle),
      };
    } catch {
      return {
        title: fallbackTitle,
        text: raw,
        metadata: defaultMetadata(fallbackTitle),
      };
    }
  }

  if (extension === ".csv") {
    return {
      title: fallbackTitle,
      text: raw,
      metadata: defaultMetadata(fallbackTitle),
    };
  }

  const frontMatterMatch = raw.match(/^---\n([\s\S]*?)\n---\n?([\s\S]*)$/);
  if (frontMatterMatch) {
    const metadata = parseFrontMatter(frontMatterMatch[1]);
    const text = frontMatterMatch[2].trim();
    const inferredMetadata = inferMetadataFromText(text);
    return {
      title: metadata.title || extractMarkdownTitle(text) || fallbackTitle,
      text,
      metadata: {
        ...defaultMetadata(metadata.title || fallbackTitle),
        ...inferredMetadata,
        ...metadata,
      },
    };
  }

  const titleMatch = raw.match(/^#\s+(.+)$/m);
  const title = titleMatch?.[1]?.trim() || fallbackTitle;
  const inferredMetadata = inferMetadataFromText(raw);
  return {
    title,
    text: raw.trim(),
    metadata: {
      ...defaultMetadata(title),
      ...inferredMetadata,
    },
  };
}

function serializeDocument(title, metadata, content) {
  const lines = ["---"];
  const normalized = {
    title,
    ...metadata,
  };

  for (const [key, value] of Object.entries(normalized)) {
    if (value === undefined || value === null || value === "") {
      continue;
    }

    lines.push(`${key}: ${String(value)}`);
  }

  lines.push("---", "", String(content || "").trim(), "");
  return lines.join("\n");
}

function parseFrontMatter(raw) {
  const metadata = {};
  for (const line of raw.split(/\r?\n/)) {
    const index = line.indexOf(":");
    if (index === -1) {
      continue;
    }

    const key = line.slice(0, index).trim();
    const value = line.slice(index + 1).trim();
    metadata[key] = value;
  }
  return metadata;
}

function defaultMetadata(title) {
  return {
    effective_date: "",
    version: "1.0",
    source_type: "manual",
    owner: "unknown",
    owner_department: "",
    access_level: "basic",
    allowed_departments: "",
    status: "active",
    topic_key: normalizeTopicKey(title),
    owner_user_id: "",
    shared_with_users: "",
    storage_folder: "",
    canonical: "false",
    added_at: "",
    synced_at: "",
    source_url: "",
    content_fingerprint: "",
    reviewed_at: "",
    reviewed_by_user_id: "",
  };
}

function inferMetadataFromText(text) {
  const metadata = {};
  const sourceUrl = extractSheetSourceUrl(text);
  const sourceFileName = extractUploadedSourceName(text);

  if (sourceUrl) {
    metadata.source_type = "sheet";
    metadata.source_url = sourceUrl;
    metadata.source_label = "sheet";
    return metadata;
  }

  if (sourceFileName) {
    metadata.source_type = "file_upload";
    metadata.source_label = "file_upload";
  }

  return metadata;
}

function splitIntoChunks(text) {
  const normalized = text.replace(/\r/g, "").trim();
  if (!normalized) {
    return [];
  }

  const blocks = normalized
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);

  const chunks = [];
  let current = "";

  for (const block of blocks) {
    const next = current ? `${current}\n\n${block}` : block;
    if (next.length > 900 && current) {
      chunks.push(current);
      current = block;
    } else if (block.length > 900) {
      chunks.push(...hardSplitBlock(block, 850));
      current = "";
    } else {
      current = next;
    }
  }

  if (current) {
    chunks.push(current);
  }

  return chunks;
}

function hardSplitBlock(block, maxLength) {
  const sentences = block.split(/(?<=[.!?])\s+/);
  const pieces = [];
  let current = "";

  for (const sentence of sentences) {
    const next = current ? `${current} ${sentence}` : sentence;
    if (next.length > maxLength && current) {
      pieces.push(current);
      current = sentence;
    } else {
      current = next;
    }
  }

  if (current) {
    pieces.push(current);
  }

  return pieces;
}

function slugify(value) {
  return value
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function toDisplayPath(cwd, filePath) {
  return path.relative(cwd, filePath).replace(/\\/g, "/");
}

async function extractUploadedText(extension, buffer, fileName) {
  if (extension === ".txt" || extension === ".md" || extension === ".csv") {
    return buffer.toString("utf8");
  }

  if (extension === ".json") {
    const raw = buffer.toString("utf8");
    try {
      const value = JSON.parse(raw);
      return JSON.stringify(value, null, 2);
    } catch {
      return raw;
    }
  }

  if (extension === ".docx") {
    const result = await mammoth.extractRawText({ buffer });
    return ensureImportedText(result.value, fileName);
  }

  if (extension === ".xlsx") {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheetTexts = workbook.SheetNames.map((sheetName) => {
      const csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      const normalized = csvToText(csv);
      return normalized ? `## ${sheetName}\n${normalized}` : "";
    }).filter(Boolean);

    return ensureImportedText(sheetTexts.join("\n\n"), fileName);
  }

  if (extension === ".pptx") {
    const zip = await JSZip.loadAsync(buffer);
    const slideEntries = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort(naturalCompare);

    const slideTexts = [];
    for (const entry of slideEntries) {
      const xml = await zip.files[entry].async("string");
      const text = extractPptxText(xml);
      if (text) {
        slideTexts.push(`## ${path.basename(entry, ".xml")}\n${text}`);
      }
    }

    return ensureImportedText(slideTexts.join("\n\n"), fileName);
  }

  throw new Error("Dinh dang file chua duoc ho tro.");
}

function ensureImportedText(value, fileName) {
  const text = String(value || "").replace(/\r/g, "").trim();
  if (!text) {
    throw new Error(`Khong trich xuat duoc noi dung tu file ${fileName}.`);
  }

  return text;
}

function extractPptxText(xml) {
  return (
    xml
      .match(/<a:t[^>]*>(.*?)<\/a:t>/g)
      ?.map((part) => part.replace(/<[^>]+>/g, "").trim())
      .filter(Boolean)
      .join("\n") || ""
  );
}

function naturalCompare(left, right) {
  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

function resolveSheetDownloadUrls(value) {
  const url = new URL(value);

  if (url.hostname === "docs.google.com" && url.pathname.includes("/spreadsheets/d/")) {
    const match = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    const gidFromHash = url.hash.match(/gid=(\d+)/i)?.[1] || "";
    const gid = url.searchParams.get("gid") || gidFromHash || "0";
    if (match?.[1]) {
      const spreadsheetId = match[1];
      return [
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`,
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?gid=${gid}&format=csv`,
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv`,
      ];
    }
  }

  return [value];
}

function normalizeImportedText(raw, contentType, title) {
  const text = raw.trim();
  if (!text) {
    throw new Error("Sheet khong co du lieu de nap.");
  }

  if (contentType.includes("text/csv") || looksLikeCsv(text)) {
    return `# ${title}\n\n${csvToText(text)}`;
  }

  return `# ${title}\n\n${text}`;
}

async function extractSheetText(rawUrl, title, fetchImpl = fetch) {
  const candidateUrls = resolveSheetDownloadUrls(rawUrl).map(addNoCacheQuery);
  let lastStatus = 0;

  for (const candidateUrl of candidateUrls) {
    const response = await fetchImpl(candidateUrl, {
      cache: "no-store",
      headers: {
        "cache-control": "no-cache, no-store, max-age=0",
        pragma: "no-cache",
      },
    });

    if (!response.ok) {
      lastStatus = response.status;
      continue;
    }

    const contentType = response.headers.get("content-type") || "";
    const raw = await response.text();
    return normalizeImportedText(raw, contentType, title);
  }

  throw new Error(`Khong the tai du lieu tu sheet (${lastStatus || "unknown"}).`);
}

function addNoCacheQuery(value) {
  const url = new URL(value);
  url.searchParams.set("_sync", Date.now().toString());
  return url.toString();
}

function buildSheetContent(text, sourceUrl) {
  return `${text}\n\nNguon sheet: ${sourceUrl}`;
}

function extractSheetSourceUrl(text) {
  const match = String(text || "").match(/(?:^|\n)Nguon sheet:\s*(https?:\/\/\S+)/i);
  return match?.[1]?.trim() || "";
}

function extractUploadedSourceName(text) {
  const match = String(text || "").match(/(?:^|\n)Nguon file:\s*(.+)/i);
  return match?.[1]?.trim() || "";
}

function shouldAutoSyncSheet(metadata = {}) {
  return (
    String(metadata.source_type || "").toLowerCase() === "sheet" &&
    String(metadata.status || "").toLowerCase() === "active" &&
    Boolean(String(metadata.source_url || "").trim())
  );
}

function looksLikeCsv(value) {
  const firstLine = value.split(/\r?\n/, 1)[0] || "";
  return firstLine.includes(",");
}

function csvToText(value) {
  const lines = value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  return lines.map((line) => line.replace(/,/g, " | ")).join("\n");
}

function normalizeTopicKey(value) {
  return slugify(String(value || ""));
}

function normalizeStatus(value) {
  const status = String(value || "").trim().toLowerCase();
  return ["active", "superseded", "draft", "rejected"].includes(status) ? status : "";
}

function normalizeAccessLevel(value) {
  const level = String(value || "").trim().toLowerCase();
  return ["basic", "advanced", "sensitive"].includes(level) ? level : "basic";
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

function normalizeAllowedDepartments(value) {
  return String(value || "")
    .split(",")
    .map((item) => normalizeDepartment(item))
    .filter(Boolean)
    .join(",");
}

function normalizeSharedUsers(value) {
  return String(value || "")
    .split(",")
    .map((item) => String(item || "").trim())
    .filter(Boolean)
    .join(",");
}

function normalizeStorageFolder(value) {
  return String(value || "")
    .split(/[\\/]+/)
    .map((segment) => String(segment || "").trim())
    .filter((segment) => segment && segment !== "." && segment !== "..")
    .join("/");
}

function resolveStorageFolderPath(rootDir, storageFolder) {
  const normalized = normalizeStorageFolder(storageFolder);
  if (!normalized) {
    throw new Error("Thu muc khong hop le.");
  }

  const rootPath = path.resolve(rootDir);
  const targetPath = path.resolve(rootPath, ...normalized.split("/"));
  const relative = path.relative(rootPath, targetPath);

  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error("Thu muc khong hop le.");
  }

  return targetPath;
}

function resolveInternalDocumentPath(rootDir, cwd, displayPath) {
  const relativePath = String(displayPath || "").trim();
  if (!relativePath) {
    throw new Error("Khong tim thay tai lieu.");
  }

  const rootPath = path.resolve(rootDir);
  const targetPath = path.resolve(cwd, relativePath);
  const relative = path.relative(rootPath, targetPath);

  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error("Tai lieu khong hop le.");
  }

  return targetPath;
}

function resolveDocumentDirectory(dir, storageFolder) {
  const normalized = normalizeStorageFolder(storageFolder);
  if (!normalized) {
    return path.join(dir, CANONICAL_DIR);
  }

  return path.join(dir, normalized);
}

function normalizeSourceValue(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function fingerprintContent(content) {
  return createHash("sha256").update(String(content || "").trim()).digest("hex");
}

function buildDuplicateDocumentError(duplicate) {
  const reason =
    duplicate.reason === "duplicate_source"
      ? "Link Google Sheet nay da duoc them truoc do."
      : "Noi dung nay da ton tai trong kho du lieu.";

  return new Error(
    `${reason}\nLink: ${duplicate.sourceUrl}\nTieu de: ${duplicate.title}\nDuong dan: ${duplicate.path}\nThoi diem them: ${duplicate.addedAt}`,
  );
}

function formatDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "Khong ro";
  }

  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "medium",
    timeStyle: "medium",
  }).format(date);
}

function currentDateValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

async function ensureDirectoryExists(targetDir) {
  try {
    const details = await stat(targetDir);
    if (!details.isDirectory()) {
      throw new Error("Thu muc khong ton tai.");
    }
  } catch {
    throw new Error("Thu muc khong ton tai.");
  }
}

async function ensureDirectoryMissing(targetDir) {
  try {
    await stat(targetDir);
    throw new Error("Thu muc moi da ton tai.");
  } catch (error) {
    if (error instanceof Error && error.message === "Thu muc moi da ton tai.") {
      throw error;
    }
  }
}

async function resolveCustomFolderLevel(folderPath) {
  const metaPath = path.join(folderPath, CUSTOM_FOLDER_META_FILE);
  try {
    const raw = await readFile(metaPath, "utf8");
    const normalized = String(raw || "").trim().toLowerCase();
    if (["basic", "advanced"].includes(normalized)) {
      return normalized;
    }
  } catch {}

  const files = await walkDirectorySafe(folderPath);
  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    return normalizeAccessLevelValue(parsed.metadata.access_level || "basic");
  }

  return "";
}

async function writeCustomFolderMeta(folderPath, level) {
  await writeFile(path.join(folderPath, CUSTOM_FOLDER_META_FILE), `${level}\n`, "utf8");
}

async function rewriteDocumentsForFolder(targetDir, currentFolder, nextFolder) {
  const files = await walkDirectorySafe(targetDir);

  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(extension)) {
      continue;
    }

    const raw = await readFile(filePath, "utf8");
    const parsed = parseDocument(filePath, raw);
    const nextStorageFolder = remapStorageFolder(parsed.metadata.storage_folder, currentFolder, nextFolder);
    if (!nextStorageFolder || nextStorageFolder === parsed.metadata.storage_folder) {
      continue;
    }

    parsed.metadata.storage_folder = nextStorageFolder;
    await writeFile(filePath, serializeDocument(parsed.title, parsed.metadata, parsed.text), "utf8");
  }
}

async function rewriteSourcesForFolderRename(sourceDir, docsDir, cwd, currentFolder, nextFolder) {
  const sources = await listSyncSources(sourceDir, docsDir, cwd);

  for (const source of sources) {
    const nextStorageFolder = remapStorageFolder(source.storage_folder, currentFolder, nextFolder);
    if (!nextStorageFolder || nextStorageFolder === source.storage_folder) {
      continue;
    }

    const recordPath = path.resolve(cwd, source.path);
    const nextDocumentPath = remapDocumentPath(source.document_path, cwd, currentFolder, nextFolder);
    const nextSource = {
      ...source,
      storage_folder: nextStorageFolder,
      document_path: nextDocumentPath,
      updated_at: formatDate(new Date()),
    };
    await writeFile(recordPath, `${JSON.stringify(nextSource, null, 2)}\n`, "utf8");
  }
}

async function deleteSourcesForFolder(sourceDir, docsDir, cwd, folderName) {
  const sources = await listSyncSources(sourceDir, docsDir, cwd);

  for (const source of sources) {
    if (normalizeStorageFolder(source.storage_folder || "") !== folderName) {
      continue;
    }

    const recordPath = path.resolve(cwd, source.path);
    await rm(recordPath, { force: true });
  }
}

async function deleteSourcesForDocument(sourceDir, docsDir, cwd, documentPath) {
  const sources = await listSyncSources(sourceDir, docsDir, cwd);

  for (const source of sources) {
    if (String(source.document_path || "").trim() !== String(documentPath || "").trim()) {
      continue;
    }

    const recordPath = path.resolve(cwd, source.path);
    await rm(recordPath, { force: true });
  }
}

function remapStorageFolder(value, currentFolder, nextFolder) {
  const normalizedValue = normalizeStorageFolder(value);
  if (!normalizedValue) {
    return "";
  }
  if (normalizedValue === currentFolder) {
    return nextFolder;
  }
  if (normalizedValue.startsWith(`${currentFolder}/`)) {
    return `${nextFolder}${normalizedValue.slice(currentFolder.length)}`;
  }
  return normalizedValue;
}

function remapDocumentPath(value, cwd, currentFolder, nextFolder) {
  const relativePath = String(value || "").trim();
  if (!relativePath) {
    return "";
  }

  const docsRoot = path.resolve(cwd, "data", "internal");
  const currentDir = resolveStorageFolderPath(docsRoot, currentFolder);
  const absolutePath = path.resolve(cwd, relativePath);
  const relativeToFolder = path.relative(currentDir, absolutePath);
  if (relativeToFolder.startsWith("..") || path.isAbsolute(relativeToFolder)) {
    return relativePath;
  }

  const nextDir = resolveStorageFolderPath(docsRoot, nextFolder);
  return toDisplayPath(cwd, path.join(nextDir, relativeToFolder));
}

function normalizeAccessLevelValue(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "advanced") {
    return "advanced";
  }
  if (normalized === "sensitive") {
    return "sensitive";
  }
  return "basic";
}

function extractMarkdownTitle(text) {
  return text.match(/^#\s+(.+)$/m)?.[1]?.trim() || "";
}

