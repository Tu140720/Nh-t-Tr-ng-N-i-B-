const STOP_WORDS = new Set([
  "a",
  "an",
  "and",
  "bao",
  "bi",
  "boi",
  "cho",
  "co",
  "cua",
  "da",
  "de",
  "den",
  "duoc",
  "for",
  "from",
  "hay",
  "how",
  "is",
  "khi",
  "khong",
  "la",
  "lam",
  "mot",
  "neu",
  "nhung",
  "o",
  "or",
  "sao",
  "sau",
  "se",
  "tai",
  "the",
  "thi",
  "this",
  "to",
  "trong",
  "tu",
  "va",
  "ve",
  "voi",
]);

const BM25_K1 = 1.5;
const BM25_B = 0.75;

export function searchDocuments(query, chunks, options = {}) {
  const limit = options.limit ?? 6;
  const queryTokens = uniqueTokens(query);
  const filteredChunks = options.activeOnly === false ? chunks : chunks.filter(isActiveChunk);

  if (queryTokens.length === 0 || filteredChunks.length === 0) {
    return [];
  }

  const preparedChunks = filteredChunks.map((chunk) => {
    const tokens = tokenize(chunk.text);
    const titleTokens = new Set(tokenize(chunk.title));
    const frequencies = countFrequencies(tokens);

    return {
      ...chunk,
      tokens,
      titleTokens,
      frequencies,
      length: Math.max(tokens.length, 1),
    };
  });

  const averageLength =
    preparedChunks.reduce((sum, chunk) => sum + chunk.length, 0) / preparedChunks.length;
  const documentFrequencies = buildDocumentFrequencies(preparedChunks);

  return preparedChunks
    .map((chunk) => {
      let score = 0;
      let overlapCount = 0;

      for (const token of queryTokens) {
        const frequency = chunk.frequencies.get(token) ?? 0;
        if (frequency === 0) {
          continue;
        }

        overlapCount += 1;
        const docsWithToken = documentFrequencies.get(token) ?? 0;
        const idf = Math.log(
          1 + (preparedChunks.length - docsWithToken + 0.5) / (docsWithToken + 0.5),
        );
        const normalized =
          frequency + BM25_K1 * (1 - BM25_B + BM25_B * (chunk.length / averageLength));
        score += idf * ((frequency * (BM25_K1 + 1)) / normalized);
      }

      let titleBoost = 0;
      for (const token of queryTokens) {
        if (chunk.titleTokens.has(token)) {
          titleBoost += 0.8;
        }
      }

      const coverage = overlapCount / queryTokens.length;
      const metadataBoost = computeMetadataBoost(chunk.metadata);
      return {
        ...chunk,
        score: Number((score + titleBoost + metadataBoost).toFixed(4)),
        overlapCount,
        coverage: Number(coverage.toFixed(3)),
      };
    })
    .filter((chunk) => chunk.overlapCount > 0 && chunk.score > 0)
    .sort(compareMatches)
    .slice(0, limit);
}

export function selectTopDocumentMatches(results, limit = 4) {
  const bestByDocument = new Map();

  for (const result of results) {
    const current = bestByDocument.get(result.documentId);
    if (!current || compareMatches(result, current) < 0) {
      bestByDocument.set(result.documentId, result);
    }
  }

  return [...bestByDocument.values()].sort(compareMatches).slice(0, limit);
}

export function hasConfidentInternalMatch(question, matches) {
  const top = matches[0];
  if (!top) {
    return false;
  }

  if (looksLikeFreshWebQuery(question) && (top.coverage < 0.55 || top.overlapCount < 4)) {
    return false;
  }

  if (top.overlapCount < 4 && top.coverage < 0.5) {
    return false;
  }

  return (
    top.score >= 1.2 ||
    (top.score >= 0.85 && top.overlapCount >= 2 && top.coverage >= 0.24) ||
    (top.score >= 0.65 && top.overlapCount >= 3)
  );
}

export function isConversationalQuery(question) {
  const normalized = normalizeForSearch(question)
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!normalized) {
    return false;
  }

  return [
    "chao",
    "chao ban",
    "xin chao",
    "hello",
    "hi",
    "alo",
    "ban la ai",
    "ban co o day khong",
    "cam on",
    "ok",
    "oke",
  ].includes(normalized);
}

export function detectConflictingMatches(matches) {
  const byTopic = new Map();

  for (const match of matches) {
    const topicKey = match.metadata?.topic_key;
    if (!topicKey) {
      continue;
    }

    if (!byTopic.has(topicKey)) {
      byTopic.set(topicKey, []);
    }

    byTopic.get(topicKey).push(match);
  }

  for (const topicMatches of byTopic.values()) {
    if (topicMatches.length < 2) {
      continue;
    }

    const signatures = new Set(topicMatches.map((match) => numberSignature(match.text)));
    if (signatures.size > 1) {
      return topicMatches.sort(compareMatches).slice(0, 3);
    }
  }

  return [];
}

export function tokenize(text) {
  return normalizeForSearch(text)
    .match(/[\p{L}\p{N}]+/gu)
    ?.filter((token) => token.length > 1 && !STOP_WORDS.has(token)) ?? [];
}

function normalizeForSearch(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase();
}

function looksLikeFreshWebQuery(question) {
  const normalized = normalizeForSearch(question);
  return [
    "hien nay",
    "hom nay",
    "moi nhat",
    "the gioi",
    "thi truong",
    "gia vang",
    "ty gia",
    "co phieu",
    "thoi tiet",
    "tin tuc",
  ].some((term) => normalized.includes(term));
}

function uniqueTokens(text) {
  return [...new Set(tokenize(text))];
}

function countFrequencies(tokens) {
  const frequencies = new Map();
  for (const token of tokens) {
    frequencies.set(token, (frequencies.get(token) ?? 0) + 1);
  }
  return frequencies;
}

function buildDocumentFrequencies(chunks) {
  const frequencies = new Map();
  for (const chunk of chunks) {
    const unique = new Set(chunk.tokens);
    for (const token of unique) {
      frequencies.set(token, (frequencies.get(token) ?? 0) + 1);
    }
  }
  return frequencies;
}

function computeMetadataBoost(metadata = {}) {
  let score = 0;

  if (String(metadata.status || "active") === "active") {
    score += 0.8;
  }

  if (String(metadata.canonical || "") === "true") {
    score += 0.6;
  }

  score += sourcePriority(metadata.source_type) * 0.08;
  score += ownerPriority(metadata.owner) * 0.04;
  score += effectiveDateBoost(metadata.effective_date);

  return score;
}

function compareMatches(left, right) {
  return (
    right.score - left.score ||
    right.coverage - left.coverage ||
    right.overlapCount - left.overlapCount ||
    compareMetadataPriority(right.metadata, left.metadata)
  );
}

function compareMetadataPriority(left = {}, right = {}) {
  return (
    effectiveDateScore(left.effective_date) - effectiveDateScore(right.effective_date) ||
    sourcePriority(left.source_type) - sourcePriority(right.source_type) ||
    ownerPriority(left.owner) - ownerPriority(right.owner)
  );
}

function effectiveDateBoost(value) {
  const score = effectiveDateScore(value);
  return score ? Math.min(score / 10_000_000_000_000, 0.4) : 0;
}

function effectiveDateScore(value) {
  const time = Date.parse(String(value || ""));
  return Number.isNaN(time) ? 0 : time;
}

function sourcePriority(value) {
  const source = String(value || "").toLowerCase();
  if (source === "sheet") {
    return 4;
  }
  if (source === "file_upload") {
    return 3;
  }
  if (source === "manual") {
    return 2;
  }
  return 1;
}

function ownerPriority(value) {
  const owner = normalizeForSearch(value);
  if (owner.includes("finance") || owner.includes("tai chinh")) {
    return 4;
  }
  if (owner.includes("sales") || owner.includes("kinh doanh")) {
    return 3;
  }
  if (owner.includes("operations") || owner.includes("van hanh")) {
    return 2;
  }
  return owner ? 1 : 0;
}

function numberSignature(text) {
  return (
    String(text || "")
      .match(/\d+(?:[.,]\d+)?/g)
      ?.join("|") || ""
  );
}

function isActiveChunk(chunk) {
  return String(chunk.metadata?.status || "active") === "active";
}
