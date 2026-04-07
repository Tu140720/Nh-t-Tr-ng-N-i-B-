export function isGeminiEnabled(config) {
  return Boolean(config.gemini.apiKey);
}

export async function extractTextFromImage(config, payload) {
  const models = resolveInternalModels(config);
  let lastError = null;

  for (let index = 0; index < models.length; index += 1) {
    const model = models[index];

    try {
      const response = await generateContent(config, {
        systemInstruction:
          "Ban la bo OCR tieng Viet. Hay doc va trich xuat toan bo van ban nhin thay trong anh. Giữ nguyen so lieu, bang bieu va cac dong quan trong. Khong giai thich them.",
        parts: buildImageOcrParts(payload),
        model,
      });

      return extractText(response);
    } catch (error) {
      lastError = error;

      const hasNextModel = index < models.length - 1;
      if (!hasNextModel || !isRetryableGeminiError(error)) {
        break;
      }
    }
  }

  throw lastError ?? new Error("Không thể OCR ảnh bằng Gemini.");
}

export async function answerWithInternalContext(config, payload) {
  const context = payload.matches
    .slice(0, 3)
    .map(
      (match, index) =>
        `[Tai lieu ${index + 1}] ${match.title} (${match.relativePath})\n${match.text}`,
    )
    .join("\n\n---\n\n");

  const models = resolveInternalModels(config);
  let lastError = null;

  for (let index = 0; index < models.length; index += 1) {
    const model = models[index];

    try {
      const response = await generateContent(config, {
        systemInstruction:
          "Bạn là trợ lý tri thức doanh nghiệp. Hãy trả lời bằng tiếng Việt có dấu, ngắn gọn, ưu tiên độ chính xác. Chỉ được sử dụng ngữ cảnh nội bộ đã cung cấp. Nếu ngữ cảnh chưa đủ, hãy nói rõ giới hạn.",
        contents: `Cau hoi: ${payload.question}\n\nNgu canh noi bo:\n${context}`,
        model,
      });

      return extractText(response);
    } catch (error) {
      lastError = error;

      const hasNextModel = index < models.length - 1;
      if (!hasNextModel || !isRetryableGeminiError(error)) {
        break;
      }
    }
  }

  throw lastError ?? new Error("Không thể gọi Gemini cho nhánh nội bộ.");
}

export async function answerWithWebSearch(config, payload) {
  const models = resolveWebSearchModels(config);
  let lastError = null;

  for (let index = 0; index < models.length; index += 1) {
    const model = models[index];

    try {
      const response = await generateContent(config, {
        systemInstruction:
          "Bạn là trợ lý nghiên cứu. Hãy trả lời bằng tiếng Việt có dấu, tổng hợp ngắn gọn và hữu ích. Nếu được grounding từ web, hãy ưu tiên thông tin cập nhật và để nguồn rõ ràng.",
        contents: payload.question,
        tools: [{ google_search: {} }],
        model,
      });

      return {
        answer: extractText(response),
        sources: extractGroundedSources(response),
        model,
      };
    } catch (error) {
      lastError = error;

      const hasNextModel = index < models.length - 1;
      if (!hasNextModel || !isRetryableGeminiError(error)) {
        break;
      }
    }
  }

  throw lastError ?? new Error("Không thể gọi Gemini web search.");
}

async function generateContent(config, payload) {
  let response = await callGenerateContent(config, payload, { skipThinkingConfig: false });

  if (response.ok) {
    return response.json();
  }

  let error = await buildGeminiError(response, payload.model);

  if (supportsThinkingFallback(error)) {
    response = await callGenerateContent(config, payload, { skipThinkingConfig: true });
    if (response.ok) {
      return response.json();
    }
    error = await buildGeminiError(response, payload.model);
  }

  throw error;
}

async function callGenerateContent(config, payload, options) {
  const body = {
    contents: [
      {
        role: "user",
        parts: payload.parts || [{ text: payload.contents }],
      },
    ],
  };

  if (payload.systemInstruction) {
    body.system_instruction = {
      parts: [{ text: payload.systemInstruction }],
    };
  }

  if (payload.tools?.length) {
    body.tools = payload.tools;
  }

  if (config.gemini.thinkingLevel && !options.skipThinkingConfig) {
    body.generationConfig = {
      thinkingConfig: {
        thinkingLevel: config.gemini.thinkingLevel,
      },
    };
  }

  return fetch(
    `${config.gemini.baseUrl}/models/${payload.model}:generateContent`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": config.gemini.apiKey,
      },
      body: JSON.stringify(body),
    },
  );
}

function buildImageOcrParts(payload) {
  const prompt =
    payload.prompt ||
    "Trich xuat van ban co trong anh. Neu anh la bang du lieu, giu cau truc dong cot o dang de doc.";

  return [
    { text: prompt },
    {
      inline_data: {
        mime_type: payload.mimeType,
        data: payload.dataBase64,
      },
    },
  ];
}

async function buildGeminiError(response, model) {
  let message = "";

  try {
    const payload = await response.json();
    message = payload?.error?.message || "";
  } catch {
    message = "";
  }

  if (!message) {
    try {
      message = await response.text();
    } catch {
      message = "";
    }
  }

  const error = new Error(
    `Gemini API loi ${response.status} (${model}): ${message.trim() || response.statusText || "Khong ro nguyen nhan."}`,
  );

  error.name = "GeminiApiError";
  error.status = response.status;
  error.model = model;
  return error;
}

function extractText(response) {
  const parts = response?.candidates?.[0]?.content?.parts ?? [];
  return parts
    .map((part) => part?.text)
    .filter((text) => typeof text === "string" && text.trim())
    .join("\n\n")
    .trim();
}

function extractGroundedSources(response) {
  const groundingMetadata = response?.candidates?.[0]?.groundingMetadata;
  const chunks = groundingMetadata?.groundingChunks ?? [];
  const sources = [];
  const seen = new Set();

  for (const chunk of chunks) {
    const web = chunk.web;
    if (!web?.uri || seen.has(web.uri)) {
      continue;
    }

    seen.add(web.uri);
    sources.push({
      type: "web",
      title: web.title || web.uri,
      url: web.uri,
      snippet: "",
    });
  }

  return sources.slice(0, 8);
}

function resolveWebSearchModels(config) {
  const candidates = [
    config.gemini.webModel,
    config.gemini.model,
    ...(config.gemini.fallbackModels || []),
    config.gemini.webFallbackModel,
    ...(config.gemini.webFallbackModels || []),
  ];

  return [...new Set(candidates.filter(Boolean))];
}

function resolveInternalModels(config) {
  const candidates = [config.gemini.model, ...(config.gemini.fallbackModels || [])];
  return [...new Set(candidates.filter(Boolean))];
}

function isRetryableGeminiError(error) {
  return Boolean(error && [429, 500, 502, 503, 504].includes(error.status));
}

function supportsThinkingFallback(error) {
  return (
    error?.status === 400 &&
    typeof error.message === "string" &&
    error.message.toLowerCase().includes("thinking level is not supported")
  );
}
