import { tokenize } from "./retrieval.mjs";

export function buildInternalFallbackAnswer(question, matches) {
  const focus = matches[0];
  if (!focus) {
    return "Toi chua tim thay tai lieu noi bo du manh de tra loi cau hoi nay.";
  }

  const evidence = matches
    .slice(0, 2)
    .map((match) => {
      const snippet = pickSnippet(match.text, question);
      return `Trong tai lieu "${match.title}", thong tin lien quan nhat la: ${snippet}`;
    })
    .join(" ");

  return [
    "Toi tim thay noi dung phu hop trong kho du lieu noi bo.",
    evidence,
    "Neu ban can cau tra loi tu nhien hon, hay cau hinh GEMINI_API_KEY de bat lop tong hop AI.",
  ].join(" ");
}

export function buildNoSearchAnswer(reason = "") {
  const suffix = reason ? ` Chi tiết: ${reason}` : "";
  return [
    "Tôi không tìm thấy dữ liệu nội bộ phù hợp.",
    "Nhánh Search Internet của hệ thống này đang dùng Gemini API với Google Search grounding, nhưng hiện tại chưa có cấu hình hợp lệ để thực thi.",
    `Hãy thêm GEMINI_API_KEY vào file .env rồi thử lại.${suffix}`,
  ].join(" ");
}

export function buildWebSearchErrorAnswer(reason = "") {
  const detail = reason ? ` Chi tiết: ${reason}` : "";
  return [
    "Tôi không tìm thấy dữ liệu nội bộ phù hợp.",
    "Gemini đã được cấu hình, nhưng nhánh Search Internet không thể hoàn tất yêu cầu lúc này.",
    `Hệ thống đã fallback về chế độ thông báo lỗi để tránh trả lời sai.${detail}`,
  ].join(" ");
}

export function buildSuggestions(route, question, sources) {
  if (route === "internal") {
    const first = sources[0]?.title;
    return [
      first
        ? `Trich dan chinh xac hon tu tai lieu "${first}".`
        : "Trich dan lai tai lieu noi bo lien quan nhat.",
      "Hoi sau ve dieu kien ap dung, thoi han hoac nguoi phe duyet.",
      `Tom tat cau tra loi cho ban lanh dao tu cau hoi: "${question}".`,
    ];
  }

  if (route === "web") {
    const first = sources[0]?.title || "nguon web dau tien";
    return [
      `So sanh thong tin giua ${first} va cac nguon con lai.`,
      "Yeu cau ban tom tat ngan hon hoac chi ra rui ro can luu y.",
      "Gioi han pham vi tim kiem theo quoc gia, linh vuc hoac thoi gian.",
    ];
  }

  return [
    "Them GEMINI_API_KEY vao .env de bat Internet search va AI tong hop.",
    "Nap them tai lieu noi bo vao kho de tang kha nang tra loi tai cho.",
    "Dat cau hoi cu the hon de he thong truy hoi tai lieu tot hon.",
  ];
}

export function buildWorkflow({
  route,
  aiEnabled,
  usedAi,
  usedWebSearch = false,
  topMatchTitle = "",
  documentCount = 0,
}) {
  const base = [
    {
      label: "AI nhan cau hoi",
      status: "completed",
      detail: "Da chuan hoa truy van va bat dau xu ly.",
    },
    {
      label: "Tim trong Internal Company Data Repository",
      status: "completed",
      detail: `Da quet ${documentCount} tai lieu noi bo.`,
    },
  ];

  if (route === "internal") {
    return [
      ...base,
      {
        label: "Co du lieu",
        status: "completed",
        detail: `Tai lieu phu hop nhat: ${topMatchTitle}.`,
      },
      {
        label: usedAi ? "RAG + AI xu ly" : "RAG + xu ly cuc bo",
        status: usedAi ? "completed" : "warning",
        detail: usedAi
          ? "Da tong hop cau tra loi dua tren tai lieu noi bo."
          : "Dang dung bo tom tat cuc bo vi chua bat Gemini.",
      },
      {
        label: "Tra loi chinh xac",
        status: "completed",
        detail: "Da tra ve cau tra loi va goi y tiep theo.",
      },
    ];
  }

  return [
    ...base,
    {
      label: "Không có dữ liệu",
      status: "completed",
      detail: "Không có tài liệu nội bộ vượt ngưỡng liên quan.",
    },
    {
      label: "Search Internet",
      status: usedWebSearch ? "completed" : "warning",
      detail: usedWebSearch
        ? "Đã gọi Google Search grounding thông qua Gemini API."
        : aiEnabled
          ? "Có AI nhưng không gọi được Google Search grounding."
          : "Chưa cấu hình GEMINI_API_KEY nên không thể tìm web.",
    },
    {
      label: "AI tong hop",
      status: usedAi ? "completed" : "warning",
      detail: usedAi
        ? "Đã tổng hợp từ kết quả Internet và trích nguồn."
        : "Không thể tổng hợp vì nhánh Internet chưa sẵn sàng.",
    },
    {
      label: "Tra loi + goi y",
      status: "completed",
      detail: "Đã trả về kết quả cuối cùng cho người dùng.",
    },
  ];
}

export function pickSnippet(text, question) {
  const sentences = text
    .replace(/\r/g, " ")
    .split(/(?<=[.!?])\s+/)
    .map((sentence) => sentence.trim())
    .filter(Boolean);

  if (sentences.length === 0) {
    return shorten(text, 180);
  }

  const keywords = new Set(tokenize(question));
  const ranked = sentences
    .map((sentence) => {
      const score = tokenize(sentence).reduce(
        (sum, token) => sum + (keywords.has(token) ? 1 : 0),
        0,
      );
      return { sentence, score };
    })
    .sort((left, right) => right.score - left.score || right.sentence.length - left.sentence.length);

  return shorten(ranked[0]?.sentence || sentences[0], 220);
}

function shorten(text, maxLength) {
  const normalized = text.replace(/\s+/g, " ").trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }

  return `${normalized.slice(0, maxLength - 1).trim()}...`;
}
