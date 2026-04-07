import test from "node:test";
import assert from "node:assert/strict";

import {
  hasConfidentInternalMatch,
  isConversationalQuery,
  searchDocuments,
  selectTopDocumentMatches,
} from "../lib/retrieval.mjs";

const chunks = [
  {
    id: "1",
    documentId: "handbook",
    title: "So tay nhan vien 2026",
    relativePath: "data/internal/employee-handbook.md",
    text: "Nhan vien chinh thuc duoc huong 15 ngay nghi phep nam cho moi nam lam viec day du.",
  },
  {
    id: "2",
    documentId: "sla",
    title: "Chinh sach SLA cham soc khach hang",
    relativePath: "data/internal/customer-support-sla.md",
    text: "Voi khach hang Enterprise, doi ngu ho tro phai gui phan hoi dau tien trong vong 30 phut doi voi muc do su co P1.",
  },
];

test("uu tien tai lieu nghi phep cho cau hoi nghi phep", () => {
  const matches = selectTopDocumentMatches(
    searchDocuments("Nhan vien duoc nghi phep bao nhieu ngay?", chunks),
    3,
  );

  assert.equal(matches[0].documentId, "handbook");
  assert.equal(hasConfidentInternalMatch("Nhan vien duoc nghi phep bao nhieu ngay?", matches), true);
});

test("uu tien tai lieu SLA cho cau hoi ve thoi gian phan hoi", () => {
  const matches = selectTopDocumentMatches(
    searchDocuments("Khach hang Enterprise duoc phan hoi dau tien trong bao lau?", chunks),
    3,
  );

  assert.equal(matches[0].documentId, "sla");
  assert.equal(
    hasConfidentInternalMatch("Khach hang Enterprise duoc phan hoi dau tien trong bao lau?", matches),
    true,
  );
});

test("khong nhan nham cau hoi ngoai mien tri thuc noi bo", () => {
  const matches = selectTopDocumentMatches(
    searchDocuments("Nui cao nhat the gioi cao bao nhieu met?", chunks),
    3,
  );

  assert.equal(hasConfidentInternalMatch("Nui cao nhat the gioi cao bao nhieu met?", matches), false);
});

test("fallback web cho cau hoi can du lieu thi truong hien tai", () => {
  const matches = selectTopDocumentMatches(
    searchDocuments("Gia vang the gioi hien nay dang tang hay giam?", chunks),
    3,
  );

  assert.equal(hasConfidentInternalMatch("Gia vang the gioi hien nay dang tang hay giam?", matches), false);
});

test("uu tien do lien quan thuc su hon metadata boost", () => {
  const ranked = searchDocuments(
    "Nhan vien chinh thuc duoc huong bao nhieu ngay nghi phep nam?",
    [
      {
        id: "sheet-1",
        documentId: "sheet",
        title: "123",
        relativePath: "data/internal/123.md",
        text: "Ngay cap nhat bang gia va ngay giao hang cho nha phan phoi.",
        metadata: {
          source_type: "sheet",
          status: "active",
          canonical: "true",
        },
      },
      {
        id: "handbook-1",
        documentId: "handbook",
        title: "So tay nhan vien 2026",
        relativePath: "data/internal/employee-handbook.md",
        text: "Nhan vien chinh thuc duoc huong 15 ngay nghi phep nam cho moi nam lam viec day du.",
        metadata: {
          source_type: "manual",
          status: "active",
          canonical: "true",
        },
      },
    ],
    { limit: 5 },
  );

  assert.equal(ranked[0].documentId, "handbook");
  assert.equal(ranked.some((item) => item.documentId === "sheet" && item.overlapCount === 0), false);
});

test("khong coi match yeu la du tu tin de tra loi noi bo", () => {
  const weakMatches = [
    {
      documentId: "sheet",
      score: 4.9341,
      overlapCount: 3,
      coverage: 0.3,
      metadata: {
        source_type: "sheet",
        status: "active",
        canonical: "true",
      },
    },
  ];

  assert.equal(
    hasConfidentInternalMatch(
      "Nhan vien chinh thuc duoc huong bao nhieu ngay nghi phep nam?",
      weakMatches,
    ),
    false,
  );
});

test("nhan dien loi chao de khong dua vao RAG", () => {
  assert.equal(isConversationalQuery("chao ban"), true);
  assert.equal(isConversationalQuery("xin chao"), true);
  assert.equal(isConversationalQuery("Nhan vien duoc nghi phep bao nhieu ngay?"), false);
});
