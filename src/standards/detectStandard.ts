import { STANDARDS } from "./index";
import type { StandardId, StandardProfile } from "./types";

/**
 * Nhận diện chuẩn của tài liệu theo dòng tiêu đề đầu văn bản (quốc hiệu hoặc tiêu đề Đảng) và quyết định
 * dùng chuẩn nào. Nguyên tắc: tiêu đề rõ ràng mới dùng; mọi trường hợp còn lại là "không chắc" để hỏi người
 * dùng, không bao giờ tự đoán rồi áp dụng định dạng.
 */

export type UnsureReason = "noTitle" | "bothTitles" | "similarTitle";

export type Detection =
  | { status: "empty" }
  | { status: "detected"; standardId: StandardId }
  | { status: "unsure"; reason: UnsureReason; suggestion?: StandardId };

/** Số đoạn có chữ đầu tiên được xét là "phần đầu" văn bản. */
const HEADER_PARAGRAPHS = 30;
/** Tiêu đề có thể bị ngắt thành tối đa bấy nhiêu đoạn liền nhau. */
const MAX_JOINED_LINES = 3;
/** Tỉ lệ sai khác tối đa (so với độ dài tiêu đề) vẫn coi là "tiêu đề viết sai chính tả". */
const SIMILAR_RATIO = 0.2;

/** Chữ hoa, bỏ dấu, bỏ dấu câu, gộp khoảng trắng: so khớp không phụ thuộc cách đặt dấu thanh. */
export function normalizeTitle(text: string): string {
  return text
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .replace(/[đĐ]/g, "D")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

export function editDistance(a: string, b: string): number {
  let previous = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i += 1) {
    const current = [i];
    for (let j = 1; j <= b.length; j += 1) {
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    previous = current;
  }
  return previous[b.length];
}

export function detectStandard(headerTexts: readonly string[], standards: readonly StandardProfile[] = STANDARDS): Detection {
  const lines = headerTexts.map(normalizeTitle).filter((line) => line !== "").slice(0, HEADER_PARAGRAPHS);
  if (lines.length === 0) {
    return { status: "empty" };
  }
  const windows: string[] = [];
  for (let start = 0; start < lines.length; start += 1) {
    for (let length = 1; length <= MAX_JOINED_LINES && start + length <= lines.length; length += 1) {
      windows.push(lines.slice(start, start + length).join(" "));
    }
  }

  const exact = standards.filter((standard) => windows.includes(normalizeTitle(standard.identifyingTitle)));
  if (exact.length === 1) {
    return { status: "detected", standardId: exact[0].id };
  }
  if (exact.length > 1) {
    return { status: "unsure", reason: "bothTitles" };
  }

  let best: { id: StandardId; ratio: number } | undefined;
  for (const standard of standards) {
    const title = normalizeTitle(standard.identifyingTitle);
    for (const window of windows) {
      const ratio = editDistance(window, title) / title.length;
      if (ratio <= SIMILAR_RATIO && (!best || ratio < best.ratio)) {
        best = { id: standard.id, ratio };
      }
    }
  }
  return best ? { status: "unsure", reason: "similarTitle", suggestion: best.id } : { status: "unsure", reason: "noTitle" };
}

export type Resolution =
  | { action: "use"; standardId: StandardId; source: "document" | "detected" | "default" }
  | { action: "ask"; reason: UnsureReason; suggestion: StandardId | undefined };

export interface ResolveInput {
  /** Chuẩn đã ghi trong thuộc tính tuỳ chỉnh của tài liệu (có thể không hợp lệ). */
  stored: string | null;
  /** Người dùng đã chọn "Không hỏi lại cho tài liệu này". */
  noAsk: boolean;
  detection: Detection;
  /** Chuẩn mặc định cho tài liệu mới (cài đặt cục bộ). */
  fallbackId: StandardId;
}

export function resolveStandard(input: ResolveInput, standards: readonly StandardProfile[] = STANDARDS): Resolution {
  const stored = standards.find((standard) => standard.id === input.stored);
  if (stored) {
    return { action: "use", standardId: stored.id, source: "document" };
  }
  const { detection } = input;
  if (detection.status === "detected") {
    return { action: "use", standardId: detection.standardId, source: "detected" };
  }
  if (detection.status === "empty" || input.noAsk) {
    return { action: "use", standardId: input.fallbackId, source: "default" };
  }
  return { action: "ask", reason: detection.reason, suggestion: detection.suggestion };
}
