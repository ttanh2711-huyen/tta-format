/**
 * Phát hiện tính năng (T23): bảng các tính năng của add-in và bộ API Word mỗi tính năng cần. Dùng để hiện cho người dùng
 * tính năng nào chạy được trên bản Word đang mở và tính năng nào không (kèm lý do), thay vì để lỗi xuất hiện sau khi bấm.
 * Mỗi mô-đun `office/*` vẫn tự kiểm tra trước khi gọi; bảng này phải khớp với chúng (có test quét mã nguồn).
 */

export type ApiSet = "WordApi" | "WordApiDesktop";

export interface Requirement {
  set: ApiSet;
  version: string;
  /** Bản Word (Microsoft 365) đầu tiên có bộ API này, hiện cho người dùng. */
  build: string;
}

export interface Feature {
  id: string;
  label: string;
  /** Cần đủ mọi yêu cầu. */
  requires: readonly Requirement[];
  /** Tính năng chạy một phần được khi thiếu các yêu cầu này (dùng đường thay thế). */
  optional?: readonly Requirement[];
}

const WORD_API_1_3: Requirement = { set: "WordApi", version: "1.3", build: "2016 trở lên" };
const WORD_API_1_6: Requirement = { set: "WordApi", version: "1.6", build: "2308" };
const DESKTOP_1_1: Requirement = { set: "WordApiDesktop", version: "1.1", build: "2408" };
const DESKTOP_1_3: Requirement = { set: "WordApiDesktop", version: "1.3", build: "2507" };
const DESKTOP_1_4: Requirement = { set: "WordApiDesktop", version: "1.4", build: "2508" };

export const FEATURES: readonly Feature[] = [
  { id: "pageSetup", label: "Thiết lập trang (khổ giấy, lề, khoảng cách header)", requires: [DESKTOP_1_3] },
  { id: "styles", label: "Cài bộ style của chuẩn", requires: [WORD_API_1_6], optional: [DESKTOP_1_1] },
  { id: "formatBody", label: "Định dạng phần thân và bảng biểu", requires: [WORD_API_1_6], optional: [DESKTOP_1_3, DESKTOP_1_4] },
  { id: "pageNumber", label: "Đánh số trang", requires: [DESKTOP_1_3] },
  { id: "check", label: "Kiểm tra văn bản và sửa lỗi", requires: [DESKTOP_1_3] },
  { id: "insertBlocks", label: "Chèn khối đầu, khối cuối, dấu khẩn, phụ lục", requires: [WORD_API_1_3] },
  { id: "createDocument", label: "Tạo văn bản mới", requires: [WORD_API_1_6, DESKTOP_1_3] },
  { id: "documentStandard", label: "Ghi và đọc chuẩn của tài liệu", requires: [WORD_API_1_3] },
];

export type IsSetSupported = (set: ApiSet, version: string) => boolean;

export interface FeatureStatus {
  feature: Feature;
  supported: boolean;
  /** Yêu cầu còn thiếu (bắt buộc). */
  missing: Requirement[];
  /** Yêu cầu tuỳ chọn còn thiếu: tính năng vẫn chạy nhưng có bước dùng đường thay thế. */
  missingOptional: Requirement[];
}

export function evaluateFeatures(isSupported: IsSetSupported, features: readonly Feature[] = FEATURES): FeatureStatus[] {
  return features.map((feature) => {
    const missing = feature.requires.filter((r) => !isSupported(r.set, r.version));
    const missingOptional = (feature.optional ?? []).filter((r) => !isSupported(r.set, r.version));
    return { feature, supported: missing.length === 0, missing, missingOptional };
  });
}

/** Câu tiếng Việt cho người dùng khi tính năng không chạy được; chuỗi rỗng nếu chạy được. */
export function describeMissing(status: FeatureStatus): string {
  if (status.supported) return "";
  const needs = status.missing.map((r) => `${r.set} ${r.version} (Microsoft 365 bản ${r.build} trở lên)`).join(", ");
  return `"${status.feature.label}" cần ${needs}; bản Word này chưa có nên tính năng bị tắt.`;
}

/** Câu cảnh báo khi tính năng chạy nhưng thiếu bộ API tuỳ chọn; rỗng nếu không thiếu. */
export function describeDegraded(status: FeatureStatus): string {
  if (!status.supported || status.missingOptional.length === 0) return "";
  const needs = status.missingOptional.map((r) => `${r.set} ${r.version}`).join(", ");
  return `"${status.feature.label}" vẫn chạy nhưng thiếu ${needs}, một số bước dùng cách thay thế hoặc bị bỏ qua.`;
}

/** Đọc bộ API của Word đang mở; ngoài Word thì coi như không có gì. */
export function officeIsSupported(): IsSetSupported {
  return (set, version) => {
    try {
      return Office.context.requirements.isSetSupported(set, version);
    } catch {
      return false;
    }
  };
}
