import type { DocType, StandardProfile } from "../standards/types";
import type { DecisionKind } from "./quyetDinh";

/**
 * Registry loại văn bản (T17): mỗi loại trong `docTypes` của chuẩn đang chọn được gán một họ mẫu. Wizard chỉ đọc
 * registry này để dựng ô chọn loại; thêm loại mới là thêm một mục vào `FAMILIES`, không sửa wizard. Loại chưa có
 * họ mẫu hiện trong danh sách kèm lý do "chưa hỗ trợ", không bị ẩn.
 */

/**
 * - congVan: công văn (V/v và Kính gửi trong bảng đầu văn bản);
 * - titled: tên loại + trích yếu dưới khối đầu, rồi nội dung (Mẫu 1.4 của Chính quyền, thể loại của Đảng);
 * - toTrinh: như titled nhưng có "Kính trình" (chuẩn Đảng);
 * - decision: nghị quyết, quyết định cá biệt (căn cứ, dòng chốt, Điều đánh số tự động);
 * - regulation: quy chế, quy định (Chương, Mục, Điều, Khoản đánh số tự động);
 * - minutes: biên bản (các mục đầu, khối ký chủ tọa và thư ký);
 * - agreement: hợp đồng, bản ghi nhớ, bản thỏa thuận (các bên, các Điều, khối ký 2 đến 4 bên).
 */
export type TemplateFamily = "congVan" | "titled" | "toTrinh" | "decision" | "regulation" | "minutes" | "agreement";

export interface TemplateEntry {
  /** Khoá duy nhất trong ô chọn; thường là mã loại, loại có hai biến thể thì kèm hậu tố. */
  key: string;
  docTypeId: string;
  label: string;
  /** Tên loại ghi vào văn bản (chữ thường đầu; khối dựng sẽ in hoa). */
  typeName: string;
  /** Chữ viết tắt gợi ý cho ký hiệu văn bản; null nếu loại không có. */
  abbr: string | null;
  /** Mẫu trong văn bản gốc, nếu có. */
  templateRef: string | null;
  family: TemplateFamily | null;
  decisionKind?: DecisionKind;
  supported: boolean;
  /** Lý do chưa hỗ trợ, hiện cho người dùng. */
  reason?: string;
}

interface FamilySpec {
  family: TemplateFamily;
  decisionKinds?: readonly DecisionKind[];
}

type Headers = StandardProfile["headerKind"];

/** Họ mẫu theo kiểu khối đầu của chuẩn rồi theo họ bố cục (`layout`) của loại. */
const FAMILIES: Record<Headers, Record<string, FamilySpec>> = {
  quocHieu: {
    "cong-van": { family: "congVan" },
    "co-ten-loai": { family: "titled" },
    "nghi-quyet": { family: "decision", decisionKinds: ["nghiQuyet"] },
    "quyet-dinh": { family: "decision", decisionKinds: ["quyetDinhTrucTiep", "quyetDinhGianTiep"] },
    "bien-ban": { family: "minutes" },
  },
  dang: {
    "cong-van": { family: "congVan" },
    "co-ten-loai": { family: "titled" },
  },
};

/** Riêng một số loại có họ mẫu khác họ bố cục của chúng. */
const BY_ID: Record<Headers, Record<string, FamilySpec>> = {
  quocHieu: {
    "quy-che": { family: "regulation" },
    "quy-dinh": { family: "regulation" },
    "hop-dong": { family: "agreement" },
    "ban-ghi-nho": { family: "agreement" },
    "ban-thoa-thuan": { family: "agreement" },
  },
  dang: { "to-trinh": { family: "toTrinh" } },
};

const DECISION_LABEL: Record<DecisionKind, string> = {
  nghiQuyet: "Nghị quyết (cá biệt)",
  quyetDinhTrucTiep: "Quyết định cá biệt, quy định trực tiếp",
  quyetDinhGianTiep: "Quyết định cá biệt, quy định gián tiếp (ban hành kèm theo)",
};

function unsupportedReason(type: DocType): string {
  return `chưa có mẫu cho họ bố cục "${type.layout}"${type.template ? ` (${type.template})` : ""}`;
}

function entryFor(type: DocType, spec: FamilySpec | undefined): TemplateEntry[] {
  const base = { docTypeId: type.id, typeName: type.name, abbr: type.abbr, templateRef: type.template };
  if (!spec) return [{ ...base, key: type.id, label: type.name, family: null, supported: false, reason: unsupportedReason(type) }];
  if (spec.family === "decision" && spec.decisionKinds) {
    return spec.decisionKinds.map((kind) => ({
      ...base,
      key: spec.decisionKinds!.length > 1 ? `${type.id}:${kind}` : type.id,
      label: DECISION_LABEL[kind],
      family: "decision" as const,
      decisionKind: kind,
      supported: true,
    }));
  }
  return [{ ...base, key: type.id, label: type.name, family: spec.family, supported: true }];
}

/** Mọi loại của chuẩn, theo thứ tự trong `docTypes`; mỗi loại có đúng một mục (quyết định cá biệt có hai biến thể). */
export function templatesFor(profile: StandardProfile): TemplateEntry[] {
  const byLayout = FAMILIES[profile.headerKind];
  const byId = BY_ID[profile.headerKind];
  return profile.docTypes.flatMap((type) => entryFor(type, byId[type.id] ?? byLayout[type.layout]));
}

export function findTemplate(profile: StandardProfile, key: string): TemplateEntry | undefined {
  return templatesFor(profile).find((entry) => entry.key === key);
}
