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
 * - agreement: hợp đồng, bản ghi nhớ, bản thỏa thuận (các bên, các Điều, khối ký 2 đến 4 bên);
 * - form: giấy tờ và mẫu nhỏ (giấy mời, giấy giới thiệu, công điện, giấy nghỉ phép, thư công): các mục cần điền;
 * - congress: văn bản đại hội của Đảng (tên đại hội trên thể loại, ký hiệu ĐH);
 * - joint: văn bản liên cơ quan của Đảng (hai khối ký).
 */
export type TemplateFamily = "congVan" | "titled" | "toTrinh" | "decision" | "regulation" | "minutes" | "agreement" | "form" | "congress" | "joint";

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
    "giay-to": { family: "form" },
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
    "giay-moi": { family: "form" },
    "giay-gioi-thieu": { family: "form" },
    "cong-dien": { family: "form" },
    "giay-nghi-phep": { family: "form" },
    "thu-cong": { family: "form" },
  },
  dang: {
    "to-trinh": { family: "toTrinh" },
    "nghi-quyet": { family: "regulation" },
    "quyet-dinh": { family: "regulation" },
    "quy-dinh": { family: "regulation" },
    "quy-che": { family: "regulation" },
    "bien-ban": { family: "minutes" },
  },
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

/** Mục không đi theo một loại trong `docTypes` (văn bản đại hội, văn bản liên cơ quan của Đảng). */
const EXTRAS: Record<Headers, TemplateEntry[]> = {
  quocHieu: [],
  dang: [
    { key: "van-ban-dai-hoi", docTypeId: "van-ban-dai-hoi", label: "Văn bản đại hội (ký hiệu ĐH)", typeName: "", abbr: "ĐH", templateRef: null, family: "congress", supported: true },
    { key: "lien-co-quan", docTypeId: "lien-co-quan", label: "Văn bản liên cơ quan (hai khối ký)", typeName: "", abbr: null, templateRef: null, family: "joint", supported: true },
  ],
};

/** Mọi loại của chuẩn, theo thứ tự trong `docTypes`, rồi các mục thêm; mỗi loại có đúng một mục (quyết định cá biệt có hai biến thể). */
export function templatesFor(profile: StandardProfile): TemplateEntry[] {
  const byLayout = FAMILIES[profile.headerKind];
  const byId = BY_ID[profile.headerKind];
  return [...profile.docTypes.flatMap((type) => entryFor(type, byId[type.id] ?? byLayout[type.layout])), ...EXTRAS[profile.headerKind]];
}

export function findTemplate(profile: StandardProfile, key: string): TemplateEntry | undefined {
  return templatesFor(profile).find((entry) => entry.key === key);
}
