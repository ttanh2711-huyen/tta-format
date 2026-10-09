import { MAX_PARENT_LINES } from "../ooxml/headerBlock";
import { MAX_ISSUER_LINES, MAX_PARENT_LINES as MAX_PARTY_PARENT_LINES } from "../standards/hd05/headerBlock";
import { assertAbbreviation } from "../standards/soKyHieu";
import type { StandardId, StandardProfile } from "../standards/types";
import type { KeyValueStore } from "./storage";

/**
 * Hồ sơ cơ quan (T22): bộ thông tin dùng lại khi tạo văn bản mới, thuộc một chuẩn. Gồm các dòng cơ quan chủ quản (cấp trên),
 * tên cơ quan ban hành, địa danh, chữ viết tắt cơ quan trong ký hiệu, chức vụ và họ tên người ký thường dùng. Mỗi chuẩn có
 * thể đặt một hồ sơ mặc định để điền sẵn khi mở trình tạo văn bản.
 */

export interface AgencyProfile {
  id: string;
  standardId: StandardId;
  /** Tên gợi nhớ trong danh sách. */
  label: string;
  parentLines: string[];
  /** Cơ quan ban hành: một dòng với chuẩn Chính quyền, từ một đến ba dòng với chuẩn Đảng. */
  issuerLines: string[];
  place: string;
  /** Chữ viết tắt cơ quan nối bằng gạch nối, ví dụ "CCTL-ĐĐPCTT". */
  abbrs: string;
  positionLines: string[];
  signerName: string;
  isDefault?: boolean;
}

export type ProfileInput = Omit<AgencyProfile, "id" | "isDefault"> & { id?: string };
export type SaveResult = { ok: true; profile: AgencyProfile } | { ok: false; reasons: string[] };

export const PROFILES_KEY = "nd30.agencyProfiles";
const MAX_POSITION_LINES = 3;

function clean(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/** Số dòng tối đa của cơ quan chủ quản và của cơ quan ban hành theo chuẩn. */
export function agencyLimits(profile: StandardProfile): { parent: number; issuer: number } {
  return profile.headerKind === "dang" ? { parent: MAX_PARTY_PARENT_LINES, issuer: MAX_ISSUER_LINES } : { parent: MAX_PARENT_LINES, issuer: 1 };
}

/** Lý do không lưu được hồ sơ; rỗng nếu hợp lệ. Giải thích bằng tiếng Việt, nêu rõ giới hạn của chuẩn. */
export function validateProfile(profile: StandardProfile, input: ProfileInput): string[] {
  const errors: string[] = [];
  const limits = agencyLimits(profile);
  const parents = input.parentLines.map(clean).filter((line) => line !== "");
  const issuers = input.issuerLines.map(clean).filter((line) => line !== "");
  if (clean(input.label) === "") errors.push("Hãy đặt tên gợi nhớ cho hồ sơ.");
  if (issuers.length === 0) errors.push("Hãy nhập tên cơ quan ban hành.");
  if (issuers.length > limits.issuer) errors.push(`Chuẩn ${profile.name} cho tối đa ${limits.issuer} dòng cơ quan ban hành, hồ sơ đang có ${issuers.length}.`);
  if (parents.length > limits.parent) errors.push(`Chuẩn ${profile.name} cho tối đa ${limits.parent} dòng cơ quan ${profile.headerKind === "dang" ? "cấp trên" : "chủ quản"}, hồ sơ đang có ${parents.length}.`);
  const positions = input.positionLines.map(clean).filter((line) => line !== "");
  if (positions.length > MAX_POSITION_LINES) errors.push(`Chức vụ người ký tối đa ${MAX_POSITION_LINES} dòng, hồ sơ đang có ${positions.length}.`);
  const abbrs = clean(input.abbrs);
  if (abbrs !== "") {
    for (const part of abbrs.split("-").map((item) => item.trim())) {
      try {
        assertAbbreviation(part, "Chữ viết tắt cơ quan");
      } catch (error) {
        errors.push(error instanceof Error ? error.message : String(error));
        break;
      }
    }
  }
  return errors;
}

function isProfile(value: unknown): value is AgencyProfile {
  if (typeof value !== "object" || value === null) return false;
  const p = value as Record<string, unknown>;
  const strings = (v: unknown) => Array.isArray(v) && v.every((item) => typeof item === "string");
  return typeof p.id === "string" && typeof p.standardId === "string" && typeof p.label === "string" && strings(p.parentLines) && strings(p.issuerLines) && typeof p.place === "string" && typeof p.abbrs === "string" && strings(p.positionLines) && typeof p.signerName === "string";
}

function readAll(store: KeyValueStore): AgencyProfile[] {
  try {
    const parsed: unknown = JSON.parse(store.get(PROFILES_KEY) ?? "[]");
    return Array.isArray(parsed) ? parsed.filter(isProfile) : [];
  } catch {
    return [];
  }
}

function writeAll(store: KeyValueStore, profiles: AgencyProfile[]): void {
  store.set(PROFILES_KEY, JSON.stringify(profiles));
}

/** Hồ sơ của một chuẩn, theo thứ tự lưu. */
export function listProfiles(store: KeyValueStore, profile: StandardProfile): AgencyProfile[] {
  return readAll(store).filter((item) => item.standardId === profile.id);
}

export function defaultProfile(store: KeyValueStore, profile: StandardProfile): AgencyProfile | undefined {
  return listProfiles(store, profile).find((item) => item.isDefault);
}

function newId(): string {
  return `agency:${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

/** Thêm hồ sơ mới hoặc cập nhật hồ sơ có `id` trùng. Giá trị ngoài giới hạn của chuẩn thì không lưu. Tên gợi nhớ không trùng trong cùng chuẩn. */
export function saveProfile(store: KeyValueStore, profile: StandardProfile, input: ProfileInput): SaveResult {
  const reasons = validateProfile(profile, input);
  const all = readAll(store);
  const label = clean(input.label);
  if (all.some((item) => item.standardId === profile.id && item.id !== input.id && item.label.toLowerCase() === label.toLowerCase())) reasons.push(`Đã có hồ sơ tên "${label}" của chuẩn ${profile.name}.`);
  if (input.id !== undefined && !all.some((item) => item.id === input.id && item.standardId === profile.id)) reasons.push("Không tìm thấy hồ sơ cần sửa.");
  if (reasons.length > 0) return { ok: false, reasons };
  const existing = all.find((item) => item.id === input.id);
  const saved: AgencyProfile = {
    id: input.id ?? newId(),
    standardId: profile.id,
    label,
    parentLines: input.parentLines.map(clean).filter((line) => line !== ""),
    issuerLines: input.issuerLines.map(clean).filter((line) => line !== ""),
    place: clean(input.place),
    abbrs: clean(input.abbrs),
    positionLines: input.positionLines.map(clean).filter((line) => line !== ""),
    signerName: clean(input.signerName),
    ...(existing?.isDefault ? { isDefault: true as const } : {}),
  };
  writeAll(store, existing ? all.map((item) => (item.id === saved.id ? saved : item)) : [...all, saved]);
  return { ok: true, profile: saved };
}

export function deleteProfile(store: KeyValueStore, id: string): void {
  writeAll(store, readAll(store).filter((item) => item.id !== id));
}

/** Đặt một hồ sơ làm mặc định của chuẩn (bỏ mặc định của hồ sơ khác cùng chuẩn). */
export function setDefaultProfile(store: KeyValueStore, id: string): boolean {
  const all = readAll(store);
  const target = all.find((item) => item.id === id);
  if (!target) return false;
  const withoutDefault = (item: AgencyProfile): AgencyProfile => {
    const copy = { ...item };
    delete copy.isDefault;
    return copy;
  };
  writeAll(store, all.map((item) => (item.standardId !== target.standardId ? item : item.id === id ? { ...item, isDefault: true as const } : withoutDefault(item))));
  return true;
}

/** Mọi hồ sơ của mọi chuẩn, dùng cho xuất cấu hình. */
export function allProfiles(store: KeyValueStore): AgencyProfile[] {
  return readAll(store);
}

/** Thay toàn bộ hồ sơ (dùng khi nhập cấu hình). */
export function replaceProfiles(store: KeyValueStore, profiles: AgencyProfile[]): void {
  writeAll(store, profiles);
}

export function isAgencyProfile(value: unknown): value is AgencyProfile {
  return isProfile(value);
}
