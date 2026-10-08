import { formatNgay } from "../standards/dateNumber";
import type { StandardProfile } from "../standards/types";

/** Chỗ trống chờ điền tay (số chưa vào sổ, ngày chờ ký); dùng khoảng trắng như file mẫu. */
const GAP = "      ";

export type TextResult = { ok: true; text: string } | { ok: false; error: string };

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export interface SoInput {
  /** null: để trống chờ vào sổ. */
  so: number | null;
  typeAbbr: string | null;
  agencyAbbrs: readonly string[];
}

/** Dòng số và ký hiệu theo cách ghi của chuẩn; số để trống thì chừa chỗ trống tại vị trí số. */
export function soText(profile: StandardProfile, input: SoInput): TextResult {
  try {
    if (input.so === null) {
      // Dùng số 1 để chuẩn kiểm tra ký hiệu và dựng dòng, rồi thay số bằng chỗ trống.
      return { ok: true, text: profile.formatSoKyHieu(1, input.typeAbbr, input.agencyAbbrs).replace(/\d+/, GAP) };
    }
    return { ok: true, text: profile.formatSoKyHieu(input.so, input.typeAbbr, input.agencyAbbrs) };
  } catch (error) {
    return { ok: false, error: messageOf(error) };
  }
}

export interface NgayInput {
  place: string;
  day: number | null;
  month: number | null;
  year: number | null;
}

/** Năm nhuận dùng khi chưa biết năm, để 29/2 vẫn được chấp nhận. */
const LEAP_YEAR = 2024;

/** "Địa danh, ngày 05 tháng 01 năm 2020"; phần nào để trống thì chừa chỗ trống. */
export function placeDateText(input: NgayInput): TextResult {
  const place = input.place.replace(/\s+/g, " ").trim();
  if (place === "") return { ok: false, error: "Hãy nhập địa danh." };
  const { day, month, year } = input;
  if (day !== null && (!Number.isInteger(day) || day < 1 || day > 31)) return { ok: false, error: "Ngày phải từ 1 đến 31." };
  if (month !== null && (!Number.isInteger(month) || month < 1 || month > 12)) return { ok: false, error: "Tháng phải từ 1 đến 12." };
  if (year !== null && (!Number.isInteger(year) || year < 1000 || year > 9999)) return { ok: false, error: "Năm phải là số nguyên gồm 4 chữ số." };
  try {
    const parts = { day: GAP, month: GAP, year: GAP };
    // formatNgay kiểm tra ngày có tồn tại trong tháng và năm (kể cả năm nhuận) và thêm số 0.
    if (day !== null && month !== null) {
      const match = /^ngày (\d+) tháng (\d+) năm /.exec(formatNgay(day, month, year ?? LEAP_YEAR));
      if (!match) return { ok: false, error: "Không dựng được ngày tháng." };
      parts.day = match[1];
      parts.month = match[2];
    } else {
      if (day !== null) parts.day = day < 10 ? `0${day}` : String(day);
      if (month !== null) parts.month = month < 3 ? `0${month}` : String(month);
    }
    if (year !== null) parts.year = String(year);
    return { ok: true, text: `${place}, ngày ${parts.day} tháng ${parts.month} năm ${parts.year}` };
  } catch (error) {
    return { ok: false, error: messageOf(error) };
  }
}
