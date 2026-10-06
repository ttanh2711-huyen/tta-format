/**
 * Cách ghi số văn bản và thời gian ban hành. Hai chuẩn quy định giống nhau:
 * NĐ 30 Phụ lục I Phần I Mục II.3.c và II.4.b; HD 05 Phần II Mục I.3.2.2 và I.4.2.2.
 */

/** Số văn bản: số nhỏ hơn 10 ghi thêm số 0 phía trước ("07"). */
export function formatSo(n: number): string {
  if (!Number.isInteger(n) || n < 1) {
    throw new RangeError("Số văn bản phải là số nguyên dương.");
  }
  return n < 10 ? `0${n}` : String(n);
}

function daysInMonth(month: number, year: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** Thời gian ban hành: ngày nhỏ hơn 10 và tháng 1, 2 ghi thêm số 0 ("ngày 05 tháng 01 năm 2020"). */
export function formatNgay(day: number, month: number, year: number): string {
  if (!Number.isInteger(year) || year < 1000 || year > 9999) {
    throw new RangeError("Năm phải là số nguyên gồm 4 chữ số.");
  }
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new RangeError("Tháng phải từ 1 đến 12.");
  }
  if (!Number.isInteger(day) || day < 1 || day > daysInMonth(month, year)) {
    throw new RangeError(`Tháng ${month} năm ${year} không có ngày ${day}.`);
  }
  const dd = day < 10 ? `0${day}` : String(day);
  const mm = month < 3 ? `0${month}` : String(month);
  return `ngày ${dd} tháng ${mm} năm ${year}`;
}
