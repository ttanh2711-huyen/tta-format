import type { LevelSpec } from "../../ooxml/numbering";

/**
 * Các cấp đánh số của văn bản Đảng (TD6): "Phần" hoặc "Chương" với số La Mã hoặc số bằng chữ ("Chương thứ hai"), "Mục"
 * với số Ả Rập, "Điều N." chạy liên tục, Khoản "1.", Điểm "a)". Hướng dẫn 05 cho phép ghi số Phần, Chương bằng chữ số La
 * Mã hoặc bằng chữ; số bằng chữ dùng định dạng đếm tiếng Việt của Word (kết quả hiển thị cần kiểm tra trên máy thật).
 */

export type PartyTopLevel = "Phần" | "Chương";
export type PartyNumbering = "roman" | "words";

export function partyRegulationLevels(top: PartyTopLevel, numbering: PartyNumbering, firstLineTwips: number): LevelSpec[] {
  return [
    numbering === "roman"
      ? { id: "chuong", fmt: "upperRoman", text: `${top} %1`, bold: true, labelOnly: true }
      : { id: "chuong", fmt: "vietnameseCounting", text: `${top} thứ %1`, bold: true, labelOnly: true },
    { id: "muc", fmt: "decimal", text: "Mục %2", bold: true, labelOnly: true },
    { id: "dieu", fmt: "decimal", text: "Điều %3.", bold: true, continuous: true, firstLineTwips },
    { id: "khoan", fmt: "decimal", text: "%4.", bold: false, firstLineTwips },
    { id: "diem", fmt: "lowerLetter", text: "%5)", bold: false, firstLineTwips },
  ];
}
