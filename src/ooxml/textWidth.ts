/**
 * Ước lượng độ rộng chữ Times New Roman từ bảng độ rộng ký tự (đơn vị 1/1000 em) để tính độ dài đường kẻ
 * dưới tên cơ quan và tiêu ngữ. Không đo phông thật nên sai số vài phần trăm; ký tự lạ tính bằng chữ n.
 * ponytail: nâng lên đo bằng canvas nếu cần chính xác hơn.
 */

const REGULAR: Record<string, number> = {
  " ": 250, "-": 333, ".": 250, ",": 250, ":": 278, ";": 278, "/": 278, "(": 333, ")": 333,
  a: 444, b: 500, c: 444, d: 500, e: 444, f: 333, g: 500, h: 500, i: 278, j: 278, k: 500, l: 278, m: 778,
  n: 500, o: 500, p: 500, q: 500, r: 333, s: 389, t: 278, u: 500, v: 500, w: 722, x: 500, y: 500, z: 444,
  A: 722, B: 667, C: 667, D: 722, E: 611, F: 556, G: 722, H: 722, I: 333, J: 389, K: 722, L: 611, M: 889,
  N: 722, O: 722, P: 556, Q: 722, R: 667, S: 556, T: 611, U: 722, V: 722, W: 944, X: 722, Y: 722, Z: 611,
};

const BOLD: Record<string, number> = {
  " ": 250, "-": 333, ".": 250, ",": 250, ":": 333, ";": 333, "/": 278, "(": 333, ")": 333,
  a: 500, b: 556, c: 444, d: 556, e: 444, f: 333, g: 500, h: 556, i: 278, j: 333, k: 556, l: 278, m: 833,
  n: 556, o: 500, p: 556, q: 556, r: 444, s: 389, t: 333, u: 556, v: 500, w: 722, x: 500, y: 500, z: 444,
  A: 722, B: 667, C: 722, D: 722, E: 667, F: 611, G: 778, H: 778, I: 389, J: 500, K: 778, L: 667, M: 944,
  N: 722, O: 778, P: 611, Q: 778, R: 722, S: 556, T: 667, U: 722, V: 722, W: 1000, X: 722, Y: 722, Z: 667,
};

const DIGIT = 500;
const POINT_MM = 25.4 / 72;

function baseLetter(char: string): string {
  if (char === "Đ") return "D";
  if (char === "đ") return "d";
  return char.normalize("NFD").replace(/\p{M}/gu, "");
}

export function textWidthMm(text: string, sizePt: number, bold: boolean): number {
  const table = bold ? BOLD : REGULAR;
  let units = 0;
  for (const char of text) {
    const base = baseLetter(char);
    units += /\d/.test(base) ? DIGIT : (table[base] ?? table.n);
  }
  return (units / 1000) * sizePt * POINT_MM;
}
