/**
 * Chuyển mã VBA viết bằng tiếng Việt có dấu thành mã chỉ gồm ký tự ASCII (T29): trình soạn VBA của Word 2007 trở về
 * sau lưu mã theo bảng mã ANSI nên chữ có dấu bị hỏng khi nhập tệp .bas. Mỗi chuỗi có ký tự ngoài ASCII được thay bằng
 * U("...\uXXXX...") (hàm U giải mã lúc chạy); chú thích được bỏ dấu.
 */

/** Bỏ dấu tiếng Việt (dùng cho chú thích). */
export function stripDiacritics(text: string): string {
  return text
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .replace(/[^\x20-\x7e\t]/g, "?");
}

const escapeLiteral = (body: string): string =>
  [...body]
    .map((ch) => {
      const code = ch.codePointAt(0) ?? 0;
      if (code < 128) return ch;
      if (code > 0xffff) return "?";
      return `\\u${code.toString(16).toUpperCase().padStart(4, "0")}`;
    })
    .join("");

/** Một dòng mã VBA: chuỗi có dấu thành U("..."), phần chú thích (sau dấu ' ngoài chuỗi) bỏ dấu. */
function convertLine(line: string): string {
  let out = "";
  let i = 0;
  while (i < line.length) {
    const ch = line[i];
    if (ch === '"') {
      let j = i + 1;
      let body = "";
      while (j < line.length) {
        if (line[j] === '"' && line[j + 1] === '"') {
          body += '""';
          j += 2;
        } else if (line[j] === '"') {
          break;
        } else {
          body += line[j];
          j += 1;
        }
      }
      out += [...body].some((ch) => ch.charCodeAt(0) > 127) ? `U("${escapeLiteral(body)}")` : `"${body}"`;
      i = j + 1;
    } else if (ch === "'") {
      out += stripDiacritics(line.slice(i));
      break;
    } else {
      out += ch;
      i += 1;
    }
  }
  return out;
}

/** Cả mô-đun VBA thành ASCII, xuống dòng kiểu Windows (CRLF) như trình soạn VBA cần. */
export function toAsciiVba(source: string): string {
  return source
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map(convertLine)
    .join("\r\n");
}
