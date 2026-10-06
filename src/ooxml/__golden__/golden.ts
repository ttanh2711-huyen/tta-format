import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

/**
 * So kết quả với file mẫu đối chiếu (golden) trong thư mục này. Chạy với biến môi trường
 * UPDATE_GOLDEN=1 để tạo hoặc cập nhật file, rồi **đọc lại diff trước khi commit**.
 */
export function expectMatchesGolden(fileName: string, actual: string): void {
  const path = join(__dirname, fileName);
  if (process.env.UPDATE_GOLDEN === "1") {
    writeFileSync(path, actual, "utf8");
    return;
  }
  if (!existsSync(path)) {
    throw new Error(`Thiếu file mẫu ${fileName}. Chạy UPDATE_GOLDEN=1 npx jest để tạo, đọc lại rồi mới commit.`);
  }
  expect(actual).toBe(readFileSync(path, "utf8"));
}
