import { checkNd30HeaderFooter } from "../standards/nd30/checks";
import type { StandardProfile } from "../standards/types";
import type { Rule } from "./engine";

/**
 * Các bộ quy tắc khối đầu và khối cuối theo kiểu khối của chuẩn (`headerKind`). Chuẩn Đảng đăng ký ở T14B; hiện
 * chưa có nên chỉ chạy quy tắc chung về trang và đoạn.
 */
export function headerFooterRules(profile: StandardProfile): Rule[] {
  return profile.headerKind === "quocHieu" ? [checkNd30HeaderFooter] : [];
}
