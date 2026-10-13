import { buildRegistry, initialState, type FormatValues } from "../settings/formatOptions";
import type { StandardProfile } from "../standards/types";

/** Giá trị Tuỳ chọn nhanh mặc định của chuẩn, giống lúc người dùng chưa chỉnh gì. */
export function buildFormatValuesForTest(profile: StandardProfile): FormatValues {
  return initialState(buildRegistry(profile)).values;
}
