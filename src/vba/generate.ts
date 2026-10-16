import { STANDARDS } from "../standards";
import { configModule } from "./configModule";
import { FORMAT_MODULE } from "./formatModule";
import { toAsciiVba } from "./vbaText";

/** Các tệp .bas của bộ macro VBA (T29), theo tên tệp trong thư mục `vba/`. */
export function vbaFiles(): Record<string, string> {
  // TTA_ChonChuan hỏi Có/Không giữa hai chuẩn.
  if (STANDARDS.length !== 2) throw new Error("TTA_ChonChuan chỉ chọn được giữa hai chuẩn; cần sửa macro khi thêm chuẩn.");
  return { "TTA_Config.bas": toAsciiVba(configModule()), "TTA_Format.bas": toAsciiVba(FORMAT_MODULE) };
}
