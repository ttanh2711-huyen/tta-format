import { detectStandard, resolveStandard, type Detection } from "./detectStandard";

const QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
const DANG = "ĐẢNG CỘNG SẢN VIỆT NAM";

describe("detectStandard: nhận diện theo tiêu đề đầu văn bản", () => {
  it("văn bản chính quyền: có quốc hiệu thì nhận ra chuẩn nd30", () => {
    const header = ["UBND TỈNH ĐỒNG THÁP", QUOC_HIEU, "Độc lập - Tự do - Hạnh phúc", "Số: 05/UBND-VP"];
    expect(detectStandard(header)).toEqual({ status: "detected", standardId: "nd30" });
  });

  it("văn bản Đảng: có tiêu đề Đảng thì nhận ra chuẩn hd05", () => {
    const header = [DANG, "ĐẢNG ỦY SỞ NÔNG NGHIỆP", "Số 05-HD/VPTW", "HƯỚNG DẪN"];
    expect(detectStandard(header)).toEqual({ status: "detected", standardId: "hd05" });
  });

  it("không phân biệt hoa thường, dấu thanh viết kiểu cũ hay mới, dấu câu thừa", () => {
    expect(detectStandard(["Cộng hoà xã hội chủ nghĩa Việt Nam"])).toEqual({ status: "detected", standardId: "nd30" });
    expect(detectStandard(["  ĐẢNG   CỘNG SẢN VIỆT NAM  "])).toEqual({ status: "detected", standardId: "hd05" });
    expect(detectStandard(["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM."])).toEqual({ status: "detected", standardId: "nd30" });
  });

  it("tiêu đề bị ngắt dòng thành nhiều đoạn vẫn nhận ra", () => {
    expect(detectStandard(["CỘNG HÒA XÃ HỘI", "CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"])).toEqual({ status: "detected", standardId: "nd30" });
  });

  it("đoạn trống xen giữa không ảnh hưởng", () => {
    expect(detectStandard(["", "  ", DANG, ""])).toEqual({ status: "detected", standardId: "hd05" });
  });

  it("tài liệu trống hoặc chỉ có khoảng trắng", () => {
    expect(detectStandard([])).toEqual({ status: "empty" });
    expect(detectStandard(["", "   ", "\t"])).toEqual({ status: "empty" });
  });

  it("có nội dung nhưng không có tiêu đề: không chắc, không gợi ý", () => {
    expect(detectStandard(["Kính gửi: Sở Nông nghiệp", "Thực hiện chỉ đạo của UBND tỉnh..."])).toEqual({ status: "unsure", reason: "noTitle" });
  });

  it("có cả hai tiêu đề: không chắc", () => {
    expect(detectStandard([DANG, QUOC_HIEU])).toEqual({ status: "unsure", reason: "bothTitles" });
  });

  it("cụm từ nằm giữa câu văn không được tính là tiêu đề", () => {
    const header = [QUOC_HIEU, "Dưới sự lãnh đạo của Đảng Cộng sản Việt Nam, ngành nông nghiệp đạt kết quả tốt."];
    expect(detectStandard(header)).toEqual({ status: "detected", standardId: "nd30" });
    expect(detectStandard(["Dưới sự lãnh đạo của Đảng Cộng sản Việt Nam, ngành nông nghiệp đạt kết quả tốt."])).toEqual({ status: "unsure", reason: "noTitle" });
  });

  it("tiêu đề sai chính tả: không chắc nhưng gợi ý chuẩn gần giống nhất", () => {
    expect(detectStandard(["CỘNG HÒA XÃ HỘI CHỦ NGHIA VIỆT NAN"])).toEqual({ status: "unsure", reason: "similarTitle", suggestion: "nd30" });
    expect(detectStandard(["ĐẢNG CÔNG SAN VIÊT NAM X"])).toEqual({ status: "unsure", reason: "similarTitle", suggestion: "hd05" });
  });

  it("tiêu đề nằm ngoài phần đầu (quá 30 đoạn có chữ) không được xét", () => {
    const filler = Array.from({ length: 30 }, (_, i) => `Đoạn nội dung số ${i + 1}`);
    expect(detectStandard([...filler, QUOC_HIEU])).toEqual({ status: "unsure", reason: "noTitle" });
  });
});

describe("resolveStandard: quyết định dùng chuẩn nào, không bao giờ tự đoán", () => {
  const detected: Detection = { status: "detected", standardId: "hd05" };
  const base = { stored: null as string | null, noAsk: false, fallbackId: "nd30" as const };

  it("chuẩn đã ghi trong tài liệu được ưu tiên, kể cả khi tiêu đề khác", () => {
    expect(resolveStandard({ ...base, stored: "nd30", detection: detected })).toEqual({ action: "use", standardId: "nd30", source: "document" });
  });

  it("chuẩn ghi trong tài liệu không còn hợp lệ thì bỏ qua", () => {
    expect(resolveStandard({ ...base, stored: "khong-co", detection: detected })).toEqual({ action: "use", standardId: "hd05", source: "detected" });
  });

  it("chưa ghi chuẩn: tiêu đề rõ ràng thì dùng chuẩn nhận diện", () => {
    expect(resolveStandard({ ...base, detection: detected })).toEqual({ action: "use", standardId: "hd05", source: "detected" });
  });

  it("tài liệu trống: dùng chuẩn mặc định, không hỏi", () => {
    expect(resolveStandard({ ...base, fallbackId: "hd05", detection: { status: "empty" } })).toEqual({ action: "use", standardId: "hd05", source: "default" });
  });

  it("không chắc: hỏi lại, kèm lý do và gợi ý", () => {
    expect(resolveStandard({ ...base, detection: { status: "unsure", reason: "similarTitle", suggestion: "hd05" } })).toEqual({
      action: "ask",
      reason: "similarTitle",
      suggestion: "hd05",
    });
    expect(resolveStandard({ ...base, detection: { status: "unsure", reason: "noTitle" } })).toEqual({ action: "ask", reason: "noTitle", suggestion: undefined });
  });

  it("đã chọn không hỏi lại cho tài liệu này: dùng chuẩn mặc định, không hỏi", () => {
    expect(resolveStandard({ ...base, noAsk: true, detection: { status: "unsure", reason: "bothTitles" } })).toEqual({ action: "use", standardId: "nd30", source: "default" });
  });

  it("không hỏi lại không ngăn việc dùng chuẩn đã ghi trong tài liệu", () => {
    expect(resolveStandard({ ...base, noAsk: true, stored: "hd05", detection: { status: "unsure", reason: "noTitle" } })).toEqual({ action: "use", standardId: "hd05", source: "document" });
  });
});
