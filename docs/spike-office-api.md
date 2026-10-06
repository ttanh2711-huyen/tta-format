# Spike T3: khả thi Office.js trên Word 365

Trạng thái: **đã chạy lần 1; đang chạy lần 2** cho bốn mục fail do cách chèn OOXML (xem "Kết quả lần 1"). Báo cáo đầy đủ lần 1 ở `docs/spike-results-lan-1.md`; lần 2 ghi vào `docs/spike-results.md`.

## Mục đích

Trả lời trước khi làm T4 trở đi: Office.js trên Word 365 có làm được các việc cốt lõi của add-in không (lề và khổ giấy theo mm, bảng đầu văn bản không viền, đường kẻ ngắn, số trang canh giữa ẩn trang đầu, cài và cập nhật style, giãn dòng Exactly cho chuẩn Đảng, lề đối xứng khi in hai mặt, phụ lục đánh số riêng, chụp và khôi phục cho nút "Chỉnh toàn bộ"). Nếu một mục bắt buộc không làm được thì phải đổi hướng (VSTO hoặc `.dotm`) trước khi tốn công.

## Điều đã biết từ tài liệu Microsoft (chưa phải kết quả chạy thật)

Nguồn: [Word JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) và trang từng lớp, đọc ngày 06/10/2026.

| Tính năng cần dùng | Requirement set | Bản Microsoft 365 Windows tối thiểu | Ghi chú |
|---|---|---|---|
| `Word.PageSetup` (khổ giấy, lề, `mirrorMargins`, `headerDistance`, `differentFirstPageHeaderFooter`, `sectionStart`) | WordApiDesktop 1.3 | 2507 (Build 19029.20208) | Không có trên bản Office mua vĩnh viễn hoặc LTSC |
| Tab stop (`TabStopCollection`) | WordApiDesktop 1.4 | 2508 (Build 19127.20264) | Add-in không dùng tab stop (khối đầu/cuối dựng bằng bảng) nên bỏ qua |
| `importStylesFromJson` | WordApi 1.6 | 2308 | Tham số xử lý trùng tên (`Overwrite`, `CreateNew`, `Ignore`) cần WordApiDesktop 1.1 (2408) |
| `Style`, `ParagraphFormat` | WordApi 1.5 | 2302 | `lineSpacing` tính bằng pt, **không có thuộc tính chọn kiểu giãn dòng** (Exactly hay bội số) |
| `insertField` | WordApi 1.5 | 2302 | Dùng cho trường số trang |
| `Table.getBorder`, `insertTable` | WordApi 1.3 | 1612 | |
| `insertOoxml`, `getOoxml` | WordApi 1.1 | 1509 | Style đã có trong tài liệu giữ định nghĩa cũ, không bị `insertOoxml` ghi đè (theo tài liệu "Use Office Open XML in your Word add-in") |

Hệ quả cho kế hoạch (cần phép thử xác nhận):
- Mọi tính năng page setup đều cần **Microsoft 365 bản 2507 trở lên**. Máy chưa cập nhật tới bản đó thì add-in phải báo rõ, không chạy ngầm (T23).
- Giãn dòng Exactly của chuẩn Đảng nhiều khả năng chỉ đặt được qua OOXML (style có `w:lineRule="exact"`), vì API không có thuộc tính kiểu giãn dòng. Phép thử p11 kiểm chứng.
- Cập nhật style: `importStylesFromJson` với `Overwrite` là cách chính; `insertOoxml` chỉ dùng để cài lần đầu. Phép thử p10 kiểm chứng.
- Không tìm thấy API đặt "đánh số trang lại từ 1" cho phần. Phương án là OOXML (`pgNumType start="1"` trong `sectPr`). Phép thử p12 kiểm chứng.

## Cách chạy

1. Mở một tài liệu Word **trống mới** (không phải văn bản thật; các phép thử xoá nội dung và đổi lề).
2. Chạy trong thư mục dự án:

```bash
npm start
```

   Lần đầu Windows hỏi cài chứng chỉ tin cậy cho `localhost`; bấm đồng ý.
3. Trong Word: tab **NĐ 30** > **Mở NĐ 30** > liên kết **Kiểm tra khả thi Office.js (T3)**.
4. Đánh dấu "Đây là tài liệu thử", bấm **Chạy tất cả**. Phép thử p14 và p15 cần bạn bấm Ctrl+Z để đếm số lần.
5. Báo cáo tự lưu vào `docs/spike-results.json` và `docs/spike-results.md`. Nếu thông báo không lưu được, bấm **Sao chép báo cáo** rồi dán cho mình.
6. Với p07, p08, p12 (cần nhìn bằng mắt): bấm **Chạy riêng** từng mục, xem tài liệu và cho mình biết đường kẻ và bảng trông thế nào.
7. Dừng máy chủ: `npm run stop`.

## Các phép thử

Mục có dấu * là bắt buộc cho quyết định go/no-go.

| Mã | Nội dung | Chuẩn |
|---|---|---|
| p02 * | Đặt A4 và lề theo mm bằng `Word.PageSetup`, đọc lại | chung |
| p03 | Lề riêng từng phần (section) | chung |
| p04 | Lề đối xứng khi in hai mặt (`mirrorMargins`) | Đảng |
| p05 | Khoảng cách header 10 mm, trang đầu khác biệt | Đảng |
| p06 * | Số trang canh giữa ở header (`insertField` hoặc OOXML), header trang đầu trống | chung |
| p07 | Đường kẻ ngắn bằng viền dưới của đoạn rỗng có thụt hai bên | chung |
| p08 * | Bảng đầu văn bản 3 x 2 không viền bằng `insertOoxml`, vừa vùng chữ; thử cả bảng rộng và thụt âm như file mẫu cũ | chung |
| p09 | Bảng tạo bằng API `insertTable` rồi bỏ viền (phương án thay thế p08) | chung |
| p10 * | Cài và cập nhật style: `importStylesFromJson` (kể cả `Overwrite`) hoặc `insertOoxml` | chung |
| p11 | Giãn dòng Exactly 18-22 pt: style OOXML và API `lineSpacing` | Đảng |
| p12 | Phụ lục đánh số trang riêng (`pgNumType start=1`) | chung |
| p13 | Chụp bản trước khi chỉnh và khôi phục (thân, bảng, header, lề) cho "Chỉnh toàn bộ" | chung |
| p14, p15 | Hoàn tác: 3 thay đổi trong 1 lượt `Word.run` so với 3 lượt, đếm số lần Ctrl+Z | chung (thủ công) |
| p16 | Hiệu năng: 600 đoạn định dạng trong một lượt | chung |
| p17 | Tab stop | bỏ qua có chủ đích |

Quy ước kết luận: mục bắt buộc không đạt hoặc không hỗ trợ thì **no-go**; chỉ mục khác không đạt thì **go-with-fallback** kèm phương án thay thế; còn lại là **go**.

## Kết quả lần 1 (06/10/2026)

Môi trường: Word cho Windows, Microsoft 365, phiên bản 16.0.20527.20016. Hỗ trợ đủ WordApi 1.1 đến 1.9 và WordApiDesktop 1.1 đến 1.5, tức là thừa điều kiện bản 2507 cần cho page setup.

Kết quả máy: pass 10, fail 3, thủ công 2, bỏ qua 1. **Cả 4 mục bắt buộc đều đạt** (p02, p06, p08, p10), nên theo quy ước là go-with-fallback.

| Hạng mục | Kết quả lần 1 | Kết luận tạm |
|---|---|---|
| Khổ A4 và lề theo mm bằng `Word.PageSetup` (p02) | Đạt, sai số dưới 0,03 mm | Dùng API này cho T4 |
| Lề riêng từng phần (p03) | Đạt | Dùng được cho phụ lục, văn bản nhiều phần |
| Lề đối xứng (`mirrorMargins`, p04) và header 10 mm (p05) | Đạt | Chuẩn Đảng in hai mặt làm được bằng API; cần xem Print Preview |
| Số trang canh giữa ở header, header trang đầu trống (p06) | Đạt cả `insertField` lẫn `insertOoxml` | Dùng `insertField` (API chính thức), OOXML dự phòng |
| Bảng đầu văn bản 3x2 không viền bằng `insertOoxml` (p08) | Đạt; rộng 453,55 pt đúng vùng chữ 160 mm; bảng rộng thụt -800 twips đọc lại đúng | Dựng bảng bằng OOXML, như kế hoạch |
| Bảng bằng API `insertTable` rồi bỏ viền (p09) | Đạt | Phương án dự phòng cho p08 |
| Cài và cập nhật style (p10) | `importStylesFromJson` đạt cả lần đầu lẫn `Overwrite` (cập nhật cỡ chữ, giãn dòng); `insertOoxml` kèm style **không cài được** | **Dùng `importStylesFromJson` với `Overwrite` cho T5** |
| Chụp bản trước khi chỉnh và khôi phục (p13) | Đạt: văn bản, định dạng đoạn, bảng, header, lề đều về đúng | "Chỉnh toàn bộ" (T6C) khôi phục được bằng `getOoxml` và `insertOoxml(Replace)` |
| Hiệu năng định dạng hàng loạt (p16) | 600 đoạn trong 542 ms (khoảng 3 giây cho 3000 đoạn). **Nhưng chèn 600 đoạn bằng vòng `insertParagraph` mất 121 giây** (khoảng 200 ms mỗi đoạn) | Định dạng văn bản có sẵn rất nhanh. Khi **tạo** văn bản phải chèn hàng loạt bằng một lệnh `insertText` hoặc một gói OOXML, không chèn từng đoạn |
| Giãn dòng Exactly cho chuẩn Đảng (p11) | API `lineSpacing = 20` cho ra quy tắc **auto** (bội số), **không phải Exactly**. OOXML: không giữ được ở lần thử đầu | Xác nhận API không đặt được Exactly, bắt buộc qua OOXML. Cách chèn OOXML đang thử lại |
| Đường kẻ bằng viền đoạn (p07), phụ lục đánh số riêng (p12) | Fail ở lần thử đầu | Nghi ngờ chung nguyên nhân với p10 và p11 (xem dưới), đang thử lại |
| Hoàn tác một lượt hay nhiều lượt (p14, p15) | Chưa ghi số lần Ctrl+Z | Chờ người chạy báo |

### Nguyên nhân nghi ngờ của các mục fail

Bốn mục fail (p07, p10 phần OOXML, p11, p12) cùng một mẫu: gói OOXML **chỉ có một đoạn** cần giữ thuộc tính riêng (viền, style, thụt dòng, ngắt phần). Hai mục chèn bảng (p08) và `importStylesFromJson` thì đạt. Giả thuyết: Word bỏ thuộc tính của đoạn cuối cùng trong gói vì dấu kết đoạn cuối không được chèn kèm. Lần chạy 2 thử ba cách chèn cho từng mục: đoạn đơn (như lần 1), kèm một đoạn rỗng phía sau, và `Paragraph.insertOoxml` với `Replace`. Giả thuyết đúng khi hai cách sau đạt còn cách đầu fail.

## Kết quả lần 2

*Chưa có.*
