# Spike T3: khả thi Office.js trên Word 365

> **Lưu ý (T26):** trang thăm dò `spike.html` đã bỏ khỏi dự án; kiểm thử trong Word nay dùng `src/selftest`. Tài liệu này và các báo cáo `spike-results*` chỉ giữ làm hồ sơ kết quả đã đo; các bước chạy trang thăm dò bên dưới không còn áp dụng.

Trạng thái: **đã chạy 3 lần trên Word 365; kết luận GO với Office.js**. Còn vài mục kiểm chứng bổ sung không chặn (xem cuối tài liệu). Báo cáo máy: `docs/spike-results-lan-1.md`, `lan-2`, `lan-3`; kết quả mới nhất ở `docs/spike-results.md`.

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
3. Trong Word: tab **TTA-Format** > **Mở TTA-Format** > liên kết **Kiểm tra khả thi Office.js (T3)**.
4. Đánh dấu "Đây là tài liệu thử", bấm **Chạy tất cả**. Phép thử p14 và p15 cần bạn bấm Ctrl+Z để đếm số lần.
5. Báo cáo tự lưu vào `docs/spike-results.json` và `docs/spike-results.md`. Nếu thông báo không lưu được, bấm **Sao chép báo cáo** rồi dán cho mình.
6. Với p07, p08, p12 (cần nhìn bằng mắt): bấm **Chạy riêng** từng mục, xem tài liệu và cho mình biết đường kẻ và bảng trông thế nào.
7. Dừng máy chủ: `npm run stop`.

### Nếu Word tự đóng giữa lúc chạy

Báo cáo được lưu **trước và sau từng phép thử**, kèm mã phép thử đang chạy. Mở lại Word, mở `docs/spike-results.md`: dòng "Dừng đột ngột khi đang chạy: pXX" cho biết phép thử gây sập. Phép thử đánh dấu ⚠ (p18) có nguy cơ cao nên **không** nằm trong "Chạy tất cả"; chỉ chạy riêng khi cần. Lượt "Chạy riêng" lưu vào `docs/spike-rieng.md`, không đè báo cáo đầy đủ.

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
| p18 ⚠ | Phụ lục đánh số riêng bằng cách sửa `sectPr` của cả thân rồi `Replace` (có thể làm Word đóng, chỉ chạy riêng) | chung |
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

## Kết quả lần 2 (06/10/2026)

Kết quả máy: pass 12, fail 1, thủ công 2, bỏ qua 1. Bốn mục bắt buộc vẫn đạt.

**Giả thuyết ở lần 1 đã được xác nhận cho p07 và p10.** Gói OOXML chỉ có một đoạn thì Word bỏ thuộc tính của đoạn cuối (viền, thụt dòng, style). Chỉ cách **kèm một đoạn rỗng phía sau** giữ được; cách đoạn đơn và `Paragraph.insertOoxml(Replace)` đều mất.

| Hạng mục | Kết quả lần 2 | Kết luận |
|---|---|---|
| Đường kẻ bằng viền dưới và thụt hai bên (p07) | Chỉ "kèm đoạn rỗng phía sau" đạt | **Quy tắc chung: chèn OOXML luôn kèm một đoạn rỗng sau, rồi xoá đoạn thừa.** Cần người xem bằng mắt: đường kẻ có ngắn hơn dòng chữ và nằm giữa |
| Cài style bằng OOXML (p10) | Chỉ "kèm đoạn rỗng" cài được; chèn lại với định nghĩa khác **không** cập nhật style đã có | Xác nhận: dùng `importStylesFromJson` với `Overwrite` (đã đạt cả cài mới lẫn cập nhật) cho T5 |
| Hiệu năng (p16) | Chèn 600 đoạn bằng một lệnh `insertText`: 161 ms; định dạng 600 đoạn: 846 ms (khoảng 4 giây cho 3000 đoạn); `insertParagraph` từng đoạn: khoảng 22 ms mỗi đoạn | Lần 1 mất 121 giây cho 600 `insertParagraph` trong một lượt, nên chèn nhiều đoạn bằng vòng lặp trong một lượt là **tăng nhanh hơn tuyến tính**. Luôn chèn hàng loạt bằng một lệnh |
| Giãn dòng Exactly (p11) | **Kết quả không dùng được.** Mọi cách đều báo đạt, kể cả API, trong khi chính dòng đó ghi API cho ra quy tắc "auto" | Phép thử quét cả phần styles của gói, mà tài liệu còn các style Exactly do p10 cài trước đó. Đã sửa: chỉ xét phần thân và đúng định nghĩa style đang dùng, mã style theo từng lần chạy. Chạy lại ở lần 3 |
| Đánh số trang phụ lục (p12) | Cả ba cách chèn đoạn chứa `sectPr` đều mất `pgNumType` | Đã thêm cách thứ tư (lấy OOXML cả thân, thêm `pgNumType` vào phần phụ lục, `Replace` lại) và ghi lại các đoạn `sectPr` để chẩn đoán. Chạy lại ở lần 3 |
| Hoàn tác (p14, p15) | Chưa ghi số lần Ctrl+Z | Còn chờ |

## Kết quả lần 3 (06/10/2026) và kết luận

Chạy trọn 16 phép thử không sập (p18 bị loại khỏi lượt chạy). Báo cáo ở `docs/spike-results-lan-3.md`. Kết quả máy: pass 12, fail 1 (p12), thủ công 2, bỏ qua 1.

**Về vụ Word tự đóng ở lần chạy trước:** lần này không sập, và khác biệt duy nhất so với lần sập là p18 (sửa `sectPr` của cả thân rồi `Replace`) không nằm trong lượt chạy. p18 là thủ phạm có khả năng cao nhất nhưng **chưa xác nhận**, vì chưa chạy riêng. Không dùng cách này trong add-in cho tới khi có kiểm chứng.

### Giãn dòng Exactly (p11), dữ liệu sạch lần đầu

| Cách | Kết quả |
|---|---|
| API `lineSpacing = 20` trên đoạn | Cho ra quy tắc **auto** (bội số), không phải Exactly |
| `importStylesFromJson` với `lineSpacing = 20` | Cho ra quy tắc **atLeast** (tối thiểu), không phải Exactly |
| Đặt trực tiếp trên đoạn bằng OOXML | **Đạt** nếu kèm đoạn rỗng phía sau; mất nếu chèn đoạn đơn hoặc `Replace` |
| Style Exactly cài bằng OOXML | **Chưa kết luận.** Phép thử tra mã style sai (Word bỏ dấu gạch dưới khi tạo mã từ tên: `ND30_J...` thành `ND30J...`), nên báo "đoạn mất style" dù có thể style vẫn tồn tại. Đã sửa phép thử, cần chạy bổ sung |

Hệ quả: chuẩn Đảng cần Exactly, mà **API và `importStylesFromJson` đều không đặt được**. Chỉ còn OOXML. Nếu style Exactly cài bằng OOXML hoạt động (chạy bổ sung), thì T5 cài style `HD05_*` lần đầu bằng OOXML (kèm đoạn rỗng), gán cho đoạn bằng API, và đổi cỡ giãn dòng bằng `style.paragraphFormat.lineSpacing` (phép thử mới kiểm tra quy tắc Exactly có còn nguyên sau khi đổi 20 sang 22 pt). Nếu không, phương án dự phòng đã chứng minh là đặt Exactly **trực tiếp trên từng đoạn** bằng OOXML (thay đoạn kèm đoạn rỗng), tốn kém hơn cho "Chỉnh toàn bộ".

### Đánh số trang phụ lục (p12)

Cả ba cách chèn đoạn chứa `sectPr` có `pgNumType start="1"` đều **mất `pgNumType`**, dù Word vẫn tạo đúng số phần (ngắt phần được giữ, thuộc tính đánh số thì không). Phương án còn lại: (1) p18, sửa OOXML cả thân rồi `Replace`, rủi ro sập; (2) trường công thức ở chân trang của phần phụ lục, kiểu `{ = { PAGE } - { PAGEREF DauPhuLuc } + 1 }` với dấu trang (bookmark) đặt đầu phụ lục; (3) hướng dẫn người dùng đặt "Start at" bằng tay. **Không chặn quyết định công nghệ**; để T21 và TD8 kiểm chứng phương án (2) khi làm phụ lục.

### Bảng tính năng trên Word 365 cho Windows (16.0.20527)

| Tính năng | Kết quả | Cách dùng trong add-in |
|---|---|---|
| Khổ A4, lề theo mm, lề từng phần | Đạt | `Word.PageSetup` (WordApiDesktop 1.3, cần bản 2507 trở lên) |
| Lề đối xứng, header 10 mm, trang đầu khác biệt | Đạt | `PageSetup`; Print Preview cần xem bằng mắt |
| Số trang canh giữa, header trang đầu trống | Đạt (hai cách) | `insertField`, OOXML dự phòng |
| Bảng đầu văn bản không viền, vừa vùng chữ, bảng rộng thụt âm | Đạt | `insertOoxml` kèm đoạn rỗng |
| Bảng API `insertTable` bỏ viền | Đạt | Dự phòng |
| Đường kẻ viền đoạn có thụt hai bên | Đạt, chỉ khi kèm đoạn rỗng phía sau | `insertOoxml` kèm đoạn rỗng |
| Cài và cập nhật style | `importStylesFromJson` Overwrite đạt; `insertOoxml` chỉ cài mới, không cập nhật | T5: `importStylesFromJson` (trừ giãn dòng Exactly, xem trên) |
| Giãn dòng Exactly | Chỉ qua OOXML (trực tiếp trên đoạn đạt; qua style chờ kiểm chứng) | Xem mục Exactly |
| Đánh số trang phụ lục riêng | Không đạt bằng `insertOoxml` | Xem mục p12 |
| Chụp và khôi phục cho "Chỉnh toàn bộ" | Đạt (văn bản, định dạng, bảng, header, lề) | `getOoxml` và `insertOoxml(Replace)` |
| Hiệu năng | Chèn 600 đoạn bằng một lệnh 166 ms; định dạng 600 đoạn 642 ms (khoảng 3 giây cho 3000 đoạn) | Chèn hàng loạt bằng `insertText` hoặc một gói OOXML; định dạng hàng loạt trong một lượt |
| Tab stop | Bỏ qua có chủ đích | Không dùng |
| Hoàn tác (p14, p15) | **Chưa có số liệu** | Thiết kế mặc định: gộp thay đổi trong ít lượt `Word.run` nhất |

### Quy tắc kỹ thuật rút ra cho các task sau

1. Mọi lần chèn OOXML chứa đoạn có thuộc tính riêng (viền, thụt dòng, style, Exactly) **phải kèm một đoạn rỗng phía sau**, rồi xoá đoạn thừa.
2. Khi tạo văn bản, chèn hàng loạt bằng một lệnh; không chèn từng đoạn trong vòng lặp lớn.
3. Mọi tính năng page setup cần kiểm tra `isSetSupported("WordApiDesktop", "1.3")` và báo rõ yêu cầu Microsoft 365 bản 2507 trở lên.
4. Style cài lần đầu một lần và cập nhật bằng `importStylesFromJson` với `Overwrite`; mã style trong OOXML là mã Word tạo (bỏ dấu gạch dưới), không phải mã mình đặt.

### Kết luận go/no-go

**GO với Office.js trên Word 365.** Cả 4 mục bắt buộc (p02, p06, p08, p10) đạt ở cả ba lần chạy. Các điểm chưa có giải pháp đã kiểm chứng (giãn dòng Exactly qua style, đánh số phụ lục) có phương án dự phòng và không làm đổi hướng công nghệ.

### Việc còn mở sau T3 (không chặn Checkpoint 0)

- [ ] Chạy bổ sung p11 (đã sửa) để biết style Exactly qua OOXML có dùng được không; kết quả quyết định cách làm Exactly ở T5 và T6C.
- [ ] Số lần Ctrl+Z của p14, p15; quan sát bằng mắt p07, p08.
- [ ] Chạy riêng p18 nếu cần (có thể làm Word đóng); chỉ làm khi p12 trở thành chặn.
