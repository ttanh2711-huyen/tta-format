# Kết quả spike Office.js

Chạy lúc: 2026-10-06T08:02:16.573Z

## Môi trường

- Ứng dụng: Word
- Nền tảng: PC
- Phiên bản: 16.0.20527.20016
- WordApi 1.1: có
- WordApi 1.2: có
- WordApi 1.3: có
- WordApi 1.4: có
- WordApi 1.5: có
- WordApi 1.6: có
- WordApi 1.7: có
- WordApi 1.8: có
- WordApi 1.9: có
- WordApiDesktop 1.1: có
- WordApiDesktop 1.2: có
- WordApiDesktop 1.3: có
- WordApiDesktop 1.4: có
- WordApiDesktop 1.5: có

## Tổng hợp

- pass: 12, fail: 1, unsupported: 0, manual: 2, skipped: 1
- **Kết luận: go-with-fallback** (Phụ lục đánh số trang riêng: ngắt phần có pgNumType start=1 qua insertOoxml (ba cách chèn): fail)

## Chi tiết

| Mã | Nội dung kiểm tra | Nhóm | Kết quả | Chi tiết |
| --- | --- | --- | --- | --- |
| p02 | Đặt A4 và lề theo mm bằng Word.PageSetup, đọc lại | chung | pass | rộng 210.03 mm; cao 297.02 mm; lề trên 20 mm; lề dưới 20 mm; lề trái 30 mm; lề phải 20 mm. |
| p03 | Lề riêng từng phần (section): đổi phần 2, phần 1 giữ nguyên | chung | pass | lề trái phần 1 30 mm; lề trái phần 2 14.99 mm. |
| p04 | Chuẩn Đảng: lề đối xứng khi in hai mặt (mirrorMargins) | hd05 | pass | lề trong (leftMargin) 30 mm; lề ngoài (rightMargin) 14.99 mm. Khi bật, lề trái là lề trong; mặt sau tự đảo thành trái 15, phải 30 (cần xem Print Preview để xác nhận bằng mắt). |
| p05 | Khoảng cách header 10 mm và "trang đầu khác biệt" | hd05 | pass | khoảng cách header 10 mm. |
| p06 | Số trang canh giữa ở header (insertField hoặc OOXML), trang đầu header trống | chung | pass | insertField: đạt; insertOoxml: đạt; header trang đầu trống. |
| p07 | Đường kẻ ngắn bằng viền dưới của đoạn rỗng có thụt hai bên | chung | pass | đoạn đơn: không đạt (viền dưới mất, thụt mất); kèm đoạn rỗng phía sau: đạt; thay thế một đoạn bằng Paragraph.insertOoxml: không đạt (viền dưới mất, thụt mất). Cần xem bằng mắt: đường kẻ phải ngắn hơn dòng chữ phía trên và nằm giữa. |
| p08 | Bảng đầu văn bản 3 hàng x 2 cột không viền bằng insertOoxml, vừa vùng chữ | chung | pass | bề rộng bảng vừa vùng chữ 453.55 pt; bảng 3 hàng, thụt 0 twips, 6 cạnh không viền; bảng rộng như file mẫu cũ: rộng 520 pt, thụt -800 twips (mong đợi -800). Cần xem bằng mắt khi chạy riêng: hai bảng không có đường viền nào. |
| p09 | Bảng tạo bằng API insertTable rồi bỏ viền (phương án thay thế cho OOXML) | chung | pass | bảng 3 hàng, viền "None". |
| p10 | Cài và cập nhật style: importStylesFromJson (Overwrite) hoặc insertOoxml | chung | pass | importStylesFromJson lần 1: đạt (cỡ 14, giãn dòng 20 pt); importStylesFromJson Overwrite cập nhật style cũ: đạt (cỡ 13, giãn dòng 22 pt); insertOoxml (đoạn đơn): không cài được ({"exists":false,"font":"","size":0,"lineSpacing":0,"spaceAfter":0}); insertOoxml (kèm đoạn rỗng phía sau): cài được style, chèn lại với định nghĩa khác KHÔNG cập nhật style cũ; insertOoxml (thay thế một đoạn bằng Paragraph.insertOoxml): không cài được ({"exists":false,"font":"","size":0,"lineSpacing":0,"spaceAfter":0}). |
| p11 | Chuẩn Đảng: giãn dòng Exactly 18-22 pt (OOXML thử nhiều cách chèn, importStylesFromJson và API lineSpacing) | hd05 | pass | style OOXML (đoạn đơn): style không được tạo, đoạn mất style, định nghĩa không có Exactly 20 pt; đặt trực tiếp trên đoạn (đoạn đơn): Exactly mất; style OOXML (kèm đoạn rỗng phía sau): style có (mã ND30Cmuwe4q502), đoạn giữ style, định nghĩa có Exactly 20 pt; đổi giãn dòng style Exactly từ 20 sang 22 pt bằng API: không giữ Exactly (quy tắc đổi); đặt trực tiếp trên đoạn (kèm đoạn rỗng phía sau): Exactly đạt; style OOXML (thay thế một đoạn bằng Paragraph.insertOoxml): style không được tạo, đoạn mất style, định nghĩa không có Exactly 20 pt; đặt trực tiếp trên đoạn (thay thế một đoạn bằng Paragraph.insertOoxml): Exactly mất; importStylesFromJson lineSpacing=20: quy tắc "atLeast" (không phải Exactly); API lineSpacing=20 cho ra quy tắc "auto" (KHÔNG phải Exactly). |
| p12 | Phụ lục đánh số trang riêng: ngắt phần có pgNumType start=1 qua insertOoxml (ba cách chèn) | chung | fail | không cách chèn nào giữ được pgNumType start=1. Ghi chú: đoạn sectPr (đoạn đơn): 3 phần, pgNumType start=không có; đoạn sectPr (kèm đoạn rỗng phía sau): 3 phần, pgNumType start=không có; đoạn sectPr (thay thế một đoạn bằng Paragraph.insertOoxml): 2 phần, pgNumType start=không có. Cần xem bằng mắt: trang đầu của phụ lục phải hiện số 1 ở phần có đánh số. |
| p13 | "Chỉnh toàn bộ": chụp bản trước khi chỉnh và khôi phục (thân, bảng, header, lề) | chung | pass | lề trái sau khôi phục 30 mm; số đoạn trước 8, sau 8 (chênh lệch do đoạn rỗng cuối là bình thường). |
| p14 | Hoàn tác: 3 thay đổi trong MỘT Word.run, bấm Ctrl+Z bao nhiêu lần? | chung | manual | Đã tô đỏ và in đậm 3 đoạn trong MỘT lần Word.run. Bấm Ctrl+Z trong Word: ghi lại cần bấm mấy lần để cả 3 đoạn trở lại bình thường (mong đợi 1). |
| p15 | Hoàn tác: 3 thay đổi trong BA lần Word.run, bấm Ctrl+Z bao nhiêu lần? | chung | manual | Đã tô đỏ và in đậm 3 đoạn trong BA lần Word.run riêng. Bấm Ctrl+Z: ghi lại cần bấm mấy lần (để so với phép thử p14). |
| p16 | Hiệu năng: 600 đoạn, định dạng hàng loạt; so sánh cách chèn đoạn | chung | pass | chèn 600 đoạn bằng một lệnh insertText: 157 ms (đọc lại được 600 đoạn); tải danh sách 633 ms; định dạng 600 đoạn: 605 ms (khoảng 3 giây cho 3000 đoạn nếu tuyến tính); chèn từng đoạn bằng insertParagraph: khoảng 25 ms mỗi đoạn. |
| p17 | Tab stop (WordApiDesktop 1.4) | chung | skipped | Bỏ qua có chủ đích: các khối đầu và cuối văn bản dựng bằng bảng, không dùng tab stop. Không ảnh hưởng quyết định công nghệ. |
