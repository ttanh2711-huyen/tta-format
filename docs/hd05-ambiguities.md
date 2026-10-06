# Hướng dẫn 05-HD/VPTW: nguồn, hiệu lực và các điểm còn mơ hồ

Tài liệu đi kèm `src/standards/hd05/`. Cập nhật khi có kết luận mới.

## Nguồn đã dùng

| Nội dung | Nguồn | Mức tin cậy |
|---|---|---|
| Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng (Phần I, II, III, Phụ lục 1, 2, 3) | Bản PDF đăng trên sav.gov.vn (`05.HD.VPTW.pdf`), đọc toàn văn ngày 06/10/2026 | Cao. Phần chữ trong PDF mất nhiều chữ "Đ" và dấu khi trích, nên **bảng cỡ chữ (Phụ lục 2) và bảng chữ viết tắt cơ quan (trang 8-9) đã đối chiếu bằng ảnh trang** |
| Điều 5 và Điều 6 của Quy định số 399-QĐ/TW ngày 09/01/2026 của Ban Bí thư (danh sách thể loại) | Hai bài báo: baolamdong.vn và lsvn.vn | Trung bình. Hai nguồn cho cùng danh sách, nhưng chưa đọc nguyên văn quy định |
| Hiệu lực | HD 05 ghi hiệu lực từ ngày ký (27/5/2026), thay Hướng dẫn 36-HD/VPTW ngày 04/3/2018 | Cao |

## Lỗi đã bắt được nhờ đối chiếu ảnh

Văn bản trích chữ làm mất chữ "Đ", dẫn đến mã sai nếu chỉ đọc chữ. Đã sửa theo ảnh: ĐUQS, ĐUCA, ĐUBP, ĐU (đảng uỷ), BXDĐ (ban xây dựng đảng), BCĐ (ban chỉ đạo), ĐTHVN, ĐTNVN (đài), HĐ (hội đồng), ĐH (đại hội), QĐ (quyết định, quy định). Văn bản gốc viết "uỷ" (không phải "ủy"); `findAgencyByName` chấp nhận cả hai cách.

## Việc cần người am hiểu xác nhận (chưa làm)

- [ ] Danh sách thể loại ở Điều 5 và Điều 6 đúng nguyên văn Quy định 399-QĐ/TW. Một bài báo ghi "27 thể loại" nhưng danh sách trong bài chỉ có **25 tên**; nghi ngờ thiếu hai thể loại. Cần đọc nguyên văn Điều 5.
- [ ] Rà bảng quy tắc `rules.ts` với Phụ lục 2 và Phần II (ký xác nhận vào đây khi xong).
- [ ] Chữ viết tắt "suy ra" (đánh dấu `abbrSource: "derived"` trong `docTypes.ts`): TB (thông báo), TC (thông cáo), KH, QH, ĐA, PA, DA và các giấy tờ GGT, GCN, GĐĐ, GNP, GM, PC, PG. Hướng dẫn 05 chỉ nêu rõ 13 chữ viết tắt (NQ, CT, KL, QC, BC, QĐ, CTr, TT, TTr, CV, HD, BB và QĐ cho cả quy định).

## Điểm mơ hồ hoặc cần lưu ý

1. **Tuyên bố** có chữ cái đầu là TB, trùng với thông báo. Hướng dẫn 05 chỉ thống nhất năm trường hợp trùng (QĐ, CT, CTr, TT, TTr) nên chưa gán chữ viết tắt cho tuyên bố. Cương lĩnh, điều lệ, chiến lược, lời kêu gọi, thư công cũng không gán (không có số, ký hiệu thông thường).
2. **Quyết định và quy định cùng là QĐ** (Mục I.3.1.2.a). Vì vậy chữ viết tắt loại văn bản không duy nhất ở chuẩn Đảng; `findDocTypesByAbbr` trả về danh sách.
3. **Hướng dẫn 05 không có mẫu riêng từng thể loại**, chỉ một sơ đồ bố trí chung (Phụ lục 1) và bảng cỡ chữ (Phụ lục 2). `template` luôn là null. Bố cục từng thể loại do add-in tự dựng theo sơ đồ chung, bạn duyệt bản in thử ở Checkpoint 3B.
4. **Cỡ chữ theo khoảng, không có giá trị mặc định.** Nội dung 14-15, thể loại 15-16, trích yếu 14-15. Mặc định đề xuất (nội dung 14, thể loại 16, trích yếu 14) là lựa chọn của kế hoạch (Q13), không phải quy định.
5. **Giãn dòng Exactly 18-22 pt** (Mục I.6.2.2); mặc định đề xuất 20 pt là lựa chọn của kế hoạch.
6. **Thụt dòng đầu "khoảng 10 mm"**: `firstLineIndentToleranceMm = 1` là lựa chọn của add-in, không có trong văn bản.
7. **Dấu chỉ mức độ khẩn**: Hướng dẫn 05 chỉ nêu chữ in hoa cỡ 14, đứng, đậm, đặt dưới số và ký hiệu (công văn: dưới trích yếu). Không nêu khung viền hay kích thước (khác NĐ 30), nên `URGENCY_MARKS` không có `widthMm`, `heightMm`.
8. **Tên cơ quan cấp trên không đậm, cơ quan ban hành đậm**, cùng cỡ 14, in hoa (Phụ lục 2, mục 2). Chức vụ người ký in hoa nhưng **không đậm**; quyền hạn (T/M, K/T...) và họ tên đậm.
9. **"Nơi nhận"**: cỡ 14, đứng, gạch chân (có đường kẻ ngang liền dưới cụm từ); các mục liệt kê cỡ 12. Khác NĐ 30 (cỡ 12 nghiêng đậm; liệt kê 11).
10. **Cỡ chữ phụ lục** (tên phụ lục, số thứ tự) không có trong bảng Phụ lục 2; chưa mã hoá. T21/TD8 cần quyết định (đề xuất cùng cỡ nội dung, đậm).
11. **Quyền hạn ký uỷ quyền**: Hướng dẫn 05 nêu người được uỷ quyền ký trực tiếp, không uỷ quyền tiếp, nên không có "TUQ." như NĐ 30. Quyền hạn "Q." (quyền) chỉ thấy qua ví dụ "Q. TRƯỞNG BAN".
12. **Văn bản chứa bí mật nhà nước và bản sao văn bản** nằm ngoài phạm vi add-in (Hướng dẫn 05 Phần I.1.1.b; Phần II.III).
13. **Chữ ký**: bản giấy ký bằng bút mực xanh; bản điện tử là ảnh PNG nền trong suốt màu xanh, canh giữa chức vụ và họ tên. Add-in chỉ chừa chỗ ký, không xử lý chữ ký số.
14. **Ngày có hiệu lực của NĐ 30** ghi `2020-03-05` (ngày ký) là giả định "hiệu lực từ ngày ký", chưa kiểm chứng; chỉ dùng để hiển thị.
15. **Bảng tra ký hiệu cơ quan** chỉ có các tên Hướng dẫn 05 thống nhất. Tên cấp uỷ cụ thể (TU, TW, VPTW, BTCTU...) do từng cấp uỷ đặt theo quy tắc chữ cái đầu.
