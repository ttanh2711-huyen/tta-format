# NĐ 30/2020/NĐ-CP: nguồn, hiệu lực và các điểm còn mơ hồ

Tài liệu đi kèm `src/rules/nd30.ts` và `src/rules/docTypes.ts`. Cập nhật khi T2B hoặc người rà soát có kết luận mới.

## Nguồn đã dùng

| Nội dung | Nguồn | Mức tin cậy |
|---|---|---|
| Phụ lục I (thể thức, kỹ thuật trình bày), Phụ lục II (viết hoa), Phụ lục III (bảng chữ viết tắt, mẫu) | Bản PDF phụ lục đăng trên cdn.thuvienphapluat.vn, đọc toàn văn ngày 06/10/2026 | Cao (nguyên văn phụ lục; chưa đối chiếu với bản đăng công báo) |
| Danh sách 29 loại văn bản hành chính ở Điều 7 | Nguồn thứ cấp (luatminhkhue.vn, qua tìm kiếm), đối chiếu chéo với 27 loại có chữ viết tắt ở Phụ lục III | Trung bình: chưa đọc nguyên văn Điều 7 |
| Hiệu lực | Tìm kiếm ngày 06/10/2026: nhiều kế hoạch công tác văn thư năm 2025 và 2026 của các cơ quan vẫn viện dẫn NĐ 30/2020/NĐ-CP | Thấp: chưa tìm thấy văn bản sửa đổi, thay thế, nhưng chưa tra cứu cơ sở dữ liệu pháp luật chính thức |

## Việc cần người am hiểu văn thư xác nhận (chưa làm)

- [ ] NĐ 30/2020/NĐ-CP còn hiệu lực nguyên văn, chưa có nghị định sửa đổi phần thể thức. Gợi ý tra: Cơ sở dữ liệu quốc gia về văn bản pháp luật, mục "Lược đồ" của NĐ 30.
- [ ] Danh sách 29 loại ở Điều 7 đúng như `docTypes.ts`.
- [ ] Rà bảng quy tắc `nd30.ts` với Phụ lục I (ký xác nhận vào đây khi xong).

## Điểm mơ hồ hoặc cần lưu ý

1. **Ký hiệu in hoa nhưng chữ viết tắt tên loại có chữ thường.** Mục II.3.c nói ký hiệu "in hoa", nhưng Bảng ở Phụ lục III có "CTr", "TTr", "QyĐ", "TrS". `nd30.ts` ghi `kyHieu` là in hoa; bộ kiểm tra T14 phải cho phép các chữ viết tắt trên giữ nguyên chữ thường ở vị trí đó. Quyết định ở T8/T14.
2. **Cỡ chữ "thống nhất" trong cùng văn bản.** Chú thích 1 của Mục V nói tăng giảm cỡ chữ phải thống nhất, ví dụ (quốc hiệu 13, tiêu ngữ 14, địa danh và ngày tháng 14) hoặc (12, 13, 13). Đây là ví dụ, không phải danh sách đóng. File mẫu của bạn dùng (13, 14, 13), nằm trong khoảng từng thành phần nhưng không trùng ví dụ nào. Add-in coi khoảng từng thành phần là quy tắc cứng và bộ ba "thống nhất" là gợi ý (mức cảnh báo thấp).
3. **Trích yếu công văn.** Sau "V/v" in thường, cỡ chữ **12 đến 13**, kiểu đứng (Mục II.5.b và bảng Mục V). Kế hoạch ban đầu ghi 13 đến 14 là sai; đã sửa. Căn giữa dưới số và ký hiệu, cách dòng 6 pt với số và ký hiệu (khoảng cách này chưa có trong `nd30.ts`, T9 xử lý).
4. **Loại văn bản chưa có mẫu ở Phụ lục III:** hợp đồng, bản ghi nhớ, bản thỏa thuận, thư công. T17 đến T20 cần quyết định khối chung cho các loại này (đề xuất: dùng khối đầu và khối ký của Mẫu 1.4, không có "Nơi nhận" cố định).
5. **"Bản cam kết" không có trong NĐ 30.** Danh sách loại văn bản trong kế hoạch ban đầu có "bản cam kết"; không có trong Điều 7 và Phụ lục III nên đã bỏ.
6. **Giấy ủy quyền, phiếu gửi, phiếu chuyển, phiếu báo** dùng chung Mẫu 1.4 (có tên loại và trích yếu), không phải "giấy tờ mẫu nhỏ" riêng như kế hoạch ban đầu. Nhóm "giấy tờ mẫu nhỏ" thực sự (mẫu riêng) chỉ có giấy mời, giấy giới thiệu, giấy nghỉ phép, công điện.
7. **Khoảng cách giữa các dòng** "tối thiểu là dòng đơn, tối đa là 1,5 lines" (Mục II.6.e). File mẫu của bạn dùng 1,17 (line = 280 trong OOXML), hợp lệ.
8. **Thụt dòng đầu 1 cm hoặc 1,27 cm.** `nd30.ts` ghi `[10, 12.7]` mm; chỉ hai giá trị này là chuẩn, giá trị khác (ví dụ 1,25 cm) bị coi là sai ở mức cảnh báo. File mẫu của bạn dùng 1,27 cm.
9. **Số trang.** Cỡ 13 đến 14 (Mục I.7); ví dụ ở bảng Mục V ghi 14. File mẫu dùng 14.
10. **Dấu chỉ mức độ khẩn và chỉ dẫn phạm vi lưu hành** là con dấu đóng bằng tay trong khung viền đơn, không phải văn bản gõ. Add-in dựng khung chữ nhật chỉ để thay con dấu khi in nội bộ; kích thước khung chỉ áp dụng cho dấu khẩn (30x8, 40x8, 20x8 mm), không có kích thước cho chỉ dẫn phạm vi lưu hành.
11. **Phụ lục II (viết hoa)** chưa mã hoá. Có thể dùng cho bộ kiểm tra chính tả viết hoa ở giai đoạn sau; không nằm trong kế hoạch hiện tại.
12. **Ngày tháng.** Chỉ ngày nhỏ hơn 10 và tháng 1, 2 phải thêm số 0 (Mục II.4.b); tháng 3 đến 9 không thêm. `formatNgay` làm đúng như vậy ("ngày 29 tháng 6 năm 2019" nhưng "ngày 05 tháng 01 năm 2020").
13. **Số văn bản.** Chỉ yêu cầu số nhỏ hơn 10 thêm số 0 (Mục II.3.c). `formatSo` từ chối số 0 và số âm vì số văn bản bắt đầu từ 01.
14. **Dấu cuối các căn cứ của quyết định, nghị quyết (T16).** Add-in ghi mỗi căn cứ kết thúc bằng ";" và căn cứ cuối bằng "," vì sau đó là dòng "QUYẾT ĐỊNH:". Đây là cách đọc theo thông lệ; chưa đối chiếu được nguyên văn Mục II.6.a và Mẫu 1.2 — nếu mẫu ghi khác (ví dụ dấu chấm phẩy hoặc chấm), đổi ở `basisLines` (`src/templates/quyetDinh.ts`) và `checkNd30Basis` (`src/standards/nd30/decisionChecks.ts`).
15. **Điểm đánh số bằng chữ cái (T16).** Danh sách tự động dùng bảng a-z của Word, nên sau "d" là "e"; chữ "đ" của quy định (a, b, c, d, đ, e, g…) phải gõ tay hoặc sửa tay. Điều, Khoản và Điểm của quyết định nằm trong một danh sách nhiều cấp; số Khoản bắt đầu lại ở mỗi Điều.
16. **Dòng "(Ban hành kèm theo ...)" của quy chế, quy định (T18).** Add-in đặt giữa trang, nghiêng, trong ngoặc đơn, ngay dưới trích yếu; chưa đối chiếu nguyên văn Phụ lục III (cỡ chữ, nghiêng hay không). Nếu mẫu ghi khác, đổi `attachedParagraph` trong `src/templates/quyChe.ts`.
17. **Bản sao y, trích sao, sao lục (T26).** Phụ lục III chỉ nêu chữ viết tắt SY, TrS, SL; bố cục bản sao (khối đầu, tên "SAO Y", "TRÍCH SAO", "SAO LỤC", mục dẫn văn bản gốc, khối ký) do add-in đề xuất theo họ mẫu giấy tờ (`src/templates/smallForms.ts`), chưa đối chiếu mẫu gốc. Nếu mẫu ghi khác, sửa `SMALL_FORMS`.
