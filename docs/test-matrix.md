# Ma trận kiểm thử (T23)

Hai cột kiểm thử: **Tự động** là phần `npm test` đã kiểm bằng mã (không có Word); **Word thật** là phần chỉ bạn kiểm được trên Microsoft 365 và **chưa thử**: mỗi ô còn trống `[ ]`, hãy đánh dấu `[x]` khi đạt hoặc ghi lỗi vào cột Ghi chú. Mọi dòng "Word thật" hiện đều chưa thử vì add-in chưa từng chạy trong Word ở các task từ T7 trở đi.

## Môi trường thử

| Mục | Giá trị (điền khi thử) |
| --- | --- |
| Hệ điều hành | |
| Microsoft 365, số bản (File > Account > About Word) | |
| Bộ API đủ (mục "Môi trường Word" ở đầu add-in báo "đủ tính năng") | |
| Ngày thử | |

Cần bản có WordApiDesktop 1.3 (Microsoft 365 bản 2507 trở lên). Word trên Mac và Word Online chỉ thử nếu tiện; tính năng thiếu bộ API phải hiện thông báo rõ chứ không được lỗi im lặng.

## 1. Nền tảng

| Tính năng | Tự động | Word thật: Chính quyền | Word thật: Đảng | Ghi chú |
| --- | --- | --- | --- | --- |
| Chọn và nhận diện chuẩn, ghi chuẩn vào tài liệu | Có (nhận diện, quyết định chuẩn) | [ ] | [ ] | Mở tài liệu có quốc hiệu, tài liệu có tiêu đề Đảng, tài liệu trộn cả hai |
| Thiết lập trang (khổ, lề, khoảng cách header) | Có (kế hoạch, kiểm tra giá trị) | [ ] | [ ] | Chuẩn Đảng: thử cả lề đối xứng khi in hai mặt |
| Cài bộ style (tiền tố `ND30_` / `HD05_`) | Có (đặc tả style, XML) | [ ] | [ ] | Chạy hai lần: không tạo trùng; giãn dòng Exactly của chuẩn Đảng |
| Định dạng nội dung và bảng biểu | Có (phân loại đoạn) | [ ] | [ ] | Giữ và xoá định dạng cục bộ |
| Đánh số trang (trang 1 ẩn, cỡ, căn giữa) | Có (header OOXML) | [ ] | [ ] | In PDF kiểm tra 10 mm cách mép trên (Đảng) |
| Chỉnh toàn bộ (7 bước) và khôi phục bản trước khi chỉnh | Có (khung chạy bước) | [ ] | [ ] | Văn bản có lỗi cài sẵn; bấm Ctrl+Z một lần |

## 2. Khối đầu, khối cuối, tạo văn bản

| Tính năng | Tự động | Word thật: Chính quyền | Word thật: Đảng | Ghi chú |
| --- | --- | --- | --- | --- |
| Chèn khối đầu (bảng, đường kẻ, số, ngày, Kính gửi) | Có (golden OOXML) | [ ] | [ ] | Đường kẻ nổi (Insert > Shapes > Line) có còn sau khi chèn? |
| Chèn khối cuối (Nơi nhận, chức vụ, họ tên) | Có (golden) | [ ] | [ ] | |
| Tạo công văn | Có (ghép khối) | [ ] | [ ] | Tài liệu có sẵn nội dung: hộp xác nhận và khôi phục |
| Nhóm có tên loại và trích yếu (15 loại Chính quyền, các thể loại Đảng) | Có (mọi loại, 0 lỗi) | [ ] | [ ] | Thử ít nhất 3 loại, in PDF đối chiếu mẫu |
| Nghị quyết, quyết định cá biệt (căn cứ, Điều tự đánh số) | Có (golden, kiểm tra căn cứ) | [ ] | — | Thêm, xoá Điều: số tự cập nhật; Tab xuống Khoản |
| Quy chế, quy định: Chương, Mục, Điều | Có (golden, đánh số năm cấp) | [ ] | [ ] | Đảng: thử "Phần", "Chương", số La Mã và số bằng chữ |
| Biên bản; hợp đồng, bản ghi nhớ, bản thỏa thuận (khối ký 2 đến 4 bên) | Có (golden 2, 3, 4 bên) | [ ] | [ ] | Tên bên rất dài; Đảng: dòng xác nhận chữ ký |
| Giấy mời, giấy giới thiệu, công điện, giấy nghỉ phép, thư công | Có (golden) | [ ] | [ ] | Bố cục chưa đối chiếu nguyên văn Phụ lục III và HD 05 |
| Văn bản đại hội, văn bản liên cơ quan | Có (0 lỗi) | — | [ ] | |

## 3. Kiểm tra và sửa

| Tính năng | Tự động | Word thật: Chính quyền | Word thật: Đảng | Ghi chú |
| --- | --- | --- | --- | --- |
| Kiểm tra văn bản (trang, đoạn, khối đầu, khối cuối) | Có (hàng trăm ca) | [ ] | [ ] | Dùng một văn bản thật, ghi lại báo sai |
| Nút "Đến vị trí" | Không | [ ] | [ ] | Bảng lồng nhau và khối `w:sdt` có thể lệch chỉ số đoạn |
| Nút "Sửa" từng lỗi | Có (giả lập trên bản sao) | [ ] | [ ] | Lỗi giãn dòng Exactly ghi lại đoạn bằng OOXML |
| Sửa tất cả: xem trước, áp dụng, hoàn tác, Ctrl+Z | Có (gom, giả lập) | [ ] | [ ] | Đo thời gian văn bản 100 trang |
| Xuất báo cáo (sao chép, tải về .txt) | Có (định dạng chữ) | [ ] | [ ] | Tải về có chạy trong Word không? |
| Cảnh báo trộn dấu hiệu hai chuẩn | Có | [ ] | [ ] | |

## 4. Thành phần bổ sung và cấu hình

| Tính năng | Tự động | Word thật: Chính quyền | Word thật: Đảng | Ghi chú |
| --- | --- | --- | --- | --- |
| Dấu chỉ mức độ khẩn (khung viền đơn) | Có (golden) | [ ] | [ ] | Kích thước 30×8, 40×8, 20×8 mm (Chính quyền) |
| Trang phụ lục, nhãn tự đánh số nối tiếp | Có (golden) | [ ] | [ ] | Chèn hai, ba phụ lục liên tiếp; Đảng: đánh số trang riêng đặt tay |
| Phạm vi lưu hành, dự thảo, không phổ biến, thông tin liên hệ | Có (golden) | — | [ ] | |
| Hồ sơ cơ quan: thêm, sửa, xoá, mặc định, dùng | Có | [ ] | [ ] | Đóng mở lại Word: hồ sơ còn |
| Xuất, nhập cấu hình sang máy khác | Có (khứ hồi, từ chối file sai) | [ ] | [ ] | Gồm preset |
| Preset, tuỳ chọn nhanh, ngoại lệ ngoài khoảng | Có | [ ] | [ ] | |

## 5. Trợ năng và xử lý lỗi

| Mục | Tự động | Word thật | Ghi chú |
| --- | --- | --- | --- |
| Mọi nút, ô nhập, ô chọn tới được bằng Tab và dùng được bằng Enter, Space | Không | [ ] | Duyệt từ trên xuống dưới, không cần chuột |
| Có khung nhìn thấy rõ khi điều khiển đang được chọn | Có (`:focus-visible` có trong CSS) | [ ] | |
| Tên cho trình đọc màn hình (Narrator) | Có (quy tắc đặt tên) | [ ] | Mở Narrator, nghe vài ô nhập |
| Tương phản màu chữ ≥ 4,5:1 | Có (tính từ CSS) | [ ] | Thử chế độ tương phản cao của Windows |
| Tiếng Việt: không còn chuỗi tiếng Anh, chính tả, dấu thanh | Không | [ ] | Rà tay toàn bộ giao diện |
| Thiếu bộ API: tính năng bị tắt kèm lý do, không lỗi im lặng | Có (mô phỏng Word cũ) | [ ] | Thử trên Word bản cũ hoặc Word Online nếu có |
| Lỗi không ai bắt hiện ở dòng trạng thái | Không | [ ] | Gây lỗi có chủ ý (ví dụ đóng tài liệu khi đang chạy) |

## Kết luận lượt thử

Điền sau khi thử: số dòng đạt, số dòng lỗi, danh sách lỗi cần sửa (ghi kèm bước tái hiện). Lỗi nào chặn việc dùng hằng ngày thì ghi mức "chặn".
