# Bộ macro TTA-Format cho Word 2007 trở về sau

Dùng khi máy chỉ có Word 2007, 2010 hoặc 2013, là các bản không chạy được add-in. Bộ macro lấy cùng thông số với add-in, gồm khổ giấy, lề, phông, cỡ chữ, style, giãn dòng và style theo vai trò đoạn. Bộ có hai chuẩn: Chính quyền theo Nghị định 30/2020/NĐ-CP và Đảng theo Hướng dẫn 05-HD/VPTW.

## Cài một lần

1. Tải hai tệp `TTA_Config.bas` và `TTA_Format.bas` trong thư mục này về máy. Trên GitHub: mở từng tệp, bấm **Download raw file**.
2. Mở Word, bấm **Alt+F11** để mở trình soạn VBA.
3. Bên trái chọn **Normal** (Normal.dotm). Nếu muốn tách riêng, chọn một tệp .dotm đặt trong thư mục STARTUP của Word.
4. Bấm **File > Import File...**, chọn `TTA_Config.bas`. Làm lại bước này với `TTA_Format.bas`.
5. Bấm **Debug > Compile Normal**. Không có thông báo lỗi là được.
6. Đóng trình soạn VBA. Khi thoát Word, chọn **Save** nếu Word hỏi lưu Normal.dotm.

Nếu Word chặn macro:
- vào **File > Options > Trust Center > Trust Center Settings > Macro Settings** (trên Word 2007 là nút Office > Word Options);
- chọn "Disable all macros with notification".

Bạn tự chỉnh cài đặt bảo mật này trên máy mình.

### Thêm nút lên thanh công cụ

- Vào **File > Options > Quick Access Toolbar** (trên Word 2007 là Word Options > Customize).
- Ở "Choose commands from", chọn **Macros**.
- Thêm các macro `TTA_...` cần dùng, rồi bấm **Modify** để đặt biểu tượng và tên ngắn.

## Các macro

| Macro | Việc làm |
|---|---|
| `TTA_ChonChuan` | Chọn chuẩn: **Có (Yes)** là Chính quyền, **Không (No)** là Đảng. Lựa chọn được nhớ cho các lần sau. |
| `TTA_ThietLapTrang` | Đặt khổ A4, lề và khoảng cách đầu, chân trang theo chuẩn cho mọi phần của văn bản. Phần đang xoay ngang vẫn giữ chiều ngang. |
| `TTA_CaiStyle` | Cài hoặc cập nhật style `ND30_...` hay `HD05_...` vào văn bản đang mở, kể cả giãn dòng Exactly. |
| `TTA_DinhDangNoiDung` | Áp style theo vai trò đoạn: chương, tên chương, mục, điều, khoản, điểm, đề mục La Mã, ý gạch đầu dòng, thân bài. Nếu đang bôi đen thì chỉ áp cho phần đó, không thì áp cho cả văn bản. Các đoạn sau được bỏ qua: đoạn trong bảng, đoạn đã có style của chuẩn, Heading và danh sách tự động. |
| `TTA_DanhSoTrang` | Đặt số trang ở giữa lề trên, không hiện số ở trang 1. |
| `TTA_KhoiDau` | Chèn khối đầu ở đầu văn bản, gồm tên cơ quan, số ký hiệu, quốc hiệu và tiêu ngữ (hoặc tiêu đề Đảng), địa danh ngày tháng, tên loại và trích yếu, với chữ để sẵn để bạn sửa. |
| `TTA_KhoiCuoi` | Chèn khối cuối ở cuối văn bản, gồm nơi nhận, quyền hạn và chức vụ người ký, họ tên. |
| `TTA_KiemTra` | Kiểm tra khổ giấy, lề, số trang, phông, cỡ chữ, thụt dòng đầu, giãn dòng, rồi ghi kết quả ra một văn bản mới. |

Mọi macro hoàn tác được bằng **Ctrl+Z**. Riêng `TTA_DinhDangNoiDung` có thể cần bấm nhiều lần.

## Giới hạn so với add-in

- Không kẻ đường dưới tên cơ quan, tiêu ngữ, trích yếu. Hãy vẽ bằng **Insert > Shapes > Line**.
- Kiểm tra chỉ gồm phần trình bày cơ bản. Số và ký hiệu, viết hoa, nơi nhận, căn cứ… chỉ add-in kiểm tra được, trên Word 2016 trở lên.
- Không có mẫu văn bản, danh sách cơ quan hay sổ đăng ký văn bản đi.

## Cập nhật

Khi repo có bản mới:
1. Trong trình soạn VBA, bấm chuột phải vào hai mô-đun `TTA_Config` và `TTA_Format`, chọn **Remove** (không cần Export).
2. Nhập lại hai tệp mới như bước cài.

Hai tệp .bas được sinh tự động từ mã add-in (`src/vba/`), nên đừng sửa tay. Nếu sửa mã nguồn, chạy `UPDATE_VBA=1 npx jest src/vba` để sinh lại.
