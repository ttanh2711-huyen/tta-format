# Hướng dẫn cài đặt và sử dụng

Add-in "NĐ 30 - Định dạng văn bản" chạy trong Microsoft Word 365 trên Windows. Add-in định dạng văn bản theo hai chuẩn: Chính quyền (Nghị định 30/2020/NĐ-CP) và Đảng (Hướng dẫn 05-HD/VPTW). Cần Microsoft 365 bản **2507 trở lên** để dùng đủ tính năng (mục "Môi trường Word" ở đầu add-in cho biết tính năng nào chạy được trên máy bạn).

Add-in là các file tĩnh đặt trên GitHub Pages; máy bạn chỉ cần file **manifest.xml** để Word biết nạp add-in từ đâu. Không cài thêm phần mềm nào khác.

## 1. Lấy file manifest

1. Mở trình duyệt, vào địa chỉ phát hành: `https://<tên tài khoản GitHub>.github.io/<tên repo>/manifest.xml`.
2. Lưu file thành `manifest.xml` vào một thư mục cố định, ví dụ `C:\AddinWord\`. Chỉ cần đúng một file này trong thư mục.

## 2. Chia sẻ thư mục và khai báo với Word

Word đọc add-in từ "danh mục tin cậy" (một thư mục chia sẻ). Làm một lần:

1. Chuột phải thư mục `C:\AddinWord`, chọn **Properties > Sharing > Share…**, chia sẻ cho chính bạn rồi **Share**. Ghi lại đường dẫn mạng hiện ra, dạng `\\TÊN-MÁY\AddinWord`.
2. Mở Word, chọn **File > Options > Trust Center > Trust Center Settings… > Trusted Add-in Catalogs**.
3. Ở ô **Catalog Url**, dán đường dẫn mạng ở bước 1, chọn **Add catalog**, đánh dấu **Show in Menu**, bấm **OK** hai lần.
4. Đóng hẳn Word rồi mở lại.

## 3. Thêm add-in vào Word

1. Chọn **Insert > Add-ins > My Add-ins** (hoặc **Get Add-ins > My Add-ins**).
2. Chọn thẻ **SHARED FOLDER**, chọn **NĐ 30 - Định dạng văn bản**, bấm **Add**.
3. Trên thanh công cụ có thẻ **NĐ 30**, nhóm **Định dạng văn bản**; bấm nút trong nhóm để mở bảng add-in ở bên phải.

Add-in nhớ bạn đã thêm: lần sau mở Word chỉ cần bấm nút trong thẻ **NĐ 30**.

## 4. Dùng hằng ngày

- **Văn bản mới**: mục **Tạo văn bản mới**, chọn loại văn bản, điền Khối đầu và Khối cuối, bấm tạo. Add-in thiết lập trang, cài style, dựng khối đầu, thân, khối cuối và đánh số trang; sau đó bạn gõ nội dung vào chỗ con trỏ. Dùng **Khôi phục** nếu muốn quay lại bản trước khi tạo.
- **Văn bản có sẵn**: **Chỉnh toàn bộ** áp dụng các bước (trang, style, thân, bảng, khối đầu và cuối, số trang, lỗi còn lại) có xem trước. Add-in chụp bản trước khi chỉnh để khôi phục.
- **Kiểm tra văn bản**: liệt kê lỗi thể thức theo chuẩn đang chọn; "Sửa" từng lỗi hoặc "Sửa tất cả" có xem trước và hoàn tác; "Xuất báo cáo" để lưu danh sách lỗi.
- **Hồ sơ cơ quan**: lưu cơ quan, địa danh, ký hiệu, người ký để điền sẵn. Dùng **Xuất cấu hình** để sao lưu hoặc chuyển sang máy khác.
- **Chuẩn Chính quyền hay Đảng**: chọn ở đầu bảng. Add-in tự nhận diện chuẩn của tài liệu từ quốc hiệu hoặc tiêu đề đầu văn bản; không chắc thì hỏi, không tự áp định dạng.

Bấm Ctrl+Z trong Word hoàn tác được một lần sửa của add-in.

## 5. Phiên bản và cập nhật

- Số phiên bản hiện ở dòng đầu của add-in ("Phiên bản x.y.z · Môi trường Word…").
- Mỗi lần có bản mới, GitHub Pages tự cập nhật; bạn **không phải cài lại**. Nếu Word còn dùng bản cũ do bộ nhớ đệm: đóng Word, xoá nội dung thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, mở lại Word.
- Chỉ khi file `manifest.xml` đổi (ví dụ thêm nút) mới cần tải lại file này, ghi đè file trong thư mục `C:\AddinWord`, rồi đóng và mở lại Word.

## 6. Gỡ cài đặt

1. Trong Word, **Insert > My Add-ins**, chuột phải add-in, chọn **Remove** (nếu có), hoặc xoá file `manifest.xml` khỏi thư mục `C:\AddinWord`.
2. Tuỳ chọn: bỏ mục khỏi **Trust Center > Trusted Add-in Catalogs** và bỏ chia sẻ thư mục.
3. Hồ sơ và cài đặt của add-in lưu trong bộ nhớ cục bộ của Word trên máy này; **Xuất cấu hình** trước nếu muốn giữ. Xoá thư mục `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` để xoá sạch.

## 7. Khi gặp sự cố

| Hiện tượng | Cách xử lý |
| --- | --- |
| Không thấy thẻ **SHARED FOLDER** | Kiểm tra bước 2.3 đã đánh dấu **Show in Menu**; đóng hẳn Word và mở lại |
| Bảng add-in trắng hoặc báo không tải được | Kiểm tra mạng; mở `https://<tên tài khoản>.github.io/<tên repo>/taskpane.html` bằng trình duyệt xem có hiện không |
| Dòng đầu báo tính năng "không chạy được" | Cập nhật Microsoft 365 lên bản 2507 trở lên (File > Account > Update Options) |
| Nút "Tải về .txt" không chạy | Dùng nút **Sao chép** và dán vào Notepad |
| Dùng bản cũ sau khi cập nhật | Xoá bộ nhớ đệm như mục 5 |

## 8. Kiểm tra add-in không gửi nội dung văn bản ra ngoài

Xem `docs/privacy.md` để biết add-in làm gì với dữ liệu. Muốn tự kiểm: mở add-in rồi chạy công cụ theo dõi mạng (ví dụ Fiddler, hoặc **Developer Tools > Network** của WebView2 nếu có bật gỡ lỗi), dùng các tính năng với một văn bản thử có chữ dễ nhận ra (ví dụ "CHUOI-THU-123") và tìm chuỗi đó trong các yêu cầu mạng; kết quả mong đợi là không có. Chỉ có yêu cầu tải file tĩnh của add-in từ GitHub Pages và thư viện Office.js từ `appsforoffice.microsoft.com`.
