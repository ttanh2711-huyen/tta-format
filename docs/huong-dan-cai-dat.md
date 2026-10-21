# Hướng dẫn cài đặt và sử dụng

Add-in "TTA-Format" chạy trong Microsoft Word 365 trên Windows. Add-in định dạng văn bản theo hai chuẩn: Chính quyền (Nghị định 30/2020/NĐ-CP) và Đảng (Hướng dẫn 05-HD/VPTW). Dùng đủ tính năng trên Microsoft 365 bản **2507 trở lên**. Word 2016, 2019, 2021 bản mua đứt cũng chạy được hầu hết tính năng (style chỉ thêm mới, không cập nhật style đã có; số trang ẩn ở trang 1 bằng trường IF), riêng **thiết lập trang** phải đặt tay trong Layout > Margins. Word 2007, 2010, 2013 không chạy được add-in: dùng **bộ macro VBA** trong thư mục `vba/` của repo. Mục "Môi trường Word" ở đầu add-in cho biết tính năng nào chạy được trên máy bạn.

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
2. Chọn thẻ **SHARED FOLDER**, chọn **TTA-Format**, bấm **Add**.
3. Trên thanh công cụ có thẻ **TTA-Format**, nhóm **Định dạng văn bản**; bấm nút trong nhóm để mở bảng add-in ở bên phải.

Add-in nhớ bạn đã thêm: lần sau mở Word chỉ cần bấm nút trong thẻ **TTA-Format**.

## 4. Dùng hằng ngày

Đầu bảng TTA-Format có nút gạt **Nghị định 30 | Hướng dẫn 05** để chọn chuẩn đang dùng. Dưới đó là thanh ribbon năm tab theo trình tự công việc:

| Tab | Mục con |
|---|---|
| **Soạn thảo** | Tạo mới, Tạo mẫu, Chèn thêm (dấu khẩn, phụ lục) |
| **Khối** | Khối đầu, Khối cuối |
| **Định dạng** | Nội dung, Trang và style, Tuỳ chọn |
| **Kiểm tra** | Kiểm tra lỗi, Chỉnh toàn bộ |
| **Cơ quan** | Hồ sơ (cả sổ đăng ký, xuất nhập cấu hình), Môi trường, Thử nghiệm |

Mỗi lần chỉ hiện một mục con nên ít phải cuộn; add-in nhớ tab và mục con bạn mở lần trước. Dùng phím mũi tên trái, phải, Home, End để chuyển tab hoặc mục con. Kết quả kiểm tra gom lỗi cùng loại thành một nhóm, xếp theo khu vực (trang, phông, đoạn văn, khối đầu, căn cứ, khối cuối, viết hoa); bấm nhóm để xem từng chỗ, sửa từng chỗ hoặc **Sửa cả nhóm**.

- **Văn bản mới**: mục **Tạo văn bản mới**, chọn loại văn bản, điền Khối đầu và Khối cuối, bấm tạo. Add-in thiết lập trang, cài style, dựng khối đầu, thân, khối cuối và đánh số trang; sau đó bạn gõ nội dung vào chỗ con trỏ. Dùng **Khôi phục** nếu muốn quay lại bản trước khi tạo.
- **Người ký có hai chức danh** (ví dụ Trưởng ban kiêm Thủ tướng): ở Khối cuối, nhóm Người ký, ghi chức danh thứ nhất vào ô Chức vụ (TRƯỞNG BAN, nằm trên chữ ký) và chức danh thứ hai vào ô **Chức danh thứ hai, dưới chữ ký** (THỦ TƯỚNG, nằm dưới chữ ký, ngay trên họ tên). Cả hai dòng in hoa theo chuẩn.
- **Văn bản có sẵn**: **Chỉnh toàn bộ** áp dụng các bước (trang, style, thân, bảng, khối đầu và cuối, số trang, lỗi còn lại) có xem trước. Add-in chụp bản trước khi chỉnh để khôi phục.
- **Biến văn bản này thành mẫu** (tab Soạn thảo, mục Tạo mẫu): thay số, ngày tháng, trích yếu, căn cứ, thân, nơi nhận, người ký bằng ô điền có tiêu đề, giữ bố cục và style; lưu bằng File > Save As > Word Template (*.dotx). Nên làm trên bản sao; có nút Khôi phục bản gốc.
- **Định dạng nội dung** (tab Định dạng, mục Nội dung) nhận ra đề mục (Chương, Mục, Điều, khoản "1.", điểm "a)", đề mục "I.") và áp style đề mục tương ứng thay vì style thân bài. Có hai nút:
  - **Định dạng phần tô đen**: bôi đen các đoạn cần sửa rồi bấm; phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn, giãn dòng về đúng chuẩn, kể cả đoạn đã có style chuẩn; chữ đậm, nghiêng trong đoạn giữ nguyên.
  - **Định dạng cả văn bản**: áp cho mọi đoạn chưa dùng style chuẩn.
  - Tài liệu chưa có bộ style thì add-in tự cài trước. Hoàn tác bằng Ctrl+Z.
- **Kiểm tra văn bản**: liệt kê lỗi thể thức theo chuẩn đang chọn; "Sửa" từng lỗi hoặc "Sửa tất cả" có xem trước và hoàn tác; "Sửa nhanh" chạy Chỉnh toàn bộ rồi Sửa tất cả chỉ bằng một nút (hoàn tác được); "Xuất báo cáo" để lưu danh sách lỗi, mở được thành tệp Word.
- **Sổ đăng ký văn bản đi** (tab Cơ quan, mục Hồ sơ): mỗi lần tạo văn bản, add-in ghi một dòng (số, ký hiệu, ngày tháng, tên loại, trích yếu, nơi nhận, người ký); xem theo năm, xoá dòng, **Tải Excel (.xlsx)** hoặc **Mở bảng trong Word** để in, lưu.
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
