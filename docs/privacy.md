# Tuyên bố quyền riêng tư

Add-in "TTA-Format" dùng cá nhân, chạy trong Microsoft Word. Tóm tắt: **nội dung văn bản của bạn không rời khỏi máy; add-in không có thống kê, theo dõi hay báo cáo lỗi gửi đi.**

## Add-in đọc và ghi gì

- Đọc nội dung, định dạng và thiết lập trang của tài liệu đang mở **trong bộ nhớ của bảng add-in** để kiểm tra, sửa và dựng văn bản. Mọi xử lý diễn ra trên máy bạn.
- Ghi vào tài liệu đang mở đúng những gì bạn yêu cầu (định dạng, khối đầu, khối cuối, số trang…). Ghi thêm một thuộc tính nhỏ trong tài liệu cho biết chuẩn đang dùng (Chính quyền hoặc Đảng).
- Lưu cài đặt (hồ sơ cơ quan, preset, chuẩn mặc định, danh sách cơ quan) trong **bộ nhớ cục bộ (localStorage) của Word trên máy này**. Không có tài khoản, không đồng bộ lên máy chủ. Bạn chủ động xuất ra file để sao lưu hoặc chuyển máy.

## Add-in kết nối mạng ở đâu

Chỉ có hai việc tải về, không gửi gì lên:

1. **Các file của add-in** (HTML, JavaScript, biểu tượng) từ GitHub Pages của chủ repo. GitHub, như mọi máy chủ web, thấy yêu cầu tải file (địa chỉ IP, thời gian).
2. **Thư viện Office.js** từ `appsforoffice.microsoft.com` (Microsoft), thư viện bắt buộc của mọi add-in Office.

Mã nguồn **không có** lời gọi `fetch`, `XMLHttpRequest`, `WebSocket`, `sendBeacon` hay dịch vụ thống kê nào; điều này được kiểm bằng test tự động (`src/privacy.test.ts`) mỗi lần phát hành, và bản build được quét địa chỉ ngoài (`npm run check:release`) trước khi đưa lên trang công khai.

## Dữ liệu trên trang công khai

Trang GitHub Pages chỉ chứa bản build của add-in (`taskpane.html`, `taskpane.js`, `manifest.xml`, biểu tượng). Không có văn bản mẫu thật, không có thông tin cá nhân. Quy trình phát hành tự dừng nếu bản build chứa đường dẫn máy, địa chỉ thư hoặc file ngoài danh sách cho phép.

## Bạn có thể tự kiểm

Làm theo mục 8 của `huong-dan-cai-dat.md` để theo dõi mạng và xác nhận không có yêu cầu nào chứa nội dung văn bản.

## Liên hệ

Add-in dùng cá nhân, không có máy chủ hỗ trợ. Mã nguồn nằm trong repo cùng trang phát hành; mọi thay đổi cách xử lý dữ liệu sẽ được ghi ở đây và ở lịch sử commit.
