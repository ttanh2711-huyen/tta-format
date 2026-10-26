# TTA-Format – Add-in Word định dạng văn bản

Add-in Microsoft Word 365 (Office.js, TypeScript) soạn và chuẩn hoá văn bản theo hai chuẩn, dùng cá nhân:

- **Chính quyền**: Nghị định 30/2020/NĐ-CP (công văn, quyết định, quy chế, biên bản, giấy tờ… theo Phụ lục III).
- **Đảng**: Hướng dẫn 05-HD/VPTW theo Quy định 399-QĐ/TW (công văn, các thể loại, nghị quyết có Phần/Chương/Điều, biên bản, văn bản đại hội, liên cơ quan…).

Giao diện tiếng Việt. Không gửi nội dung văn bản ra ngoài (xem [docs/privacy.md](docs/privacy.md)).

## Giao diện

Nút gạt chuẩn **NĐ 30 | HD 05** ở đầu bảng; ribbon sáu tab **Nhanh** (việc hay làm, mỗi việc một lần bấm), **Soạn** (tạo mới, tạo mẫu, chèn thêm), **Khối** (khối đầu, khối cuối), **Định dạng** (nội dung, trang và style, tuỳ chọn), **Kiểm tra** (kiểm tra lỗi gom theo nhóm, chỉnh toàn bộ), **Cơ quan** (hồ sơ, môi trường, thử nghiệm); mỗi lần hiện một mục con, nhớ tab và mục đã mở. Xem thử giao diện ngoài Word: `npm run build:dev` rồi mở `dist/preview.html` qua một máy chủ tĩnh (trang này giả lập Office, không có trong bản phát hành).

## Tính năng chính

- Thiết lập trang, cài bộ style `ND30_*` / `HD05_*`, định dạng thân và bảng, đánh số trang.
- Dựng khối đầu, khối cuối và tạo văn bản mới theo loại (registry loại văn bản, danh sách đánh số Điều/Khoản/Điểm tự động).
- **Chỉnh toàn bộ** nhiều bước có xem trước và khôi phục bản trước.
- **Kiểm tra văn bản**: bộ quy tắc theo từng chuẩn, nút "Sửa" từng lỗi, "Sửa tất cả" có xem trước và hoàn tác, xuất báo cáo (.txt hoặc mở thành tệp Word), "Sửa nhanh" một nút (Chỉnh toàn bộ rồi Sửa tất cả); kiểm tra và sửa cả khối đầu, khối cuối.
- Dấu chỉ mức độ khẩn, phụ lục, các thành phần bổ sung của văn bản Đảng.
- Hồ sơ cơ quan, preset, xuất và nhập cấu hình.

## Cài đặt và sử dụng

Xem [docs/huong-dan-cai-dat.md](docs/huong-dan-cai-dat.md). Cần Microsoft 365 bản 2507 trở lên để dùng đủ tính năng.

## Phát triển

Cần Node.js 24 trở lên.

```bash
npm ci                 # cài thư viện
npm run dev-server     # máy chủ phát triển https://localhost:3000
npm start              # mở Word với bản thử từ manifest.dev.xml (manifest.xml là bản phát hành)
npm run typecheck && npm run lint && npm test
npm run build          # bản build thường
```

Cấu trúc: `src/standards` (quy tắc từng chuẩn), `src/ooxml` (dựng OOXML), `src/templates` và `src/standards/hd05/templates` (họ mẫu văn bản), `src/checker` (bộ kiểm tra), `src/office` (gọi Word), `src/ui` (giao diện), `src/settings` (cấu hình cục bộ), `tasks/` (kế hoạch và tiến độ), `docs/` (tài liệu).

## Phát hành

Đẩy lên nhánh `main` thì GitHub Actions ([.github/workflows/pages.yml](.github/workflows/pages.yml)) chạy kiểm tra, build bản phát hành, tạo manifest với địa chỉ GitHub Pages, quét bản build rồi đưa lên GitHub Pages. Làm tay trên máy:

```bash
npm run build:release
npm run manifest:release -- "https://<tài khoản>.github.io/<repo>"
npm run check:release
```

Đổi số phiên bản trong `package.json` trước khi phát hành (manifest lấy theo đó).

## Kiểm thử

`npm test` chạy kiểm thử tự động. Phần phải thử trong Word thật nằm ở [docs/test-matrix.md](docs/test-matrix.md).
