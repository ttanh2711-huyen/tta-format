# Kế hoạch triển khai: Add-in Word định dạng văn bản theo Nghị định 30/2020/NĐ-CP

> Trạng thái: **DỰ THẢO v5 – chờ duyệt**. v5: đã chốt Q12–Q14 (không có mẫu văn bản Đảng; giữ mặc định đề xuất của chuẩn Đảng nhưng mọi tuỳ chọn đều có ô "Tuỳ chỉnh…"; tự nhận diện chuẩn theo tiêu đề đầu văn bản và hỏi lại khi không chắc) và thêm nút **Chỉnh toàn bộ** (T6C). Chưa có `SPEC.md` và chưa có codebase (thư mục trống). Đã chốt: Office.js, Word 365, dùng cá nhân, cỡ chữ mặc định 14, chọn nhanh định dạng theo nhóm, host tĩnh miễn phí, tuỳ chọn cơ quan chủ quản/cơ quan ban hành, bảng đầu văn bản vừa vùng chữ, "Nơi nhận" và chữ ký chung một bảng. **v4: hai chuẩn văn bản, chọn được trong add-in:** (1) **Chính quyền** theo Nghị định 30/2020/NĐ-CP, (2) **Đảng** theo Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng (thay Hướng dẫn 36-HD/VPTW). File mẫu công văn của người dùng là đáp án chuẩn cho chuẩn Chính quyền (mục 6A); Phụ lục 1–2 của HD 05 là đáp án chuẩn cho chuẩn Đảng (mục 6B). Chi tiết từng task nằm ở `tasks/todo.md`.

## 1. Tổng quan

Xây add-in cho Microsoft Word giúp soạn thảo và chuẩn hoá văn bản theo **thể thức và kỹ thuật trình bày** của **hai chuẩn**, chọn bằng một công tắc "Chuẩn văn bản" ở đầu task pane:

| Chuẩn | Căn cứ | Dùng cho |
|---|---|---|
| **Chính quyền** | Nghị định 30/2020/NĐ-CP về công tác văn thư (Phụ lục I, mẫu ở Phụ lục III) | Văn bản hành chính của cơ quan nhà nước |
| **Đảng** | Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng (thi hành Quy định 399-QĐ/TW ngày 09/01/2026 của Ban Bí thư; thay Hướng dẫn 36-HD/VPTW) | Văn bản của cấp uỷ, tổ chức, cơ quan đảng từ Trung ương đến chi bộ |

Chuẩn được chọn quyết định toàn bộ: quy tắc, khoảng giá trị cho phép, mặc định chọn nhanh, bộ style, bố cục khối đầu và khối cuối, cách ghi số/ký hiệu, danh sách loại văn bản và bộ kiểm tra lỗi. Chuẩn của từng tài liệu được ghi vào thuộc tính tuỳ chỉnh của tài liệu và tự nhận diện được khi mở tài liệu có sẵn (có "ĐẢNG CỘNG SẢN VIỆT NAM" thì là chuẩn Đảng). Ba năng lực chính:

1. **Chuẩn hoá** văn bản có sẵn: khổ giấy, lề, phông chữ, cỡ chữ, giãn dòng/đoạn, thụt đầu dòng, đánh số trang. Nút **Chỉnh toàn bộ** chuẩn hoá cả tài liệu bằng một lần bấm theo chuẩn đang chọn: trang, style, phần thân, bảng biểu, khối đầu và khối cuối nhận diện được, số trang. Có hộp xem trước các bước sắp làm, báo cáo sau khi chạy và khôi phục lại bản trước khi chỉnh.
2. **Chọn nhanh tuỳ chỉnh định dạng, chia nhóm:** mọi tuỳ chọn định dạng đều chọn được một lần bấm, xếp theo nhóm: **Trang** (khổ, lề, khoảng cách header/footer), **Chữ** (phông, cỡ chữ nội dung), **Đoạn** (giãn dòng, giãn đoạn, thụt dòng đầu, căn lề), **Đầu văn bản** (cỡ quốc hiệu, tiêu ngữ, tên cơ quan, số/ký hiệu, ngày tháng, V/v, độ dài đường kẻ), **Cuối văn bản** (cỡ nơi nhận, số dòng trống chỗ ký, quyền hạn/chức vụ), **Bảng biểu**, **Số trang**, **Hành vi** (giữ/xoá định dạng cục bộ, phạm vi áp dụng). Có preset tự đặt tên. Mỗi tuỳ chọn có thêm ô "Tuỳ chỉnh…" để nhập giá trị riêng: giá trị nằm trong khoảng của chuẩn đang chọn thì áp dụng ngay; giá trị ngoài khoảng thì cảnh báo và chỉ áp dụng khi bạn xác nhận (đánh dấu là ngoại lệ chủ ý, bộ kiểm tra không báo lại mục đó).
3. **Tạo văn bản mới** từ biểu mẫu cho **mọi loại văn bản hành chính** theo NĐ 30 (công văn, quyết định, tờ trình, báo cáo, thông báo, kế hoạch, biên bản, hợp đồng, giấy mời... danh sách chính xác chốt ở T2): dựng khối đầu (quốc hiệu, tiêu ngữ, **cơ quan chủ quản và cơ quan ban hành**, số/ký hiệu, địa danh ngày tháng, tên loại, trích yếu, "Kính gửi" với công văn) và khối cuối (quyền hạn, chức vụ, họ tên người ký, nơi nhận). Cơ quan chủ quản là tuỳ chọn (không có / một dòng / nhiều dòng, như mẫu của bạn có hai dòng), cơ quan ban hành chọn từ danh sách đã lưu hoặc nhập mới.
4. **Kiểm tra lỗi thể thức** và sửa nhanh từng lỗi hoặc sửa tất cả, kèm báo cáo.

Giao diện tiếng Việt, dùng cá nhân trên Word 365. Toàn bộ xử lý chạy cục bộ trong Word; không gửi nội dung văn bản ra ngoài.

## 2. Giả định (cần xác nhận)

| # | Giả định | Nếu sai thì |
|---|---|---|
| A1 | **ĐÃ CHỐT:** Office Add-in (Office.js + TypeScript), task pane + ribbon | Spike T3 vẫn là cổng go/no-go nếu API thiếu |
| A2 | **ĐÃ CHỐT:** Word 365 (Microsoft 365 Apps) trên Windows là mục tiêu chính. Word Online và Word for Mac chỉ kiểm tra nhẹ ở T3/T23, không cam kết | Nếu cần Mac/Online đầy đủ thì thêm task kiểm thử |
| A3 | **ĐÃ CHỐT:** mọi loại văn bản hành chính theo NĐ 30 (các tên loại ở Điều 7 và Phụ lục III). **Không** gồm văn bản quy phạm pháp luật (nghị định, thông tư...) | Bổ sung giai đoạn riêng |
| A8 | **ĐÃ CHỐT:** hỗ trợ hai chuẩn: Chính quyền (NĐ 30) và Đảng (HD 05-HD/VPTW 27/5/2026). Phạm vi chuẩn Đảng: bản giấy và bản điện tử của văn bản cấp uỷ, tổ chức, cơ quan đảng; **không** gồm bản sao văn bản, văn bản mật (theo HD 05 thực hiện theo quy định bảo vệ bí mật nhà nước) và chữ ký số/dấu | Thêm giai đoạn riêng |
| A9 | Danh sách thể loại văn bản Đảng lấy từ Quy định 399-QĐ/TW (chưa đọc, làm ở T2B). HD 05 nhắc các ký hiệu NQ, CT, KL, QC, BC, QĐ, CTr, TT (thông tri), TTr, CV, v.v. | Điều chỉnh danh sách ở T2B |
| A7 | **ĐÃ CHỐT:** file mẫu công văn của người dùng (Chi cục Thủy lợi, công văn cử công chức tham dự tập huấn) là "đáp án chuẩn" cho công văn. File chứa họ tên, số điện thoại, email thật nên **không đưa nguyên bản vào repo**; chỉ lưu bản ẩn danh (T2A) | Dùng Phụ lục III làm đáp án thay thế |
| A4 | Không bao gồm: chữ ký số, con dấu, QR, bản sao văn bản, quản lý số văn bản, đồng bộ hệ thống quản lý văn bản | Ngoài phạm vi |
| A5 | **ĐÃ CHỐT:** dùng cá nhân, một người một máy; cấu hình lưu cục bộ, không backend, không tài khoản | Thêm backend thì đổi kiến trúc |
| A6 | **ĐÃ CHỐT:** cỡ chữ nội dung mặc định 14 (cho phép đổi 13 trong khoảng NĐ 30) | Đổi một hằng số trong profile |

## 3. Quyết định kiến trúc

1. **Tách logic thuần khỏi Office.js.** Quy tắc NĐ 30, bộ định dạng số/ngày, bộ kiểm tra và bộ dựng OOXML là hàm thuần (TypeScript), kiểm thử bằng Jest không cần Word. Lớp mỏng `office/` gọi Word API. Lý do: Office.js khó kiểm thử tự động; càng nhiều logic nằm ngoài lớp này càng tốt.
2. **Quy tắc là dữ liệu** (`src/rules/nd30.ts`): mỗi quy tắc có id, điều khoản tham chiếu, giá trị/khoảng cho phép, mức độ. Cùng một bảng quy tắc cung cấp cho cả bộ định dạng lẫn bộ kiểm tra, tránh lệch nhau.
3. **Định dạng bằng Style, không bằng định dạng trực tiếp.** Cài bộ style `ND30_*` vào tài liệu (idempotent). Kiểm tra lỗi dựa trên style + mẫu văn bản. Người dùng đổi một style là đổi cả văn bản.
4. **Khối đầu/cuối dựng bằng OOXML** (`insertOoxml`): bảng 2 cột không viền, đường kẻ ngang bằng border đoạn. Đây là cách ổn định nhất qua các phiên bản Word cho bố cục dạng bảng.
5. **Page setup** (lề, A4) dùng `Word.PageSetup` (WordApiDesktop 1.3). Tab stop (WordApiDesktop 1.4). API này chỉ có trên Word desktop bản mới, nên **task spike T3 phải chứng minh trước** (fail fast). Nếu không đạt, đổi hướng công nghệ trước khi tốn công.
6. **Cấu hình theo profile**: NĐ 30 cho phép khoảng giá trị (lề trái 30–35 mm, cỡ chữ 13–14...). Mặc định chọn giá trị phổ biến, người dùng đổi trong khoảng cho phép; ngoài khoảng thì cảnh báo.
7. **Phân phối cho dùng cá nhân (đã chốt host tĩnh miễn phí)**: web add-in là file tĩnh, host HTTPS miễn phí. Mặc định **GitHub Pages** (dự phòng: Cloudflare Pages). Nạp manifest qua shared-folder catalog trong Word. Máy chủ chỉ phát mã, không nhận nội dung văn bản; trang GitHub Pages công khai nên **không đưa dữ liệu cá nhân hay file mẫu thật vào repo/trang**. Không cần Centralized Deployment. Thực hiện ở T24.
8. **Chọn nhanh định dạng theo nhóm (presets)**: một registry `FormatOptions` điều khiển bằng dữ liệu, mỗi tuỳ chọn khai báo nhóm, nhãn tiếng Việt, các giá trị chọn nhanh, khoảng NĐ 30 cho phép và giá trị mặc định lấy từ file mẫu. Giao diện sinh từ registry nên thêm tuỳ chọn không phải viết thêm màn hình. Là nguồn duy nhất cho T4, T6, các khối đầu/cuối, wizard và bộ kiểm tra. Preset đặt tên của người dùng lưu cục bộ.
9. **Văn bản theo họ mẫu, cấu hình thay vì code**: các loại văn bản nhóm theo bố cục (có tên loại + trích yếu; công văn; Chương/Điều; nhiều bên ký; giấy tờ mẫu nhỏ). Mỗi loại là một mục trong registry; thêm loại mới không sửa wizard.
10. **File mẫu là đáp án chuẩn có kiểm thử tự động**: T2A trích thuộc tính định dạng từ file mẫu thành fixture ẩn danh; golden test của khối đầu/cuối và công văn hoàn chỉnh so với fixture này.
11. **`StandardProfile` là trục xuyên suốt:** mỗi chuẩn (`nd30`, `hd05`) là một đối tượng khai báo bộ quy tắc, registry tuỳ chọn nhanh (khoảng và mặc định riêng), bộ style (`ND30_*`, `HD05_*`), bộ dựng khối đầu/khối cuối, định dạng số/ký hiệu, danh sách loại văn bản, bộ kiểm tra. Mọi task dùng chung (page setup, style, chọn nhanh, wizard, checker) nhận `StandardProfile`; thêm chuẩn thứ ba về sau không phải sửa wizard hay checker. Chuẩn của tài liệu lưu ở thuộc tính tuỳ chỉnh của tài liệu và tự nhận diện được.
12. **Chuẩn Đảng có những khác biệt kỹ thuật không phải chỉ đổi con số** (mục 6B): bố cục khối đầu khác (tiêu đề ở góc phải, dấu `*` dưới tên cơ quan), giãn dòng kiểu Exactly 18–22 pt, định dạng "Số 05-HD/VPTW" (không có dấu hai chấm), 5 dấu gạch nối dưới trích yếu, "Nơi nhận" gạch chân, số trang cách mép trên 10 mm, lề đối xứng khi in hai mặt. Vì vậy chuẩn Đảng có các task khối đầu/cuối riêng (giai đoạn 3B), không nhồi vào task của chuẩn Chính quyền.
13. **Đường kẻ dưới tên cơ quan/tiêu ngữ/trích yếu dùng viền đoạn (inline), không dùng hình vẽ nổi.** File mẫu dùng hai đường thẳng nổi (floating connector) dễ lệch khi sửa nội dung; viền đoạn đi theo văn bản. Cần kiểm chứng ở T3 là điều khiển được độ dài đường kẻ.

## 4. Đồ thị phụ thuộc

```
T1 Scaffold ──┬── T2 Quy tắc NĐ30 (dữ liệu) ──────────────┐
              ├── T2B Quy tắc HD05 + StandardProfile ──────┤ (hai chuẩn dùng chung khung)
              ├── T2A Fixture ẩn danh từ file mẫu ──────────┤ (đáp án chuẩn cho golden test)
              └── T3 Spike khả thi API  ──► [Checkpoint 0: go/no-go công nghệ]
                                                            │
        T4 Page setup ◄─────────────────────────────────────┤
        T5 Bộ style NĐ30 ◄──────────────────────────────────┤
        T6 Áp định dạng nội dung (T4,T5)
        T6A Chọn nhanh định dạng / presets (T4,T6)
        T6B Bộ chọn chuẩn + nhận diện chuẩn tài liệu (T2B,T6A)
        T6C Nút "Chỉnh toàn bộ" (khung + bước trang/style/thân/bảng) (T4,T5,T6,T6A,T6B) ──► [Checkpoint 1]
              ▲ các bước đăng ký thêm: số trang (T11), khối đầu/cuối NĐ30 (T14), khối đầu/cuối Đảng (T14B), sửa lỗi còn lại (T15)
                                                            │
        T7 Quốc hiệu+Cơ quan ──► T8 Số/ký hiệu+Địa danh ──► T9 Tên loại+Trích yếu ──► [Checkpoint 2]
                                                            │
        T10 Chữ ký + Nơi nhận ; T11 Số trang
                       └──► T12 Wizard "Tạo Công văn" (T6A,T7–T11) ──► [Checkpoint 3 = MVP Chính quyền]
                                                            │
        TD1 Tiêu đề Đảng + cơ quan + dấu * ──► TD2 Số/ký hiệu + địa danh ngày ──► TD3 Thể loại + trích yếu + Kính gửi
        TD4 Chữ ký + Nơi nhận (Đảng) ; số trang Đảng (T11)
                       └──► TD5 Wizard "Tạo văn bản Đảng" (T12,TD1–TD4) ──► [Checkpoint 3B = MVP Đảng]
                                                            │
        T13 Engine kiểm tra (T2,T2B,T5) ──► T14 Kiểm tra NĐ30 (T7–T10) ; T14B Kiểm tra HD05 (TD1–TD4) ──► T15 Sửa tất cả ──► [Checkpoint 4]
                                                            │
        T16 Quyết định/Nghị quyết (cá biệt)
        T17 Registry + nhóm "tên loại + trích yếu"  (T12)
        T18 Quy chế/Quy định (Chương/Điều)          (T16)
        T19 Biên bản + văn bản nhiều bên ký         (T17)
        T20 Giấy tờ và mẫu nhỏ                      (T17)
        T21 Phụ lục + dấu khẩn/mật                  (T12)
        TD6 Đảng: Phần/Chương/Mục/Điều (nghị quyết, quyết định, quy định, quy chế) (TD5)
        TD7 Đảng: các thể loại còn lại theo Quy định 399 (TD5,T17)
        TD8 Đảng: phụ lục, dấu khẩn, phạm vi lưu hành, dự thảo, thông tin liên hệ (TD5)
                                                            │
        T22 Profile cơ quan ; T23 Kiểm thử Word 365 ; T24 Đóng gói/phân phối ──► [Checkpoint cuối]
```

Song song được: T7–T9 với T13 (sau T5); T16, T17, T21 độc lập nhau sau T12; T18–T20 song song sau khi registry T17 xong; tài liệu hướng dẫn song song với T23.

## 5. Danh sách task (chi tiết ở `tasks/todo.md`)

### Giai đoạn 0 – Nền tảng và kiểm chứng rủi ro
- [ ] T1: Khởi tạo add-in, chạy được trong Word, có hạ tầng test/lint
- [ ] T2: Mã hoá quy tắc NĐ 30 và danh sách loại văn bản thành dữ liệu, đối chiếu văn bản gốc
- [ ] T2B: Mã hoá HD 05 và thể loại văn bản Đảng thành dữ liệu, dựng khung `StandardProfile`
- [ ] T2A: Trích file mẫu thành fixture ẩn danh (đáp án chuẩn cho chuẩn Chính quyền)
- [ ] T3: Spike khả thi Office.js (lề, đường kẻ, tab stop, số trang, style) trên Word 365
- **Checkpoint 0:** xác nhận go công nghệ

### Giai đoạn 1 – Lát cắt đầu tiên: chuẩn hoá văn bản có sẵn
- [ ] T4: Thiết lập trang chuẩn (A4, lề, phông mặc định)
- [ ] T5: Cài bộ style NĐ 30
- [ ] T6: Áp định dạng nội dung cho văn bản có sẵn
- [ ] T6A: Chọn nhanh tuỳ chỉnh định dạng theo nhóm (registry theo chuẩn + giao diện + preset)
- [ ] T6B: Bộ chọn chuẩn Chính quyền/Đảng, tự nhận diện chuẩn theo tiêu đề đầu văn bản và hỏi lại khi không chắc
- [ ] T6C: Nút "Chỉnh toàn bộ" (khung chạy nhiều bước, xem trước, báo cáo, khôi phục)
- **Checkpoint 1:** văn bản mẫu lộn xộn thành đúng thể thức phần thân, đổi tuỳ chọn một lần bấm

### Giai đoạn 2 – Khối đầu văn bản
- [ ] T7: Quốc hiệu, tiêu ngữ, cơ quan chủ quản (tuỳ chọn, nhiều dòng) và cơ quan ban hành (chọn từ danh sách)
- [ ] T8: Số, ký hiệu, địa danh, ngày tháng
- [ ] T9: Tên loại văn bản, trích yếu, "Kính gửi"
- **Checkpoint 2:** đầu công văn khớp mẫu trong Phụ lục III

### Giai đoạn 3 – Khối cuối văn bản và MVP
- [ ] T10: Chữ ký (quyền hạn, chức vụ, họ tên) và nơi nhận
- [ ] T11: Đánh số trang
- [ ] T12: Wizard "Tạo công văn mới" end-to-end
- **Checkpoint 3 (MVP Chính quyền):** tạo công văn hoàn chỉnh theo NĐ 30, in PDF, đối chiếu mẫu

### Giai đoạn 3B – Văn bản Đảng theo HD 05-HD/VPTW
- [ ] TD1: Tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM", cơ quan cấp trên trực tiếp, cơ quan ban hành, dấu `*`
- [ ] TD2: Số và ký hiệu (dạng "Số 05-HD/VPTW"), địa danh, ngày tháng
- [ ] TD3: Thể loại, trích yếu (5 dấu gạch nối), "về việc" của công văn, "Kính gửi"/"Kính trình", căn cứ ban hành
- [ ] TD4: Quyền hạn (T/M, K/T, T/L, Q.), chức vụ, họ tên, "Nơi nhận" gạch chân, ký hiệu người soạn
- [ ] TD5: Wizard "Tạo văn bản Đảng" (báo cáo, công văn, thông báo, quyết định...)
- **Checkpoint 3B (MVP Đảng):** tạo văn bản Đảng hoàn chỉnh, in PDF, đối chiếu Phụ lục 1–2 của HD 05

### Giai đoạn 4 – Bộ kiểm tra lỗi thể thức
- [ ] T13: Engine kiểm tra + quy tắc cấp trang/phông chữ
- [ ] T14: Kiểm tra khối đầu/cuối và mẫu văn bản chuẩn Chính quyền (số/ký hiệu, ngày tháng, nơi nhận)
- [ ] T14B: Kiểm tra theo chuẩn Đảng (HD 05)
- [ ] T15: "Sửa tất cả" có xem trước và hoàn tác
- **Checkpoint 4:** báo cáo lỗi đúng trên bộ văn bản mẫu có lỗi cài sẵn

### Giai đoạn 5 – Phủ mọi loại văn bản hành chính (theo họ mẫu)
- [ ] T16: Quyết định, nghị quyết (cá biệt): căn cứ, Điều/Khoản/Điểm
- [ ] T17: Registry loại văn bản + nhóm "tên loại + trích yếu" (tờ trình, báo cáo, thông báo, kế hoạch, chương trình, hướng dẫn, đề án, dự án, phương án, thông cáo)
- [ ] T18: Quy chế, quy định (Chương/Mục/Điều)
- [ ] T19: Biên bản và văn bản nhiều bên ký (hợp đồng, bản ghi nhớ, bản thỏa thuận; ba loại này không có mẫu ở Phụ lục III)
- [ ] T20: Giấy tờ mẫu riêng (giấy mời, giấy giới thiệu, giấy nghỉ phép, công điện, thư công)
- [ ] T21: Phụ lục và dấu chỉ mức độ khẩn/mật
- [ ] TD6: Đảng: Phần/Chương/Mục/Điều/Khoản/Điểm (nghị quyết, quyết định, quy định, quy chế)
- [ ] TD7: Đảng: các thể loại còn lại theo Quy định 399-QĐ/TW
- [ ] TD8: Đảng: phụ lục, dấu khẩn, chỉ dẫn phạm vi lưu hành, dự thảo, thông tin liên hệ
- **Checkpoint 5:** mỗi họ mẫu có ít nhất một văn bản mẫu đúng thể thức; đối chiếu danh sách loại ở T2 không sót loại nào

### Giai đoạn 6 – Hoàn thiện và phát hành (dùng cá nhân)
- [ ] T22: Danh sách cơ quan đã lưu, profile và xuất/nhập cấu hình
- [ ] T23: Kiểm thử trên Word 365, trợ năng, xử lý lỗi/fallback
- [ ] T24: Đóng gói, GitHub Pages, nạp qua shared-folder catalog, hướng dẫn, tuyên bố riêng tư
- **Checkpoint cuối:** đạt toàn bộ tiêu chí, dùng được hằng ngày

## 6. Phụ lục tham chiếu: tóm tắt quy tắc thể thức (cần đối chiếu ở T2)

Nguồn: Phụ lục I và III của NĐ 30/2020/NĐ-CP. Phần trích dưới đây một phần đã đối chiếu với kết quả tra cứu (khổ A4, lề trên/dưới 20–25 mm, lề trái 30–35 mm, lề phải 15–20 mm, Times New Roman Unicode màu đen) và phần còn lại là **hiểu biết hiện có, phải đối chiếu với văn bản gốc trước khi mã hoá**.

| Thành phần | Quy tắc cần mã hoá (chờ xác minh ở T2) |
|---|---|
| Trang | A4 dọc; lề trên/dưới 20–25 mm; trái 30–35 mm; phải 15–20 mm |
| Phông | Times New Roman, Unicode, màu đen |
| Quốc hiệu | In hoa, cỡ 12–13, đứng, đậm |
| Tiêu ngữ | In thường (viết hoa chữ cái đầu mỗi cụm), cỡ 13–14, đứng, đậm; gạch nối giữa các cụm; có đường kẻ ngang bên dưới dài bằng dòng chữ |
| Tên cơ quan | In hoa, cỡ 12–13, đứng; tên cơ quan ban hành đậm; đường kẻ ngang bên dưới dài khoảng 1/3–1/2 dòng chữ |
| Số, ký hiệu | "Số:" cỡ 13, đứng; số nhỏ hơn 10 thêm số 0 phía trước; ký hiệu có "/" và "-" đúng quy cách |
| Địa danh, ngày tháng | Cỡ 13–14, nghiêng; ngày < 10 và tháng 1–2 thêm số 0 |
| Tên loại | In hoa, cỡ 13–14, đứng, đậm |
| Trích yếu | In thường, cỡ 14, đứng, đậm, có đường kẻ ngang bên dưới; công văn dùng "V/v" cỡ 12–13, đứng |
| Nội dung | Cỡ 13–14, đứng, dàn đều hai lề; thụt dòng đầu 1–1,27 cm; giãn đoạn tối thiểu 6 pt; giãn dòng từ đơn đến 1,5 |
| Phần/Chương/Mục/Điều/Khoản/Điểm | Quy cách cỡ chữ, kiểu chữ riêng từng cấp |
| Chức vụ người ký | Quyền hạn (TM., KT., TL., TUQ.) và chức vụ: in hoa, đứng, đậm, cỡ 13–14 |
| Họ tên người ký | In thường, đứng, đậm, cỡ 13–14; không ghi học hàm, học vị |
| Nơi nhận | "Nơi nhận:" cỡ 12, nghiêng, đậm; phần liệt kê cỡ 11, đứng |
| Số trang | Chữ số Ả Rập, cỡ 13–14, đứng, canh giữa lề trên; trang 1 không đánh số |

**Lưu ý hiệu lực:** cần xác nhận NĐ 30/2020 hiện còn hiệu lực nguyên văn, có văn bản sửa đổi/hướng dẫn mới nào không, và cách ghi tên cơ quan sau các đợt sắp xếp bộ máy. Việc này giao cho T2, dùng quy trình kiểm chứng nguồn pháp lý.

## 6A. Giá trị đo được từ file mẫu công văn của bạn (mặc định của add-in)

Đo trực tiếp từ OOXML của file mẫu. Đây là mặc định cho T4–T12; mọi giá trị nằm trong khoảng NĐ 30 (T2 sẽ xác nhận lại).

| Thành phần | Giá trị trong file mẫu |
|---|---|
| Khổ giấy, lề | A4 dọc; lề trên 20, dưới 20, trái 30, **phải 20 mm**; header 12,5 mm; footer khoảng 5,3 mm |
| Số trang | Header mặc định, trường PAGE ở giữa, trang đầu khác biệt (không hiện số trang 1) |
| Phông | Times New Roman, đặt bằng định dạng trực tiếp (Normal của file mẫu là Calibri 11) |
| Nội dung | 14 pt, căn đều hai lề, thụt dòng đầu 1,27 cm, giãn trước và sau 6 pt, giãn dòng 1,17 |
| "Kính gửi:" | 14 pt, căn giữa, giãn trước và sau 12 pt |
| Bố cục đầu văn bản | Một bảng 2 cột (khoảng 45% / 55%), không viền, 3 hàng: (cơ quan | quốc hiệu và tiêu ngữ), (số | địa danh ngày), (V/v | trống) |
| Cơ quan chủ quản | Tuỳ chọn; trong mẫu có **2 dòng** (Sở... / Thành phố...), in hoa, 13 pt, không đậm |
| Cơ quan ban hành | In hoa, 13 pt, **đậm** |
| Quốc hiệu / tiêu ngữ | Quốc hiệu in hoa 13 pt đậm; tiêu ngữ 14 pt đậm |
| Số, ký hiệu | 13 pt, căn giữa ô, giãn trước 6 pt; để trống số chờ vào sổ ("Số:      /CCTL-ĐĐPCTT") |
| Địa danh, ngày | 13 pt, nghiêng, căn giữa ô; để trống ngày chờ ký |
| V/v | 13 pt, căn giữa ô bên trái, đứng |
| Đường kẻ dưới cơ quan và tiêu ngữ | Hai đường thẳng nổi (floating), không phải viền đoạn |
| Bảng dữ liệu trong thân | 14 pt, tiêu đề đậm căn giữa, ô giãn trước và sau 3 pt |
| Nơi nhận | "Nơi nhận:" 12 pt đậm nghiêng; danh sách mặc định 11 pt; gạch đầu dòng, cuối mục `;`, mục cuối "Lưu: VT, ..." |
| Chữ ký | Quyền hạn + chức vụ in hoa đậm 14 pt, 6 dòng trống chỗ ký, họ tên đậm 14 pt |

Chênh lệch so với đề xuất ban đầu: lề phải mặc định là **20 mm** (theo mẫu) thay vì 15 mm; bảng đầu văn bản trong mẫu **rộng hơn vùng chữ** (thụt âm khoảng 1,4 cm, rộng khoảng 18,3 cm so với 16 cm vùng chữ), còn add-in sẽ dựng vừa vùng chữ (xem Q9).

## 6B. Chuẩn Đảng: giá trị lấy từ Hướng dẫn 05-HD/VPTW (27/5/2026)

Đọc từ bản PDF Hướng dẫn 05-HD/VPTW (nguồn: [sav.gov.vn](https://www.sav.gov.vn/Files/Others/761e611b-48d7-4b86-a5a2-4c17bac0f753/05.HD.VPTW.pdf)), riêng bảng cỡ chữ ở Phụ lục 2 đối chiếu bằng ảnh các trang của bảng. Hướng dẫn này thay Hướng dẫn 36-HD/VPTW (04/3/2018), thi hành Quy định 399-QĐ/TW ngày 09/01/2026. T2B sẽ mã hoá lại và rà soát một lần nữa.

**Trang:** A4 dọc; lề trên 20, dưới 20, trái 30, phải 15 mm. Khi in hai mặt thì mặt sau: trên 20, dưới 20, **trái 15, phải 30** (lề đối xứng). Bảng biểu được xoay ngang. Phông Times New Roman, Unicode TCVN 6909:2001, màu đen. **Số trang:** chữ số Ả Rập, giữa trang, cách mép trên 10 mm, không hiện ở trang 1, phụ lục đánh số riêng.

**Nội dung:** căn đều hai lề; dòng đầu thụt khoảng 10 mm; giãn đoạn tối thiểu 6 pt; giãn dòng **18–22 pt kiểu Exactly**; kết thúc bằng dấu chấm. Căn cứ ban hành: mỗi căn cứ một dòng, đầu dòng có gạch ngang, cuối dòng có dấu chấm phẩy, căn cứ cuối có dấu chấm.

| Thành phần | Loại chữ | Cỡ | Kiểu | Ghi chú vị trí |
|---|---|---|---|---|
| Tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM" | In hoa | 15 | Đứng, đậm | Góc phải, dòng đầu; dưới có đường kẻ ngang nét liền dài bằng tiêu đề |
| Cơ quan cấp trên trực tiếp | In hoa | 14 | Đứng | Góc trái, ngang tiêu đề |
| Cơ quan ban hành | In hoa | 14 | Đứng, đậm | Dưới cơ quan cấp trên; dưới có dấu `*` |
| Số và ký hiệu | In thường | 14 | Đứng | Căn giữa dưới tên cơ quan; dạng "Số 127-QĐ/TW" (**không có dấu hai chấm**) |
| Địa danh, ngày tháng | In thường | 14 | Nghiêng | Đối diện số và ký hiệu; "Hà Nội, ngày 20 tháng 10 năm 2025" |
| Thể loại văn bản | In hoa | 15–16 | Đứng, đậm | Giữa trang đầu |
| Trích yếu | In thường | 14–15 | Đứng, đậm | Dưới thể loại, dưới có **5 dấu gạch nối** (-----) |
| Trích yếu của công văn ("về việc...") | In thường | 12 | Nghiêng | Căn đối diện số và ký hiệu; công văn không ghi tên loại |
| Nội dung | In thường | 14–15 | Đứng | |
| "Phần", "Chương" + số | In thường | 14–15 | Đứng, đậm | Giữa; số La Mã hoặc chữ ("Phần thứ nhất") |
| Tên phần, chương | In hoa | 14–15 | Đứng, đậm | Dòng liền dưới |
| "Mục" + số; tên mục | In thường; in hoa | 14–15 | Đứng, đậm | Số Ả Rập |
| "Điều", số, tên điều | In thường | 14–15 | Đứng, đậm | Số Ả Rập, sau số có dấu chấm |
| Khoản; Điểm | In thường | 14–15 | Đứng | "1." ; "a)" |
| Quyền hạn ký (T/M, K/T, T/L, Q.) | In hoa | 14 | Đứng, đậm | Góc phải, dưới nội dung |
| Chức vụ người ký | In hoa | 14 | Đứng | **Không đậm** |
| Họ tên người ký | In thường | 14 | Đứng, đậm | Dưới chữ ký; không ghi học hàm, học vị |
| "Kính gửi", "Kính trình" | In thường | 14 | Nghiêng | Giữa trang đầu (công văn, tờ trình) |
| Tên cơ quan sau "Kính gửi"/"Kính trình" | In thường | 14–15 | Đứng | Mỗi nơi một dòng, gạch đầu dòng, cuối dòng chấm phẩy |
| "Nơi nhận" | In thường | 14 | Đứng, **gạch chân** | Góc trái, dưới nội dung; có đường kẻ ngang dưới cụm từ |
| Tên cơ quan sau "Nơi nhận" | In thường | 12 | Đứng | Mỗi mục kết thúc `;`, mục "Lưu" kết thúc `.` |
| Dấu chỉ mức độ khẩn (KHẨN, THƯỢNG KHẨN, HỎA TỐC) | In hoa | 14 | Đứng, đậm | Dưới số và ký hiệu |
| Chỉ dẫn phạm vi lưu hành ("TÀI LIỆU HỘI NGHỊ") | In hoa | 12 | Đứng, đậm | Góc phải, trang đầu, cách mép trên 10 mm |
| Chỉ dẫn dự thảo ("Dự thảo lần 2") | In thường | 14 | Nghiêng | Dưới số và ký hiệu |
| Ký hiệu người soạn, số bản phát hành | In thường | 11 | Đứng | Cuối "Nơi nhận" |
| Thông tin liên hệ cơ quan ban hành | In thường | 10 | Đứng | Trang cuối, dưới đường kẻ ngang hết vùng trình bày |
| Số trang | In thường | 14 | Đứng | |

**Ký hiệu văn bản:** số viết bằng chữ số Ả Rập, nhỏ hơn 10 thì thêm số 0; số và ký hiệu nối bằng "-"; tên loại và tên cơ quan viết tắt nối bằng "/"; liên cơ quan nối các tên bằng "-". Ký hiệu loại: NQ (nghị quyết), CT (chỉ thị), KL (kết luận), QC (quy chế), BC (báo cáo), QĐ (quyết định, quy định), CTr (chương trình), TT (thông tri), TTr (tờ trình). Ký hiệu cơ quan có danh mục riêng (ví dụ UBKT, BTGDV, BXDĐ, CB, TU) nên T2B mã hoá thành bảng tra. Văn bản đại hội dùng ký hiệu chung "ĐH". Số đánh liên tục từ 01 theo từng tên loại trong một nhiệm kỳ cấp uỷ.

**Quyền hạn ký:** T/M (thay mặt: đại hội, cấp uỷ, uỷ ban kiểm tra, hội đồng), K/T (ký thay: cơ quan chuyên trách làm việc theo chế độ thủ trưởng), T/L (thừa lệnh), Q. (quyền). Một số trường hợp có dòng "Xác nhận chữ ký của đồng chí..." (văn bản đại hội, biên bản hội nghị không có con dấu riêng). Chữ ký trên bản giấy bằng mực xanh, trên bản điện tử là ảnh PNG nền trong suốt màu xanh, canh giữa chức vụ và họ tên; add-in chỉ chừa chỗ ký, không xử lý chữ ký số.

### So sánh nhanh hai chuẩn (các điểm khác nhau cần mã hoá riêng)

| Điểm | Chính quyền (NĐ 30, theo file mẫu) | Đảng (HD 05) |
|---|---|---|
| Tiêu đề đầu văn bản | "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" + tiêu ngữ "Độc lập - Tự do - Hạnh phúc" (góc phải, 13/14 pt) | "ĐẢNG CỘNG SẢN VIỆT NAM" 15 pt đậm (góc phải), không có tiêu ngữ |
| Tên cơ quan | Cơ quan chủ quản (thường) + cơ quan ban hành (đậm), 13 pt, đường kẻ dưới | Cơ quan cấp trên (thường) + cơ quan ban hành (đậm), 14 pt, dấu `*` dưới |
| Số và ký hiệu | "Số: 05/CCTL-ĐĐPCTT", cỡ 13 | "Số 05-HD/VPTW", cỡ 14, không dấu hai chấm |
| Địa danh, ngày | Cỡ 13, nghiêng | Cỡ 14, nghiêng |
| Trích yếu | Công văn "V/v" cỡ 13; loại khác đậm có đường kẻ dưới | 5 dấu gạch nối dưới; công văn "về việc" 12 pt nghiêng |
| Lề phải | 15–20 mm (mẫu của bạn 20) | 15 mm cố định; in hai mặt thì đối xứng |
| Giãn dòng | Đơn đến 1,5 (mẫu 1,17) | 18–22 pt Exactly |
| Thụt dòng đầu | 1–1,27 cm | Khoảng 10 mm |
| Cỡ chữ nội dung | 13–14 | 14–15 |
| Chức vụ người ký | In hoa, đậm | In hoa, **không đậm** |
| "Nơi nhận" | 12 pt nghiêng đậm | 14 pt đứng, gạch chân |
| Tên cơ quan sau "Nơi nhận" | 11 pt | 12 pt |
| Số trang | Giữa trang, lề trên, trang 1 ẩn | Cùng vị trí, cách mép trên 10 mm, trang 1 ẩn |
| Loại văn bản | Điều 7 NĐ 30 | Quy định 399-QĐ/TW |

## 7. Rủi ro và giảm thiểu

| Rủi ro | Mức độ | Giảm thiểu |
|---|---|---|
| Office.js chưa điều khiển đủ (lề, tab stop, border) kể cả trên Word 365; một số API (WordApiDesktop 1.3/1.4) có thể còn ở bản preview | Trung bình (đã giảm nhờ chốt Word 365) | T3 spike sớm; Checkpoint 0; phương án dự phòng cho từng tính năng (OOXML) hoặc VSTO |
| Phải có mạng/host để nạp add-in (dùng cá nhân nhưng web add-in vẫn cần URL HTTPS) | Thấp | Host tĩnh miễn phí; T24 ghi hướng dẫn; có thể chạy `localhost` với chứng chỉ dev nếu cần offline |
| Hiểu sai quy tắc NĐ 30 (kích thước, ngoại lệ theo loại văn bản) | Cao | T2 đối chiếu văn bản gốc; có người am hiểu văn thư duyệt checkpoint 2, 3 và 5 |
| Thêm chuẩn Đảng gần như nhân đôi khối lượng (khối đầu/cuối, checker, thể loại) | Cao | `StandardProfile` dựng từ đầu (T2B, T6B); NĐ 30 làm trước để chứng minh khung; chuẩn Đảng đi theo giai đoạn riêng 3B với checkpoint MVP riêng; có thể hoãn TD6–TD8 nếu cần |
| Chuẩn Đảng không có file mẫu thật (Q12 đã chốt: không có); mặc định giãn dòng, cỡ tên loại, độ rộng cột phải tự chọn trong khoảng cho phép | Trung bình | Dùng khoảng của HD 05, ghi rõ "mặc định đề xuất"; mọi tuỳ chọn có ô "Tuỳ chỉnh…"; bạn duyệt in thử ở Checkpoint 3B |
| "Chỉnh toàn bộ" làm hỏng định dạng người dùng đã đặt, hoặc nhận diện sai đoạn (nhầm thân văn bản với khối ký, nhầm chuẩn) | Cao | Hộp xem trước liệt kê các bước và chuẩn sẽ áp dụng; xác nhận trước khi chạy; chỉ động vào phần nhận diện được với độ tin cậy cao, phần còn lại báo "bỏ qua"; chụp bản trước khi chỉnh để khôi phục; kiểm thử trên tài liệu lộn xộn (T3 kiểm chứng khôi phục khả thi) |
| Danh sách thể loại và danh mục chữ viết tắt cơ quan đảng nằm ở Quy định 399-QĐ/TW và Phụ lục, chưa đọc | Trung bình | T2B đọc nguyên văn trước khi mã hoá; mục chưa chắc đánh dấu "cần xác minh" |
| Hướng dẫn 05 mới ban hành (27/5/2026), có thể được sửa hoặc có công văn hướng dẫn thêm | Thấp | Quy tắc là dữ liệu có số phiên bản và ngày hiệu lực, cập nhật không phải sửa code |
| Phủ mọi loại văn bản làm phình phạm vi | Trung bình | Nhóm theo họ mẫu, registry cấu hình (T17–T20); chỉ làm mẫu có trong Phụ lục III; loại ít dùng làm sau |
| Preset cho phép người dùng chọn giá trị vi phạm NĐ 30 | Thấp | `FormatOptions` kiểm tra với khoảng quy tắc; chặn hoặc cảnh báo rõ (T6A) |
| `insertOoxml` ghi đè/trộn style hoặc làm hỏng định dạng sẵn có | Trung bình | Style tiền tố `ND30_`, thao tác idempotent, test trên tài liệu có nhiều định dạng |
| Bộ kiểm tra báo sai (false positive) làm người dùng mất tin | Trung bình | Mức độ lỗi (lỗi/cảnh báo/gợi ý), cho phép bỏ qua; bộ văn bản mẫu có/không lỗi |
| Hiệu năng với văn bản dài (100+ trang) | Trung bình | Xử lý theo lô, `context.sync()` tối thiểu; đo ở T15 |
| Phân phối web add-in (host HTTPS) | Thấp | Host tĩnh miễn phí, nạp qua shared-folder catalog; chốt ở T24 |

## 8. Quyết định đã chốt và câu hỏi còn mở

**Đã chốt (người dùng):** Office.js · Word 365 · dùng cá nhân · mọi loại văn bản hành chính theo NĐ 30, không cần văn bản quy phạm pháp luật · cỡ chữ 14 · chọn nhanh **tất cả** tuỳ chọn định dạng, chia nhóm phân loại · host tĩnh miễn phí · tuỳ chọn cơ quan chủ quản và cơ quan ban hành · file mẫu công văn làm đáp án chuẩn.

**Đã chốt thêm:** Q9 bảng đầu văn bản dựng vừa vùng chữ (tuỳ chọn "rộng như mẫu cũ" vẫn giữ trong nhóm Đầu văn bản) · Q10 không có mẫu cho các loại khác, dùng Phụ lục III của NĐ 30 và Phụ lục 1–2 của HD 05 · Q11 "Nơi nhận" và chữ ký chung một bảng 2 cột · **hai chuẩn: Chính quyền (NĐ 30) và Đảng (HD 05-HD/VPTW)**.

**Đã chốt thêm (v5):** Q12 không có mẫu văn bản Đảng, đối chiếu Phụ lục 1–2 của HD 05 và bạn duyệt bản in thử ở Checkpoint 3B · Q13 giữ mặc định đề xuất cho chuẩn Đảng (nội dung 14 pt, giãn dòng Exactly 20 pt, thụt dòng đầu 10 mm, giãn đoạn 6 pt, thể loại 16 pt, trích yếu 14 pt, lề 20/20/30/15 mm) **và mọi tuỳ chọn đều có thêm ô "Tuỳ chỉnh…"** · Q14 tự nhận diện chuẩn theo tiêu đề đầu văn bản (có "ĐẢNG CỘNG SẢN VIỆT NAM" là chuẩn Đảng, có "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" là chuẩn Chính quyền) và hỏi lại khi không chắc · thêm nút "Chỉnh toàn bộ".

**Còn mở (không chặn T1–T6B):**
- **Q15 – Giá trị tuỳ chỉnh ngoài khoảng của chuẩn:** Mình hiểu "thêm tuỳ chỉnh" là mỗi tuỳ chọn có ô nhập giá trị riêng. Đề xuất: trong khoảng của chuẩn thì áp dụng ngay; **ngoài khoảng thì cảnh báo và chỉ áp dụng khi bạn xác nhận**, đánh dấu là ngoại lệ chủ ý. Nếu bạn muốn chặn hẳn giá trị ngoài khoảng, hoặc có tuỳ chỉnh cụ thể nào chưa có trong bảng ở T6A, cho mình biết.
