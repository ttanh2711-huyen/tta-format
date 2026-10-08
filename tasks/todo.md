# Danh sách task: Add-in Word theo Nghị định 30/2020/NĐ-CP

Xem bối cảnh, kiến trúc, rủi ro ở `tasks/plan.md`. Quy ước kiểm chứng sau T1:
- `npm test` = chạy Jest (logic thuần)
- `npm run build` = build TypeScript/webpack
- `npm start` = sideload add-in vào Word để kiểm tra thủ công
- Quy đổi: 1 mm = 2,8346 pt; 1 cm = 28,346 pt

Kích cỡ: XS 1 file, S 1–2, M 3–5. Không task nào quá 5 file.

---

## Giai đoạn 0 – Nền tảng và kiểm chứng rủi ro

### T1: Khởi tạo add-in và hạ tầng
**Mô tả:** Tạo dự án Office Add-in (TypeScript) cho Word, có ribbon tab "NĐ 30" và task pane tiếng Việt, cấu hình Jest, ESLint, thư mục `src/rules`, `src/ooxml`, `src/office`, `src/ui`.
**Tiêu chí nghiệm thu:**
- [ ] Sideload được vào Word desktop; ribbon hiện nút "NĐ 30" mở task pane (chờ bạn chạy `npm start` và xác nhận; manifest đã qua `npm run validate`)
- [ ] Task pane có nút thử nghiệm chèn một đoạn văn bản tiếng Việt có dấu đúng (mã đã viết, chờ thử trong Word)
- [x] Có 1 test Jest mẫu chạy qua (3 test)
**Kiểm chứng:**
- [x] `npm test` xanh; `npm run build` thành công; `npm run lint` và `npm run typecheck` sạch
- [ ] Thủ công: mở Word, bấm nút, thấy văn bản chèn đúng dấu
**Trạng thái:** mã và hạ tầng xong, 2 commit (`3485b1b`, `06ef12d`) tại `C:\Users\The Anh\dev\nd30-word-addin`. Còn 3 mục thủ công trong Word chờ xác nhận.
**Phụ thuộc:** không
**File dự kiến:** `package.json`, `manifest.xml`, `src/taskpane/*`, `jest.config.js`, `tsconfig.json`
**Kích cỡ:** M

### T2: Mã hoá quy tắc NĐ 30 và danh sách loại văn bản thành dữ liệu
**Mô tả:** Đọc văn bản gốc Phụ lục I và III, xác minh hiệu lực/sửa đổi, rồi mã hoá bảng quy tắc ở `plan.md` mục 6 thành `src/rules/nd30.ts` (id, điều khoản tham chiếu, khoảng/giá trị, mức độ, ghi chú ngoại lệ). Lập danh sách đầy đủ các tên loại văn bản hành chính (Điều 7) kèm chữ viết tắt ký hiệu (Phụ lục III) và họ mẫu bố cục thành `src/rules/docTypes.ts`. Ghi các điểm mơ hồ vào `docs/nd30-ambiguities.md`.
**Tiêu chí nghiệm thu:**
- [x] Mỗi quy tắc có trích dẫn vị trí trong Phụ lục I/III (trường `source`, có test)
- [x] `docTypes.ts` liệt kê 29 loại ở Điều 7 (27 loại có chữ viết tắt đúng Phụ lục III + công văn + thư công), mỗi loại có mẫu và họ mẫu. Điều 7 mới đối chiếu qua nguồn thứ cấp
- [ ] Đã xác nhận NĐ 30 còn hiệu lực và nêu rõ văn bản sửa đổi (nếu có): tìm kiếm chưa thấy văn bản sửa đổi, nhiều kế hoạch 2025–2026 vẫn viện dẫn; chờ bạn hoặc người am hiểu văn thư xác nhận (`docs/nd30-ambiguities.md`)
- [x] Có hàm thuần: `mmToPt`, `ptToMm`, `isMarginValid`, `formatSo`, `formatNgay` (thêm số 0) với test
**Kiểm chứng:**
- [x] `npm test`: test biên theo đúng quy tắc gốc (số 7 → "07"; ngày < 10 và tháng 1–2 thêm số 0; ngày tháng không hợp lệ bị từ chối). 70 test, `typecheck`, `lint`, `build` sạch
- [ ] Người am hiểu văn thư rà bảng quy tắc, ký xác nhận trong `docs/nd30-ambiguities.md`
**Trạng thái:** mã, test và tài liệu xong; còn hai mục chờ người xác nhận (hiệu lực, rà bảng quy tắc).
**Phụ thuộc:** T1
**File thực tế:** `src/standards/nd30/rules.ts`, `src/standards/nd30/docTypes.ts`, `src/standards/dateNumber.ts`, `src/standards/units.ts` (chuyển từ `src/rules/` ở T2B), `docs/nd30-ambiguities.md`
**Kích cỡ:** M

### T2A: Trích file mẫu thành fixture ẩn danh (đáp án chuẩn)
**Mô tả:** Viết script đọc file mẫu công văn của bạn (`Congvan_  Dang ky tap huan kiem soat TTHC.docx`, nằm ngoài repo) và xuất ra `fixtures/congvan-mau.json`: thuộc tính định dạng từng đoạn/ô/bảng (phông, cỡ, đậm/nghiêng, căn lề, thụt, giãn, độ rộng cột, lề trang, header/footer) và cấu trúc khối. **Ẩn danh toàn bộ nội dung cá nhân** (họ tên, số điện thoại, email, số hiệu văn bản thật) bằng chuỗi giả; không sao chép file `.docx` gốc vào repo.
**Tiêu chí nghiệm thu:**
- [x] `fixtures/congvan-mau.json` phản ánh đúng các giá trị ở mục 6A của `plan.md` (test `src/rules/fixture.test.ts` đối chiếu từng giá trị với mặc định của chuẩn Chính quyền: lề, header/footer, nội dung, Kính gửi, bảng dữ liệu, bảng đầu văn bản, cỡ chữ khối đầu, nơi nhận, chữ ký)
- [x] Không còn họ tên, số điện thoại, email thật trong repo: chỉ giữ chữ thể thức công khai (quốc hiệu, tiêu ngữ, "Nơi nhận:", "STT") và nhãn đầu dòng ("Kính gửi:", "Số:", "V/v"…), mọi chữ khác thành "[…]"; có test quét số điện thoại, email, tên và cơ quan cụ thể
- [x] Có hàm `compareToFixture(docModel, fixture, {ignore})` trả về danh sách khác biệt (chỉ so thuộc tính fixture khai báo, dung sai 0,06), dùng lại cho golden test ở T7–T12
**Kiểm chứng:**
- [x] `npm test`: 458 test; script chạy trên docx dựng sẵn (không dùng file thật vì nằm ngoài repo), hai lần chạy cho cùng kết quả
- [ ] Thủ công: **chờ bạn đọc `fixtures/congvan-mau.json` và xác nhận không lộ thông tin cá nhân**
**Phụ thuộc:** T1
**Trạng thái:** xong; `typecheck`, `lint`, `build` sạch. Chạy lại bằng `node tools/extract-sample.js "<file .docx>" fixtures/congvan-mau.json`. Mô hình `DocModel` (`src/rules/docModel.ts`) đọc định dạng trực tiếp từ OOXML (không giải style) và sẽ được bộ kiểm tra T13 dùng lại; bỏ qua khối `w:sdt`. Thêm devDependency `adm-zip` (đã có sẵn như phụ thuộc gián tiếp). Fixture ghi lại đúng điều kế hoạch đã nêu: bảng đầu văn bản rộng 183,4 mm thụt âm 14,1 mm (T7 dựng vừa vùng chữ nên phải bỏ qua cột/độ rộng khi so), đường kẻ dưới cơ quan và tiêu ngữ là hình nổi dài 17,5 mm và 59 mm.
**File thực tế:** `tools/extract-sample.js` (+ test), `fixtures/congvan-mau.json`, `src/rules/{docModel,compareToFixture}.ts` (+ test), `src/rules/fixture.test.ts`
**Kích cỡ:** M

### T2B: Mã hoá Hướng dẫn 05-HD/VPTW và thể loại văn bản Đảng, dựng khung `StandardProfile`
**Mô tả:** (1) Đọc nguyên văn Hướng dẫn 05-HD/VPTW ngày 27/5/2026 (Phần II, Phụ lục 1–2) và **Quy định 399-QĐ/TW ngày 09/01/2026** (thể loại, thẩm quyền ban hành), mã hoá bảng ở mục 6B của `plan.md` thành `src/standards/hd05/rules.ts`, danh sách thể loại và chữ viết tắt thành `src/standards/hd05/docTypes.ts`, bảng tra ký hiệu cơ quan đảng (UBKT, BTGDV, BXDĐ, CB, TU...) thành `src/standards/hd05/agencyCodes.ts`. (2) Dựng khung `StandardProfile` dùng chung và chuyển phần NĐ 30 của T2 vào `src/standards/nd30/` theo cùng khung: quy tắc, khoảng giá trị, mặc định, tên bộ style, định dạng số/ký hiệu, danh sách loại văn bản, điểm mở rộng cho bộ dựng khối đầu/cuối và bộ kiểm tra.
**Tiêu chí nghiệm thu:**
- [x] Mọi giá trị ở mục 6B có trích dẫn vị trí (trường `source` kiểm bằng test); cỡ chữ và chữ viết tắt cơ quan đã đối chiếu bằng ảnh trang PDF gốc. Điểm chưa chắc ghi ở `docs/hd05-ambiguities.md`
- [ ] Danh sách thể loại văn bản Đảng đủ theo Quy định 399-QĐ/TW, mỗi loại có ký hiệu viết tắt: **đã mã hoá 25 thể loại (Điều 5) và 8 giấy tờ (Điều 6) từ hai nguồn thứ cấp trùng nhau; một bài báo ghi 27 thể loại nên cần đối chiếu nguyên văn Điều 5**. Họ mẫu bố cục: HD 05 chỉ có sơ đồ chung (Phụ lục 1), không có mẫu riêng từng thể loại
- [x] `StandardProfile` có hai bản cài đặt (`nd30`, `hd05`); mã dùng chung chỉ nhận `StandardProfile`; test quét mã nguồn xác nhận không có `=== "nd30"`, `case "hd05"`... ngoài thư mục riêng của từng chuẩn
- [x] Hàm thuần theo chuẩn: định dạng số/ký hiệu ("Số 05-HD/VPTW" so với "Số: 05/BNV-VP"), ngày tháng, kiểm tra lề, cỡ chữ, giãn dòng, thụt dòng đầu (`src/standards/validate.ts`)
**Kiểm chứng:**
- [x] `npm test`: 156 test cho hai chuẩn (số 7 → "07"; "Số 05-HD/VPTW"; liên cơ quan "Số 05-QC/BTCTU-BNCTU"; giãn dòng 17 pt bị từ chối, 20 pt hợp lệ; lề trái 15 mm hợp lệ ở mặt sau khi in hai mặt). `typecheck`, `lint`, `build` sạch
- [ ] Thủ công: bạn rà bảng quy tắc HD 05 với bản PDF gốc, ký xác nhận trong `docs/hd05-ambiguities.md`
**Trạng thái:** mã, test và tài liệu xong; còn hai mục chờ người xác nhận (đối chiếu nguyên văn Điều 5 Quy định 399-QĐ/TW; rà bảng quy tắc).
**Phụ thuộc:** T1, T2
**File thực tế:** `src/standards/types.ts` (kiểu `StandardProfile`), `src/standards/index.ts`, `src/standards/validate.ts`, `src/standards/soKyHieu.ts`, `src/standards/{nd30,hd05}/profile.ts`, `src/standards/hd05/{rules,docTypes,agencyCodes}.ts`
**Kích cỡ:** M

### T3: Spike khả thi Office.js
**Mô tả:** Viết mã thử nghiệm tối thiểu cho: (a) đặt A4 và lề theo mm, (b) đường kẻ ngang bằng border đoạn có độ dài tuỳ chỉnh (thụt hai bên), so sánh với đường thẳng nổi mà file mẫu đang dùng rồi chọn phương án ổn định hơn, (c) bảng 2 cột không viền bằng `insertOoxml`, kể cả bảng thụt âm/rộng hơn vùng chữ như file mẫu, (d) tab stop, (e) số trang ở header, bỏ trang 1, (f) thêm style qua `insertOoxml`, (g) cho chuẩn Đảng: lề đối xứng khi in hai mặt (mặt sau trái 15, phải 30 mm), giãn dòng kiểu Exactly 18–22 pt qua style, khoảng cách số trang cách mép trên 10 mm, gạch chân cụm "Nơi nhận", dòng dấu `*` và 5 dấu gạch nối canh giữa, khối "Phụ lục" đánh số trang riêng, (h) cho nút "Chỉnh toàn bộ": gộp nhiều thay đổi trong ít lượt `Word.run` nhất để Ctrl+Z hoàn tác được trọn một lần, và chụp bản trước khi chỉnh (nội dung chính, header/footer, thuộc tính trang) rồi khôi phục được; ghi rõ phần nào khôi phục được hoàn toàn và phần nào không. Chạy trên Word 365 desktop (Windows) là chính; nếu tiện thì thử thêm Word Online và Word for Mac để biết giới hạn, không bắt buộc. Ghi kết quả vào `docs/spike-office-api.md` kèm đề xuất go/no-go.
**Tiêu chí nghiệm thu:**
- [x] Bảng "tính năng × môi trường" ghi rõ chạy được/không/ghi chú (Word 365 Windows; Online và Mac không thử vì không bắt buộc)
- [ ] Với mỗi tính năng không chạy, có phương án thay thế hoặc kết luận cần đổi công nghệ: đã có cho mọi mục trừ hai mục còn mở là Exactly qua style (có phương án dự phòng đã kiểm chứng) và đánh số phụ lục (có ba phương án chưa kiểm chứng)
- [x] Nêu rõ trạng thái GA/preview của từng requirement set và bản Microsoft 365 tối thiểu: WordApiDesktop 1.3 (page setup) cần bản 2507 (Build 19029.20208), 1.4 cần 2508; đều là bản phát hành chính thức, không có trên Office mua vĩnh viễn/LTSC (xem `docs/spike-office-api.md`)
**Kiểm chứng:**
- [ ] Thủ công trên Word 365 desktop: đã chạy; chưa có ảnh chụp và quan sát bằng mắt
- [ ] Người duyệt đọc tài liệu và xác nhận go (Checkpoint 0)
**Trạng thái:** đã chạy 3 lần trên Word 365 (16.0.20527); **kết luận GO với Office.js** (cả 4 mục bắt buộc đạt). Còn mở, không chặn: (1) chạy bổ sung p11 để biết style Exactly qua OOXML dùng được không; (2) đánh số trang phụ lục riêng chưa có cách đã kiểm chứng (để T21/TD8); (3) số lần Ctrl+Z (p14, p15) và quan sát bằng mắt (p07, p08) chờ bạn báo. Chi tiết và các quy tắc kỹ thuật rút ra ở `docs/spike-office-api.md`.
**Phụ thuộc:** T1
**File thực tế:** `src/spike/{probes,runner,report,ooxmlBuilders,ooxmlFacts,spike}.ts`, `src/spike/spike.html`, `tools/receiveSpikeReport.js`, `docs/spike-office-api.md`
**Kích cỡ:** M

### Checkpoint 0 – Xác nhận công nghệ
- [ ] T1, T2, T2A, T2B, T3 hoàn tất; `npm test`, `npm run build` xanh
- [ ] Spike cho thấy các tính năng cốt lõi chạy được trên Word 365
- [ ] **Người duyệt xác nhận tiếp tục với Office.js (hoặc nêu rõ tính năng cần phương án thay thế); cập nhật `plan.md`**

---

## Giai đoạn 1 – Lát cắt đầu: chuẩn hoá văn bản có sẵn

### T4: Thiết lập trang chuẩn
**Mô tả:** Nút "Thiết lập trang" đặt A4 dọc, lề mặc định theo file mẫu của bạn (trên 20, dưới 20, trái 30, phải 20 mm; header 12,5 mm, footer khoảng 5,3 mm), phông mặc định Times New Roman. Người dùng chỉnh lề trong khoảng cho phép; ngoài khoảng thì cảnh báo. Giá trị chọn nhanh đầy đủ nằm ở T6A. Nhận `StandardProfile`: chuẩn Đảng dùng lề 20/20/30/15 mm và tuỳ chọn "In hai mặt" (lề đối xứng, `mirrorMargins`), cho phép trang ngang cho bảng biểu; chuẩn Chính quyền dùng giá trị theo file mẫu.
**Tiêu chí nghiệm thu:**
- [ ] Sau khi bấm, mọi section của tài liệu có đúng khổ và lề (sai số ≤ 0,1 mm): mã áp dụng lên từng phần và đọc lại, so với kế hoạch bằng `findPageSetupMismatches` (có test); **chờ bạn thử trong Word**
- [x] Nhập lề ngoài khoảng của chuẩn thì hiện cảnh báo bằng tiếng Việt, không áp dụng (có test; áp dụng cả hai chuẩn, chuẩn Đảng báo "đúng 15 mm")
- [ ] Hoạt động với tài liệu nhiều section: vòng lặp áp dụng từng phần (spike p03 đã chứng minh đặt riêng từng phần), thông báo ghi "cho N phần"; **chờ bạn thử trong Word**
**Kiểm chứng:**
- [x] `npm test`: 244 test; validator lề, đổi mm → pt, mặc định từng chuẩn, từ chối lề sai, in hai mặt, đọc số có dấu phẩy, so sánh sau khi đọc lại, câu thông báo
- [ ] Thủ công: mở tài liệu Letter (hoặc A4 có nhiều phần), bấm nút, kiểm tra Page Setup trong Word
**Trạng thái:** mã, giao diện và test xong, `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bước "phông mặc định Times New Roman" (style Normal) là bước phụ, lỗi ở bước này không làm hỏng thiết lập trang và chưa kiểm chứng trong Word. Chuẩn được chọn tạm trong bảng này; T6B sẽ chuyển công tắc chuẩn lên đầu task pane; ô "Tuỳ chỉnh" ngoài khoảng của chuẩn (Q15) làm ở T6A.
**Phụ thuộc:** T2, T2B, T3
**File thực tế:** `src/office/pageSetupPlan.ts` (kế hoạch, kiểm tra, so sánh, thông báo), `src/office/pageSetup.ts` (áp dụng bằng Word.PageSetup), `src/ui/PageSetupPanel.ts`, thêm `pageDefaults` vào `StandardProfile`
**Kích cỡ:** S

### T5: Cài bộ style theo chuẩn (`ND30_*`, `HD05_*`)
**Mô tả:** Thêm các style `ND30_NoiDung`, `ND30_KinhGui`, `ND30_BangDuLieu`, `ND30_Phan`, `ND30_Chuong`, `ND30_Dieu`, `ND30_Khoan`, `ND30_Diem` v.v. theo quy tắc và giá trị ở mục 6A của `plan.md`, cài idempotent (chạy lại không nhân đôi, cập nhật nếu khác). Giá trị style lấy từ `FormatOptions` hiện hành (T6A) để đổi tuỳ chọn là đổi cả văn bản. Mỗi chuẩn có bộ style riêng với tiền tố khác nhau (`ND30_*` và `HD05_*`) để hai chuẩn cùng tồn tại trong một tài liệu không đè lên nhau; bộ `HD05_*` gồm cả giãn dòng Exactly và các cấp Phần/Chương/Mục/Điều/Khoản/Điểm theo mục 6B.
**Tiêu chí nghiệm thu:**
- [ ] Các style xuất hiện trong thư viện style, thuộc tính đúng bảng quy tắc: mã cài bằng `importStylesFromJson` rồi đọc lại từng style so với spec (`findStyleMismatches`, có test); **chờ bạn thử trong Word**. **Riêng giãn dòng Exactly của chuẩn Đảng chưa đạt**: `importStylesFromJson` cho ra quy tắc atLeast (spike p11), nên kết quả cài `HD05_*` sẽ báo rõ "cần Exactly, Word ghi atLeast" cho các style nội dung
- [ ] Chạy lần hai không tạo bản sao, không làm mất định dạng của đoạn đã dùng style: dùng `Overwrite` cập nhật tại chỗ (đã chứng minh ở p10); **chờ bạn thử trong Word**
- [x] Không đổi style Normal/Heading của người dùng: chỉ nhập các style do spec khai báo, test kiểm tên có tiền tố ND30_/HD05_ và không trùng tên style có sẵn của Word
**Kiểm chứng:**
- [x] `npm test`: 280 test; spec từng style, bộ dựng JSON và OOXML khớp file mẫu đối chiếu (`src/ooxml/__golden__/{nd30,hd05}-styles.{json,xml}`), so sánh sau khi đọc lại, câu thông báo
- [ ] Thủ công: cài hai lần, kiểm tra danh sách style
**Trạng thái:** mã, giao diện (nút "Cài bộ style" trong bảng thiết lập trang) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Giãn dòng Exactly cho `HD05_*`: bộ dựng OOXML đã có (`toStyleDefinitionsXml`, giãn dòng Exactly 400 twips, là cách duy nhất biết đặt được Exactly) nhưng **chưa dùng để cài** vì phải chèn một đoạn rỗng vào tài liệu rồi xoá, và chưa kiểm chứng style Exactly qua OOXML (chờ chạy bổ sung p11 sau khi sửa). Các style tiêu đề khối đầu/cuối (quốc hiệu, nơi nhận, họ tên...) thêm ở T7 đến T10; giá trị lấy từ `FormatOptions` (T6A) nối sau. `ND30_Phan`/`ND30_Chuong` gộp thành `ND30_PhanChuong`; phần chữ in hoa dựa vào chữ gõ vào, không dùng định dạng IN HOA của Word.
**Phụ thuộc:** T2, T2B, T3
**File thực tế:** `src/ooxml/{styleSpecs,styleJson,styleXml,facts}.ts`, `src/ooxml/__golden__/*`, `src/office/{installStyles,styleVerify}.ts`, `src/standards/{nd30,hd05}/styles.ts`, nút trong `src/ui/PageSetupPanel.ts`
**Kích cỡ:** M

### T6: Áp định dạng nội dung cho văn bản có sẵn
**Mô tả:** Nút "Định dạng nội dung" áp style `ND30_NoiDung` (mặc định theo mẫu: dàn đều, thụt dòng đầu 1,27 cm, giãn trước và sau 6 pt, giãn dòng 1,17, cỡ 14; đổi được ở T6A) cho các đoạn thân văn bản trong vùng chọn hoặc toàn bộ; bỏ qua bảng, header/footer và các đoạn đã dùng style `ND30_*` khác.
**Tiêu chí nghiệm thu:**
- [ ] Đoạn thân văn bản đạt đúng thuộc tính; bảng và đoạn có style `ND30_*`/`HD05_*` khác không bị đổi (logic chọn đoạn đã có test; **chờ bạn thử trong Word**)
- [ ] Không làm mất in đậm/nghiêng cục bộ do người dùng cố ý (cấu hình được: giữ/xoá; chế độ xoá dùng `Font.reset()` + `Paragraph.reset()`, cần Word 2508; **chờ bạn thử**)
- [ ] Hoàn tác (Ctrl+Z) quay lại trạng thái trước
**Kiểm chứng:**
- [x] `npm test`: 315 test; logic chọn đoạn áp dụng/bỏ qua, câu thông báo, đặt Exactly trong OOXML, bỏ sectPr cấp thân, kiểm tra kết quả Exactly
- [ ] Thủ công: văn bản mẫu lộn xộn trước/sau
**Trạng thái:** mã, giao diện (bảng "Định dạng nội dung" dưới bảng thiết lập trang, dùng chuẩn đang chọn ở đó) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Chuẩn Đảng (Exactly): theo lựa chọn của bạn, đặt Exactly trực tiếp trên từng đoạn bằng OOXML (`getOoxml` → `withExactLineSpacing` → `insertOoxml(..., "Replace")`). Vì chưa thử được trong Word, mã tự chạy một phép thử trên tài liệu tạm (`createDocument()`, không mở ra) rồi mới làm trên tài liệu thật; không đạt thì chỉ áp style và báo cảnh báo. Sau khi ghi, đọc lại số đoạn, chữ và số đoạn Exactly; lệch thì báo và nhắc Ctrl+Z. **Rủi ro chưa kiểm chứng**: (1) cách ghi lại đoạn có thể chèn thêm đoạn trống hay làm mất bình luận/đánh dấu gắn trong đoạn; (2) tài liệu tạm từ `createDocument()` chưa biết có tự đóng không; (3) Ctrl+Z có hoàn tác trọn một lần bấm không (như p14/p15). Chuẩn NĐ 30 không đi qua OOXML, chỉ áp style. Lựa chọn đoạn: bỏ qua trống, trong bảng, đã dùng style bộ chuẩn, tiêu đề có sẵn của Word, mục danh sách tự động.
**Phụ thuộc:** T4, T5
**File thực tế:** `src/office/{paragraphClassifier,formatBody}.ts`, `src/ooxml/{exactSpacing,exactVerify}.ts`, `src/ui/FormatPanel.ts`; `mountPageSetupPanel` trả về `currentStandard` để dùng chung (T6B sẽ chuyển chọn chuẩn lên đầu task pane)
**Kích cỡ:** M

### T6A: Chọn nhanh tuỳ chỉnh định dạng theo nhóm (registry + giao diện + preset)
**Mô tả:** Tuỳ chọn nhanh ở đầu task pane, **tất cả** thuộc tính định dạng đều chọn được một lần bấm, xếp trong các nhóm thu gọn/mở rộng. Mỗi tuỳ chọn là một mục trong registry `FormatOptions` (nhóm, nhãn tiếng Việt, các giá trị chọn nhanh, khoảng NĐ 30, mặc định theo file mẫu); giao diện sinh từ registry. Mọi lựa chọn áp ngay vào tài liệu hiện tại và đồng thời là mặc định cho các khối và wizard.

| Nhóm | Tuỳ chọn (giá trị chọn nhanh, **mặc định**) |
|---|---|
| **Trang** | Khổ giấy A4 cố định · Lề: **Theo mẫu 20/20/30/20 mm**, Lề phải 15 mm, Đóng gáy 20/20/35/20 mm, Tuỳ chọn (từng cạnh trong khoảng cho phép) · Khoảng cách header **12,5 mm**, footer **5,3 mm** |
| **Chữ** | Phông **Times New Roman** · Cỡ chữ nội dung 13 \| **14** · Màu **đen** |
| **Đoạn** | Căn lề **đều hai bên** \| trái · Thụt dòng đầu 1,0 \| **1,27 cm** · Giãn trước/sau **6 pt** \| 8 \| 12 (không dưới 6) · Giãn dòng Đơn \| **1,17 (như mẫu)** \| 1,5 \| Tối thiểu 15 pt |
| **Đầu văn bản** | Cỡ quốc hiệu 12 \| **13** · Cỡ tiêu ngữ 13 \| **14** · Cỡ tên cơ quan 12 \| **13** · Cỡ số/ký hiệu **13** \| 14 · Cỡ địa danh ngày **13** \| 14 · Cỡ V/v 12 \| **13** · Độ dài đường kẻ dưới cơ quan (1/3 \| **1/2** dòng chữ) · Bảng đầu văn bản: **vừa vùng chữ** \| rộng như mẫu cũ (Q9) |
| **Cuối văn bản** | Cỡ "Nơi nhận" **12** · Cỡ danh sách nơi nhận **11** · Số dòng trống chỗ ký 4 \| 5 \| **6** \| 7 · Cỡ quyền hạn/chức vụ/họ tên 13 \| **14** |
| **Bảng biểu** | Cỡ chữ trong bảng 13 \| **14** · Giãn ô trước/sau 0 \| **3 pt** · Viền: **có** \| không · Hàng tiêu đề **đậm, căn giữa** |
| **Số trang** | Vị trí **giữa lề trên** · Trang 1 **không hiện số** \| hiện · Cỡ **14** |
| **Hành vi** | Phạm vi áp dụng: **vùng chọn** \| toàn văn bản · Định dạng cục bộ (đậm/nghiêng do người dùng đặt): **giữ** \| xoá |

**Chuẩn Đảng (HD 05):** cùng 8 nhóm nhưng registry riêng, khoảng và mặc định khác (ghi "đề xuất" chờ Q13). Bảng trên là của chuẩn Chính quyền; bộ chọn ở T6B quyết định registry nào hiện ra.

| Nhóm | Tuỳ chọn chuẩn Đảng (giá trị chọn nhanh, **mặc định đề xuất**) |
|---|---|
| **Trang** | Lề **20/20/30/15 mm** · In hai mặt: **tắt** \| bật (mặt sau trái 15, phải 30 mm, lề đối xứng) · Hướng trang: **dọc** \| ngang (cho bảng biểu) · Khoảng cách số trang cách mép trên **10 mm** |
| **Chữ** | Phông **Times New Roman** · Cỡ nội dung 14 \| 15 (**14**) |
| **Đoạn** | Căn **đều hai bên** · Thụt dòng đầu **10 mm** · Giãn trước/sau **6 pt** (tối thiểu) · Giãn dòng **Exactly** 18 \| 19 \| **20** \| 21 \| 22 pt |
| **Đầu văn bản** | Cỡ tiêu đề Đảng **15** (cố định) · Cỡ cơ quan **14** (cố định) · Cỡ số/ký hiệu **14** · Cỡ địa danh ngày **14** · Cỡ thể loại 15 \| **16** · Cỡ trích yếu **14** \| 15 · Cỡ "về việc" của công văn **12** · Số gạch nối dưới trích yếu **5** |
| **Cuối văn bản** | Cỡ quyền hạn/chức vụ/họ tên **14** · Chức vụ **không đậm** · Số dòng trống chỗ ký 4 \| 5 \| **6** · Cỡ "Nơi nhận" **14** gạch chân · Cỡ danh sách nơi nhận **12** · Cỡ ký hiệu người soạn **11** |
| **Bảng biểu** | Cỡ chữ trong bảng 13 \| **14** · Giãn ô trước/sau 0 \| **3 pt** · Viền **có** · Hàng tiêu đề đậm căn giữa |
| **Số trang** | Giữa trang · Trang 1 **không hiện số** · Cỡ **14** · Phụ lục đánh số riêng **bật** |
| **Hành vi** | Như chuẩn Chính quyền |

Preset đặt tên ("Công văn thường", "Báo cáo dài"...): lưu, đổi tên, xoá, nhân bản, áp lại một lần bấm; **mỗi preset thuộc một chuẩn**. Chuẩn Chính quyền có sẵn preset "Theo mẫu công văn của tôi"; chuẩn Đảng có sẵn preset "Theo HD 05 (mặc định đề xuất)". Ba nhóm Đầu văn bản, Cuối văn bản, Số trang chỉ có hiệu lực khi chèn hoặc tạo khối ở T7–T11; trước đó hiển thị "áp dụng khi tạo văn bản".
**Tiêu chí nghiệm thu:**
- [x] Mọi tuỳ chọn ở bảng trên có mặt, đúng nhóm, đúng nhãn tiếng Việt; thêm tuỳ chọn mới chỉ cần thêm mục registry (`buildRegistry`); trừ các mục ghi ở Trạng thái
- [ ] Bấm một lựa chọn thuộc nhóm Trang/Chữ/Đoạn/Bảng biểu thì tài liệu đổi ngay và hoàn tác được (Ctrl+Z): mã xong (Trang qua `applyPageSetup`, Chữ/Đoạn/Bảng qua cập nhật bộ style tại chỗ); **chờ bạn thử trong Word**
- [x] Mỗi tuỳ chọn số có ô "Tuỳ chỉnh…" nhập giá trị riêng (đã chạy thử trong trình duyệt giả lập Office: chọn nhanh, nhập tay, hộp xác nhận ngoại lệ, số sai) (đúng đơn vị: pt, mm, cm). Giá trị trong khoảng của chuẩn đang chọn thì áp dụng ngay; giá trị ngoài khoảng hiện cảnh báo kèm lý do bằng tiếng Việt và chỉ áp dụng khi người dùng xác nhận, được đánh dấu là "ngoại lệ chủ ý" (bộ kiểm tra T13 không báo lại mục đó)
- [x] Nút "Về mặc định" cho từng tuỳ chọn, "Đặt lại nhóm", "Đặt lại tất cả"; "Lưu làm mặc định" ghi vào localStorage của task pane (kiểm tra trong trình duyệt giả lập; trong Word thật **chờ bạn thử**, nếu localStorage bị chặn thì rơi về bộ nhớ tạm)
- [x] Preset lưu, đổi tên, nhân bản, xoá, áp lại; giá trị lỗi thời bị bỏ qua có thông báo (có test). "Còn nguyên sau khi đóng mở Word" **chờ bạn thử**
- [ ] Mặc định chuẩn Chính quyền khớp `fixtures/congvan-mau.json` (T2A): **chưa làm được vì T2A chưa xong**; hiện test đối chiếu với `pageDefaults` và `styles.defaults` của chuẩn (vốn lấy từ file mẫu). Mặc định chuẩn Đảng khớp bảng quy tắc HD 05 (đã có test)
- [x] Mỗi chuẩn có registry riêng; giá trị hợp lệ của chuẩn này bị từ chối ở chuẩn kia khi cần (lề phải 20 mm ở chuẩn Đảng cần xác nhận ngoại lệ; Exactly 20 pt ở chuẩn Chính quyền bị từ chối hẳn) (ví dụ giãn dòng Exactly 20 pt chỉ có ở chuẩn Đảng; lề phải 20 mm bị từ chối ở chuẩn Đảng)
**Kiểm chứng:**
- [x] `npm test`: 372 test; registry hợp lệ (mỗi mục có mặc định nằm trong khoảng), lưu/đọc preset, ánh xạ lựa chọn sang style và thiết lập trang
- [ ] Thủ công: đổi cỡ chữ 14 → 13, giãn dòng 1,17 → đơn, áp preset, Ctrl+Z
**Phụ thuộc:** T2A, T2B, T4, T6
**Trạng thái:** mã, giao diện (bảng "Tuỳ chọn nhanh" ở trên bảng thiết lập trang) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Điểm khác kế hoạch: (1) bỏ **màu chữ** (chuẩn chỉ cho đen, một lựa chọn), **hướng trang ngang**, **bộ lề đặt tên** ("Đóng gáy 20/20/35/20"; thay bằng chọn nhanh từng cạnh: nhỏ nhất, mặc định, lớn nhất), **giãn dòng "Tối thiểu 15 pt"** (trái quy định bội số), **độ dài đường kẻ dưới cơ quan**, **số gạch nối dưới trích yếu**; (2) giãn đoạn trước và sau gộp một tuỳ chọn; (3) nhóm Đầu/Cuối văn bản, Số trang và hai tuỳ chọn viền/hàng tiêu đề bảng chỉ ghi nhận, dùng khi tạo văn bản ở T7 trở đi (đã ghi chú trên giao diện); (4) căn trái là ngoại lệ cần xác nhận vì chuẩn quy định dàn đều; (5) giá trị sai hẳn (không phải số, ngoài giới hạn cứng như Exactly 20 pt ở chuẩn bội số) bị từ chối, không cho xác nhận. Nút "Cài bộ style" và "Định dạng nội dung" dùng bộ style theo tuỳ chọn hiện hành; nhóm Hành vi (phạm vi, định dạng cục bộ) điều khiển nút "Định dạng nội dung". Chọn chuẩn vẫn ở bảng thiết lập trang (T6B sẽ chuyển lên đầu); bảng thiết lập trang còn các ô lề riêng trùng với nhóm Trang, T6B sẽ gộp. Chuẩn Đảng: đổi giãn dòng làm style đổi (atLeast) còn đoạn đã đặt Exactly giữ giá trị cũ cho đến khi bấm lại "Định dạng nội dung". Cập nhật style khi chưa cài bộ style thì chỉ ghi nhận, không tự cài. Cỡ chữ nội dung chuẩn Đảng chọn 14 hoặc 15 theo `rules.ts`. Cần bổ sung vào dữ liệu chuẩn: `defaultPresetName`, `styles.blockSizes`, `StyleOptions.alignment/table`.
**File thực tế:** `src/settings/{formatOptions,presets,storage}.ts` (+ test), `src/office/applyOptions.ts`, `src/ui/QuickOptions.ts`, sửa `src/office/formatBody.ts`, `src/ui/{FormatPanel,PageSetupPanel}.ts`, `src/ooxml/styleSpecs.ts`, `src/standards/*`
**Kích cỡ:** M

### T6B: Bộ chọn chuẩn (Chính quyền / Đảng) và tự nhận diện chuẩn của tài liệu
**Mô tả:** Công tắc "Chuẩn văn bản" ở đầu task pane với hai lựa chọn "Chính quyền (NĐ 30)" và "Đảng (HD 05)". Chuẩn đang chọn quyết định registry tuỳ chọn nhanh (T6A), bộ style (T5), nút "Thiết lập trang" (T4), nút "Định dạng nội dung" (T6) và sau này wizard, bộ kiểm tra. Ghi chuẩn vào thuộc tính tuỳ chỉnh của tài liệu (`nd30.standard`). Mở tài liệu có sẵn thì tự nhận diện: tìm "ĐẢNG CỘNG SẢN VIỆT NAM" hoặc "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" trong phần đầu; không chắc thì hỏi lại, không tự áp dụng. Chuẩn mặc định cho tài liệu mới lưu trong cài đặt cục bộ.
**Tiêu chí nghiệm thu:**
- [x] Đổi chuẩn thì toàn bộ tuỳ chọn nhanh, nhãn và giá trị mặc định đổi theo (kiểm tra trong trình duyệt giả lập: Giãn dòng 1,17 ↔ Exactly 20 pt); các ô lề riêng của bảng thiết lập trang đã bỏ để không còn hai nơi chỉnh
- [ ] Chuẩn của tài liệu được nhớ qua lần mở sau (ghi vào thuộc tính tuỳ chỉnh `nd30.standard` khi người dùng chọn; **chờ bạn thử: lưu, đóng, mở lại trong Word**); tài liệu chưa có chuẩn thì nhận diện hoặc hỏi (xong, có test)
- [ ] Nhận diện đúng với tài liệu chuẩn Chính quyền, chuẩn Đảng và tài liệu trống: có test với dòng tiêu đề dựng sẵn; **chưa thử trên file mẫu thật** (T2A chưa làm) và chưa có văn bản mẫu chuẩn Đảng thật
- [x] Đổi chuẩn của tài liệu đã có nội dung hiện cảnh báo và chỉ đổi khi người dùng bấm "Đổi chuẩn"; add-in không bao giờ tự sửa nội dung khi đổi chuẩn
- [x] Khi không chắc (không có tiêu đề, có cả hai tiêu đề, tiêu đề sai chính tả) hiện hộp thoại (tiêu đề ngắt dòng lạ thì vẫn nhận ra được, không phải hỏi)  "Văn bản này theo chuẩn nào?" với gợi ý và hai lựa chọn; có tuỳ chọn "Không hỏi lại cho tài liệu này"; không bao giờ tự đoán rồi áp dụng định dạng
**Kiểm chứng:**
- [x] `npm test`: 393 test; nhận diện chuẩn với các ca có/không có tiêu đề, lẫn lộn, sai chính tả, ngắt dòng, cụm từ giữa câu, tài liệu trống; quyết định dùng/hỏi; chuẩn mặc định cho tài liệu mới
- [ ] Thủ công: mở file mẫu công văn, mở văn bản thử chuẩn Đảng, đổi qua lại và xem tuỳ chọn đổi theo
**Phụ thuộc:** T2B, T6A
**Trạng thái:** mã, giao diện (công tắc ở đầu task pane) và test xong; `typecheck`, `lint`, `build` sạch; đã chạy trong trình duyệt giả lập Office với 7 kịch bản (trống, chính quyền, Đảng, không tiêu đề, cả hai tiêu đề, sai chính tả, đã ghi chuẩn); **chưa thử trong Word**. Quy tắc: tiêu đề phải là cả một dòng (hoặc tối đa 3 đoạn liền nhau ghép lại) sau khi bỏ dấu và dấu câu, nên cụm "Đảng Cộng sản Việt Nam" nằm giữa câu không bị coi là tiêu đề; chỉ xét 30 đoạn có chữ đầu tiên. Chuẩn đã ghi trong tài liệu được ưu tiên hơn tiêu đề (không cảnh báo nếu hai bên khác nhau). Tiêu đề rõ ràng thì dùng ngay nhưng **không ghi vào tài liệu** (tránh sửa tài liệu chỉ vì mở); chỉ ghi khi người dùng chọn hoặc trả lời hộp thoại. Chuẩn người dùng chọn gần nhất lưu làm mặc định cho tài liệu mới (localStorage). Có thêm nút "Để sau" trong hộp thoại. Bảng thiết lập trang chỉ còn nút "Thiết lập trang" (lấy lề từ Tuỳ chọn nhanh) và "Cài bộ style"; hàm `planPageSetup` (T4) không còn được giao diện gọi nhưng vẫn giữ cùng test.
**File thực tế:** `src/standards/detectStandard.ts`, `src/settings/defaultStandard.ts`, `src/office/documentStandard.ts`, `src/ui/StandardSwitcher.ts` (+ test), sửa `src/ui/PageSetupPanel.ts`, `src/taskpane/taskpane.{ts,html}`
**Kích cỡ:** S

### T6C: Nút "Chỉnh toàn bộ"
**Mô tả:** Nút nổi bật ở đầu task pane, chuẩn hoá **cả tài liệu** bằng một lần bấm theo chuẩn đang chọn (T6B) và các tuỳ chọn hiện hành (T6A). Dựng một **khung chạy nhiều bước** (pipeline) cho phép các task sau đăng ký thêm bước. Bản đầu có các bước: (1) thiết lập trang (T4), (2) cài/cập nhật style (T5), (3) định dạng phần thân và bảng biểu (T6). Luồng: bấm nút → hộp xem trước nêu chuẩn đã nhận diện, danh sách bước sẽ chạy, số đoạn/bảng sẽ bị đổi và phần sẽ bỏ qua → xác nhận → chạy → báo cáo từng bước (đã làm, bỏ qua, lỗi) → nút "Khôi phục bản trước khi chỉnh". Phần nào không nhận diện được với độ tin cậy cao thì bỏ qua và báo, không đoán.
**Tiêu chí nghiệm thu:**
- [ ] Văn bản có sẵn lộn xộn bấm một lần ra đúng thể thức phần thân và trang theo chuẩn đang chọn; chạy lại lần hai không đổi gì thêm (idempotent): mã xong (mỗi bước tự kiểm tra trạng thái hiện tại và bỏ qua khi đã đúng; đoạn đã dùng style của bộ chuẩn không bị áp lại); **chờ bạn thử trong Word**
- [x] Hộp xem trước hiện đúng chuẩn, đúng bước và số lượng thay đổi; huỷ ở bước này không đổi gì trong tài liệu (lập kế hoạch chỉ đọc; đã thử luồng trong trình duyệt giả lập; số liệu thật **chờ bạn thử**)
- [ ] Hoàn tác được: nút "Khôi phục bản trước khi chỉnh" thay lại thân văn bản và lề bằng bản chụp (cách đã đạt ở spike p13); **Ctrl+Z quay lại trọn một lần thì chưa kiểm chứng** (p14/p15 chưa có số liệu), nhiều bước nên có thể cần bấm nhiều lần. Chờ bạn thử
- [x] Bước nào lỗi thì dừng, các bước trước giữ nguyên, báo rõ bước lỗi, các bước sau ghi "chưa chạy", và hiện nút Khôi phục (có test và đã thử luồng giả lập)
- [x] Task sau đăng ký bước mới bằng `StepRegistry.register` với `order`, truyền registry vào `mountFixAllButton`; không sửa khung (có test thứ tự và trùng mã)
**Kiểm chứng:**
- [x] `npm test`: 424 test; khung chạy bước (thứ tự, lỗi giữa chừng, bỏ qua, chạy lại, tiến độ), hộp xem trước, báo cáo, câu chữ kế hoạch từng bước, phân loại bảng
- [ ] Thủ công: văn bản mẫu lộn xộn trước/sau cho cả hai chuẩn; Ctrl+Z; nút Khôi phục
**Phụ thuộc:** T4, T5, T6, T6A, T6B (và kết quả T3 về khôi phục)
**Trạng thái:** mã, giao diện (nút xanh dưới công tắc chuẩn) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bốn bước: (1) thiết lập trang theo Tuỳ chọn nhanh và đặt phông mặc định, (2) cài/cập nhật style, (3) định dạng phần thân cả tài liệu (chuẩn Đảng đặt Exactly như T6), (4) áp style chữ bảng cho **bảng có viền**. Bảng không viền (đầu văn bản, nơi nhận, chữ ký) và bảng lồng nhau được giữ nguyên và báo trong xem trước; viền và hàng tiêu đề của bảng chưa được đổi (hai tuỳ chọn này là "áp khi tạo văn bản"). Chụp bản gốc (thân văn bản + lề từng phần) ngay trước khi chạy; nếu không chụp được thì không chạy gì. Khôi phục: thay lại thân bằng `insertOoxml(Replace)` và đặt lại lề; **không khôi phục** header/footer (chưa bước nào sửa chúng; task sửa header/footer phải mở rộng `snapshot.ts`) và định nghĩa style ND30_/HD05_ (giữ giá trị mới); bản chụp nằm trong bộ nhớ task pane, đóng pane là mất; khôi phục có thể ảnh hưởng bình luận, đánh dấu. Bước chuẩn Đảng ghi lại từng đoạn nên kế thừa rủi ro của T6. Idempotent nhờ mỗi bước tự đọc trạng thái; riêng quy tắc giãn dòng Exactly của style do JSON cài không được tính là "cần cập nhật" (giới hạn đã biết). Có `snapshots` và `registry` truyền vào được để thử giao diện không cần Word.
**File thực tế:** `src/pipeline/{runner,steps,planText}.ts` (+ test), `src/office/{snapshot,formatTables}.ts`, `src/ui/FixAllButton.ts`; thêm `checkPageSetup` (pageSetup.ts), `planStyleInstall` (installStyles.ts), `previewFormatBody` (formatBody.ts), `classifyTable`/`describeSkipped` (paragraphClassifier.ts)
**Kích cỡ:** M

### Checkpoint 1 – Chuẩn hoá phần thân
- [ ] `npm test`, `npm run build` xanh
- [ ] Văn bản mẫu lộn xộn qua T4+T6 đúng thể thức phần thân (đối chiếu bảng quy tắc)
- [ ] Chọn nhanh (T6A) đổi định dạng đúng và không cho vượt khoảng của chuẩn đang chọn
- [ ] Nút "Chỉnh toàn bộ" (T6C) chạy trọn vẹn trên văn bản mẫu lộn xộn, xem trước đúng, hoàn tác và khôi phục hoạt động
- [ ] Bộ chọn chuẩn (T6B) hoạt động: cùng một văn bản lộn xộn định dạng thân đúng cả theo chuẩn Chính quyền lẫn chuẩn Đảng (giãn dòng Exactly 18–22 pt, lề 20/20/30/15 mm)
- [ ] Người duyệt xem demo và đồng ý hướng đi trước khi làm khối đầu/cuối

---

## Giai đoạn 2 – Khối đầu văn bản

### T7: Quốc hiệu, tiêu ngữ, cơ quan chủ quản và cơ quan ban hành
**Mô tả:** Form nhập cơ quan và chèn bảng 2 cột không viền (tỷ lệ khoảng 45%/55% như mẫu): trái là tên cơ quan, phải là quốc hiệu và tiêu ngữ, mỗi bên có đường kẻ dưới bằng viền đoạn (độ dài theo tuỳ chọn T6A).
- **Cơ quan chủ quản (tuỳ chọn):** công tắc "Có cơ quan chủ quản"; khi bật nhập 1–3 dòng (nút "Thêm dòng"), mỗi dòng tự in hoa, không đậm, cỡ 13 (như mẫu có hai dòng: Sở... / Thành phố...). Tắt thì khối co lại, cơ quan ban hành lên đầu.
- **Cơ quan ban hành:** ô chọn từ danh sách cơ quan đã lưu hoặc nhập mới, có nút "Lưu vào danh sách" (lưu bằng kho cục bộ của T6A; T22 bổ sung màn hình quản lý, xuất/nhập); tự in hoa, đậm, cỡ 13.
**Tiêu chí nghiệm thu:**
- [x] Quốc hiệu/tiêu ngữ đúng chữ, cỡ, kiểu, in hoa/thường theo quy tắc và theo `fixtures/congvan-mau.json` (test dùng `compareToFixture` với các đoạn của hàng đầu trong fixture)
- [x] Cơ quan chủ quản bật/tắt, 0 đến 3 dòng đều sinh khối đúng (có test); cơ quan ban hành chọn từ danh sách đã lưu hoặc nhập mới (đã thử trong trình duyệt giả lập)
- [ ] Đường kẻ có độ dài đúng tỷ lệ cho phép, đi theo văn bản (viền dưới của một đoạn riêng, không hình nổi): độ dài tính bằng **ước lượng bề rộng chữ** (bảng độ rộng Times New Roman, lệch vài %, đã đối chiếu với file mẫu: tiêu ngữ 59 mm trong ±6%); **chờ bạn xem bản in thử**
- [x] Tên cơ quan dài tự xuống dòng không làm vỡ bảng (bảng bố cục cố định, tên không thụt; đường kẻ tính theo bề rộng ô); **xem bằng mắt chờ bạn**
**Kiểm chứng:**
- [x] `npm test`: 491 test; bộ dựng OOXML khớp golden file (`__golden__/header-block.xml`); `compareToFixture` không khác biệt; 0/1/2/3 dòng chủ quản
- [ ] Thủ công: in PDF, so với file mẫu của bạn và Phụ lục III
**Phụ thuộc:** T2A, T5, T6A
**Trạng thái:** mã, giao diện (form "Khối đầu văn bản" trên Tuỳ chọn nhanh) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bảng 1 hàng 2 cột 45%/55%, không viền, lề ô 0, vừa vùng chữ (theo lề trong Tuỳ chọn nhanh), thụt 0; cơ quan chủ quản không đậm, cơ quan ban hành đậm, quốc hiệu 13 đậm, tiêu ngữ 14 đậm, đều in hoa trừ tiêu ngữ. Hàng số/ký hiệu, địa danh và V/v thêm ở T8, T9 (cần mở rộng `buildHeaderBlock`). Chèn vào đầu thân bằng `insertOoxml(..., "Start")` kèm một đoạn trống, đọc lại bảng đầu để kiểm tra; tài liệu đã bắt đầu bằng bảng thì hỏi trước khi chèn thêm; tài liệu trống sẽ có hai đoạn trống sau bảng. Thêm tuỳ chọn `agencyLineRatio` (1/3 hoặc 1/2, mặc định 1/2) và hai trường `headerKind`, `agencyLineRatio` cho chuẩn; chuẩn Đảng hiện chỉ hiện ghi chú "TD1". `flatOpc` chuyển từ spike sang `src/ooxml/flatOpc.ts`. Danh sách cơ quan lưu theo chuẩn trong localStorage (T22 bổ sung xuất/nhập).
**File thực tế:** `src/ooxml/{headerBlock,textWidth,flatOpc}.ts` (+ test, golden), `src/office/insertHeader.ts`, `src/ui/HeaderForm.ts`, `src/settings/agencies.ts` (+ test)
**Kích cỡ:** M

### T8: Số, ký hiệu, địa danh, ngày tháng
**Mô tả:** Thêm hàng thứ hai của bảng đầu văn bản: "Số: …/KÝ-HIỆU" (trái, cỡ 13, giãn trước 6 pt) và "Địa danh, ngày … tháng … năm …" (phải, nghiêng, cỡ 13). Có kiểm tra định dạng số/ký hiệu và ngày tháng; cho phép để trống số và ngày như file mẫu ("Số:      /CCTL-ĐĐPCTT").
**Tiêu chí nghiệm thu:**
- [x] Số < 10 tự thêm số 0; ký hiệu đúng quy cách ("/" giữa số và ký hiệu, "-" giữa nhóm viết tắt, không cách chữ): dùng `formatSoKyHieu` của chuẩn nên chuẩn Đảng cũng đúng ("Số 05-HD/VPTW")
- [x] Ngày < 10 và tháng 1–2 thêm số 0; ngày không hợp lệ (31/2, 29/2 năm không nhuận) bị từ chối
- [x] Cho phép để trống số và từng phần của ngày (chưa vào sổ, chưa ký): chừa 6 khoảng trắng đúng vị trí, giữ nguyên trong tài liệu (`xml:space="preserve"`)
**Kiểm chứng:**
- [x] `npm test`: 516 test; bảng biên cho số, ký hiệu, ngày, và hàng thứ hai khớp fixture
- [ ] Thủ công: kiểm tra bố cục với công văn mẫu
**Phụ thuộc:** T7
**Trạng thái:** mã, giao diện (mục "Thêm hàng số, ký hiệu và ngày tháng" trong form khối đầu) và test xong; `typecheck`, `lint`, `build` sạch; đã chạy trong trình duyệt giả lập (đủ, để trống, ngày sai, ký hiệu sai); **chưa thử trong Word**. Hàng thứ hai nằm trong cùng bảng của T7 (số căn giữa 13 pt, giãn trước 6 pt; địa danh nghiêng 13 pt). Ký hiệu nhập ở hai ô: "viết tắt loại" (để trống cho công văn) và "viết tắt cơ quan" nối bằng "-". Địa danh được nhớ cho lần sau (localStorage). Công văn chuẩn Chính quyền không có chữ viết tắt loại; việc chọn loại văn bản và tự điền ký hiệu thuộc wizard T12.
**File thực tế:** `src/rules/validators.ts` (+ test), mở rộng `src/ooxml/headerBlock.ts` (`numberRow`), `src/ui/HeaderForm.ts`
**Kích cỡ:** M

### T9: Tên loại văn bản, trích yếu, "Kính gửi"
**Mô tả:** Chèn tên loại (in hoa, đậm) và trích yếu (đậm, đường kẻ dưới) cho các loại có tên (quyết định, tờ trình...); công văn dùng dòng "V/v ..." (cỡ 13, không đậm, căn giữa ô bên trái như mẫu) và dòng "Kính gửi: ..." (14 pt, căn giữa, giãn trước và sau 12 pt), cho nhập nhiều nơi nhận ("Kính gửi" nhiều dòng).
**Tiêu chí nghiệm thu:**
- [ ] Hai biến thể (có tên loại / công văn "V/v" + "Kính gửi") đúng cỡ, kiểu, căn giữa/căn trái đúng quy tắc và khớp fixture
- [ ] Đường kẻ dưới trích yếu đúng độ dài tỷ lệ
- [ ] Trích yếu nhiều dòng vẫn đúng
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML cho hai biến thể
- [ ] Thủ công: so với mẫu Phụ lục III
**Phụ thuộc:** T8
**File dự kiến:** `src/ooxml/titleBlock.ts`, `src/office/insertHeader.ts`, `src/ooxml/titleBlock.test.ts`
**Kích cỡ:** S

### Checkpoint 2 – Đầu văn bản
- [ ] Khối đầu công văn khớp `fixtures/congvan-mau.json` (không khác biệt ngoài các khác biệt đã chấp nhận ở Q9) và khớp file mẫu của bạn khi in PDF
- [ ] Khối đầu quyết định khớp mẫu Phụ lục III
- [ ] **Bạn duyệt trực quan** (và người am hiểu văn thư nếu có)
- [ ] `npm test`, `npm run build` xanh

---

## Giai đoạn 3 – Khối cuối văn bản và MVP

### T10: Chữ ký và nơi nhận
**Mô tả:** Chèn bảng 2 cột không viền (như mẫu): trái là "Nơi nhận:" (nghiêng, đậm, cỡ 12) và danh sách (cỡ 11, gạch đầu dòng, kết thúc "Lưu: VT, …"); phải là quyền hạn (TM./KT./TL./TUQ.), chức vụ (in hoa, đậm, 14 pt), số dòng trống chỗ ký (mặc định 6, đổi ở T6A), họ tên (đậm, không học hàm/học vị). Hỗ trợ chức vụ hai dòng như mẫu ("KT. CHI CỤC TRƯỞNG" / "PHÓ CHI CỤC TRƯỞNG").
**Tiêu chí nghiệm thu:**
- [ ] Các biến thể quyền hạn hoạt động; chức vụ tự in hoa
- [ ] Danh sách nơi nhận: gạch đầu dòng, `;` cuối mục, dấu chấm cuối mục cuối
- [ ] Nhập họ tên có tiền tố "TS.", "ThS." thì cảnh báo
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML; hàm phát hiện học hàm/học vị
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T5
**File dự kiến:** `src/ooxml/footerBlock.ts`, `src/office/insertFooter.ts`, `src/ui/FooterForm.ts`, `src/ooxml/footerBlock.test.ts`
**Kích cỡ:** M

### T11: Đánh số trang
**Mô tả:** Đánh số trang ở lề trên (header mặc định, trường PAGE), canh giữa, chữ số Ả Rập, đứng, cỡ 14 như mẫu; bật "trang đầu khác biệt" để trang 1 không hiển thị số. Tuỳ chọn ở nhóm Số trang (T6A). Nhận `StandardProfile`: chuẩn Đảng đặt khoảng cách header 10 mm (số trang cách mép trên 10 mm) và đánh số riêng cho từng phụ lục (T3 xác nhận khả thi).
**Tiêu chí nghiệm thu:**
- [ ] Từ trang 2 có số trang đúng cỡ và vị trí; trang 1 không có
- [ ] Chạy lại không tạo trùng trường số trang
- [ ] Hoạt động với tài liệu một section và nhiều section
- [ ] Đăng ký thành bước "Đánh số trang" của "Chỉnh toàn bộ" (T6C), idempotent
**Kiểm chứng:**
- [ ] `npm test`: bộ dựng OOXML header
- [ ] Thủ công: tài liệu 3 trang, in preview
**Phụ thuộc:** T3, T4
**File dự kiến:** `src/office/pageNumber.ts`, `src/ooxml/pageNumberHeader.ts`, `src/ooxml/pageNumberHeader.test.ts`
**Kích cỡ:** S

### T12: Wizard "Tạo công văn mới" (chuẩn Chính quyền)
**Mô tả:** Luồng một màn hình/nhiều bước: chọn cơ quan chủ quản (tuỳ chọn) và cơ quan ban hành, nhập số, ngày, V/v, "Kính gửi", người ký, nơi nhận → tạo tài liệu mới có đủ khối đầu, thân (con trỏ ở nội dung), khối cuối, số trang, trang A4 và style theo tuỳ chọn T6A hiện hành. Là lát cắt dọc hoàn chỉnh đầu tiên.
**Tiêu chí nghiệm thu:**
- [ ] Từ tài liệu trống, một lần chạy cho ra công văn đủ thể thức, và khi nhập đúng nội dung của file mẫu thì `compareToFixture` không báo khác biệt về định dạng
- [ ] Trường bắt buộc/tuỳ chọn rõ; lỗi nhập hiện tại chỗ, bằng tiếng Việt
- [ ] Chạy trong tài liệu đã có nội dung thì hỏi xác nhận trước khi chèn
**Kiểm chứng:**
- [ ] `npm test`: golden toàn tài liệu; `npm run build` thành công
- [ ] Thủ công: tạo 3 công văn mẫu (tên dài, nhiều nơi nhận, ký thay) rồi in PDF
**Phụ thuộc:** T6A, T7, T8, T9, T10, T11
**File dự kiến:** `src/ui/NewDocWizard.ts`, `src/office/createDocument.ts`, `src/ooxml/documentAssembler.ts`, `src/ooxml/documentAssembler.test.ts`
**Kích cỡ:** M

### Checkpoint 3 – MVP chuẩn Chính quyền
- [ ] Công văn tạo từ add-in khớp file mẫu của bạn (fixture + in PDF so sánh); bạn xác nhận
- [ ] Test, build, lint xanh; không lỗi console khi chạy trong Word
- [ ] Khung `StandardProfile` đã chứng minh dùng được qua NĐ 30; bắt đầu cài chuẩn Đảng trên cùng khung

---

## Giai đoạn 3B – Văn bản Đảng theo Hướng dẫn 05-HD/VPTW

Dùng lại bảng quy tắc ở mục 6B của `plan.md`, bộ style `HD05_*` (T5), tuỳ chọn nhanh chuẩn Đảng (T6A) và danh sách cơ quan đảng (kho lưu T6A, quản lý ở T22). Đáp án chuẩn là Phụ lục 1 (sơ đồ bố trí) và Phụ lục 2 (bảng cỡ chữ) của HD 05; không có file mẫu thật (đã chốt Q12), nên bạn duyệt bản in thử ở Checkpoint 3B.

### TD1: Tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM", cơ quan cấp trên, cơ quan ban hành, dấu `*`
**Mô tả:** Chèn bảng 2 cột không viền vừa vùng chữ: **phải** là tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM" (in hoa, 15 pt, đậm, đường kẻ ngang nét liền dưới dài bằng tiêu đề); **trái** là cơ quan cấp trên trực tiếp (in hoa 14 pt, không đậm, tuỳ chọn 0–2 dòng), cơ quan ban hành (in hoa 14 pt, đậm, nhiều dòng được; liên cơ quan nối bằng dấu gạch nối), và dòng dấu `*` căn giữa dưới tên cơ quan. Ô chọn cơ quan lấy từ danh sách đã lưu như T7.
**Tiêu chí nghiệm thu:**
- [ ] Tiêu đề, cơ quan cấp trên, cơ quan ban hành đúng cỡ, kiểu, in hoa theo mục 6B; tên cơ quan tự chuyển in hoa
- [ ] Cơ quan cấp trên bật/tắt, 0 đến 2 dòng đều sinh khối đúng; tên cơ quan dài xuống dòng không vỡ bảng; dấu `*` luôn nằm ngay dưới tên cơ quan ban hành, căn giữa
- [ ] Đường kẻ dưới tiêu đề bằng độ dài dòng chữ, đi theo văn bản khi đổi nội dung (viền đoạn, không dùng hình nổi)
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML cho các ca 0/1/2 dòng cơ quan cấp trên, cơ quan nhiều dòng, liên cơ quan
- [ ] Thủ công: in PDF, đối chiếu ô 1–2 của Phụ lục 1 và ví dụ ở Phụ lục 2
**Phụ thuộc:** T2B, T5, T6A, T6B
**File dự kiến:** `src/standards/hd05/headerBlock.ts`, `src/office/insertHeader.ts`, `src/ui/HeaderForm.ts`, `src/standards/hd05/headerBlock.test.ts`
**Kích cỡ:** M

### TD2: Số và ký hiệu, địa danh, ngày tháng (chuẩn Đảng)
**Mô tả:** Hàng dưới của bảng đầu văn bản: "Số 05-HD/VPTW" (14 pt, đứng, căn giữa dưới tên cơ quan, **không có dấu hai chấm**) và "Hà Nội, ngày 03 tháng 02 năm 2026" (14 pt, nghiêng, đối diện bên phải). Có bộ kiểm tra số, ký hiệu loại, ký hiệu cơ quan (bảng tra từ T2B), liên cơ quan và ngày tháng; cho phép để trống số và ngày chờ vào sổ.
**Tiêu chí nghiệm thu:**
- [ ] Số nhỏ hơn 10 tự thêm số 0; số và ký hiệu nối bằng "-", tên loại và cơ quan viết tắt nối bằng "/", liên cơ quan nối bằng "-"; ví dụ hợp lệ "Số 127-QĐ/TW", "Số 05-QC/BTCTU-BNCTU", "Số 01-BC/ĐH"
- [ ] Ngày nhỏ hơn 10 và tháng nhỏ hơn 3 thêm số 0; ngày không hợp lệ bị từ chối; cấp hành chính trước địa danh ghi được (ví dụ "TP. Huế", "Phường Ngô Quyền") khi cần
- [ ] Gợi ý ký hiệu loại từ thể loại đã chọn (bảng tra T2B) và ký hiệu cơ quan từ danh sách đã lưu
**Kiểm chứng:**
- [ ] `npm test`: bảng ca biên cho số, ký hiệu loại, ký hiệu cơ quan, liên cơ quan, ngày tháng; golden OOXML
- [ ] Thủ công: in PDF, so ô 3–4 của Phụ lục 1 và Phụ lục 2
**Phụ thuộc:** TD1
**File dự kiến:** `src/standards/hd05/soKyHieuBlock.ts`, `src/standards/hd05/validators.ts`, `src/ui/HeaderForm.ts`, `src/standards/hd05/validators.test.ts`
**Kích cỡ:** M

### TD3: Thể loại, trích yếu, "về việc" của công văn, "Kính gửi"/"Kính trình", căn cứ ban hành
**Mô tả:** Giữa trang đầu: thể loại (in hoa 15–16 pt, đậm), trích yếu (in thường 14–15 pt, đậm) và **5 dấu gạch nối** căn giữa bên dưới. Công văn không ghi tên loại; dòng "về việc ..." (in thường 12 pt, nghiêng) căn đối diện số và ký hiệu. Dòng "Kính gửi:" (14 pt, nghiêng) và danh sách nơi nhận xử lý (mỗi nơi một dòng, gạch đầu dòng, cuối dòng `;`, dòng cuối `.`) cho công văn; "Kính trình:" cho tờ trình. Phần căn cứ ban hành: mỗi căn cứ một dòng bắt đầu bằng gạch ngang, kết thúc `;`, căn cứ cuối kết thúc `.`.
**Tiêu chí nghiệm thu:**
- [ ] Ba biến thể (có thể loại + trích yếu; công văn "về việc" + "Kính gửi"; tờ trình "Kính trình") đúng cỡ, kiểu, vị trí
- [ ] Trích yếu nhiều dòng vẫn căn giữa đúng; 5 dấu gạch nối luôn là dòng riêng dưới trích yếu
- [ ] Căn cứ ban hành nhập nhiều dòng, dấu câu tự chuẩn hoá
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML ba biến thể và danh sách căn cứ
- [ ] Thủ công: in PDF, so ô 5a, 5b, 9b của Phụ lục 1 và ví dụ Phụ lục 2
**Phụ thuộc:** TD2
**File dự kiến:** `src/standards/hd05/titleBlock.ts`, `src/standards/hd05/recipientsTop.ts`, `src/office/insertHeader.ts`, `src/standards/hd05/titleBlock.test.ts`
**Kích cỡ:** M

### TD4: Quyền hạn, chức vụ, họ tên, "Nơi nhận" gạch chân, ký hiệu người soạn
**Mô tả:** Bảng 2 cột không viền: trái là "Nơi nhận:" (in thường 14 pt, đứng, gạch chân, có đường kẻ ngang dưới cụm từ) và danh sách (12 pt, gạch đầu dòng, cuối mục `;`, mục "Lưu" kết thúc `.`; công văn có "Như trên"; tờ trình theo mục "Kính trình"); phải là quyền hạn (T/M, K/T, T/L, Q.; in hoa 14 pt, đậm), chức vụ (in hoa 14 pt, **không đậm**), chỗ ký (mặc định 6 dòng trống), họ tên (in thường 14 pt, đậm, không học hàm, học vị). Dòng "Xác nhận chữ ký của đồng chí..." tuỳ chọn. Ký hiệu người soạn thảo và số lượng bản phát hành (11 pt) ở cuối "Nơi nhận" khi bật tuỳ chọn.
**Tiêu chí nghiệm thu:**
- [ ] Bốn quyền hạn (T/M, K/T, T/L, Q.) và các tổ hợp chức vụ hai dòng ("K/T TRƯỞNG BAN" / "PHÓ TRƯỞNG BAN") sinh đúng; chức vụ in hoa, không đậm; họ tên đậm
- [ ] "Nơi nhận" gạch chân, danh sách đúng dấu câu; nhập họ tên có "TS.", "GS." thì cảnh báo
- [ ] Dòng xác nhận chữ ký bật/tắt được và đặt đúng chỗ
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML; hàm phát hiện học hàm, học vị dùng chung với T10
- [ ] Thủ công: in PDF, so ô 7a–7c, 9a, 12 của Phụ lục 1 và Phụ lục 2
**Phụ thuộc:** T2B, T5, T6A
**File dự kiến:** `src/standards/hd05/footerBlock.ts`, `src/standards/hd05/signatureForms.ts`, `src/ui/FooterForm.ts`, `src/standards/hd05/footerBlock.test.ts`
**Kích cỡ:** M

### TD5: Wizard "Tạo văn bản Đảng"
**Mô tả:** Dùng lại vỏ wizard của T12 với tham số chuẩn Đảng: chọn thể loại (công văn, báo cáo, thông báo, tờ trình, quyết định ở bản đầu; các loại còn lại ở TD7), nhập cơ quan, số, ngày, trích yếu, "Kính gửi", căn cứ, người ký, nơi nhận → tạo tài liệu có đủ khối đầu, thân (con trỏ ở nội dung), khối cuối, số trang (T11 cho chuẩn Đảng: giữa trang, cách mép trên 10 mm, trang 1 không hiện), trang A4 lề 20/20/30/15 và style `HD05_*`.
**Tiêu chí nghiệm thu:**
- [ ] Từ tài liệu trống, một lần chạy cho ra văn bản Đảng đủ thể thức; ghi chuẩn Đảng vào thuộc tính tài liệu
- [ ] Số trang đúng vị trí và không hiện ở trang 1; tuỳ chọn in hai mặt tạo lề đối xứng
- [ ] Chạy trong tài liệu đã có nội dung thì hỏi xác nhận trước khi chèn
**Kiểm chứng:**
- [ ] `npm test`: golden toàn tài liệu cho công văn, báo cáo, thông báo; `npm run build` thành công
- [ ] Thủ công: tạo ba văn bản mẫu (công văn nhiều nơi nhận, báo cáo kèm căn cứ, thông báo ký thay) rồi in PDF
**Phụ thuộc:** T11, T12, TD1, TD2, TD3, TD4
**File dự kiến:** `src/ui/NewDocWizard.ts`, `src/standards/hd05/assembleDocument.ts`, `src/office/createDocument.ts`, `src/standards/hd05/assembleDocument.test.ts`
**Kích cỡ:** M

### Checkpoint 3B – MVP chuẩn Đảng
- [ ] Văn bản Đảng tạo từ add-in khớp bảng cỡ chữ Phụ lục 2 và sơ đồ Phụ lục 1 của HD 05 khi in PDF
- [ ] **Bạn duyệt bản in thử** (không có mẫu thật nên đối chiếu Phụ lục 1–2 của HD 05); mặc định giãn dòng, cỡ thể loại, trích yếu đã chốt ở Q13 và chỉnh được bằng ô "Tuỳ chỉnh…"
- [ ] Test, build, lint xanh; chuyển đổi qua lại hai chuẩn không rò rỉ style hay tuỳ chọn

---

## Giai đoạn 4 – Bộ kiểm tra lỗi thể thức

### T13: Engine kiểm tra và quy tắc cấp trang/phông
**Mô tả:** Engine nhận `StandardProfile` (chuẩn tự nhận diện ở T6B), đọc cấu trúc tài liệu (trang, đoạn, phông) và đối chiếu bảng quy tắc của chuẩn đó, trả về danh sách lỗi (rule id, vị trí, mức độ, giá trị hiện tại/mong đợi). Task pane hiện báo cáo, bấm vào lỗi thì nhảy tới vị trí.
**Tiêu chí nghiệm thu:**
- [ ] Phát hiện đúng sai khổ giấy, lề, phông, cỡ chữ, căn lề, thụt dòng, giãn đoạn
- [ ] Mỗi lỗi nêu quy tắc và điều khoản tham chiếu bằng tiếng Việt
- [ ] Văn bản chuẩn từ T12 cho 0 lỗi
**Kiểm chứng:**
- [ ] `npm test`: engine chạy trên mô hình tài liệu giả; bộ ca có lỗi và không lỗi
- [ ] Thủ công: tài liệu cố ý sai từng quy tắc
**Phụ thuộc:** T2, T2B, T5, T6B
**File dự kiến:** `src/checker/engine.ts`, `src/checker/pageRules.ts`, `src/office/readDocument.ts`, `src/ui/ReportPanel.ts`, `src/checker/engine.test.ts`
**Kích cỡ:** M

### T14: Kiểm tra khối đầu/cuối và mẫu văn bản (chuẩn Chính quyền)
**Mô tả:** Thêm quy tắc nhận diện và kiểm tra: quốc hiệu/tiêu ngữ đúng chữ, "Số: ./." đúng mẫu, ngày tháng đúng quy cách, nơi nhận đúng cấu trúc, chức vụ in hoa, họ tên không học hàm. Mỗi lỗi có nút "Sửa".
**Tiêu chí nghiệm thu:**
- [ ] Phát hiện đúng các lỗi phổ biến (thiếu số 0, tiêu ngữ viết sai, thiếu gạch nối, quyền hạn sai)
- [ ] Nút "Sửa" sửa đúng một lỗi mà không đụng chỗ khác
- [ ] Tỷ lệ báo sai (false positive) trên bộ văn bản chuẩn = 0
- [ ] Đăng ký bước "Chuẩn hoá khối đầu/cuối (Chính quyền)" vào "Chỉnh toàn bộ" (T6C): nhận diện bảng đầu văn bản, khối ký và nơi nhận có sẵn rồi **chỉ sửa định dạng**, không đổi chữ; phần không nhận diện chắc chắn thì bỏ qua và báo
**Kiểm chứng:**
- [ ] `npm test`: bộ ca mẫu có/không lỗi (tối thiểu 30 ca)
- [ ] Thủ công: 5 văn bản thật (đã ẩn thông tin) từ người dùng văn thư
**Phụ thuộc:** T7–T10, T13
**File dự kiến:** `src/checker/headerRules.ts`, `src/checker/footerRules.ts`, `src/checker/fixes.ts`, `src/checker/headerRules.test.ts`
**Kích cỡ:** M

### T14B: Kiểm tra lỗi thể thức theo chuẩn Đảng (HD 05)
**Mô tả:** Thêm bộ quy tắc kiểm tra cho chuẩn Đảng vào engine T13: lề 20/20/30/15 mm (và đối xứng khi in hai mặt), giãn dòng Exactly 18–22 pt, thụt dòng đầu khoảng 10 mm, cỡ chữ từng thành phần theo bảng 6B, tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM" đúng chữ, dấu `*`, "Số 05-HD/VPTW" (không dấu hai chấm, số 0 phía trước, "-" và "/" đúng chỗ), ngày tháng, 5 dấu gạch nối dưới trích yếu, chức vụ không đậm, "Nơi nhận" gạch chân, số trang cách mép trên 10 mm và ẩn ở trang 1, nội dung kết thúc bằng dấu chấm, bố cục Phần/Chương/Mục/Điều/Khoản/Điểm. Mỗi lỗi có nút "Sửa" khi sửa được tự động.
**Tiêu chí nghiệm thu:**
- [ ] Phát hiện đúng các lỗi phổ biến: dùng tiêu ngữ của chính quyền, ghi "Số:" có dấu hai chấm, thiếu 5 dấu gạch nối, giãn dòng sai kiểu (Single thay vì Exactly), chức vụ in đậm, thiếu số 0
- [ ] Cảnh báo khi tài liệu trộn dấu hiệu của hai chuẩn (có cả quốc hiệu Chính quyền và tiêu đề Đảng)
- [ ] Văn bản chuẩn Đảng tạo từ TD5 cho 0 lỗi; văn bản chuẩn Chính quyền không bị báo theo luật Đảng
- [ ] Đăng ký bước "Chuẩn hoá khối đầu/cuối (Đảng)" vào "Chỉnh toàn bộ" (T6C): chỉ sửa định dạng, không đổi chữ; không tự đổi tiêu đề từ chuẩn này sang chuẩn kia (đổi chuẩn là thao tác riêng có xác nhận ở T6B)
**Kiểm chứng:**
- [ ] `npm test`: tối thiểu 30 ca có/không lỗi cho chuẩn Đảng
- [ ] Thủ công: tài liệu cố ý sai từng quy tắc; thêm một văn bản Đảng thật bạn đang có (đã ẩn thông tin) nếu tiện, không bắt buộc
**Phụ thuộc:** T13, TD1–TD4
**File dự kiến:** `src/standards/hd05/checks.ts`, `src/standards/hd05/fixes.ts`, `src/checker/engine.ts`, `src/standards/hd05/checks.test.ts`
**Kích cỡ:** M

### T15: "Sửa tất cả" có xem trước và hoàn tác
**Mô tả:** Sửa hàng loạt các lỗi có thể sửa tự động; xem trước danh sách thay đổi; hoàn tác một lần; xuất báo cáo lỗi (txt/docx).
**Tiêu chí nghiệm thu:**
- [ ] Xem trước liệt kê chính xác các thay đổi sẽ làm
- [ ] Sau khi sửa, kiểm tra lại cho 0 lỗi có thể sửa
- [ ] Đăng ký bước cuối "Sửa các lỗi còn lại" vào "Chỉnh toàn bộ" (T6C); chạy "Chỉnh toàn bộ" trên văn bản có lỗi cài sẵn của cả hai chuẩn rồi kiểm tra lại cho 0 lỗi có thể sửa
- [ ] Văn bản 100 trang xử lý trong ngưỡng chấp nhận (đặt ngưỡng cụ thể sau khi đo, ví dụ < 15 giây)
**Kiểm chứng:**
- [ ] `npm test`: logic gom/sắp xếp bản sửa
- [ ] Thủ công: tài liệu 100 trang, đo thời gian; Ctrl+Z hoàn tác
**Phụ thuộc:** T14, T14B
**File dự kiến:** `src/checker/fixAll.ts`, `src/ui/FixPreview.ts`, `src/checker/exportReport.ts`, `src/checker/fixAll.test.ts`
**Kích cỡ:** M

### Checkpoint 4 – Bộ kiểm tra
- [ ] Nút "Chỉnh toàn bộ" chạy đủ các bước (trang, style, thân, bảng, khối đầu/cuối, số trang, sửa lỗi còn lại) trên cả hai chuẩn và kiểm tra lại cho 0 lỗi sửa được
- [ ] Bộ văn bản mẫu có lỗi cài sẵn cho **cả hai chuẩn**: báo đủ lỗi, không báo sai trên văn bản chuẩn; tài liệu trộn dấu hiệu hai chuẩn được cảnh báo
- [ ] Người dùng văn thư thử trên văn bản thật và xác nhận mức độ hữu ích
- [ ] Test, build, lint xanh

---

## Giai đoạn 5 – Phủ mọi loại văn bản hành chính (theo họ mẫu)

Danh sách loại và họ mẫu lấy từ `docTypes` của từng chuẩn (NĐ 30 ở T2, HD 05 ở T2B). Mỗi loại là một mục cấu hình trong registry (T17); T18–T20 (Chính quyền) và TD6–TD8 (Đảng) chỉ thêm họ mẫu mới và dữ liệu, không sửa wizard.

### T16: Quyết định và nghị quyết (cá biệt)
**Mô tả:** Mẫu có phần "Căn cứ" (nghiêng, kết thúc ";"), dòng "QUYẾT ĐỊNH:" (hoặc tương ứng), đánh số Điều/Khoản/Điểm bằng multilevel list gắn với style `ND30_*`.
**Tiêu chí nghiệm thu:**
- [ ] Điều (đậm), Khoản (số), Điểm (chữ cái) đánh số tự động đúng cấp
- [ ] Hai loại sinh đúng thể thức: nghị quyết cá biệt (Mẫu 1.1) và quyết định cá biệt quy định trực tiếp (Mẫu 1.2) hoặc gián tiếp (Mẫu 1.3). Chỉ thị dùng Mẫu 1.4, làm ở T17
- [ ] Bộ kiểm tra T14 nhận biết và kiểm tra các thành phần mới
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML numbering
- [ ] Thủ công: so mẫu Phụ lục III; thêm/xoá Điều thì số tự cập nhật
**Phụ thuộc:** T12, T14
**File dự kiến:** `src/ooxml/numbering.ts`, `src/templates/quyetDinh.ts`, `src/checker/headerRules.ts`, `src/templates/quyetDinh.test.ts`
**Kích cỡ:** M

### T17: Registry loại văn bản và nhóm "tên loại + trích yếu"
**Mô tả:** Dựng registry cấu hình (mỗi loại khai báo khối đầu/cuối, trường bắt buộc, ký hiệu viết tắt) và đưa vào wizard bước chọn loại. Làm nhóm dùng Mẫu 1.4 (có tên loại + trích yếu), gồm 17 loại: chỉ thị, quy chế, quy định, thông cáo, thông báo, hướng dẫn, chương trình, kế hoạch, phương án, đề án, dự án, báo cáo, tờ trình, giấy ủy quyền, phiếu gửi, phiếu chuyển, phiếu báo (đọc từ `layout: "co-ten-loai"` trong `src/rules/docTypes.ts`). Quy chế và quy định có thêm cấu trúc Chương/Điều ở T18.
**Tiêu chí nghiệm thu:**
- [ ] Thêm một loại mới chỉ cần thêm một mục cấu hình, không sửa wizard
- [ ] Các loại trong nhóm sinh đúng thể thức so với mẫu
- [ ] Ký hiệu viết tắt gợi ý đúng cho từng loại
**Kiểm chứng:**
- [ ] `npm test`: golden cho từng loại; test registry (mọi loại ở `docTypes.ts` đều có cấu hình hợp lệ hoặc được đánh dấu "chưa hỗ trợ")
- [ ] Thủ công: tạo từng loại, in PDF, đối chiếu mẫu
**Phụ thuộc:** T12
**File dự kiến:** `src/templates/registry.ts`, `src/templates/titledDocs.ts`, `src/ui/NewDocWizard.ts`, `src/templates/registry.test.ts`
**Kích cỡ:** M

### T18: Quy chế, quy định (Chương/Mục/Điều)
**Mô tả:** Họ mẫu có cấu trúc Chương, Mục, Điều, Khoản, Điểm với cỡ chữ và kiểu chữ từng cấp theo quy tắc; tên văn bản dạng "QUY CHẾ" kèm trích yếu và phần ban hành kèm theo.
**Tiêu chí nghiệm thu:**
- [ ] Chương/Mục/Điều đánh số tự động, đúng cấp và đúng kiểu chữ
- [ ] Hai loại (quy chế, quy định) sinh đúng thể thức
- [ ] Bộ kiểm tra nhận biết cấu trúc Chương/Mục
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML numbering nhiều cấp
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T16, T17
**File dự kiến:** `src/templates/quyChe.ts`, `src/ooxml/numbering.ts`, `src/checker/headerRules.ts`, `src/templates/quyChe.test.ts`
**Kích cỡ:** M

### T19: Biên bản và văn bản nhiều bên ký
**Mô tả:** Biên bản (thời gian, địa điểm, thành phần, nội dung, khối ký chủ trì/thư ký), hợp đồng, bản ghi nhớ, bản thỏa thuận (khối ký hai hoặc nhiều bên, có thể không có "Nơi nhận"). Ba loại sau **không có mẫu ở Phụ lục III** nên cần bạn duyệt bố cục đề xuất (xem `docs/nd30-ambiguities.md`, mục 4). Bản cam kết không có trong NĐ 30.
**Tiêu chí nghiệm thu:**
- [ ] Khối ký nhiều bên (2–4 bên) dàn đều đúng bố cục, không vỡ khi tên dài
- [ ] Các loại sinh đúng thể thức so với mẫu
- [ ] Bộ kiểm tra không báo sai "thiếu nơi nhận" cho loại không yêu cầu
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML khối ký 2/3/4 bên
- [ ] Thủ công: in PDF, so mẫu
**Phụ thuộc:** T17
**File dự kiến:** `src/ooxml/multiSignBlock.ts`, `src/templates/bienBan.ts`, `src/templates/hopDong.ts`, `src/ooxml/multiSignBlock.test.ts`
**Kích cỡ:** M

### T20: Giấy tờ và mẫu nhỏ
**Mô tả:** Giấy mời (Mẫu 1.7), giấy giới thiệu (Mẫu 1.8), giấy nghỉ phép (Mẫu 1.10), công điện (Mẫu 1.6) và thư công (không có mẫu ở Phụ lục III). Giấy ủy quyền, phiếu gửi, phiếu chuyển, phiếu báo dùng Mẫu 1.4 nên làm ở T17. Mỗi mẫu có bố cục riêng, thường ngắn và có phần tuỳ biến (khung, bảng nhỏ). Nếu một mẫu quá phức tạp thì tách thành task riêng sau khi duyệt.
**Tiêu chí nghiệm thu:**
- [ ] Mỗi mẫu có trong registry và sinh đúng thể thức so với Phụ lục III
- [ ] Mẫu chưa hỗ trợ (nếu có) được đánh dấu rõ trong giao diện, không im lặng
- [ ] Bộ kiểm tra không báo sai cho các mẫu này
**Kiểm chứng:**
- [ ] `npm test`: golden từng mẫu
- [ ] Thủ công: tạo từng mẫu, in PDF
**Phụ thuộc:** T17
**File dự kiến:** `src/templates/smallForms.ts`, `src/ooxml/smallForms.ts`, `src/templates/registry.ts`, `src/templates/smallForms.test.ts`
**Kích cỡ:** M

### T21: Phụ lục và dấu chỉ mức độ khẩn/mật
**Mô tả:** Chèn trang phụ lục (ngắt trang, tiêu đề "PHỤ LỤC", tên phụ lục) và khung dấu mức độ khẩn (KHẨN/THƯỢNG KHẨN/HỎA TỐC) và độ mật theo quy tắc.
**Tiêu chí nghiệm thu:**
- [ ] Dấu khẩn trong khung chữ nhật viền đơn, in hoa, đậm, đúng vị trí
- [ ] Phụ lục bắt đầu trang mới, đánh số thứ tự tự động
- [ ] Bộ kiểm tra nhận ra hai thành phần này
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T12
**File dự kiến:** `src/ooxml/urgencyMark.ts`, `src/ooxml/appendix.ts`, `src/ui/InsertPanel.ts`, `src/ooxml/urgencyMark.test.ts`
**Kích cỡ:** S

### TD6: Đảng: Phần/Chương/Mục/Điều/Khoản/Điểm (nghị quyết, quyết định, quy định, quy chế)
**Mô tả:** Cấu trúc nội dung theo HD 05: "Phần"/"Chương" và số (La Mã hoặc chữ, ví dụ "Phần I", "Chương thứ hai") cùng dòng, giữa trang, in thường đậm 14–15 pt, tên phần/chương in hoa đậm ở dòng dưới; "Mục" + số Ả Rập đậm, tên mục in hoa đậm; "Điều 1. ..." in thường đậm, số Ả Rập có dấu chấm; Khoản "1." và Điểm "a)" in thường đứng. Đánh số tự động bằng multilevel list gắn style `HD05_*`. Mẫu cho nghị quyết, quyết định, quy định, quy chế (có phần "Căn cứ" từ TD3).
**Tiêu chí nghiệm thu:**
- [ ] Các cấp đánh số tự động đúng, đúng kiểu chữ theo mục 6B; thêm/xoá Điều thì số tự cập nhật
- [ ] Bốn thể loại sinh đúng thể thức; bộ kiểm tra T14B nhận biết các cấp
- [ ] Số La Mã và số bằng chữ cho Phần/Chương chọn được
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML numbering nhiều cấp
- [ ] Thủ công: in PDF, so bảng Phụ lục 2 mục 6
**Phụ thuộc:** TD5
**File dự kiến:** `src/standards/hd05/numbering.ts`, `src/standards/hd05/templates/regulations.ts`, `src/standards/hd05/checks.ts`, `src/standards/hd05/numbering.test.ts`
**Kích cỡ:** M

### TD7: Đảng: các thể loại còn lại theo Quy định 399-QĐ/TW
**Mô tả:** Dựa trên danh sách thể loại đọc ở T2B, bổ sung các họ mẫu còn lại vào registry T17: nghị quyết, chỉ thị, kết luận, thông tri, chương trình, kế hoạch, biên bản (hội nghị cấp uỷ: người ghi biên bản và chủ trì hội nghị ký, có dòng xác nhận chữ ký khi không đóng dấu), văn bản đại hội (tên đại hội, đoàn chủ tịch/đoàn thư ký/ban thẩm tra tư cách đại biểu/ban kiểm phiếu, ký hiệu "ĐH"), văn bản liên cơ quan (hai khối ký, tên cơ quan nối bằng gạch nối). Loại nào quá phức tạp thì tách thành task riêng sau khi duyệt.
**Tiêu chí nghiệm thu:**
- [ ] Mỗi thể loại trong danh sách T2B hoặc có mẫu đúng thể thức hoặc ghi rõ "chưa hỗ trợ" và lý do trong giao diện
- [ ] Biên bản, văn bản đại hội và văn bản liên cơ quan có khối ký đúng như ví dụ ở mục 7 của HD 05
- [ ] Ký hiệu loại gợi ý đúng (NQ, CT, KL, QC, BC, QĐ, CTr, TT, TTr, CV, ...)
**Kiểm chứng:**
- [ ] `npm test`: golden từng mẫu; test registry (mọi thể loại đều có cấu hình hoặc cờ "chưa hỗ trợ")
- [ ] Thủ công: tạo từng mẫu, in PDF, đối chiếu ví dụ trong HD 05
**Phụ thuộc:** T17, TD5
**File dự kiến:** `src/standards/hd05/templates/*.ts`, `src/standards/hd05/signatureForms.ts`, `src/templates/registry.ts`, `src/standards/hd05/templates/templates.test.ts`
**Kích cỡ:** M

### TD8: Đảng: phụ lục, dấu khẩn, phạm vi lưu hành, dự thảo, thông tin liên hệ
**Mô tả:** Thành phần bổ sung theo mục II của HD 05: trang phụ lục ("Phụ lục" + số, dòng chỉ dẫn số/ký hiệu văn bản, đánh số trang riêng cho từng phụ lục); dấu chỉ mức độ khẩn (KHẨN, THƯỢNG KHẨN, HỎA TỐC; in hoa 14 pt, đậm, dưới số và ký hiệu, công văn thì dưới trích yếu); chỉ dẫn phạm vi lưu hành (in hoa 12 pt, đậm, góc phải trang đầu, cách mép trên 10 mm; "Không phổ biến trên phương tiện thông tin đại chúng" ở chân trang cuối); chỉ dẫn dự thảo ("Dự thảo lần 2", 14 pt nghiêng, dưới số và ký hiệu); ký hiệu người soạn và số lượng bản (11 pt); thông tin liên hệ (10 pt, trang cuối, dưới đường kẻ ngang hết vùng trình bày).
**Tiêu chí nghiệm thu:**
- [ ] Từng thành phần đúng cỡ, kiểu, vị trí theo mục 6B và Phụ lục 1 (ô 10, 11a, 11b, 12, 13)
- [ ] Phụ lục bắt đầu trang mới, đánh số trang riêng; chạy lại không tạo trùng
- [ ] Bộ kiểm tra T14B nhận biết các thành phần này
**Kiểm chứng:**
- [ ] `npm test`: golden OOXML từng thành phần
- [ ] Thủ công: in PDF, so Phụ lục 1 và bảng Phụ lục 2
**Phụ thuộc:** TD5
**File dự kiến:** `src/standards/hd05/extras.ts`, `src/standards/hd05/appendix.ts`, `src/ui/InsertPanel.ts`, `src/standards/hd05/extras.test.ts`
**Kích cỡ:** M

### Checkpoint 5 – Phủ loại văn bản
- [ ] Đối chiếu `docTypes` của từng chuẩn với Điều 7 NĐ 30 và Quy định 399-QĐ/TW: mỗi loại hoặc có mẫu đúng thể thức hoặc ghi rõ lý do chưa hỗ trợ
- [ ] Mỗi họ mẫu có ít nhất một văn bản mẫu được duyệt trực quan
- [ ] Test, build, lint xanh

---

## Giai đoạn 6 – Hoàn thiện và phát hành (dùng cá nhân)

### T22: Danh sách cơ quan đã lưu (theo chuẩn), profile và xuất/nhập cấu hình
**Mô tả:** Quản lý danh sách cơ quan (thêm, sửa, xoá, đặt mặc định) cho ô chọn ở T7 và TD1; **mỗi mục thuộc một chuẩn** (chuẩn Đảng: cơ quan cấp trên trực tiếp 0–2 dòng và cơ quan ban hành, ký hiệu cơ quan đảng tra từ T2B). Cài đặt chung thêm "Chuẩn mặc định cho văn bản mới". Mỗi mục gồm các dòng cơ quan chủ quản (0–3), tên cơ quan ban hành, địa danh, ký hiệu mặc định (ví dụ "CCTL-ĐĐPCTT"), chức vụ và họ tên người ký thường dùng. Dùng chung kho lưu trữ cục bộ với preset ở T6A; có xuất/nhập file cấu hình (gồm cả preset) để chuyển sang máy khác.
**Tiêu chí nghiệm thu:**
- [ ] Danh sách cơ quan hiện trong ô chọn của T7 và wizard; chọn một mục thì điền sẵn cả cơ quan chủ quản lẫn cơ quan ban hành
- [ ] Mở wizard thì điền sẵn từ profile
- [ ] Giá trị ngoài khoảng NĐ 30 không lưu được kèm giải thích
- [ ] Xuất rồi nhập trên máy khác giữ nguyên profile và preset
**Kiểm chứng:**
- [ ] `npm test`: lưu/đọc/kiểm tra/nhập xuất cấu hình, phiên bản định dạng file
- [ ] Thủ công: đóng mở lại Word, giá trị còn
**Phụ thuộc:** T6A, T6B, T12, TD1
**File dự kiến:** `src/settings/profile.ts`, `src/settings/exportImport.ts`, `src/ui/SettingsPanel.ts`, `src/settings/profile.test.ts`
**Kích cỡ:** S

### T23: Kiểm thử trên Word 365, trợ năng, xử lý lỗi
**Mô tả:** Chạy kiểm thử thủ công theo ma trận tính năng × chuẩn (Chính quyền, Đảng) trên Word 365 (Windows; thử Mac/Online nếu tiện); thêm phát hiện tính năng (`isSetSupported`) và thông báo/fallback khi API không có; kiểm tra bàn phím, tương phản; rà soát chuỗi tiếng Việt.
**Tiêu chí nghiệm thu:**
- [ ] Mọi tính năng hoặc chạy đúng hoặc hiện thông báo rõ khi môi trường không hỗ trợ, không bao giờ lỗi im lặng
- [ ] Thao tác chính dùng được chỉ bằng bàn phím
- [ ] Biên bản kiểm thử lưu ở `docs/test-matrix.md`
**Kiểm chứng:**
- [ ] `npm test`, `npm run build`, `npm run lint` xanh
- [ ] Thủ công theo ma trận tính năng
**Phụ thuộc:** T15, T21, T22, TD8, và các task T16–T20, TD6–TD7 đã làm
**File dự kiến:** `src/office/capabilities.ts`, `src/ui/*`, `docs/test-matrix.md`
**Kích cỡ:** M

### T24: Đóng gói, GitHub Pages, hướng dẫn
**Mô tả:** Build bản phát hành, đưa lên **GitHub Pages** (host tĩnh miễn phí, dự phòng Cloudflare Pages) bằng quy trình tự động (GitHub Actions), nạp manifest qua shared-folder catalog trong Word, viết hướng dẫn cài đặt/sử dụng tiếng Việt, tuyên bố riêng tư (không gửi nội dung văn bản ra ngoài, không telemetry). Trang công khai nên kiểm tra repo và bản build không chứa file mẫu thật hay thông tin cá nhân.
**Tiêu chí nghiệm thu:**
- [ ] Cài được theo hướng dẫn trên Word 365 sạch, dùng được hằng ngày
- [ ] Xác nhận bằng giám sát mạng: không có request chứa nội dung văn bản
- [ ] Có số phiên bản, cách cập nhật và gỡ cài đặt
**Kiểm chứng:**
- [ ] `npm run build` bản production; kiểm tra manifest hợp lệ
- [ ] Cài thử trên hồ sơ Word sạch theo đúng hướng dẫn
**Phụ thuộc:** T23
**File dự kiến:** `manifest.xml`, `docs/huong-dan-cai-dat.md`, `docs/privacy.md`, `README.md`
**Kích cỡ:** S

### Checkpoint cuối
- [ ] Mọi tiêu chí nghiệm thu đạt; ma trận kiểm thử hoàn thành
- [ ] Bạn dùng thử trên văn bản thật và xác nhận
- [ ] Rà soát bảo mật và quyền riêng tư xong
- [ ] Sẵn sàng dùng hằng ngày
