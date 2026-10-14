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
**Trạng thái:** mã và hạ tầng xong, 2 commit (`3485b1b`, `06ef12d`). Còn 3 mục thủ công trong Word chờ xác nhận.
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
- [ ] Đường kẻ có độ dài đúng tỷ lệ cho phép, đi theo đoạn chữ (**đã đổi theo yêu cầu của bạn ngày 09/10/2026: là đường thẳng nổi Insert > Shapes > Line neo trong đoạn, như file mẫu, không còn là viền đoạn riêng**): độ dài tính bằng **ước lượng bề rộng chữ** (bảng độ rộng Times New Roman, lệch vài %, đã đối chiếu với file mẫu: tiêu ngữ 59 mm trong ±6%); **chờ bạn xem bản in thử**
- [x] Tên cơ quan dài tự xuống dòng không làm vỡ bảng (bảng bố cục cố định, tên không thụt; đường kẻ tính theo bề rộng ô); **xem bằng mắt chờ bạn**
**Kiểm chứng:**
- [x] `npm test`: 491 test; bộ dựng OOXML khớp golden file (`__golden__/header-block.xml`); `compareToFixture` không khác biệt; 0/1/2/3 dòng chủ quản
- [ ] Thủ công: in PDF, so với file mẫu của bạn và Phụ lục III
**Phụ thuộc:** T2A, T5, T6A
**Cập nhật 09/10/2026:** đường kẻ dưới cơ quan ban hành và dưới tiêu ngữ nay là hình đường thẳng neo (`ooxml/lineShape.ts`, cùng cấu trúc với file mẫu), độ dài vẫn theo tỉ lệ ước lượng bề rộng chữ; đường kẻ dưới trích yếu (T9) vẫn là viền đoạn. Chưa kiểm chứng trong Word việc `insertOoxml` giữ được hình nổi này.
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
- [x] Hai biến thể (có tên loại / công văn "V/v" + "Kính gửi") đúng cỡ, kiểu, căn giữa/căn trái và khớp fixture (V/v ở hàng ba của bảng đầu văn bản; Kính gửi 14 pt, giãn 12/12 pt, giãn dòng 1,17). Còn **chưa so được với mẫu Phụ lục III** vì chưa có bản mẫu trong repo
- [x] Đường kẻ dưới trích yếu đúng độ dài tỷ lệ (1/3 hoặc 1/2, tính bằng ước lượng bề rộng chữ như T7)
- [x] Trích yếu nhiều dòng vẫn đúng: là một đoạn tự xuống dòng; đường kẻ tính theo dòng chữ nhưng không vượt vùng chữ (chưa giới hạn bề rộng đoạn trích yếu, chờ bạn xem bằng mắt)
**Kiểm chứng:**
- [x] `npm test`: 531 test; golden OOXML `title-named.xml`, `title-letter.xml`, `title-letter-multi.xml`
- [ ] Thủ công: so với mẫu Phụ lục III
**Phụ thuộc:** T8
**Cập nhật 09/10/2026 (theo yêu cầu của bạn):** dòng "Kính gửi" của **công văn** không còn là đoạn dưới bảng mà là **hàng thứ tư của bảng đầu văn bản**: nhãn "Kính gửi:" ở ô trái (dưới ô V/v) canh phải, các nơi nhận ở ô phải canh trái; cả hai ô bắt đầu từ dòng thứ hai (dòng đầu để trống). Một nơi nhận ghi nguyên tên và **nằm ngang hàng với nhãn "Kính gửi:"**; từ hai nơi nhận trở lên thì danh sách gạch đầu dòng nằm **ngay dưới dòng của nhãn một dòng** (cập nhật tiếp ngày 09/10/2026). Biến thể có tên loại không có hàng này. Golden `title-letter*.xml` bỏ, thêm `header-block-recipients.xml`. Chưa so được với fixture mẫu vì mẫu cũ đặt Kính gửi ở đoạn giữa trang.
**Trạng thái:** mã, giao diện (mục "Thêm tên loại, trích yếu, Kính gửi" trong form khối đầu) và test xong; `typecheck`, `lint`, `build` sạch; đã chạy trong trình duyệt giả lập (công văn một và nhiều nơi nhận, có tên loại, thiếu dữ liệu, tắt mục); **chưa thử trong Word**. **Giả định cần bạn xác nhận**: nhiều nơi nhận thì in "Kính gửi:" căn giữa rồi mỗi nơi một dòng "- Tên;" (dòng cuối ".") căn trái, thụt vào ngang chữ "Kính gửi:" với thụt treo 4 mm, theo cách nhớ quy tắc NĐ 30 II.9.d, chưa đối chiếu với văn bản gốc. Tên loại 14 pt đậm in hoa, trích yếu 14 pt đậm (thêm hai cỡ này vào nhóm Đầu văn bản và tuỳ chọn `titleLineRatio`). Chọn loại văn bản từ danh sách và tự điền ký hiệu thuộc wizard T12; ở đây gõ tay tên loại.
**File thực tế:** `src/ooxml/titleBlock.ts` (+ test, 3 golden), mở rộng `headerBlock.ts` (`subjectRow`, xuất `lineParagraph`), `HeaderForm.ts`, `styles.ts` của chuẩn Chính quyền, `formatOptions.ts`
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
- [x] Các biến thể quyền hạn (TM., Q., KT., TL., TUQ., hoặc không) đứng trước dòng chức vụ đầu; chức vụ tự in hoa, một hoặc hai dòng
- [x] Danh sách nơi nhận: gạch đầu dòng, `;` cuối mục, dấu chấm cuối mục cuối (mục "Lưu: VT, …" luôn ở cuối); chuẩn hoá dấu người dùng đã gõ
- [x] Nhập họ tên có tiền tố "TS.", "ThS."... thì hiện cảnh báo và hỏi "Vẫn chèn / Sửa lại"
**Kiểm chứng:**
- [x] `npm test`: 567 test; golden `footer-block.xml`, khớp fixture (nơi nhận, chức vụ, họ tên), hàm phát hiện học hàm, học vị
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T5
**Trạng thái:** mã, giao diện (form "Khối cuối văn bản") và test xong; `typecheck`, `lint`, `build` sạch; đã chạy trong trình duyệt giả lập (thiếu dữ liệu, học hàm, chèn đủ); **chưa thử trong Word**. Bảng 2 cột không viền vừa vùng chữ, cột trái 51,4% như mẫu (4660/9060); "Nơi nhận:" 12 pt đậm nghiêng, danh sách 11 pt; quyền hạn và chức vụ 14 pt đậm in hoa căn giữa, số dòng trống lấy từ tuỳ chọn `signGap` (mặc định 6), họ tên đậm 14 pt; có đoạn đệm cao 1 pt sau bảng như mẫu. Chèn vào **cuối** thân văn bản bằng `insertOoxml(..., "End")` rồi đọc lại bảng cuối để kiểm tra; **chưa biết Word đặt bảng thế nào so với đoạn trống cuối tài liệu** (có thể còn một đoạn trống phía trên hoặc dưới). Không có nơi nhận thì ô trái để trống. "Nơi nhận" căn trái (mẫu đặt dàn đều hai bên, không khác khi một dòng).
**File thực tế:** `src/ooxml/footerBlock.ts` (+ test, golden), `src/office/insertFooter.ts`, `src/ui/FooterForm.ts`
**Kích cỡ:** M

### T11: Đánh số trang
**Mô tả:** Đánh số trang ở lề trên (header mặc định, trường PAGE), canh giữa, chữ số Ả Rập, đứng, cỡ 14 như mẫu; bật "trang đầu khác biệt" để trang 1 không hiển thị số. Tuỳ chọn ở nhóm Số trang (T6A). Nhận `StandardProfile`: chuẩn Đảng đặt khoảng cách header 10 mm (số trang cách mép trên 10 mm) và đánh số riêng cho từng phụ lục (T3 xác nhận khả thi).
**Tiêu chí nghiệm thu:**
- [ ] Từ trang 2 có số trang đúng cỡ (theo Tuỳ chọn nhanh) và vị trí; trang 1 không có: mã xong; vị trí theo khoảng cách header của bước thiết lập trang (12,5 mm mẫu, 10 mm chuẩn Đảng). **Chờ bạn xem bản in thử 3 trang**
- [x] Chạy lại không tạo trùng trường số trang: mỗi phần chỉ bị ghi khi header chưa có PAGE căn giữa đúng cỡ (có test kế hoạch idempotent)
- [ ] Hoạt động với tài liệu một section và nhiều section: kế hoạch theo từng phần (chỉ phần đầu ẩn số ở trang đầu, các phần sau luôn hiện số) có test; **chưa thử nhiều phần trong Word**, chưa biết header "liên kết với phần trước" xử lý ra sao khi ghi
- [x] Đăng ký thành bước 5 "Đánh số trang" của "Chỉnh toàn bộ" (idempotent, có xem trước); bản chụp để khôi phục nay gồm cả header mặc định, header trang đầu và cờ "trang đầu khác biệt"
**Kiểm chứng:**
- [x] `npm test`: 582 test; bộ dựng OOXML, đọc trạng thái số trang (kể cả trường PAGE phức hợp, loại NUMPAGES), kế hoạch từng phần, câu chữ xem trước
- [ ] Thủ công: tài liệu 3 trang, in preview
**Phụ thuộc:** T3, T4
**Trạng thái:** mã, nút "Đánh số trang" (cạnh nút Thiết lập trang) và bước trong "Chỉnh toàn bộ" xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Ghi số trang bằng `insertOoxml(..., "Replace")` vào header mặc định (cách dự phòng của spike p06; đoạn căn giữa, Times New Roman, trường PAGE), đặt "trang đầu khác biệt" qua `PageSetup`, đọc lại và báo chỗ chưa đúng. Chưa kiểm chứng: header sau khi ghi có thêm đoạn trống không (làm lệch vị trí số trang), tài liệu có header liên kết giữa các phần. **Chưa làm**: đánh số phụ lục riêng của chuẩn Đảng (T21, TD8). Header cũ có nội dung khác (tên cơ quan...) sẽ bị thay bằng số trang.
**File thực tế:** `src/ooxml/pageNumberHeader.ts` (+ test), `src/office/pageNumber.ts`, mở rộng `src/pipeline/{steps,planText}.ts`, `src/office/snapshot.ts`, `src/ui/PageSetupPanel.ts`
**Kích cỡ:** S

### T12: Wizard "Tạo công văn mới" (chuẩn Chính quyền)
**Mô tả:** Luồng một màn hình/nhiều bước: chọn cơ quan chủ quản (tuỳ chọn) và cơ quan ban hành, nhập số, ngày, V/v, "Kính gửi", người ký, nơi nhận → tạo tài liệu mới có đủ khối đầu, thân (con trỏ ở nội dung), khối cuối, số trang, trang A4 và style theo tuỳ chọn T6A hiện hành. Là lát cắt dọc hoàn chỉnh đầu tiên.
**Tiêu chí nghiệm thu:**
- [ ] Từ tài liệu trống, một lần chạy cho ra công văn đủ thể thức: mã xong, **chờ bạn thử trong Word**. Khi nhập đúng nội dung của file mẫu, thân tài liệu dựng ra khớp fixture (khối đầu, Kính gửi, khối cuối) qua `compareToFixture`: đã có test
- [x] Lỗi nhập hiện bằng tiếng Việt, gom lỗi Khối đầu và Khối cuối trong một thông báo (đã thử trong trình duyệt giả lập). Lỗi hiện ở thông báo chung chứ chưa tô đỏ từng ô
- [x] Chạy trong tài liệu đã có nội dung (hoặc họ tên người ký có học hàm, học vị) thì hỏi xác nhận "Vẫn tạo / Huỷ" trước khi làm; có nút khôi phục bản trước khi tạo
**Kiểm chứng:**
- [x] `npm test`: 593 test; golden toàn tài liệu `document-letter.xml`; `npm run build` thành công
- [ ] Thủ công: tạo 3 công văn mẫu (tên dài, nhiều nơi nhận, ký thay) rồi in PDF
**Phụ thuộc:** T6A, T7, T8, T9, T10, T11
**Trạng thái:** mã, giao diện (nút xanh "Tạo công văn" trên đầu các form khối đầu, khối cuối) và test xong; `typecheck`, `lint`, `build` sạch; đã thử trong trình duyệt giả lập phần kiểm tra lỗi và hộp xác nhận; **chưa thử toàn bộ trong Word**. Khác dự kiến: không mở tài liệu mới (cửa sổ mới không có task pane, các API thiết lập chưa kiểm chứng) mà làm **trong tài liệu đang mở**: (1) thiết lập trang, (2) cài style, (3) thay toàn bộ thân bằng công văn dựng từ các ô của Khối đầu và Khối cuối (có chỗ đánh dấu nội dung, thay bằng đoạn trống style `ND30_NoiDung` và đặt con trỏ vào đó), (4) đánh số trang. Không chạy bước "định dạng thân" của Chỉnh toàn bộ vì nó sẽ áp lại style lên dòng Kính gửi. Dùng lại các bước của T6C, chụp bản gốc trước khi làm. Chưa kiểm chứng: thay toàn bộ thân bằng `insertOoxml(..., "Replace")` có để lại đoạn trống thừa cuối tài liệu hay không. Muốn tạo tài liệu mới thì hiện phải mở tài liệu trống trước.
**File thực tế:** `src/ooxml/documentAssembler.ts` (+ test, golden), `src/office/createDocument.ts`, `src/ui/NewDocWizard.ts`; `HeaderForm` và `FooterForm` thêm hàm `collect()`
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
- [x] Tiêu đề 15 pt đậm, cơ quan cấp trên 14 pt không đậm, cơ quan ban hành 14 pt đậm, đều in hoa tự động (cỡ lấy từ Tuỳ chọn nhanh)
- [x] Cơ quan cấp trên bật/tắt, 0 đến 2 dòng đều sinh khối đúng; cơ quan ban hành 1 đến 3 dòng (nhập mỗi dòng một cơ quan); dấu `*` luôn nằm ngay dưới dòng cuối của cơ quan ban hành, căn giữa; tên dài xuống dòng trong ô
- [ ] Đường kẻ dưới tiêu đề bằng độ dài dòng chữ (viền đoạn riêng, không hình nổi); độ dài tính bằng ước lượng bề rộng chữ như T7; **chờ bạn xem bản in thử**
**Kiểm chứng:**
- [x] `npm test`: 609 test; golden `party-header-block.xml`, `party-header-block-multi.xml`; ca 0/1/2 dòng cấp trên, nhiều dòng cơ quan, liên cơ quan
- [ ] Thủ công: in PDF, đối chiếu ô 1–2 của Phụ lục 1 và ví dụ ở Phụ lục 2
**Phụ thuộc:** T2B, T5, T6A, T6B
**Cập nhật 09/10/2026:** đường kẻ dưới tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM" nay là hình đường thẳng neo trong đoạn tiêu đề (theo yêu cầu của bạn), không còn là viền đoạn.
**Trạng thái:** mã, giao diện (form Khối đầu đổi theo chuẩn: chuẩn Đảng có ô cơ quan cấp trên 0–2 dòng và ô cơ quan ban hành nhiều dòng) và test xong; `typecheck`, `lint`, `build` sạch; đã thử form trong trình duyệt giả lập; **chưa thử trong Word**. Dùng lại bảng 2 cột 45%/55% của T7 (tách các hàm dùng chung `borderlessTable`, `textParagraph`, `cell` ra khỏi `ooxml/headerBlock.ts`; golden T7 không đổi). Hàng số, ký hiệu, ngày tháng (TD2) và thể loại, trích yếu (TD3) chưa có. "Liên cơ quan" hiện hiểu là nhiều dòng cơ quan ban hành, người dùng tự gõ dấu gạch nối nếu muốn một dòng.
**File thực tế:** `src/standards/hd05/headerBlock.ts` (+ test, 2 golden), `src/ooxml/headerBlock.ts`, `src/ui/HeaderForm.ts` (kết quả `collect()` có thêm `kind`), `src/ui/NewDocWizard.ts` (từ chối chuẩn Đảng tới TD5)
**Kích cỡ:** M

### TD2: Số và ký hiệu, địa danh, ngày tháng (chuẩn Đảng)
**Mô tả:** Hàng dưới của bảng đầu văn bản: "Số 05-HD/VPTW" (14 pt, đứng, căn giữa dưới tên cơ quan, **không có dấu hai chấm**) và "Hà Nội, ngày 03 tháng 02 năm 2026" (14 pt, nghiêng, đối diện bên phải). Có bộ kiểm tra số, ký hiệu loại, ký hiệu cơ quan (bảng tra từ T2B), liên cơ quan và ngày tháng; cho phép để trống số và ngày chờ vào sổ.
**Tiêu chí nghiệm thu:**
- [x] Số nhỏ hơn 10 tự thêm số 0; số và ký hiệu nối bằng "-", tên loại và cơ quan viết tắt nối bằng "/", liên cơ quan nối bằng "-"; ví dụ hợp lệ (có test cho cả ba ví dụ) "Số 127-QĐ/TW", "Số 05-QC/BTCTU-BNCTU", "Số 01-BC/ĐH"
- [x] Ngày nhỏ hơn 10 và tháng nhỏ hơn 3 thêm số 0; ngày không hợp lệ bị từ chối; cấp hành chính trước địa danh ghi được (địa danh là ô tự do; có test "TP. Huế", "Phường Ngô Quyền")
- [ ] Gợi ý ký hiệu: bấm để điền chữ viết tắt loại (từ bảng loại văn bản của chuẩn), chữ viết tắt cơ quan đã lưu và bảng ký hiệu cơ quan của chuẩn. **Gợi ý theo thể loại đã chọn chưa làm** vì thể loại chỉ được chọn ở TD3/TD5
**Kiểm chứng:**
- [x] `npm test`: 616 test; ca biên cho số, ký hiệu, liên cơ quan, ngày tháng; golden `party-header-block-number.xml`
- [ ] Thủ công: in PDF, so ô 3–4 của Phụ lục 1 và Phụ lục 2
**Phụ thuộc:** TD1
**Trạng thái:** mã, giao diện và test xong; `typecheck`, `lint`, `build` sạch; đã thử trong trình duyệt giả lập (thiếu ký hiệu loại, điền gợi ý, lưu viết tắt cơ quan, chèn); **chưa thử trong Word**. Không cần bộ kiểm tra riêng của chuẩn Đảng: dùng lại `soText`/`placeDateText` của T8 vì `formatSoKyHieu` của chuẩn đã mã hoá cách ghi ("Số 05-HD/VPTW", bắt buộc có ký hiệu loại). Số trống thì chừa chỗ trống tại vị trí số. Thêm hàng thứ hai vào khối đầu chuẩn Đảng (số căn giữa 14 pt dưới tên cơ quan, ngày nghiêng 14 pt bên phải), trường `agencyCodes` của chuẩn và danh sách "viết tắt cơ quan" đã lưu (tách khỏi danh sách tên cơ quan). Trong ký hiệu liên cơ quan nhập các viết tắt nối bằng "-".
**File thực tế:** mở rộng `src/standards/hd05/headerBlock.ts` (`numberRow`), `src/ui/HeaderForm.ts`, `src/settings/agencies.ts`, `src/standards/{types.ts,hd05/profile.ts}`; thêm test vào `src/rules/validators.test.ts`
**Kích cỡ:** M

### TD3: Thể loại, trích yếu, "về việc" của công văn, "Kính gửi"/"Kính trình", căn cứ ban hành
**Mô tả:** Giữa trang đầu: thể loại (in hoa 15–16 pt, đậm), trích yếu (in thường 14–15 pt, đậm) và **5 dấu gạch nối** căn giữa bên dưới. Công văn không ghi tên loại; dòng "về việc ..." (in thường 12 pt, nghiêng) căn đối diện số và ký hiệu. Dòng "Kính gửi:" (14 pt, nghiêng) và danh sách nơi nhận xử lý (mỗi nơi một dòng, gạch đầu dòng, cuối dòng `;`, dòng cuối `.`) cho công văn; "Kính trình:" cho tờ trình. Phần căn cứ ban hành: mỗi căn cứ một dòng bắt đầu bằng gạch ngang, kết thúc `;`, căn cứ cuối kết thúc `.`.
**Tiêu chí nghiệm thu:**
- [ ] Ba biến thể (có thể loại + trích yếu; công văn "về việc" + "Kính gửi"; tờ trình "Kính trình") đúng cỡ, kiểu và thứ tự: mã và test xong; **vị trí dòng "về việc" cần bạn xem** (xem Trạng thái)
- [x] Trích yếu dài tự xuống dòng vẫn là một đoạn căn giữa; 5 dấu gạch nối luôn là dòng riêng dưới trích yếu
- [x] Căn cứ ban hành nhập nhiều dòng (mỗi dòng một căn cứ), dấu gạch đầu dòng, `;` và `.` tự chuẩn hoá
**Kiểm chứng:**
- [x] `npm test`: 637 test; golden `party-title-named.xml`, `party-title-letter.xml`, `party-title-trinh.xml` (có căn cứ)
- [ ] Thủ công: in PDF, so ô 5a, 5b, 9b của Phụ lục 1 và ví dụ Phụ lục 2
**Phụ thuộc:** TD2
**Cập nhật 09/10/2026 (theo yêu cầu của bạn):** với **công văn**, "Kính gửi:" (nghiêng, canh phải) và danh sách nơi nhận (canh trái) nằm ở **hàng thứ tư của bảng đầu văn bản**, ô dưới "về việc" và ô kế bên, cả hai bắt đầu từ dòng thứ hai. Một nơi nhận ghi nguyên tên (không gạch đầu dòng) ngang hàng với nhãn; từ hai nơi nhận trở lên thì danh sách gạch đầu dòng nằm dưới dòng của nhãn một dòng. **Tờ trình** vẫn để "Kính trình:" và danh sách ở đoạn dưới 5 dấu gạch nối vì trích yếu của tờ trình nằm ngoài bảng. Golden `party-title-letter.xml` bỏ.
**Trạng thái:** mã, giao diện (mục "Thêm thể loại, trích yếu, Kính gửi, căn cứ" của form Khối đầu chuẩn Đảng) và test xong; `typecheck`, `lint`, `build` sạch; đã thử trong trình duyệt giả lập (công văn, tờ trình, lỗi thiếu dữ liệu); **chưa thử trong Word**. **Giả định cần bạn xác nhận**: (1) dòng "về việc ..." của công văn đặt ở **hàng thứ ba của bảng đầu văn bản, ô trái dưới số và ký hiệu** (giống V/v của NĐ 30), vì hướng dẫn chỉ ghi "căn đối diện số và ký hiệu" chưa rõ cụ thể; (2) "Kính gửi:" nghiêng căn giữa rồi danh sách nơi nhận căn trái, mỗi nơi một dòng "- ...;" (một nơi thì "- ...."); (3) căn cứ ban hành đặt ngay sau phần Kính gửi hoặc trích yếu, căn đều hai bên, thụt treo 4 mm; (4) thể loại 16 pt, trích yếu 14 pt theo mặc định đề xuất Q13; giãn dòng các dòng danh sách lấy Exactly theo tuỳ chọn. Tờ trình tự điền thể loại "Tờ trình". Dùng lại `paragraph` của `ooxml/titleBlock.ts` (thêm nghiêng, căn đều, Exactly; golden NĐ 30 không đổi).
**File thực tế:** `src/standards/hd05/titleBlock.ts` (+ test, 3 golden), mở rộng `src/standards/hd05/headerBlock.ts` (`subjectRow`), `src/ooxml/titleBlock.ts`, `src/ui/HeaderForm.ts`
**Kích cỡ:** M

### TD4: Quyền hạn, chức vụ, họ tên, "Nơi nhận" gạch chân, ký hiệu người soạn
**Mô tả:** Bảng 2 cột không viền: trái là "Nơi nhận:" (in thường 14 pt, đứng, gạch chân, có đường kẻ ngang dưới cụm từ) và danh sách (12 pt, gạch đầu dòng, cuối mục `;`, mục "Lưu" kết thúc `.`; công văn có "Như trên"; tờ trình theo mục "Kính trình"); phải là quyền hạn (T/M, K/T, T/L, Q.; in hoa 14 pt, đậm), chức vụ (in hoa 14 pt, **không đậm**), chỗ ký (mặc định 6 dòng trống), họ tên (in thường 14 pt, đậm, không học hàm, học vị). Dòng "Xác nhận chữ ký của đồng chí..." tuỳ chọn. Ký hiệu người soạn thảo và số lượng bản phát hành (11 pt) ở cuối "Nơi nhận" khi bật tuỳ chọn.
**Tiêu chí nghiệm thu:**
- [x] Bốn quyền hạn (T/M, K/T, T/L, Q.) và tổ hợp chức vụ hai dòng ("K/T TRƯỞNG BAN" / "PHÓ TRƯỞNG BAN") sinh đúng; quyền hạn đậm, chức vụ in hoa không đậm, họ tên đậm (kiểm tra từng run)
- [x] "Nơi nhận" gạch chân, danh sách đúng dấu câu; họ tên có "TS.", "GS." thì cảnh báo "Vẫn chèn / Sửa lại" (dùng chung hàm của T10)
- [x] Dòng xác nhận chữ ký bật/tắt được, đặt ngay dưới họ tên, cùng ô chữ ký, cỡ 14 pt (bạn đã xác nhận 09/10/2026)
**Kiểm chứng:**
- [x] `npm test`: 655 test; golden `party-footer-block.xml`; hàm phát hiện học hàm, học vị dùng chung với T10
- [ ] Thủ công: in PDF, so ô 7a–7c, 9a, 12 của Phụ lục 1 và Phụ lục 2
**Phụ thuộc:** T2B, T5, T6A
**Đã chốt với bạn (09/10/2026):** (1) dòng xác nhận chữ ký ngay dưới họ tên, cùng ô, 14 pt; (2) "Nơi nhận" gạch chân bằng gạch chân của chữ; (3) hai cột 50/50; (4) tờ trình không có quy tắc riêng "nơi nhận theo mục Kính trình", người dùng tự nhập danh sách.
**Trạng thái:** mã, giao diện (form Khối cuối đổi theo chuẩn) và test xong; `typecheck`, `lint`, `build` sạch; đã thử trong trình duyệt giả lập; **chưa thử trong Word**. Bảng 2 cột 50%/50% không viền; "Nơi nhận:" 14 pt đứng gạch chân (gạch chân của chữ, chưa vẽ đường kẻ riêng), danh sách 12 pt; ký hiệu người soạn và số bản 11 pt là ô nhập tay ở cuối nơi nhận (không có quy tắc tự sinh); dòng xác nhận "Xác nhận chữ ký của đồng chí ..." 14 pt đặt dưới họ tên. Danh sách quyền hạn lấy từ `signAuthority` của chuẩn. Chưa có mục "Nơi nhận" riêng cho tờ trình ("theo mục Kính trình"): người dùng tự nhập. Dùng lại `recipientLines` và `findAcademicTitle` của T10 (file `signatureForms.ts` không cần).
**File thực tế:** `src/standards/hd05/footerBlock.ts` (+ test, golden), `src/ui/FooterForm.ts` (kết quả `collect()` có thêm `kind`), `src/ui/NewDocWizard.ts`
**Kích cỡ:** M

### TD5: Wizard "Tạo văn bản Đảng"
**Mô tả:** Dùng lại vỏ wizard của T12 với tham số chuẩn Đảng: chọn thể loại (công văn, báo cáo, thông báo, tờ trình, quyết định ở bản đầu; các loại còn lại ở TD7), nhập cơ quan, số, ngày, trích yếu, "Kính gửi", căn cứ, người ký, nơi nhận → tạo tài liệu có đủ khối đầu, thân (con trỏ ở nội dung), khối cuối, số trang (T11 cho chuẩn Đảng: giữa trang, cách mép trên 10 mm, trang 1 không hiện), trang A4 lề 20/20/30/15 và style `HD05_*`.
**Tiêu chí nghiệm thu:**
- [ ] Từ tài liệu trống, một lần chạy cho ra văn bản Đảng đủ thể thức: mã xong, **chờ bạn thử trong Word**; chuẩn Đảng được ghi vào thuộc tính tài liệu sau khi tạo xong
- [ ] Số trang (giữa, header cách mép trên 10 mm theo bước thiết lập trang, trang 1 không hiện) và lề đối xứng khi bật "In hai mặt" (bước thiết lập trang dùng `mirrorMargins`): dùng lại các bước đã có, **chờ bạn thử**
- [x] Chạy trong tài liệu đã có nội dung (hoặc họ tên người ký có học hàm, học vị) thì hỏi xác nhận trước khi tạo; có nút khôi phục
**Kiểm chứng:**
- [ ] `npm test`: 662 test; `npm run build` thành công; golden toàn tài liệu `document-party-letter.xml` (công văn); các ca có thể loại và căn cứ kiểm tra bằng cấu trúc, **chưa có golden riêng cho báo cáo và thông báo** (khác công văn chỉ ở các dòng thể loại, đã có golden `party-title-*.xml`)
- [ ] Thủ công: tạo ba văn bản mẫu (công văn nhiều nơi nhận, báo cáo kèm căn cứ, thông báo ký thay) rồi in PDF
**Phụ thuộc:** T11, T12, TD1, TD2, TD3, TD4
**Trạng thái:** mã, giao diện (nút "Tạo văn bản Đảng" cùng vỏ với "Tạo công văn", đổi nhãn theo chuẩn) và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Dùng chung vỏ của T12; `createDocument.ts` nay nhận hàm `buildBody` của từng chuẩn (không còn phụ thuộc kiểu của chuẩn Chính quyền). Đoạn chỗ nội dung được gắn sẵn style `HD05_NoiDung` (mã style suy từ tên, Word bỏ dấu gạch dưới) và **giãn dòng Exactly đặt trực tiếp** vì style cài bằng JSON chỉ ra "tối thiểu"; sau khi dựng, mã đọc lại và báo nếu style hay Exactly chưa đúng. Thể loại: chọn công văn, có thể loại (báo cáo, thông báo, quyết định...) hoặc tờ trình ở form Khối đầu; thể loại gõ tay (danh sách đầy đủ ở TD7). Chưa có: chọn thể loại từ danh sách, gợi ý ký hiệu loại theo thể loại.
**File thực tế:** `src/standards/hd05/assembleDocument.ts` (+ test, golden), `src/office/createDocument.ts`, `src/ooxml/documentAssembler.ts` (`markerParagraph`), `src/ui/NewDocWizard.ts`, `src/taskpane/taskpane.ts`
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
- [x] Phát hiện đúng sai khổ giấy, hướng trang, lề, khoảng cách số trang, phông, cỡ chữ, căn lề, thụt dòng đầu, giãn đoạn, giãn dòng (kể cả Single thay vì Exactly của chuẩn Đảng); chạy theo chuẩn đang chọn
- [x] Mỗi lỗi nêu quy tắc (mã), mô tả tiếng Việt, giá trị hiện tại và mong đợi, điều khoản tham chiếu lấy từ bảng quy tắc của chuẩn
- [x] Văn bản dựng từ trình tạo (công văn và quyết định chuẩn Chính quyền, văn bản có thể loại và căn cứ chuẩn Đảng, kèm style của bộ chuẩn) cho 0 lỗi trong test
**Kiểm chứng:**
- [x] `npm test`: 721 test; riêng bộ kiểm tra 42 ca có lỗi và không lỗi, 8 ca giải style
- [ ] Thủ công: **chờ bạn thử** tài liệu cố ý sai từng quy tắc
**Phụ thuộc:** T2, T2B, T5, T6B
**Trạng thái:** mã, bảng "Kiểm tra văn bản" (nút kiểm tra, danh sách lỗi, nút "Đến vị trí") và test xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Cách làm: đọc OOXML của thân + `PageSetup` từng phần, **giải style** (`ooxml/styleResolver.ts`: mặc định tài liệu, chuỗi basedOn, định dạng trực tiếp) để biết định dạng hiệu lực, vì nội dung dựng từ style không có định dạng trực tiếp. Chỉ kiểm các đoạn **ngoài bảng**; đoạn căn giữa/phải, đoạn có thụt (gạch đầu dòng), đoạn trống bị bỏ; căn lề, thụt dòng đầu, giãn đoạn, giãn dòng chỉ kiểm cho "văn xuôi" (đoạn từ 40 ký tự trở lên hoặc style `*_NoiDung`) để tránh báo sai; đoạn dùng style của chuẩn (ví dụ `ND30_Dieu`) kiểm cỡ chữ theo thành phần của style. Giãn đoạn tính tổng trước + sau ≥ 6 pt. **Giới hạn**: bảng lồng nhau và khối `w:sdt` làm lệch chỉ số đoạn khi nhảy tới vị trí; style cài bằng JSON của chuẩn Đảng ra giãn dòng "tối thiểu" nên đoạn chỉ dùng style (chưa ghi Exactly trực tiếp) sẽ bị báo, đúng thực tế; chưa kiểm khối đầu/cuối (T14, T14B), màu chữ, đánh số trang. **Bổ sung nút "Sửa" cho mọi lỗi trang/đoạn** (trừ trang ngang): khổ giấy → A4 đứng; lề/khoảng cách header → giá trị mặc định của chuẩn kẹp vào khoảng cho phép; phông → phông chuẩn; cỡ chữ → kẹp vào khoảng của thành phần; căn đều; thụt dòng đầu; giãn đoạn; giãn dòng (bội số hoặc Exactly, ghi lại đoạn bằng OOXML nên bình luận gắn trong đoạn đó có thể mất). Mỗi nút chỉ sửa đúng đoạn/phần của lỗi đó.
**File thực tế:** `src/checker/{types,documentModel,pageRules,paragraphRules,engine}.ts` (+ test), `src/ooxml/styleResolver.ts` (+ test), `src/office/readDocument.ts`, `src/ui/ReportPanel.ts`; `docModel.ts` thêm `styleId`, `hangingIndentMm`
**Kích cỡ:** M

### T14: Kiểm tra khối đầu/cuối và mẫu văn bản (chuẩn Chính quyền)
**Mô tả:** Thêm quy tắc nhận diện và kiểm tra: quốc hiệu/tiêu ngữ đúng chữ, "Số: ./." đúng mẫu, ngày tháng đúng quy cách, nơi nhận đúng cấu trúc, chức vụ in hoa, họ tên không học hàm. Mỗi lỗi có nút "Sửa".
**Tiêu chí nghiệm thu:**
- [x] Phát hiện đúng các lỗi phổ biến: quốc hiệu và tiêu ngữ viết sai (kể cả dấu gạch dài, thiếu dấu, viết thường) hoặc không đậm, sai cỡ; số và ký hiệu (thiếu số 0, thiếu dấu hai chấm, khoảng trắng thừa quanh "/" và "-"); ngày tháng (thiếu số 0 ở ngày < 10 và tháng 1, 2, thừa số 0, ngày không tồn tại, năm không đủ 4 chữ số); nơi nhận (dấu cuối mục, gạch đầu dòng, cỡ chữ, định dạng "Nơi nhận"); chức vụ (in hoa, đậm, cỡ), quyền hạn thiếu dấu chấm; họ tên (học hàm, học vị, in hoa toàn bộ, không đậm)
- [x] Nút "Sửa" sửa đúng một lỗi: mỗi cách sửa chỉ trỏ vào một đoạn, test kiểm tra áp cách sửa thì lỗi đó hết và các đoạn khác nguyên vẹn; lỗi về chữ không thể sửa chắc chắn (ngày không tồn tại, học hàm, họ tên in hoa) không có nút Sửa
- [x] Tỷ lệ báo sai trên bộ văn bản chuẩn = 0 trong test: công văn dựng từ trình tạo (nhiều nơi nhận, để trống số và ngày, sáu kiểu quyền hạn)
- [x] Đăng ký bước 5 "Chuẩn hoá khối đầu, khối cuối" vào "Chỉnh toàn bộ" (T6C): chỉ chạy các cách sửa **định dạng** (đậm, nghiêng, cỡ chữ), không đổi chữ; lỗi về chữ được đếm và báo để xem ở mục Kiểm tra văn bản. Nhận diện theo nội dung chữ ("Nơi nhận:", quốc hiệu, tiêu ngữ, số, ngày) trong các bảng; chuẩn Đảng chưa có (T14B). Đánh số trang đổi thành bước 6
**Kiểm chứng:**
- [x] `npm test`: 55 ca riêng cho khối đầu, khối cuối (có/không lỗi, cách sửa, không báo nhầm), tổng 776 test
- [ ] Thủ công: **chờ bạn thử** 5 văn bản thật (đã ẩn thông tin) từ người dùng văn thư
**Phụ thuộc:** T7–T10, T13
**Trạng thái:** mã, nút "Sửa" trong bảng Kiểm tra văn bản và bước trong "Chỉnh toàn bộ" xong; `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Quy tắc nằm ở `src/standards/nd30/checks.ts` (dùng chung khung engine T13; chuẩn Đảng sẽ có `src/standards/hd05/checks.ts`), chọn bộ quy tắc theo `headerKind` ở `checker/ruleSets.ts`. Chỉ xét đoạn trong bảng nên không đụng phần thân. Cách sửa: thay chữ (tìm chữ trong đúng đoạn, giữ định dạng) hoặc đặt đậm, nghiêng, cỡ chữ cho cả đoạn (`office/applyFix.ts`, một lần Word.run). **Giới hạn**: cách nhận diện khối ký dựa vào ô cùng hàng với "Nơi nhận:" nên khối ký không đi kèm "Nơi nhận:" thì không được kiểm; không kiểm khoảng trắng chỗ ký, vị trí đường kẻ, "V/v" và Kính gửi; ký hiệu chỉ kiểm cách viết, không biết đúng chữ viết tắt hay không.
**File thực tế:** `src/standards/nd30/checks.ts` (+ test), `src/checker/{fixes,ruleSets}.ts`, `src/office/applyFix.ts`, mở rộng `src/ui/ReportPanel.ts`, `src/pipeline/steps.ts`
**Kích cỡ:** M

### T14B: Kiểm tra lỗi thể thức theo chuẩn Đảng (HD 05)
**Mô tả:** Thêm bộ quy tắc kiểm tra cho chuẩn Đảng vào engine T13: lề 20/20/30/15 mm (và đối xứng khi in hai mặt), giãn dòng Exactly 18–22 pt, thụt dòng đầu khoảng 10 mm, cỡ chữ từng thành phần theo bảng 6B, tiêu đề "ĐẢNG CỘNG SẢN VIỆT NAM" đúng chữ, dấu `*`, "Số 05-HD/VPTW" (không dấu hai chấm, số 0 phía trước, "-" và "/" đúng chỗ), ngày tháng, 5 dấu gạch nối dưới trích yếu, chức vụ không đậm, "Nơi nhận" gạch chân, số trang cách mép trên 10 mm và ẩn ở trang 1, nội dung kết thúc bằng dấu chấm, bố cục Phần/Chương/Mục/Điều/Khoản/Điểm. Mỗi lỗi có nút "Sửa" khi sửa được tự động.
**Tiêu chí nghiệm thu:**
- [x] Phát hiện đúng các lỗi phổ biến: dùng tiêu ngữ của chính quyền, ghi "Số:" có dấu hai chấm, thiếu 5 dấu gạch nối, giãn dòng sai kiểu (Single thay vì Exactly), chức vụ in đậm, thiếu số 0
- [x] Cảnh báo khi tài liệu trộn dấu hiệu của hai chuẩn (có cả quốc hiệu Chính quyền và tiêu đề Đảng)
- [x] Văn bản chuẩn Đảng tạo từ TD5 cho 0 lỗi; văn bản chuẩn Chính quyền không bị báo theo luật Đảng
- [x] Đăng ký bước "Chuẩn hoá khối đầu/cuối (Đảng)" vào "Chỉnh toàn bộ" (T6C): chỉ sửa định dạng, không đổi chữ; không tự đổi tiêu đề từ chuẩn này sang chuẩn kia (đổi chuẩn là thao tác riêng có xác nhận ở T6B)
**Kiểm chứng:**
- [x] `npm test`: tối thiểu 30 ca có/không lỗi cho chuẩn Đảng
- [ ] Thủ công: tài liệu cố ý sai từng quy tắc; thêm một văn bản Đảng thật bạn đang có (đã ẩn thông tin) nếu tiện, không bắt buộc
**Phụ thuộc:** T13, TD1–TD4
**Trạng thái:** mã và test xong (39 ca ở `src/standards/hd05/checks.test.ts`, 2 ca ở `engine.test.ts`; 820 test); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Quy tắc: tiêu đề Đảng (chữ, đậm, cỡ 15), tiêu ngữ Chính quyền trong văn bản Đảng, cơ quan ban hành (hoa, đậm, cỡ) và dấu `*`, "Số 05-HD/VPTW" (dấu hai chấm, số 0, khoảng trắng; ghi kiểu Chính quyền chỉ báo), ngày tháng (dùng chung với Chính quyền, thêm in nghiêng), nhãn Kính gửi/Kính trình nghiêng, thể loại, trích yếu, 5 dấu gạch nối (sai số lượng sửa được; thiếu cả dòng chỉ báo), "Nơi nhận" gạch chân, danh sách nơi nhận, quyền hạn T/M, K/T, T/L (TM., KT., TL. tự sửa; TUQ. chỉ báo), chức vụ hoa và không đậm, họ tên; nội dung kết thúc bằng chữ/số thì cảnh báo thiếu dấu chấm (quy tắc chung `paragraph.endsWithPeriod`). Cảnh báo trộn hai chuẩn (`header.mixed`) và tiêu đề của chuẩn khác (`header.otherStandard`) dùng chung cho cả hai chuẩn (`checker/mixedStandards.ts`); hàm chung chuyển sang `checker/common.ts`. Bước 5 của "Chỉnh toàn bộ" dùng chung cho hai chuẩn (chọn quy tắc theo `headerKind`), chỉ sửa định dạng (đậm, nghiêng, gạch chân, cỡ chữ), nay tính cả lỗi `title.*`. **Giới hạn**: thể loại nhận diện là đoạn in hoa căn giữa đầu tiên trong 6 đoạn thân đầu; chưa kiểm số trang ẩn ở trang 1 và khoảng cách 10 mm trong header, bố cục Phần/Chương/Mục/Điều/Khoản/Điểm, ký hiệu người soạn; khối ký không đi kèm "Nơi nhận:" thì không được kiểm.
**File dự kiến:** `src/standards/hd05/checks.ts`, `src/standards/hd05/fixes.ts`, `src/checker/engine.ts`, `src/standards/hd05/checks.test.ts`
**Kích cỡ:** M

### T15: "Sửa tất cả" có xem trước và hoàn tác
**Mô tả:** Sửa hàng loạt các lỗi có thể sửa tự động; xem trước danh sách thay đổi; hoàn tác một lần; xuất báo cáo lỗi (txt/docx).
**Tiêu chí nghiệm thu:**
- [x] Xem trước liệt kê chính xác các thay đổi sẽ làm
- [x] Sau khi sửa, kiểm tra lại cho 0 lỗi có thể sửa
- [x] Đăng ký bước cuối "Sửa các lỗi còn lại" vào "Chỉnh toàn bộ" (T6C); chạy "Chỉnh toàn bộ" trên văn bản có lỗi cài sẵn của cả hai chuẩn rồi kiểm tra lại cho 0 lỗi có thể sửa
- [ ] Văn bản 100 trang xử lý trong ngưỡng chấp nhận (đặt ngưỡng cụ thể sau khi đo, ví dụ < 15 giây)
**Kiểm chứng:**
- [x] `npm test`: logic gom/sắp xếp bản sửa
- [ ] Thủ công: tài liệu 100 trang, đo thời gian; Ctrl+Z hoàn tác
**Phụ thuộc:** T14, T14B
**Trạng thái:** mã và test xong (12 ca ở `src/checker/fixAll.test.ts`; 832 test); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**; **chưa đo văn bản 100 trang** (còn mở). Cách làm: `checker/fixAll.ts` (thuần tuý) gom cách sửa của danh sách lỗi — bỏ trùng, gộp font/đoạn/lề cùng chỗ, hoãn cách thay chữ thứ hai trên cùng một đoạn sang lượt sau, sắp xếp trang trước rồi theo đoạn — và giả lập "sửa rồi kiểm tra lại" (tối đa 3 lượt). `office/fixAll.ts` làm vòng sửa thật trong Word. Bảng Kiểm tra văn bản: "Sửa tất cả (N)" mở hộp xem trước liệt kê từng thay đổi, bấm "Áp dụng" mới sửa (chụp tài liệu trước, sau đó có nút "Hoàn tác lần sửa tất cả" dùng bản chụp T3, hoặc Ctrl+Z); "Xuất báo cáo" hiện báo cáo chữ (`checker/exportReport.ts`) kèm nút Sao chép và liên kết Tải về .txt. Bước 7 "Sửa các lỗi còn lại" của "Chỉnh toàn bộ": chỉ cách sửa không đổi chữ; chọn giữ định dạng cục bộ thì bỏ qua lỗi phông, cỡ chữ, căn lề, giãn của phần thân. **Giới hạn**: chưa có xuất .docx (chỉ chữ thuần); hoàn tác bằng bản chụp thay lại cả thân văn bản nên mất mọi chỉnh sửa tay sau lần "Sửa tất cả"; footer chưa nằm trong bản chụp.
**File dự kiến:** `src/checker/fixAll.ts`, `src/ui/FixPreview.ts`, `src/checker/exportReport.ts`, `src/checker/fixAll.test.ts`
**Kích cỡ:** M

### Checkpoint 4 – Bộ kiểm tra
- [ ] Nút "Chỉnh toàn bộ" chạy đủ các bước (trang, style, thân, bảng, khối đầu/cuối, số trang, sửa lỗi còn lại) trên cả hai chuẩn và kiểm tra lại cho 0 lỗi sửa được
- [x] Bộ văn bản mẫu có lỗi cài sẵn cho **cả hai chuẩn**: báo đủ lỗi, không báo sai trên văn bản chuẩn; tài liệu trộn dấu hiệu hai chuẩn được cảnh báo
- [ ] Người dùng văn thư thử trên văn bản thật và xác nhận mức độ hữu ích
- [x] Test, build, lint xanh

**Trạng thái (tự động):** `src/checker/checkpoint4.test.ts` (20 ca, 852 test cả dự án): văn bản dựng từ trình tạo của cả hai chuẩn cho 0 lỗi; 12 lỗi cài sẵn (quốc hiệu/tiêu đề, số, ngày, chức vụ, họ tên, gạch nối, "Nơi nhận", quyền hạn…) đều được báo đúng mã và giả lập "Sửa tất cả" hết lỗi sửa được; trộn tiêu đề hai chuẩn bị cảnh báo. **Còn chờ bạn**: chạy "Chỉnh toàn bộ" thật trong Word cho cả hai chuẩn và thử trên văn bản thật (hai ô chưa tick: "Chỉnh toàn bộ" chạy đủ bước, người dùng văn thư xác nhận).

---

## Giai đoạn 5 – Phủ mọi loại văn bản hành chính (theo họ mẫu)

Danh sách loại và họ mẫu lấy từ `docTypes` của từng chuẩn (NĐ 30 ở T2, HD 05 ở T2B). Mỗi loại là một mục cấu hình trong registry (T17); T18–T20 (Chính quyền) và TD6–TD8 (Đảng) chỉ thêm họ mẫu mới và dữ liệu, không sửa wizard.

### T16: Quyết định và nghị quyết (cá biệt)
**Mô tả:** Mẫu có phần "Căn cứ" (nghiêng, kết thúc ";"), dòng "QUYẾT ĐỊNH:" (hoặc tương ứng), đánh số Điều/Khoản/Điểm bằng multilevel list gắn với style `ND30_*`.
**Tiêu chí nghiệm thu:**
- [x] Điều (đậm), Khoản (số), Điểm (chữ cái) đánh số tự động đúng cấp
- [x] Hai loại sinh đúng thể thức: nghị quyết cá biệt (Mẫu 1.1) và quyết định cá biệt quy định trực tiếp (Mẫu 1.2) hoặc gián tiếp (Mẫu 1.3). Chỉ thị dùng Mẫu 1.4, làm ở T17
- [x] Bộ kiểm tra T14 nhận biết và kiểm tra các thành phần mới
**Kiểm chứng:**
- [x] `npm test`: golden OOXML numbering
- [ ] Thủ công: so mẫu Phụ lục III; thêm/xoá Điều thì số tự cập nhật
**Phụ thuộc:** T12, T14
**Trạng thái:** mã và test xong (`src/ooxml/numbering.ts`, `src/templates/quyetDinh.ts`, `src/standards/nd30/decisionChecks.ts`; 3 golden `quyetDinh-*.xml`; 886 test); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word** (numbering qua `insertOoxml` là điểm rủi ro nhất, chưa có spike). Cách dùng: ở chuẩn Chính quyền, mục "Tạo công văn" có ô chọn *Dạng văn bản* (Công văn / Nghị quyết cá biệt / Quyết định cá biệt trực tiếp / gián tiếp); chọn dạng quyết định thì Khối đầu phải ở chế độ "Văn bản có tên loại" (add-in tự ghi tên loại QUYẾT ĐỊNH hoặc NGHỊ QUYẾT), nhập thêm thẩm quyền ban hành, các căn cứ, số Điều. Kết quả: khối đầu, tên loại + trích yếu, thẩm quyền (đậm, giữa), căn cứ (nghiêng, `;` và `,` ở căn cứ cuối), "QUYẾT ĐỊNH:"/"QUYẾT NGHỊ:", các Điều tự đánh số "Điều N." (nhãn đậm, nội dung thường, style ND30_NoiDung), con trỏ ở Điều 1 (gián tiếp: sau câu "Ban hành kèm theo Quyết định này "), khối cuối. Numbering ba cấp Điều/Khoản/Điểm đã định nghĩa (`legalParagraph("khoan"|"diem", …)`), nhưng trình tạo mới chỉ sinh cấp Điều; Khoản/Điểm người dùng thêm bằng Tab/Shift+Tab trong danh sách. Bộ kiểm tra: căn cứ liền trên dòng chốt phải nghiêng, cỡ 13–14, dấu cuối dòng đúng; dòng chốt đậm (`basis.*`, có nút Sửa). **Giới hạn / cần bạn duyệt**: (1) dấu cuối căn cứ cuối (xem `docs/nd30-ambiguities.md` mục 14); (2) chữ cái Điểm "đ" (mục 15); (3) Mẫu 1.2 và 1.3 hiện khác nhau chỉ ở câu mở Điều 1; chưa có phần "văn bản ban hành kèm theo"; (4) chưa tự động thêm Khoản/Điểm vào các Điều; (5) tên Điều (tiêu đề) và Khoản trong style ND30_Dieu/ND30_Khoan dùng cho quy chế (T18), quyết định dùng style nội dung.
**File dự kiến:** `src/ooxml/numbering.ts`, `src/templates/quyetDinh.ts`, `src/checker/headerRules.ts`, `src/templates/quyetDinh.test.ts`
**Kích cỡ:** M

### T17: Registry loại văn bản và nhóm "tên loại + trích yếu"
**Mô tả:** Dựng registry cấu hình (mỗi loại khai báo khối đầu/cuối, trường bắt buộc, ký hiệu viết tắt) và đưa vào wizard bước chọn loại. Làm nhóm dùng Mẫu 1.4 (có tên loại + trích yếu), gồm 17 loại: chỉ thị, quy chế, quy định, thông cáo, thông báo, hướng dẫn, chương trình, kế hoạch, phương án, đề án, dự án, báo cáo, tờ trình, giấy ủy quyền, phiếu gửi, phiếu chuyển, phiếu báo (đọc từ `layout: "co-ten-loai"` trong `src/rules/docTypes.ts`). Quy chế và quy định có thêm cấu trúc Chương/Điều ở T18.
**Tiêu chí nghiệm thu:**
- [x] Thêm một loại mới chỉ cần thêm một mục cấu hình, không sửa wizard
- [x] Các loại trong nhóm sinh đúng thể thức so với mẫu
- [x] Ký hiệu viết tắt gợi ý đúng cho từng loại
**Kiểm chứng:**
- [x] `npm test`: golden cho từng loại; test registry (mọi loại ở `docTypes.ts` đều có cấu hình hợp lệ hoặc được đánh dấu "chưa hỗ trợ")
- [ ] Thủ công: tạo từng loại, in PDF, đối chiếu mẫu
**Phụ thuộc:** T12
**Trạng thái:** mã và test xong (`src/templates/registry.ts`, `registry.test.ts`, `titledDocs.test.ts`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Registry gán họ mẫu cho mọi loại trong `docTypes` của chuẩn đang chọn (theo kiểu khối đầu và họ bố cục, ngoại lệ theo mã loại); thêm loại mới là thêm một mục vào `FAMILIES`/`BY_ID`, không sửa wizard. Wizard có ô "Loại văn bản" liệt kê mọi loại (loại chưa có họ mẫu hiện mờ kèm lý do "chưa hỗ trợ", không bị ẩn); chọn loại thì Khối đầu tự chuyển dạng (công văn / có tên loại / tờ trình), điền tên loại và gợi ý chữ viết tắt (`HeaderForm.preset`). 17 loại của Mẫu 1.4 (Chính quyền) và các thể loại có tên loại của chuẩn Đảng đều có test dựng văn bản và kiểm tra 0 lỗi. **Giới hạn**: quy chế và quy định chưa có cấu trúc Chương/Mục/Điều (T18); biên bản, hợp đồng, bản ghi nhớ, bản thỏa thuận, giấy tờ nhỏ hiện "chưa hỗ trợ" (T19, T20, TD7).
**File dự kiến:** `src/templates/registry.ts`, `src/templates/titledDocs.ts`, `src/ui/NewDocWizard.ts`, `src/templates/registry.test.ts`
**Kích cỡ:** M

### T18: Quy chế, quy định (Chương/Mục/Điều)
**Mô tả:** Họ mẫu có cấu trúc Chương, Mục, Điều, Khoản, Điểm với cỡ chữ và kiểu chữ từng cấp theo quy tắc; tên văn bản dạng "QUY CHẾ" kèm trích yếu và phần ban hành kèm theo.
**Tiêu chí nghiệm thu:**
- [x] Chương/Mục/Điều đánh số tự động, đúng cấp và đúng kiểu chữ
- [x] Hai loại (quy chế, quy định) sinh đúng thể thức
- [x] Bộ kiểm tra nhận biết cấu trúc Chương/Mục
**Kiểm chứng:**
- [x] `npm test`: golden OOXML numbering nhiều cấp
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T16, T17
**Trạng thái:** mã và test xong (`src/templates/quyChe.ts`, `src/ooxml/numbering.ts` nay là bộ dựng danh sách nhiều cấp dùng chung, `src/standards/nd30/structureChecks.ts`; golden `quyChe-*.xml`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Loại "quy chế", "quy định" (chuẩn Chính quyền) có họ mẫu `regulation`: dòng "(Ban hành kèm theo ...)", rồi từng Chương (nhãn "Chương I" tự đánh số La Mã, tên Chương in hoa đậm ở đoạn dưới, con trỏ ở tên Chương 1), tuỳ chọn Mục ("Mục 1" + tên), các Điều "Điều N." (số chạy liên tục qua Chương/Mục, `lvlRestart=0`), mỗi Điều một Khoản "1."; Điểm "a)" có trong danh sách. Bộ kiểm tra `structure.*`: tên Chương/Mục in hoa và đậm, Điều đậm, nhãn "Điều N." gõ tay liên tục. **Rủi ro chưa thử**: `lvlRestart`, `suff=nothing`, nhãn La Mã tự đánh số khi nhận qua `insertOoxml`. **Cần bạn duyệt**: dòng "ban hành kèm theo" (docs/nd30-ambiguities.md mục 16).
**File dự kiến:** `src/templates/quyChe.ts`, `src/ooxml/numbering.ts`, `src/checker/headerRules.ts`, `src/templates/quyChe.test.ts`
**Kích cỡ:** M

### T19: Biên bản và văn bản nhiều bên ký
**Mô tả:** Biên bản (thời gian, địa điểm, thành phần, nội dung, khối ký chủ trì/thư ký), hợp đồng, bản ghi nhớ, bản thỏa thuận (khối ký hai hoặc nhiều bên, có thể không có "Nơi nhận"). Ba loại sau **không có mẫu ở Phụ lục III** nên cần bạn duyệt bố cục đề xuất (xem `docs/nd30-ambiguities.md`, mục 4). Bản cam kết không có trong NĐ 30.
**Tiêu chí nghiệm thu:**
- [x] Khối ký nhiều bên (2–4 bên) dàn đều đúng bố cục, không vỡ khi tên dài
- [x] Các loại sinh đúng thể thức so với mẫu
- [x] Bộ kiểm tra không báo sai "thiếu nơi nhận" cho loại không yêu cầu
**Kiểm chứng:**
- [x] `npm test`: golden OOXML khối ký 2/3/4 bên
- [ ] Thủ công: in PDF, so mẫu
**Phụ thuộc:** T17
**Trạng thái:** mã và test xong (`src/ooxml/multiSignBlock.ts`, `src/templates/bienBan.ts`, `src/templates/hopDong.ts`; golden `multiSign-2|3|4.xml`, `bienBan-body.xml`, `hopDong-2|4.xml`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Khối ký nhiều bên: bảng không viền 2 đến 4 cột bằng nhau, mỗi cột tên bên/vai trò + chức vụ (in hoa đậm), chỗ ký, họ tên (đậm); tên dài tự xuống dòng trong cột; hàng "Nơi nhận" trải hết các cột chỉ có khi người dùng nhập. Biên bản: các mục Thời gian, Địa điểm, Thành phần, Chủ trì, Thư ký, Nội dung, ký mặc định chủ tọa + thư ký. Hợp đồng, bản ghi nhớ, bản thỏa thuận: **bố cục đề xuất** (không có mẫu ở Phụ lục III): dòng mở đầu, nhãn từng bên, Điều đánh số tự động, ký 2–4 bên. Trong trình tạo, hai họ mẫu này nhập khối ký bằng các dòng "Tên bên | Chức vụ | Họ tên" và không dùng Khối cuối. Bộ kiểm tra không báo sai khi không có "Nơi nhận" (có test). **Cần bạn duyệt**: bố cục hợp đồng/bản ghi nhớ/bản thỏa thuận và vị trí các mục của biên bản so với Mẫu 1.9 (tôi chưa đối chiếu nguyên văn Mẫu 1.9).
**File dự kiến:** `src/ooxml/multiSignBlock.ts`, `src/templates/bienBan.ts`, `src/templates/hopDong.ts`, `src/ooxml/multiSignBlock.test.ts`
**Kích cỡ:** M

### T20: Giấy tờ và mẫu nhỏ
**Mô tả:** Giấy mời (Mẫu 1.7), giấy giới thiệu (Mẫu 1.8), giấy nghỉ phép (Mẫu 1.10), công điện (Mẫu 1.6) và thư công (không có mẫu ở Phụ lục III). Giấy ủy quyền, phiếu gửi, phiếu chuyển, phiếu báo dùng Mẫu 1.4 nên làm ở T17. Mỗi mẫu có bố cục riêng, thường ngắn và có phần tuỳ biến (khung, bảng nhỏ). Nếu một mẫu quá phức tạp thì tách thành task riêng sau khi duyệt.
**Tiêu chí nghiệm thu:**
- [x] Mỗi mẫu có trong registry và sinh đúng thể thức so với Phụ lục III
- [ ] Mẫu chưa hỗ trợ (nếu có) được đánh dấu rõ trong giao diện, không im lặng
- [x] Bộ kiểm tra không báo sai cho các mẫu này
**Kiểm chứng:**
- [x] `npm test`: golden từng mẫu
- [ ] Thủ công: tạo từng mẫu, in PDF
**Phụ thuộc:** T17
**Trạng thái:** mã và test xong (`src/templates/smallForms.ts`; golden `smallForm-*.xml`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Năm mẫu có trong registry (họ `form`) và trong trình tạo: giấy mời, giấy giới thiệu, công điện, giấy nghỉ phép, thư công. Mỗi mẫu = khối đầu + tên loại/trích yếu + các mục cần điền (nhãn đậm, con trỏ ở mục đầu) + câu kết + khối cuối thường. Với chuẩn Chính quyền, mọi loại trong `docTypes` nay đều có họ mẫu. **Giới hạn quan trọng**: tôi chưa có nguyên văn Mẫu 1.6, 1.7, 1.8, 1.10 nên các mục và câu kết là bố cục theo cách trình bày thông dụng; trình tạo ghi rõ điều này khi chọn từng mẫu. Cần bạn đối chiếu với Phụ lục III và báo mục nào khác để tôi sửa (một chỗ: `SMALL_FORMS`). Khung viền, bảng nhỏ riêng của từng mẫu chưa làm. Công điện chưa có dấu "HỎA TỐC" (T21).
**File dự kiến:** `src/templates/smallForms.ts`, `src/ooxml/smallForms.ts`, `src/templates/registry.ts`, `src/templates/smallForms.test.ts`
**Kích cỡ:** M

### T21: Phụ lục và dấu chỉ mức độ khẩn/mật
**Mô tả:** Chèn trang phụ lục (ngắt trang, tiêu đề "PHỤ LỤC", tên phụ lục) và khung dấu mức độ khẩn (KHẨN/THƯỢNG KHẨN/HỎA TỐC) và độ mật theo quy tắc.
**Tiêu chí nghiệm thu:**
- [x] Dấu khẩn trong khung chữ nhật viền đơn, in hoa, đậm, đúng vị trí
- [x] Phụ lục bắt đầu trang mới, đánh số thứ tự tự động
- [x] Bộ kiểm tra nhận ra hai thành phần này
**Kiểm chứng:**
- [x] `npm test`: golden OOXML
- [ ] Thủ công: so mẫu Phụ lục III
**Phụ thuộc:** T12
**Trạng thái:** mã và test xong (`src/ooxml/urgencyMark.ts`, `src/ooxml/appendix.ts`, `src/office/insertExtras.ts`, `src/ui/InsertPanel.ts`, `src/checker/urgencyRules.ts`; golden `urgencyMark-khan.xml`, `appendix-1.xml`; 1012 test); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bảng mới "Chèn dấu khẩn và phụ lục": chọn dấu (HỎA TỐC 30×8, THƯỢNG KHẨN 40×8, KHẨN 20×8 mm theo quy định) rồi "Chèn khung tại con trỏ" (bảng một ô viền đơn, cao cố định, chữ in hoa đậm giữa ô); nhập tên và chỉ dẫn rồi "Chèn phụ lục ở cuối tài liệu" (nhãn "Phụ lục I, II…" tự đánh số La Mã, nối tiếp các phụ lục đã có theo số đoạn style PhuLucSo; bắt đầu trang mới bằng `pageBreakBefore`; tên in hoa; chỉ dẫn trong ngoặc, nghiêng). Bộ kiểm tra: `urgency.*` nhận ra khung (chữ trùng dấu của chuẩn, trong bảng) và kiểm in hoa, đậm, cỡ 13–14 (dùng cho cả chuẩn Đảng); `structure.phuLuc.*` kiểm tên phụ lục. **Giới hạn**: dấu khẩn chèn tại con trỏ chứ chưa tự tìm chỗ dưới "Số, ký hiệu"; "độ mật" (mật, tối mật) không có trong NĐ 30 nên chưa làm, phạm vi lưu hành chưa có mục chèn riêng; nếu con trỏ đang ở giữa chữ thì bảng tách đoạn tại đó.
**File dự kiến:** `src/ooxml/urgencyMark.ts`, `src/ooxml/appendix.ts`, `src/ui/InsertPanel.ts`, `src/ooxml/urgencyMark.test.ts`
**Kích cỡ:** S

### TD6: Đảng: Phần/Chương/Mục/Điều/Khoản/Điểm (nghị quyết, quyết định, quy định, quy chế)
**Mô tả:** Cấu trúc nội dung theo HD 05: "Phần"/"Chương" và số (La Mã hoặc chữ, ví dụ "Phần I", "Chương thứ hai") cùng dòng, giữa trang, in thường đậm 14–15 pt, tên phần/chương in hoa đậm ở dòng dưới; "Mục" + số Ả Rập đậm, tên mục in hoa đậm; "Điều 1. ..." in thường đậm, số Ả Rập có dấu chấm; Khoản "1." và Điểm "a)" in thường đứng. Đánh số tự động bằng multilevel list gắn style `HD05_*`. Mẫu cho nghị quyết, quyết định, quy định, quy chế (có phần "Căn cứ" từ TD3).
**Tiêu chí nghiệm thu:**
- [x] Các cấp đánh số tự động đúng, đúng kiểu chữ theo mục 6B; thêm/xoá Điều thì số tự cập nhật
- [x] Bốn thể loại sinh đúng thể thức; bộ kiểm tra T14B nhận biết các cấp
- [x] Số La Mã và số bằng chữ cho Phần/Chương chọn được
**Kiểm chứng:**
- [x] `npm test`: golden OOXML numbering nhiều cấp
- [ ] Thủ công: in PDF, so bảng Phụ lục 2 mục 6
**Phụ thuộc:** TD5
**Trạng thái:** mã và test xong (`src/standards/hd05/numbering.ts`, `src/standards/hd05/templates/regulations.ts`, `src/standards/hd05/structureChecks.ts`; `src/checker/structureRules.ts` dùng chung với chuẩn Chính quyền; `buildStructureBody` tách khỏi `templates/quyChe.ts`; golden `partyRegulation-*.xml`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bốn thể loại Đảng (nghị quyết, quyết định, quy định, quy chế) có họ mẫu `regulation`: khối đầu, thể loại + trích yếu + 5 gạch nối + căn cứ (TD3), rồi Phần hoặc Chương (chọn; số La Mã hoặc số bằng chữ "Chương thứ hai"), tên Phần/Chương in hoa đậm ở dòng dưới, Mục "Mục 1" (tuỳ chọn) + tên, Điều "Điều N." (đậm, số chạy liên tục), Khoản "1.", Điểm "a)"; con trỏ ở tên Phần/Chương đầu. Kiểm tra `structure.*`: tên Phần/Chương/Mục in hoa đậm, Điều đậm (có nút Sửa). **Rủi ro chưa thử**: số bằng chữ dùng định dạng `vietnameseCounting` của Word, kết quả hiển thị (chữ hoa hay thường, "Hai" hay "hai") cần xem trên máy thật; nếu không đẹp thì chuyển sang gõ tay hoặc chỉ dùng La Mã. Cấp Điểm "đ" gõ tay.
**File dự kiến:** `src/standards/hd05/numbering.ts`, `src/standards/hd05/templates/regulations.ts`, `src/standards/hd05/checks.ts`, `src/standards/hd05/numbering.test.ts`
**Kích cỡ:** M

### TD7: Đảng: các thể loại còn lại theo Quy định 399-QĐ/TW
**Mô tả:** Dựa trên danh sách thể loại đọc ở T2B, bổ sung các họ mẫu còn lại vào registry T17: nghị quyết, chỉ thị, kết luận, thông tri, chương trình, kế hoạch, biên bản (hội nghị cấp uỷ: người ghi biên bản và chủ trì hội nghị ký, có dòng xác nhận chữ ký khi không đóng dấu), văn bản đại hội (tên đại hội, đoàn chủ tịch/đoàn thư ký/ban thẩm tra tư cách đại biểu/ban kiểm phiếu, ký hiệu "ĐH"), văn bản liên cơ quan (hai khối ký, tên cơ quan nối bằng gạch nối). Loại nào quá phức tạp thì tách thành task riêng sau khi duyệt.
**Tiêu chí nghiệm thu:**
- [x] Mỗi thể loại trong danh sách T2B hoặc có mẫu đúng thể thức hoặc ghi rõ "chưa hỗ trợ" và lý do trong giao diện
- [x] Biên bản, văn bản đại hội và văn bản liên cơ quan có khối ký đúng như ví dụ ở mục 7 của HD 05
- [x] Ký hiệu loại gợi ý đúng (NQ, CT, KL, QC, BC, QĐ, CTr, TT, TTr, CV, ...)
**Kiểm chứng:**
- [x] `npm test`: golden từng mẫu; test registry (mọi thể loại đều có cấu hình hoặc cờ "chưa hỗ trợ")
- [ ] Thủ công: tạo từng mẫu, in PDF, đối chiếu ví dụ trong HD 05
**Phụ thuộc:** T17, TD5
**Trạng thái:** mã và test xong (`src/standards/hd05/templates/documents.ts`; `multiSignBlock` thêm dòng "Xác nhận chữ ký của đồng chí ..." và kiểu "Nơi nhận" gạch chân của Đảng; `smallForms` thêm giấy chứng nhận, giấy đi đường, phiếu chuyển, phiếu gửi cho Đảng; golden `partyForm-giay-chung-nhan.xml`; 1042 test); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Registry chuẩn Đảng nay không còn loại nào "chưa hỗ trợ": các thể loại có tên loại (chỉ thị, kết luận, thông tri, chương trình, kế hoạch, báo cáo… ký hiệu NQ, CT, KL, QC, BC, QĐ, CTr, TT, TTr) dùng họ `titled`; nghị quyết, quyết định, quy định, quy chế dùng cấu trúc Phần/Chương (TD6); **biên bản** (hai chữ ký "Người ghi biên bản", "Chủ trì hội nghị", dòng xác nhận chữ ký nhập ở cột thứ tư khi không đóng dấu); **văn bản đại hội** (dòng tên đại hội in hoa đậm trên thể loại, gợi ý ký hiệu ĐH, ký T/M Đoàn chủ tịch qua Khối cuối); **văn bản liên cơ quan** (đúng hai khối ký, nhiều dòng cơ quan ban hành và ký hiệu nối bằng gạch nối ở Khối đầu); 8 giấy tờ (họ `form`). Bộ kiểm tra nhận ra dòng tên đại hội, không nhầm với thể loại. **Cần bạn duyệt**: tôi chưa có nguyên văn mục 7 của HD 05 nên bố cục khối ký biên bản, đại hội, liên cơ quan và các mục của giấy tờ là đề xuất; cũng chưa tách riêng "ban thẩm tra tư cách đại biểu", "ban kiểm phiếu".
**File dự kiến:** `src/standards/hd05/templates/*.ts`, `src/standards/hd05/signatureForms.ts`, `src/templates/registry.ts`, `src/standards/hd05/templates/templates.test.ts`
**Kích cỡ:** M

### TD8: Đảng: phụ lục, dấu khẩn, phạm vi lưu hành, dự thảo, thông tin liên hệ
**Mô tả:** Thành phần bổ sung theo mục II của HD 05: trang phụ lục ("Phụ lục" + số, dòng chỉ dẫn số/ký hiệu văn bản, đánh số trang riêng cho từng phụ lục); dấu chỉ mức độ khẩn (KHẨN, THƯỢNG KHẨN, HỎA TỐC; in hoa 14 pt, đậm, dưới số và ký hiệu, công văn thì dưới trích yếu); chỉ dẫn phạm vi lưu hành (in hoa 12 pt, đậm, góc phải trang đầu, cách mép trên 10 mm; "Không phổ biến trên phương tiện thông tin đại chúng" ở chân trang cuối); chỉ dẫn dự thảo ("Dự thảo lần 2", 14 pt nghiêng, dưới số và ký hiệu); ký hiệu người soạn và số lượng bản (11 pt); thông tin liên hệ (10 pt, trang cuối, dưới đường kẻ ngang hết vùng trình bày).
**Tiêu chí nghiệm thu:**
- [x] Từng thành phần đúng cỡ, kiểu, vị trí theo mục 6B và Phụ lục 1 (ô 10, 11a, 11b, 12, 13)
- [ ] Phụ lục bắt đầu trang mới, đánh số trang riêng; chạy lại không tạo trùng
- [x] Bộ kiểm tra T14B nhận biết các thành phần này
**Kiểm chứng:**
- [x] `npm test`: golden OOXML từng thành phần
- [ ] Thủ công: in PDF, so Phụ lục 1 và bảng Phụ lục 2
**Phụ thuộc:** TD5
**Trạng thái:** mã và test xong (`src/standards/hd05/extras.ts`; `ui/InsertPanel.ts` nay dùng cho cả hai chuẩn; golden `partyContact.xml`, `partyAppendix-2.xml`); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Chuẩn Đảng có thêm trong bảng "Chèn dấu khẩn, phụ lục…": dấu khẩn (KHẨN, THƯỢNG KHẨN, HỎA TỐC; dùng chung `urgencyMark`), chỉ dẫn phạm vi lưu hành (khung canh phải 12 pt in hoa đậm), chỉ dẫn dự thảo ("Dự thảo lần N", 14 pt nghiêng), dòng "Không phổ biến trên phương tiện thông tin đại chúng" (12 pt) và thông tin liên hệ (đường kẻ ngang rồi các dòng 10 pt) ở cuối tài liệu, trang phụ lục ("Phụ lục I…" tự đánh số nối tiếp, đếm theo nhãn danh sách, bắt đầu trang mới, tên in hoa đậm, chỉ dẫn nghiêng). Ký hiệu người soạn và số bản (11 pt) đã có từ TD4. Bộ kiểm tra: `extras.draft.*`, `extras.noPublic.size`, `urgency.*` (có nút Sửa). **Chưa làm / giới hạn**: đánh số trang riêng cho từng phụ lục (cần ngắt phần và `pgNumType`, rủi ro làm hỏng thiết lập trang nên để người dùng đặt tay, giao diện có hướng dẫn); chỉ dẫn phạm vi lưu hành chèn tại con trỏ chứ chưa tự đặt cách mép trên 10 mm; dòng "Không phổ biến" và thông tin liên hệ chèn vào cuối tài liệu chứ chưa vào chân trang cuối. Tôi cũng chưa có nguyên văn mục 6B/Phụ lục 1 ô 10–13 nên cỡ chữ lấy từ `ELEMENTS` của chuẩn.
**File dự kiến:** `src/standards/hd05/extras.ts`, `src/standards/hd05/appendix.ts`, `src/ui/InsertPanel.ts`, `src/standards/hd05/extras.test.ts`
**Kích cỡ:** M

### Checkpoint 5 – Phủ loại văn bản
- [ ] Đối chiếu `docTypes` của từng chuẩn với Điều 7 NĐ 30 và Quy định 399-QĐ/TW: mỗi loại hoặc có mẫu đúng thể thức hoặc ghi rõ lý do chưa hỗ trợ
- [ ] Mỗi họ mẫu có ít nhất một văn bản mẫu được duyệt trực quan
- [x] Test, build, lint xanh


**Trạng thái (tự động):** `registry.test.ts` bảo đảm mọi loại trong `docTypes` của cả hai chuẩn có họ mẫu hoặc cờ "chưa hỗ trợ" kèm lý do; hiện **cả hai chuẩn không còn loại nào chưa hỗ trợ** (Chính quyền: 29 loại, Đảng: 33 loại + văn bản đại hội + liên cơ quan). Ba loại của Chính quyền không có mẫu ở Phụ lục III (hợp đồng, bản ghi nhớ, bản thỏa thuận) và thư công là bố cục đề xuất; giấy mời, giấy giới thiệu, công điện, giấy nghỉ phép và toàn bộ giấy tờ của Đảng là bố cục theo cách trình bày thông dụng, chưa đối chiếu nguyên văn. **Còn chờ bạn**: duyệt trực quan từng họ mẫu (tạo mẫu trong Word, in PDF, đối chiếu Phụ lục III và HD 05) và đối chiếu danh sách loại với Điều 7 NĐ 30 và Quy định 399-QĐ/TW.

---

## Giai đoạn 6 – Hoàn thiện và phát hành (dùng cá nhân)

### T22: Danh sách cơ quan đã lưu (theo chuẩn), profile và xuất/nhập cấu hình
**Mô tả:** Quản lý danh sách cơ quan (thêm, sửa, xoá, đặt mặc định) cho ô chọn ở T7 và TD1; **mỗi mục thuộc một chuẩn** (chuẩn Đảng: cơ quan cấp trên trực tiếp 0–2 dòng và cơ quan ban hành, ký hiệu cơ quan đảng tra từ T2B). Cài đặt chung thêm "Chuẩn mặc định cho văn bản mới". Mỗi mục gồm các dòng cơ quan chủ quản (0–3), tên cơ quan ban hành, địa danh, ký hiệu mặc định (ví dụ "CCTL-ĐĐPCTT"), chức vụ và họ tên người ký thường dùng. Dùng chung kho lưu trữ cục bộ với preset ở T6A; có xuất/nhập file cấu hình (gồm cả preset) để chuyển sang máy khác.
**Tiêu chí nghiệm thu:**
- [x] Danh sách cơ quan hiện trong ô chọn của T7 và wizard; chọn một mục thì điền sẵn cả cơ quan chủ quản lẫn cơ quan ban hành
- [x] Mở wizard thì điền sẵn từ profile
- [x] Giá trị ngoài khoảng NĐ 30 không lưu được kèm giải thích
- [x] Xuất rồi nhập trên máy khác giữ nguyên profile và preset
**Kiểm chứng:**
- [x] `npm test`: lưu/đọc/kiểm tra/nhập xuất cấu hình, phiên bản định dạng file
- [ ] Thủ công: đóng mở lại Word, giá trị còn
**Phụ thuộc:** T6A, T6B, T12, TD1
**Trạng thái:** mã và test xong (`src/settings/profile.ts`, `src/settings/exportImport.ts`, `src/ui/SettingsPanel.ts`; `HeaderForm.applyAgency`, `FooterForm.applyAgency`; 27 ca ở `profile.test.ts`; 1091 test cả dự án); `typecheck`, `lint`, `build` sạch; **chưa thử trong Word**. Bảng "Hồ sơ cơ quan và cấu hình": mỗi chuẩn một danh sách hồ sơ (tên gợi nhớ, 0–3 dòng chủ quản / 0–2 dòng cấp trên của Đảng, cơ quan ban hành 1 dòng / 1–3 dòng của Đảng, địa danh, chữ viết tắt cơ quan, chức vụ, họ tên người ký) với thêm, sửa, xoá, đặt mặc định, "Dùng" (điền Khối đầu và Khối cuối). Hồ sơ mặc định tự điền khi mở add-in và khi đổi chuẩn. Giá trị ngoài giới hạn của chuẩn (số dòng, chữ viết tắt có khoảng trắng hay dấu phân cách, chức vụ quá 3 dòng, trùng tên) không lưu được và có giải thích tiếng Việt. Xuất/nhập: một file JSON có `format`, `version` (hiện 1), gồm hồ sơ, preset, mặc định từng chuẩn, danh sách tên và viết tắt cơ quan, chuẩn mặc định; nhập gộp vào cấu hình sẵn có, file sai/phiên bản mới hơn bị từ chối không ghi gì, mục sai bị bỏ qua và báo lại. "Chuẩn mặc định cho văn bản mới" dùng lựa chọn gần nhất đã có (`defaultStandard`) và được xuất/nhập cùng. **Giới hạn**: xuất ra ô văn bản kèm sao chép và liên kết tải về (chưa thử xem tải về có chạy trong Word không); ô chọn cơ quan cũ (`listAgencies`) vẫn là danh sách riêng chứ chưa hợp nhất với hồ sơ.
**File dự kiến:** `src/settings/profile.ts`, `src/settings/exportImport.ts`, `src/ui/SettingsPanel.ts`, `src/settings/profile.test.ts`
**Kích cỡ:** S

### T23: Kiểm thử trên Word 365, trợ năng, xử lý lỗi
**Mô tả:** Chạy kiểm thử thủ công theo ma trận tính năng × chuẩn (Chính quyền, Đảng) trên Word 365 (Windows; thử Mac/Online nếu tiện); thêm phát hiện tính năng (`isSetSupported`) và thông báo/fallback khi API không có; kiểm tra bàn phím, tương phản; rà soát chuỗi tiếng Việt.
**Tiêu chí nghiệm thu:**
- [ ] Mọi tính năng hoặc chạy đúng hoặc hiện thông báo rõ khi môi trường không hỗ trợ, không bao giờ lỗi im lặng
- [ ] Thao tác chính dùng được chỉ bằng bàn phím
- [ ] Biên bản kiểm thử lưu ở `docs/test-matrix.md`
**Kiểm chứng:**
- [x] `npm test`, `npm run build`, `npm run lint` xanh
- [ ] Thủ công theo ma trận tính năng
**Phụ thuộc:** T15, T21, T22, TD8, và các task T16–T20, TD6–TD7 đã làm
**Trạng thái:** phần làm được bằng mã đã xong (`src/office/capabilities.ts` + `src/ui/EnvironmentPanel.ts`, `src/ui/a11y.ts`, sửa CSS, bộ bắt lỗi toàn cục, `docs/test-matrix.md`); `typecheck`, `lint`, `build` sạch; **thử thủ công trong Word chưa làm được nên hai tiêu chí cuối còn mở**. Việc đã làm: (1) bảng tính năng × bộ API Word (`FEATURES`) có test quét mã nguồn để mọi `isSetSupported` trong `src/` đều có trong bảng; mục "Môi trường Word" ở đầu add-in liệt kê tính năng nào chạy được, chạy một phần hay bị tắt kèm bản Word cần; nút chèn khối cuối nay cũng tự tắt kèm thông báo khi thiếu WordApi 1.3 (các bảng khác đã có sẵn); (2) trợ năng: màu chữ phụ `#777` (4,48:1, không đạt) đổi thành `#595959`, test tính tương phản từ CSS (chữ ≥ 4,5:1 trên nền trắng và nền nhạt), khung `:focus-visible` cho mọi điều khiển, bộ quan sát gắn `aria-label` (từ placeholder, title) cho ô nhập chưa có tên, trang khai `lang="vi"`; (3) lỗi không ai bắt (`error`, `unhandledrejection`) hiện ở dòng trạng thái thay vì im lặng; (4) `docs/test-matrix.md`: ma trận tính năng × chuẩn, cột "Tự động" đã điền theo test hiện có, cột "Word thật" để bạn điền (mọi ô đang chưa thử). **Chưa làm / cần bạn**: chạy ma trận trong Word 365 (Windows; Mac/Online nếu tiện), thử toàn bộ bằng bàn phím và Narrator, rà chính tả tiếng Việt toàn giao diện.
**File dự kiến:** `src/office/capabilities.ts`, `src/ui/*`, `docs/test-matrix.md`
**Kích cỡ:** M

### T24: Đóng gói, GitHub Pages, hướng dẫn
**Mô tả:** Build bản phát hành, đưa lên **GitHub Pages** (host tĩnh miễn phí, dự phòng Cloudflare Pages) bằng quy trình tự động (GitHub Actions), nạp manifest qua shared-folder catalog trong Word, viết hướng dẫn cài đặt/sử dụng tiếng Việt, tuyên bố riêng tư (không gửi nội dung văn bản ra ngoài, không telemetry). Trang công khai nên kiểm tra repo và bản build không chứa file mẫu thật hay thông tin cá nhân.
**Tiêu chí nghiệm thu:**
- [ ] Cài được theo hướng dẫn trên Word 365 sạch, dùng được hằng ngày
- [ ] Xác nhận bằng giám sát mạng: không có request chứa nội dung văn bản
- [x] Có số phiên bản, cách cập nhật và gỡ cài đặt
**Kiểm chứng:**
- [x] `npm run build` bản production; kiểm tra manifest hợp lệ
- [ ] Cài thử trên hồ sơ Word sạch theo đúng hướng dẫn
**Phụ thuộc:** T23
**Trạng thái:** mã, workflow và tài liệu xong (`tools/make-manifest.js`, `tools/check-release.js`, `.github/workflows/pages.yml`, `webpack.config.js` có chế độ `--env release` bỏ trang thăm dò spike và ghi số phiên bản, `README.md`, `docs/huong-dan-cai-dat.md`, `docs/privacy.md`, `src/privacy.test.ts`); `typecheck`, `lint`, `build` sạch; manifest phát hành `dist/manifest.xml` qua `office-addin-manifest validate` báo hợp lệ (lần đầu báo sai vì phiên bản 0.9.0 dưới 1.0, đã sửa); chạy thử cục bộ `build:release` → `manifest:release` → `check:release` đạt (dist chỉ gồm `taskpane.html`, `taskpane.js`, `manifest.xml`, `assets`). Số phiên bản `1.0.0` (package.json → manifest `1.0.0.0`, hiện ở dòng đầu add-in; Office không nhận manifest dưới 1.0). **Chưa làm / cần bạn**: (1) tạo repo GitHub và bật Pages (Settings > Pages > Source: GitHub Actions), đẩy `main`, lấy địa chỉ `https://<tài khoản>.github.io/<repo>/`; (2) cài thử trên hồ sơ Word sạch theo `docs/huong-dan-cai-dat.md` và báo bước nào sai; (3) giám sát mạng với văn bản thử (mục 8 của hướng dẫn) để xác nhận không có yêu cầu chứa nội dung; (4) tăng số phiên bản mỗi lần phát hành. Dự phòng Cloudflare Pages: cùng thư mục `dist`, cùng lệnh `manifest:release` với địa chỉ của Cloudflare.
**File dự kiến:** `manifest.xml`, `docs/huong-dan-cai-dat.md`, `docs/privacy.md`, `README.md`
**Kích cỡ:** S

### T25: Giao diện TTA-Format (tab, mục xổ ra, tên mới)
**Mô tả:** Đổi tên add-in thành **TTA-Format** và làm lại giao diện: bốn tab riêng **Nghị định 30**, **Hướng dẫn 05**, **Kiểm tra**, **Cơ quan**; các mục có nhiều tuỳ chọn gom lại, bấm vào mới xổ ra (mục xổ ra nhớ trạng thái mở); phong cách hiện đại (thẻ bo góc, màu nhấn, nút chính/phụ), vẫn đạt tương phản và dùng được bằng bàn phím.
**Tiêu chí nghiệm thu:**
- [x] Tên TTA-Format ở manifest (tên hiển thị, thẻ ribbon, nút), tiêu đề trang, tài liệu
- [x] Chọn tab Nghị định 30 hoặc Hướng dẫn 05 thì đổi chuẩn đang dùng (thay ô chọn chuẩn cũ); tab Kiểm tra gồm kiểm tra văn bản, sửa, chỉnh toàn bộ; tab Cơ quan gồm hồ sơ cơ quan, xuất nhập cấu hình, môi trường
- [x] Mỗi nhóm tuỳ chọn nằm trong mục xổ ra; trạng thái mở nhớ theo máy
- [x] Tab dùng được bằng bàn phím (mũi tên, Home, End) và có ARIA đúng; tương phản chữ ≥ 4,5:1
- [ ] Mọi chức năng cũ vẫn có đủ trong giao diện mới
**Kiểm chứng:**
- [x] `npm test`: logic tab và mục xổ ra, tương phản từ biến màu, tên add-in
- [x] Xem trước trên trình duyệt với Office giả (`preview.html`): chụp màn hình bốn tab
- [ ] Thủ công trong Word
**Phụ thuộc:** T22, T23
**Trạng thái:** mã, test và xem trước xong; `typecheck`, `lint`, 1153 test, `build:release` (manifest hợp lệ, `check:release` đạt) sạch; **chưa thử trong Word**. Đã làm: đổi tên TTA-Format (manifest, thẻ ribbon, nút, tiêu đề trang, tài liệu); `src/ui/Tabs.ts` (tablist ARIA, mũi tên/Home/End), `src/ui/Accordion.ts` (mục xổ ra nhớ trạng thái, logic ở `tabsLogic.ts`), `src/taskpane/app.ts` dựng bốn tab, `taskpane.css` (biến màu, thẻ bo góc, nút chính/phụ), `taskpane.html` còn khung tĩnh; ô chọn chuẩn cũ bỏ, thay bằng tab chuẩn (`StandardSwitcher.choose`, vẫn hỏi xác nhận khi tài liệu đã có nội dung theo chuẩn khác); tab Kiểm tra có chọn chuẩn nhanh; nhóm xổ ra trong Khối đầu, Khối cuối, Chèn dấu, Xuất nhập cấu hình; hồ sơ cơ quan có nhãn hiện rõ. Test tương phản nay đọc biến màu từ CSS. Có trang xem trước `preview.html` (Office giả, chỉ ở bản phát triển) và đã chụp, bấm thử cả bốn tab, không lỗi console. **Chưa kiểm**: hành vi trong Word thật (nhất là hộp xác nhận khi đổi chuẩn từ tab), giao diện ở độ rộng bảng khác 380 px, chế độ tương phản cao của Windows, đọc bằng Narrator.
**File dự kiến:** `src/ui/Tabs.ts`, `src/ui/Accordion.ts`, `src/taskpane/taskpane.css`, `src/taskpane/taskpane.html`, `src/taskpane/taskpane.ts`, `src/ui/StandardSwitcher.ts`, `manifest.xml`
**Kích cỡ:** M

### T26: Đề xuất bổ sung sau kiểm tra toàn bộ
**Mô tả:** thực hiện các đề xuất thêm bớt sau đợt kiểm tra lại toàn bộ.
**Tiêu chí nghiệm thu:**
- [x] Gộp "Tuỳ chọn nhanh", "Trang, style và số trang", "Định dạng nội dung" thành mục "Định dạng"; nút chèn khối đầu/cuối lẻ vào nhóm "Nâng cao"; bỏ trang thăm dò API (spike)
- [x] Giấy tờ ít dùng của chuẩn Đảng (giấy chứng nhận, giấy đi đường, phiếu chuyển, phiếu gửi) ẩn khỏi ô chọn loại, vẫn có cấu hình
- [x] Sổ số văn bản: gợi ý số tiếp theo theo chuẩn, năm, loại; ghi số khi tạo văn bản; sửa số trong Cơ quan; xuất, nhập cùng cấu hình
- [x] Danh bạ nơi nhận theo chuẩn, chọn từ danh sách khi nhập Kính gửi và Nơi nhận
- [x] Tự dựng sẵn Khoản và Điểm trong các Điều của quyết định, nghị quyết, quy chế, quy định
- [x] Khối đầu, khối cuối còn trống vẫn tạo được văn bản theo mẫu, mục thiếu điền chữ giả định (`src/templates/assumptions.ts`)
- [x] Kiểm tra sửa luôn khối đầu, khối cuối: thêm dấu * dưới tên cơ quan (Đảng), dòng 5 gạch nối, gạch đầu dòng nơi nhận, họ tên viết hoa chữ cái đầu, bỏ học hàm, học vị (cách sửa mới `insertAfter`)
- [x] Nút "Sửa nhanh" ở tab Kiểm tra: chạy Chỉnh toàn bộ rồi Sửa tất cả, một lần chụp để hoàn tác
- [x] Xuất báo cáo kiểm tra ra .docx (zip không nén tự dựng, mở bằng `createDocument`), vẫn giữ .txt
- [x] Bản sao y, trích sao, sao lục (SY, TrS, SL) cho chuẩn Chính quyền, họ mẫu giấy tờ
- [ ] Thủ công trong Word
**Trạng thái:** mã và test xong. **Đã thử trong Word thật (14/10/2026)**: tạo công văn bỏ trống (chữ giả định), quyết định có Khoản/Điểm (đánh số 1., a), b) đúng, khởi động lại mỗi Điều), bản sao y, Sửa nhanh + Hoàn tác (lề sai thành đúng rồi về lại), mở báo cáo .docx. Lỗi tìm thấy và đã sửa: quốc hiệu xuống dòng (Word bỏ thụt dòng/giãn đoạn 0 khi gói thiếu docDefaults của tài liệu đích; nay gói mang docDefaults), báo "đọc lại chưa đúng" số trang (Word bỏ cỡ chữ bằng mặc định), ô số Điều/Khoản/Điểm thiếu nhãn. Thử tiếp (14/10/2026): công văn và báo cáo chuẩn Đảng tạo được; xoá dấu * rồi Sửa thì dấu * được chèn lại đúng chỗ; xoá dòng 5 gạch nối rồi Sửa thì dòng được chèn lại, căn giữa (cách sửa `insertAfter` chạy đúng trong Word thật). **Chưa thử trong Word**: sửa họ tên, học hàm, gạch đầu dòng nơi nhận (chỉ có test giả lập), các họ mẫu quy chế, biên bản, hợp đồng, nghị quyết Đảng. **Lưu ý khi tự thử**: Word WebView2 đổi phiên bản thì phải cấp lại quyền xem `msedgewebview2.exe` cho công cụ điều khiển máy, nếu không pane trông như treo ở "Loading". **Lưu ý**: bố cục bản sao y, trích sao, sao lục là đề xuất (Phụ lục III chỉ nêu chữ viết tắt), cần bạn duyệt; xem `docs/nd30-ambiguities.md` mục 17.
**Kích cỡ:** L

- [ ] Mọi tiêu chí nghiệm thu đạt; ma trận kiểm thử hoàn thành
- [ ] Bạn dùng thử trên văn bản thật và xác nhận
- [x] Rà soát bảo mật và quyền riêng tư xong (14/10/2026: không `innerHTML`, `eval`, lời gọi mạng; quyền chỉ `ReadWriteDocument`; `npm audit --omit=dev` 0 lỗ hổng; nhập cấu hình kiểm kiểu từng trường; xuất .docx thoát ký tự XML)
- [ ] Sẵn sàng dùng hằng ngày
