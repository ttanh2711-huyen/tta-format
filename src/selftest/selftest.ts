import { runFixAll, checkDocument } from "../office/fixAll";
import { createLetter, type LetterContext } from "../office/createDocument";
import { appendAppendix, countAppendices, insertAtSelection } from "../office/insertExtras";
import { restoreSnapshot, takeSnapshot } from "../office/snapshot";
import { assembleDocumentBody, type AssembleInput } from "../ooxml/documentAssembler";
import { appendixNumberingXml, buildAppendix } from "../ooxml/appendix";
import { QUOC_HIEU, headerInputFromOptions } from "../ooxml/headerBlock";
import { footerInputFromOptions } from "../ooxml/footerBlock";
import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { multiSignFromOptions } from "../ooxml/multiSignBlock";
import { buildUrgencyMark } from "../ooxml/urgencyMark";
import { recipientsRowOf, subjectLine, titleInputFromOptions } from "../ooxml/titleBlock";
import { buildFormatValuesForTest } from "./values";
import { assemblePartyDocumentBody } from "../standards/hd05/assembleDocument";
import { partyFooterFromOptions } from "../standards/hd05/footerBlock";
import { partyHeaderFromOptions } from "../standards/hd05/headerBlock";
import { assemblePartyCongressBody, assemblePartyFormBody, assemblePartyJointBody, assemblePartyMinutesBody } from "../standards/hd05/templates/documents";
import { assemblePartyRegulationBody, partyRegulationFromOptions, partyRegulationNumberingXml } from "../standards/hd05/templates/regulations";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { partyRecipientsRow, partySubjectRow, partyTitleFromOptions } from "../standards/hd05/titleBlock";
import { assembleMinutesBody, defaultMinutesParties } from "../templates/bienBan";
import { agreementNumberingXml, assembleAgreementBody, defaultAgreementParties } from "../templates/hopDong";
import { assembleDecisionBody, decisionInputFromOptions, decisionNumberingXml, type DecisionKind } from "../templates/quyetDinh";
import { assembleRegulationBody, regulationInputFromOptions, regulationNumberingXml } from "../templates/quyChe";
import { assembleSmallFormBody, findSmallForm } from "../templates/smallForms";
import type { StandardProfile } from "../standards/types";

interface Result {
  id: string;
  ok: boolean;
  detail: string;
  /** Mã quy tắc còn báo sau khi tạo (kiểm tra lại bằng bộ kiểm tra thật). */
  issues?: string[];
  lists?: string[];
}

const results: Result[] = [];
const out = document.getElementById("out") as HTMLElement;

function log(result: Result): void {
  results.push(result);
  const line = document.createElement("div");
  line.className = result.ok ? "pass" : "fail";
  line.textContent = `${result.ok ? "ĐẠT" : "LỖI"} ${result.id}: ${result.detail}${result.issues?.length ? ` | còn báo: ${result.issues.join(",")}` : ""}`;
  out.append(line);
}

async function resetDoc(): Promise<void> {
  await Word.run(async (context) => {
    context.document.body.clear();
    await context.sync();
  });
}

/** Chuỗi nhãn đánh số tự động của các đoạn danh sách (đọc từ Word thật). */
async function listLabels(): Promise<string[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/isListItem");
    await context.sync();
    const items = paragraphs.items.filter((p) => p.isListItem);
    items.forEach((p) => p.listItem.load("listString,level"));
    await context.sync();
    return items.map((p) => p.listItem.listString);
  });
}

async function bodyText(): Promise<string> {
  return Word.run(async (context) => {
    context.document.body.load("text");
    await context.sync();
    return context.document.body.text;
  });
}

function baseContext(profile: StandardProfile, values: Record<string, number | string | boolean>): Pick<LetterContext, "profile" | "values" | "specs" | "localFormatting"> {
  return { profile, values, specs: buildStyleSpecs(profile), localFormatting: "clear" };
}

async function scenario(id: string, profile: StandardProfile, parts: Pick<LetterContext, "buildBody" | "headText" | "signerName" | "numberingXml" | "markerStyle">, expectLabels?: string[]): Promise<void> {
  try {
    await resetDoc();
    const values = buildFormatValuesForTest(profile);
    const report = await createLetter({ ...baseContext(profile, values), ...parts });
    const failed = report.steps.filter((s) => s.state === "failed" || s.kind === "error").map((s) => `${s.label}: ${s.text}`);
    const warned = report.steps.filter((s) => s.kind === "warning").map((s) => `${s.label}: ${s.text}`);
    const issues = (await checkDocument(profile)).map((i) => `${i.ruleId}`);
    const labels = await listLabels();
    const missingLabels = (expectLabels ?? []).filter((expected) => !labels.includes(expected));
    const ok = report.ok && failed.length === 0 && missingLabels.length === 0;
    log({ id, ok, issues, lists: labels.slice(0, 12), detail: [report.ok ? "tạo xong" : "tạo lỗi", ...failed, ...warned.map((w) => `cảnh báo ${w}`), ...(missingLabels.length ? [`thiếu nhãn ${missingLabels.join("|")}`] : [])].join("; ") });
  } catch (error) {
    log({ id, ok: false, detail: `ngoại lệ ${error instanceof Error ? error.stack ?? error.message : String(error)}` });
  }
}

const nd30Values = () => buildFormatValuesForTest(ND30);
const partyValues = () => buildFormatValuesForTest(HD05);

function nd30Assemble(variant: "letter" | "named", typeName: string): AssembleInput {
  const profile = ND30;
  const values = nd30Values();
  const titleInput = titleInputFromOptions(profile, values, { variant, typeName, subject: variant === "letter" ? "cử công chức tham dự hội nghị tập huấn" : "Về việc triển khai nhiệm vụ", recipients: ["Sở Nông nghiệp và Môi trường"] });
  return {
    header: headerInputFromOptions(profile, values, {
      parentLines: ["Sở Nông nghiệp và Môi trường"],
      issuer: "Chi cục Thủy lợi",
      numberRow: { so: "Số: 05/CCTL-VP", placeDate: "Cần Thơ, ngày 05 tháng 01 năm 2026" },
      subjectRow: variant === "letter" ? { text: subjectLine(titleInput.subject) } : undefined,
      recipientsRow: recipientsRowOf(titleInput),
    }),
    title: titleInput,
    footer: footerInputFromOptions(profile, values, { recipients: ["Như trên", "Lưu"], luu: "VT", authority: "", positionLines: ["Chi cục trưởng"], signerName: "Mai Hiếu Hiền" }),
  };
}

function partyAssemble(variant: "congVan" | "coTenLoai" | "toTrinh", typeName: string) {
  const profile = HD05;
  const values = partyValues();
  const title = partyTitleFromOptions(profile, values, { variant, typeName, subject: variant === "congVan" ? "cử cán bộ dự hội nghị" : "kết quả hội nghị", recipients: ["Ban Thường vụ Tỉnh ủy"], basis: variant === "congVan" ? [] : ["Căn cứ Điều lệ Đảng"] });
  const header = partyHeaderFromOptions(profile, values, {
    parentLines: ["Tỉnh ủy Đồng Tháp"],
    issuerLines: ["Ban Tổ chức"],
    numberRow: { so: "Số 05-BC/BTCTU", placeDate: "Đồng Tháp, ngày 03 tháng 02 năm 2026" },
    subjectRow: partySubjectRow(title),
    recipientsRow: partyRecipientsRow(title),
  });
  const footer = partyFooterFromOptions(profile, values, { recipients: ["Như trên"], luu: "VT", drafter: "", authority: "K/T", positionLines: ["Trưởng ban", "Phó Trưởng ban"], signerName: "Nguyễn Văn A", confirmation: "" });
  return { header, title, footer };
}

const decision = (kind: DecisionKind) => decisionInputFromOptions(ND30, nd30Values(), { kind, authority: "Chủ tịch Ủy ban nhân dân thành phố", basis: ["Căn cứ Luật Tổ chức chính quyền địa phương", "Xét đề nghị của Giám đốc Sở Nội vụ"], articleCount: 3 });

async function main(): Promise<void> {
  const nd = (variant: "letter" | "named", typeName = "Thông báo") => nd30Assemble(variant, typeName);
  const parts = (input: AssembleInput) => ({ headText: QUOC_HIEU, signerName: "Mai Hiếu Hiền", buildBody: (body: Parameters<typeof assembleDocumentBody>[1]) => assembleDocumentBody(input, body) });

  await scenario("nd30.congVan", ND30, parts(nd("letter")));
  await scenario("nd30.titled", ND30, parts(nd("named")));
  for (const kind of ["quyetDinhTrucTiep", "nghiQuyet", "quyetDinhGianTiep"] as const) {
    const d = decision(kind);
    await scenario(`nd30.${kind}`, ND30, { headText: QUOC_HIEU, signerName: "Mai Hiếu Hiền", buildBody: () => assembleDecisionBody(nd("named", kind === "nghiQuyet" ? "Nghị quyết" : "Quyết định"), d), numberingXml: decisionNumberingXml(d) }, ["Điều 1.", "Điều 2.", "Điều 3."]);
  }
  const reg = regulationInputFromOptions(ND30, nd30Values(), { chapters: 2, articlesPerChapter: 2, withSections: true, attachedLine: "Ban hành kèm theo Quyết định số 05/QĐ-VP" });
  await scenario("nd30.quyChe", ND30, { headText: QUOC_HIEU, signerName: "Mai Hiếu Hiền", buildBody: () => assembleRegulationBody(nd("named", "Quy chế"), reg), numberingXml: regulationNumberingXml(reg), markerStyle: "ND30_TieuDePhanChuong" }, ["Chương I", "Chương II", "Mục 1", "Điều 1.", "Điều 4."]);
  const sign = multiSignFromOptions(ND30, nd30Values(), { parties: defaultMinutesParties(), recipients: ["Như trên"], luu: "" });
  await scenario("nd30.bienBan", ND30, { headText: QUOC_HIEU, signerName: "", buildBody: () => assembleMinutesBody(nd("named", "Biên bản"), "ND30_", sign) });
  const agreement = { parties: 2, articleCount: 3, stylePrefix: "ND30_", firstLineIndentMm: 12.7 };
  const signAgreement = multiSignFromOptions(ND30, nd30Values(), { parties: defaultAgreementParties(2), recipients: [], luu: "" });
  await scenario("nd30.hopDong", ND30, { headText: QUOC_HIEU, signerName: "", buildBody: () => assembleAgreementBody(nd("named", "Hợp đồng"), agreement, signAgreement), numberingXml: agreementNumberingXml(agreement) }, ["Điều 1.", "Điều 3."]);
  const giayMoi = findSmallForm("giay-moi");
  if (giayMoi) await scenario("nd30.giayMoi", ND30, { headText: QUOC_HIEU, signerName: "Mai Hiếu Hiền", buildBody: () => assembleSmallFormBody(nd("named", "Giấy mời"), giayMoi, "ND30_") });

  const pv = partyValues();
  const pparts = (variant: "congVan" | "coTenLoai" | "toTrinh", typeName: string) => {
    const a = partyAssemble(variant, typeName);
    return { headText: HD05.identifyingTitle, signerName: "Nguyễn Văn A", buildBody: (body: Parameters<typeof assemblePartyDocumentBody>[1]) => assemblePartyDocumentBody({ header: a.header, title: a.title, footer: a.footer }, body) };
  };
  await scenario("hd05.congVan", HD05, pparts("congVan", ""));
  await scenario("hd05.titled", HD05, pparts("coTenLoai", "Báo cáo"));
  await scenario("hd05.toTrinh", HD05, pparts("toTrinh", "Tờ trình"));
  for (const [top, numbering] of [["Chương", "roman"], ["Phần", "words"]] as const) {
    const a = partyAssemble("coTenLoai", "Quy chế");
    const r = partyRegulationFromOptions(HD05, pv, { chapters: 2, articlesPerChapter: 2, withSections: true, top, numbering });
    await scenario(`hd05.quyChe.${top}.${numbering}`, HD05, { headText: HD05.identifyingTitle, signerName: "Nguyễn Văn A", buildBody: () => assemblePartyRegulationBody({ header: a.header, title: a.title, footer: a.footer }, r), numberingXml: partyRegulationNumberingXml(r), markerStyle: "HD05_TenPhanChuong" }, [numbering === "roman" ? "Chương I" : "Phần thứ Một", "Điều 1."]);
  }
  const pa = partyAssemble("coTenLoai", "Biên bản");
  const psign = multiSignFromOptions(HD05, pv, { parties: [{ title: "Người ghi biên bản", positionLines: [], signerName: "Lê Thị C", confirmation: "Trần Văn B, Bí thư" }, { title: "Chủ trì hội nghị", positionLines: ["Bí thư"], signerName: "Nguyễn Văn A" }], recipients: ["Như trên"], luu: "" });
  await scenario("hd05.bienBan", HD05, { headText: HD05.identifyingTitle, signerName: "", buildBody: () => assemblePartyMinutesBody({ header: pa.header, title: pa.title }, "HD05_", psign) });
  const pc = partyAssemble("coTenLoai", "Nghị quyết");
  await scenario("hd05.daiHoi", HD05, { headText: HD05.identifyingTitle, signerName: "Nguyễn Văn A", buildBody: (body) => assemblePartyCongressBody({ header: pc.header, title: pc.title, footer: pc.footer }, "Đại hội đại biểu Đảng bộ tỉnh lần thứ XI", 14, body) });
  const pj = partyAssemble("coTenLoai", "Quy chế");
  const jsign = multiSignFromOptions(HD05, pv, { parties: [{ title: "T/M Ban Thường vụ", positionLines: ["Bí thư"], signerName: "Nguyễn Văn A" }, { title: "T/M Ủy ban Kiểm tra", positionLines: ["Chủ nhiệm"], signerName: "Lê Văn C" }], recipients: [], luu: "" });
  await scenario("hd05.lienCoQuan", HD05, { headText: HD05.identifyingTitle, signerName: "", buildBody: (body) => assemblePartyJointBody({ header: pj.header, title: pj.title }, jsign, body) });
  const pf = partyAssemble("coTenLoai", "Giấy chứng nhận");
  const certificate = findSmallForm("giay-chung-nhan");
  if (certificate) await scenario("hd05.giayChungNhan", HD05, { headText: HD05.identifyingTitle, signerName: "Nguyễn Văn A", buildBody: () => assemblePartyFormBody({ header: pf.header, title: pf.title, footer: pf.footer }, certificate, "HD05_") });

  // Phần chèn thêm, sửa, hoàn tác chạy trên công văn Chính quyền.
  await resetDoc();
  await createLetter({ ...baseContext(ND30, nd30Values()), ...parts(nd("letter")) });
  try {
    const mark = ND30.urgencyMarks[0];
    await Word.run(async (context) => {
      const first = context.document.body.paragraphs.getLast();
      first.select("End");
      await context.sync();
    });
    const res = await insertAtSelection(buildUrgencyMark({ text: mark.text, widthMm: mark.widthMm, heightMm: mark.heightMm, sizePt: 13 }), mark.text);
    log({ id: "extras.dauKhan", ok: res.verified, detail: res.note });
    const before = await countAppendices("ND30_PhuLucSo");
    const a1 = await appendAppendix({ count: () => countAppendices("ND30_PhuLucSo"), matches: (p) => p.style === "ND30_PhuLucTieuDe" && p.text.trim() === "DANH SÁCH A", build: (n) => ({ bodyXml: buildAppendix({ title: "Danh sách A", reference: "Kèm theo công văn", number: n, stylePrefix: "ND30_" }), numberingXml: appendixNumberingXml(n) }) });
    const a2 = await appendAppendix({ count: () => countAppendices("ND30_PhuLucSo"), matches: (p) => p.style === "ND30_PhuLucTieuDe" && p.text.trim() === "DANH SÁCH B", build: (n) => ({ bodyXml: buildAppendix({ title: "Danh sách B", reference: "", number: n, stylePrefix: "ND30_" }), numberingXml: appendixNumberingXml(n) }) });
    const labels = (await listLabels()).filter((l) => l.startsWith("Phụ lục"));
    log({ id: "extras.phuLuc", ok: a1.verified && a2.verified && labels.join("|") === "Phụ lục I|Phụ lục II", detail: `trước=${before}, nhãn=${labels.join("|")}`, lists: labels });
  } catch (error) {
    log({ id: "extras.error", ok: false, detail: error instanceof Error ? error.message : String(error) });
  }

  try {
    // Làm sai có chủ ý rồi sửa tất cả và hoàn tác.
    await resetDoc();
    await createLetter({ ...baseContext(ND30, nd30Values()), ...parts(nd("letter")) });
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      const p = tables.items[0].getRange().paragraphs.getFirst();
      p.font.bold = false;
      p.font.size = 20;
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      sections.items[0].pageSetup.leftMargin = 10;
      await context.sync();
    });
    const broken = (await checkDocument(ND30)).map((i) => i.ruleId);
    const snap = await takeSnapshot();
    const fixed = await runFixAll(ND30);
    const after = fixed.remaining.map((i) => i.ruleId);
    log({ id: "sua.suaTatCa", ok: broken.length > 0 && fixed.applied > 0 && !fixed.remaining.some((i) => i.fix), detail: `trước=[${broken.join(",")}] đã sửa ${fixed.applied} trong ${fixed.rounds} lượt, còn=[${after.join(",")}]`, issues: after });
    const restored = await restoreSnapshot(snap);
    const again = (await checkDocument(ND30)).map((i) => i.ruleId);
    log({ id: "sua.hoanTac", ok: restored.problems.length === 0 && again.length >= broken.length - 1, detail: `vấn đề=${restored.problems.join(";")} | sau hoàn tác còn=[${again.join(",")}]` });
  } catch (error) {
    log({ id: "sua.error", ok: false, detail: error instanceof Error ? (error.stack ?? error.message) : String(error) });
  }

  const text = await bodyText();
  log({ id: "ket-thuc", ok: true, detail: `độ dài thân cuối cùng ${text.length}` });
  const summary = results.filter((r) => !r.ok).length === 0 ? "TẤT CẢ ĐẠT" : `${results.filter((r) => !r.ok).length} LỖI`;
  const head = document.createElement("h2");
  head.textContent = summary;
  out.prepend(head);
  await fetch("/selftest-report", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ results }) });
}

Office.onReady(() => {
  void main().catch((error: unknown) => {
    const line = document.createElement("div");
    line.className = "fail";
    line.textContent = `Hỏng chung: ${error instanceof Error ? (error.stack ?? error.message) : String(error)}`;
    out.append(line);
    void fetch("/selftest-report", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ results, fatal: String(error) }) });
  });
});
