import { buildFooterBlock, footerInputFromOptions, validateFooterInput } from "../ooxml/footerBlock";
import { buildHeaderBlock, headerInputFromOptions, validateHeaderInput } from "../ooxml/headerBlock";
import { buildTitleBody, recipientsRowOf, subjectLine, titleInputFromOptions, validateTitleInput } from "../ooxml/titleBlock";
import { placeDateText, soText } from "../rules/validators";
import { buildFormatValuesForTest } from "../selftest/values";
import { buildPartyFooterBlock, partyFooterFromOptions, validatePartyFooter } from "../standards/hd05/footerBlock";
import { buildPartyHeaderBlock, partyHeaderFromOptions, validatePartyHeader } from "../standards/hd05/headerBlock";
import { HD05 } from "../standards/hd05/profile";
import { buildPartyTitleBody, partyRecipientsRow, partySubjectRow, partyTitleFromOptions, validatePartyTitle } from "../standards/hd05/titleBlock";
import { ND30 } from "../standards/nd30/profile";
import { assumeFooter, assumeHeader, PLACEHOLDERS, type FooterDraft, type HeaderDraft } from "./assumptions";

const blankHeader = (over: Partial<HeaderDraft> = {}): HeaderDraft => ({ issuer: "", place: "", agencyAbbrs: "", subject: "", recipients: "", typeName: "", needsRecipients: true, needsTypeName: false, withNumberRow: true, withTitle: true, ...over });
const blankFooter = (over: Partial<FooterDraft> = {}): FooterDraft => ({ recipients: "", positionLines: [""], signerName: "", ...over });

describe("giả định các mục còn thiếu của khối đầu", () => {
  it("không bật giả định thì giữ nguyên, không ghi gì", () => {
    const result = assumeHeader(blankHeader(), false);
    expect(result.draft).toEqual(blankHeader());
    expect(result.assumed).toEqual([]);
  });

  it("bật giả định: điền chữ giả định cho mọi mục bắt buộc còn trống và liệt kê lại", () => {
    const { draft, assumed } = assumeHeader(blankHeader({ needsTypeName: true }), true);
    expect(draft.issuer).toBe(PLACEHOLDERS.issuer);
    expect(draft.place).toBe(PLACEHOLDERS.place);
    expect(draft.agencyAbbrs).toBe(PLACEHOLDERS.agencyAbbr);
    expect(draft.subject).toBe(PLACEHOLDERS.subject);
    expect(draft.recipients).toBe(PLACEHOLDERS.recipient);
    expect(draft.typeName).toBe(PLACEHOLDERS.typeName);
    expect(assumed).toEqual(expect.arrayContaining(["cơ quan ban hành", "địa danh", "chữ viết tắt cơ quan", "trích yếu", "nơi nhận Kính gửi", "tên loại văn bản"]));
  });

  it("mục đã nhập thì giữ nguyên; mục không dùng (tắt hàng số, tắt tên loại) thì không giả định", () => {
    const { draft, assumed } = assumeHeader(blankHeader({ issuer: "Chi cục A", subject: "x", recipients: "Sở B", withNumberRow: false, withTitle: false }), true);
    expect(draft.issuer).toBe("Chi cục A");
    expect(assumed).toEqual([]);
    expect(draft.place).toBe("");
    expect(draft.subject).toBe("x");
  });

  it("văn bản không cần Kính gửi thì không giả định nơi nhận", () => {
    expect(assumeHeader(blankHeader({ needsRecipients: false }), true).assumed).not.toContain("nơi nhận Kính gửi");
  });

  it("nhiều dòng trống hoặc chỉ khoảng trắng cũng coi là thiếu", () => {
    expect(assumeHeader(blankHeader({ issuer: " \n  ", recipients: "\n\n" }), true).draft.issuer).toBe(PLACEHOLDERS.issuer);
  });
});

describe("giả định các mục còn thiếu của khối cuối", () => {
  it("điền chức vụ, họ tên khi trống; giữ nguyên khi đã nhập", () => {
    const filled = assumeFooter(blankFooter(), true);
    expect(filled.draft.positionLines).toEqual([PLACEHOLDERS.position]);
    expect(filled.draft.signerName).toBe(PLACEHOLDERS.signer);
    expect(filled.assumed).toEqual(expect.arrayContaining(["chức vụ người ký", "họ tên người ký"]));
    const kept = assumeFooter(blankFooter({ positionLines: ["Giám đốc", ""], signerName: "Nguyễn A" }), true);
    expect(kept.draft.positionLines).toEqual(["Giám đốc", ""]);
    expect(kept.assumed).toEqual([]);
  });

  it("không bật giả định thì không đổi", () => {
    expect(assumeFooter(blankFooter(), false)).toEqual({ draft: blankFooter(), assumed: [] });
  });
});

describe("văn bản dựng hoàn toàn từ giả định hợp lệ theo từng chuẩn", () => {
  it("Chính quyền: công văn và văn bản có tên loại, khối đầu và khối cuối dựng được", () => {
    const values = buildFormatValuesForTest(ND30);
    for (const variant of ["letter", "named"] as const) {
      const header = assumeHeader(blankHeader({ needsRecipients: variant === "letter", needsTypeName: variant === "named" }), true).draft;
      const footer = assumeFooter(blankFooter(), true).draft;
      const so = soText(ND30, { so: null, typeAbbr: null, agencyAbbrs: header.agencyAbbrs.split("-") });
      const date = placeDateText({ place: header.place, day: null, month: null, year: null });
      expect(so.ok && date.ok).toBe(true);
      if (!so.ok || !date.ok) return;
      const titleInput = titleInputFromOptions(ND30, values, { variant, typeName: header.typeName, subject: header.subject, recipients: header.recipients.split("\n") });
      expect(validateTitleInput(titleInput)).toEqual([]);
      const input = headerInputFromOptions(ND30, values, { parentLines: [], issuer: header.issuer, numberRow: { so: so.text, placeDate: date.text }, subjectRow: variant === "letter" ? { text: subjectLine(header.subject) } : undefined, recipientsRow: recipientsRowOf(titleInput) });
      expect(validateHeaderInput(input)).toEqual([]);
      expect(() => buildHeaderBlock(input)).not.toThrow();
      expect(() => buildTitleBody(titleInput)).not.toThrow();
      const footerInput = footerInputFromOptions(ND30, values, { recipients: ["Như trên", "Lưu"], luu: "VT", authority: "", positionLines: footer.positionLines, signerName: footer.signerName });
      expect(validateFooterInput(footerInput)).toEqual([]);
      expect(() => buildFooterBlock(footerInput)).not.toThrow();
    }
  });

  it("Đảng: công văn, thể loại và tờ trình, khối đầu và khối cuối dựng được", () => {
    const values = buildFormatValuesForTest(HD05);
    for (const variant of ["congVan", "coTenLoai", "toTrinh"] as const) {
      const header = assumeHeader(blankHeader({ needsRecipients: variant !== "coTenLoai", needsTypeName: variant !== "congVan" }), true).draft;
      const footer = assumeFooter(blankFooter(), true).draft;
      const so = soText(HD05, { so: null, typeAbbr: "BC", agencyAbbrs: header.agencyAbbrs.split("-") });
      const date = placeDateText({ place: header.place, day: null, month: null, year: null });
      expect(so.ok && date.ok).toBe(true);
      if (!so.ok || !date.ok) return;
      const title = partyTitleFromOptions(HD05, values, { variant, typeName: header.typeName, subject: header.subject, recipients: header.recipients.split("\n"), basis: [] });
      expect(validatePartyTitle(title)).toEqual([]);
      const input = partyHeaderFromOptions(HD05, values, { parentLines: [], issuerLines: header.issuer.split("\n"), numberRow: { so: so.text, placeDate: date.text }, subjectRow: partySubjectRow(title), recipientsRow: partyRecipientsRow(title) });
      expect(validatePartyHeader(input)).toEqual([]);
      expect(() => buildPartyHeaderBlock(input)).not.toThrow();
      expect(() => buildPartyTitleBody(title)).not.toThrow();
      const footerInput = partyFooterFromOptions(HD05, values, { recipients: ["Như trên"], luu: "VT", drafter: "", authority: "", positionLines: footer.positionLines, signerName: footer.signerName, confirmation: "" });
      expect(validatePartyFooter(footerInput)).toEqual([]);
      expect(() => buildPartyFooterBlock(footerInput)).not.toThrow();
    }
  });
});
