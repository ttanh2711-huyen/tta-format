import { insertSampleParagraph } from "../office/insertSample";
import { isDocumentStandardSupported, readDocumentSnapshot, writeDocumentNoAsk, writeDocumentStandard } from "../office/documentStandard";
import { browserStore } from "../settings/storage";
import { STANDARDS } from "../standards";
import type { StandardId } from "../standards/types";
import { createAccordion, type AccordionSpec } from "../ui/Accordion";
import { watchControls } from "../ui/a11y";
import { appVersion, mountEnvironmentPanel } from "../ui/EnvironmentPanel";
import { mountFixAllButton } from "../ui/FixAllButton";
import { mountTemplateMaker } from "../ui/TemplateMaker";
import { mountFooterForm } from "../ui/FooterForm";
import { mountFormatPanel } from "../ui/FormatPanel";
import { mountHeaderForm } from "../ui/HeaderForm";
import { mountInsertPanel } from "../ui/InsertPanel";
import { mountNewDocWizard } from "../ui/NewDocWizard";
import { mountPageSetupPanel } from "../ui/PageSetupPanel";
import { mountQuickOptions } from "../ui/QuickOptions";
import { mountReportPanel } from "../ui/ReportPanel";
import { mountSettingsPanel } from "../ui/SettingsPanel";
import { addEntry, entryFromDocument } from "../settings/soDangKy";
import { mountStandardSwitcher } from "../ui/StandardSwitcher";
import { mountTabs, tabButtonId, type TabSpec } from "../ui/Tabs";
import { loadAccordionOpen, loadActiveTab, saveAccordionOpen, saveActiveTab } from "../ui/tabsLogic";

const CHECK_TAB = "check";
const AGENCY_TAB = "agency";
const PANEL_STANDARD = "panel-standard";
const PANEL_CHECK = "panel-check";
const PANEL_AGENCY = "panel-agency";

export function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (status) status.textContent = message;
}

async function onInsertSample(): Promise<void> {
  try {
    await Word.run(async (context) => {
      insertSampleParagraph(context.document.body);
      await context.sync();
    });
    setStatus("Đã chèn đoạn văn bản thử nghiệm vào cuối tài liệu.");
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    setStatus(`Không chèn được văn bản: ${reason}`);
  }
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K, id: string, props: Partial<HTMLElementTagNameMap[K]> = {}): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  node.id = id;
  Object.assign(node, props);
  return node;
}

function need(id: string): HTMLElement {
  const found = document.getElementById(id);
  if (!found) throw new Error(`Thiếu vùng hiển thị "${id}" trong trang.`);
  return found;
}

/**
 * Dựng toàn bộ giao diện TTA-Format (T25): bốn tab. Hai tab chuẩn (Nghị định 30, Hướng dẫn 05) dùng chung một vùng làm việc
 * gồm các mục xổ ra; bấm tab nào thì chuẩn đang dùng đổi theo tab đó. Tab Kiểm tra gồm kiểm tra văn bản và chỉnh toàn bộ;
 * tab Cơ quan gồm hồ sơ cơ quan, xuất nhập cấu hình, môi trường và văn bản thử nghiệm.
 */
export function mountApp(): void {
  const store = browserStore();
  const noticeEl = need("notice");
  const panelStandard = need(PANEL_STANDARD);
  const panelCheck = need(PANEL_CHECK);
  const panelAgency = need(PANEL_AGENCY);
  need("app-version").textContent = `v${appVersion()}`;

  const acc = (parent: HTMLElement, spec: AccordionSpec, content: HTMLElement): void => {
    const handle = createAccordion(spec, store);
    handle.body.append(content);
    parent.append(handle.element);
  };

  // Vùng nội dung của từng bảng (giữ mã cũ để CSS và test tìm được).
  const wizardEl = element("div", "new-doc");
  const headerEl = element("div", "header-form");
  const footerEl = element("div", "footer-form");
  const quickEl = element("div", "quick-options");
  const pageEl = element("div", "page-setup-panel");
  const formatEl = element("div", "format-panel");
  const extrasEl = element("div", "insert-extras");
  const reportEl = element("div", "report");
  const fixAllEl = element("div", "fix-all");
  const templateEl = element("div", "template-maker");
  const settingsEl = element("div", "settings-panel");
  const environmentEl = element("div", "environment");
  const sampleEl = element("div", "sample");

  acc(panelStandard, { id: "ws.tao", title: "Tạo văn bản mới", hint: "Chọn loại văn bản rồi tạo", defaultOpen: true }, wizardEl);
  // Ngay dưới "Tạo văn bản mới" để dễ thấy (issue #2): tạo mẫu từ văn bản đang mở.
  acc(panelStandard, { id: "ws.mau", title: "Biến văn bản này thành mẫu", hint: "Thay các vùng thay đổi bằng ô điền, lưu thành .dotx", defaultOpen: true }, templateEl);
  acc(panelStandard, { id: "ws.khoi-dau", title: "Khối đầu văn bản", hint: "Cơ quan, số và ký hiệu, ngày tháng, trích yếu", defaultOpen: true }, headerEl);
  acc(panelStandard, { id: "ws.khoi-cuoi", title: "Khối cuối văn bản", hint: "Nơi nhận, quyền hạn, chức vụ, họ tên người ký" }, footerEl);
  // Một mục "Định dạng" gom ba bảng chồng nhau (T26): tuỳ chọn nhanh, trang/style/số trang, định dạng nội dung.
  const formatEls = element("div", "format-group");
  for (const [key, title, content] of [
    ["tuy-chon", "Tuỳ chọn nhanh (chữ, đoạn, trang, bảng, preset)", quickEl],
    ["trang", "Trang, style và số trang", pageEl],
    ["noi-dung", "Định dạng nội dung", formatEl],
  ] as const) {
    const details = document.createElement("details");
    details.id = `fg-${key}`;
    details.open = loadAccordionOpen(store, `ws.dinh-dang.${key}`, key === "tuy-chon");
    details.addEventListener("toggle", () => saveAccordionOpen(store, `ws.dinh-dang.${key}`, details.open));
    const summary = document.createElement("summary");
    summary.textContent = title;
    details.append(summary, content);
    formatEls.append(details);
  }
  acc(panelStandard, { id: "ws.dinh-dang", title: "Định dạng", hint: "Tuỳ chọn nhanh, trang và style, định dạng nội dung" }, formatEls);
  acc(panelStandard, { id: "ws.chen", title: "Dấu khẩn, phụ lục và thành phần bổ sung", hint: "Chèn tại con trỏ hoặc cuối tài liệu" }, extrasEl);

  const chipsEl = element("div", "check-standard", { className: "chips" });
  panelCheck.append(chipsEl);
  acc(panelCheck, { id: "ck.kiem-tra", title: "Kiểm tra văn bản", hint: "Liệt kê lỗi thể thức, sửa từng lỗi hoặc sửa tất cả", defaultOpen: true }, reportEl);
  acc(panelCheck, { id: "ck.chinh-toan-bo", title: "Chỉnh toàn bộ", hint: "Áp dụng nhiều bước có xem trước và khôi phục" }, fixAllEl);

  acc(panelAgency, { id: "ag.ho-so", title: "Hồ sơ cơ quan và cấu hình", hint: "Thêm, sửa, đặt mặc định; xuất và nhập cấu hình", defaultOpen: true }, settingsEl);
  const environment = mountEnvironmentPanel(environmentEl);
  acc(panelAgency, { id: "ag.moi-truong", title: "Môi trường Word và phiên bản", hint: environment.summary, defaultOpen: environment.unsupported > 0 }, environmentEl);
  const sampleButton = element("button", "btn-sample", { type: "button", textContent: "Chèn văn bản thử nghiệm" });
  sampleButton.addEventListener("click", () => void onInsertSample());
  sampleEl.append(element("p", "sample-note", { className: "note", textContent: "Chèn một đoạn văn bản mẫu vào cuối tài liệu để thử nhanh các nút." }), sampleButton);
  acc(panelAgency, { id: "ag.thu-nghiem", title: "Văn bản thử nghiệm", hint: "Chèn đoạn mẫu để thử" }, sampleEl);

  const documentSupported = isDocumentStandardSupported();
  const switcher = mountStandardSwitcher(noticeEl, {
    store,
    readSnapshot: documentSupported ? readDocumentSnapshot : undefined,
    writeStandard: documentSupported ? writeDocumentStandard : undefined,
    writeNoAsk: documentSupported ? writeDocumentNoAsk : undefined,
  });
  const quick = mountQuickOptions(quickEl, { currentStandard: switcher.currentStandard, store });
  const header = mountHeaderForm(headerEl, { currentStandard: switcher.currentStandard, values: quick.values, store });
  const footer = mountFooterForm(footerEl, { currentStandard: switcher.currentStandard, values: quick.values, store });
  const wizard = mountNewDocWizard(wizardEl, {
    currentStandard: switcher.currentStandard,
    values: quick.values,
    styleSpecs: quick.styleSpecs,
    behavior: quick.behavior,
    collectHeader: header.collect,
    presetHeader: header.preset,
    collectFooter: footer.collect,
    writeStandard: documentSupported ? writeDocumentStandard : undefined,
    afterCreate: (doc) => {
      header.commitNumber();
      addEntry(store, switcher.currentStandard(), entryFromDocument(doc));
      settings.refresh();
    },
  });
  const extras = mountInsertPanel(extrasEl, { currentStandard: switcher.currentStandard, values: quick.values });
  const settings = mountSettingsPanel(settingsEl, {
    currentStandard: switcher.currentStandard,
    store,
    useProfile: (agency) => {
      header.applyAgency(agency);
      footer.applyAgency(agency);
    },
    onImported: () => {
      quick.refresh();
      header.refresh();
      footer.refresh();
      wizard.refresh();
    },
  });
  const report = mountReportPanel(reportEl, {
    currentStandard: switcher.currentStandard,
    pipelineContext: () => ({ profile: switcher.currentStandard(), values: quick.values(), specs: quick.styleSpecs(), localFormatting: quick.behavior().localFormatting }),
  });
  mountFixAllButton(fixAllEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs, behavior: quick.behavior });
  mountTemplateMaker(templateEl, { currentStandard: switcher.currentStandard });
  mountPageSetupPanel(pageEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs });
  mountFormatPanel(formatEl, { currentStandard: switcher.currentStandard, behavior: quick.behavior, styleSpecs: quick.styleSpecs });

  // Thanh tab: hai tab chuẩn (đổi chuẩn đang dùng), Kiểm tra, Cơ quan.
  const standardIds: readonly string[] = STANDARDS.map((standard) => standard.id);
  const tabSpecs: TabSpec[] = [
    ...STANDARDS.map((standard) => ({ id: standard.id, label: standard.tabLabel, panelId: PANEL_STANDARD })),
    { id: CHECK_TAB, label: "Kiểm tra", panelId: PANEL_CHECK },
    { id: AGENCY_TAB, label: "Cơ quan", panelId: PANEL_AGENCY },
  ];
  let pendingStandardTab: string | null = null;

  function activate(id: string): void {
    const onStandard = standardIds.includes(id);
    panelStandard.hidden = !onStandard;
    panelCheck.hidden = id !== CHECK_TAB;
    panelAgency.hidden = id !== AGENCY_TAB;
    if (onStandard) panelStandard.setAttribute("aria-labelledby", tabButtonId(id));
    panelCheck.setAttribute("aria-labelledby", tabButtonId(CHECK_TAB));
    panelAgency.setAttribute("aria-labelledby", tabButtonId(AGENCY_TAB));
    tabs.setActive(id);
    saveActiveTab(store, id);
  }

  function onTab(id: string): void {
    pendingStandardTab = null;
    if (!standardIds.includes(id) || id === switcher.currentStandard().id) {
      activate(id);
      return;
    }
    // Đổi chuẩn có thể phải hỏi xác nhận; tab chỉ đổi khi chuẩn thật sự đổi (xem onChange bên dưới).
    pendingStandardTab = id;
    switcher.choose(id as StandardId);
  }

  const savedTab = loadActiveTab(store, tabSpecs.map((tab) => tab.id));
  const initialTab = standardIds.includes(savedTab) ? switcher.currentStandard().id : savedTab;
  const tabs = mountTabs(need("tabs"), tabSpecs, initialTab, onTab);
  activate(initialTab);

  const chips = STANDARDS.map((standard) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.textContent = standard.tabLabel;
    chip.addEventListener("click", () => switcher.choose(standard.id));
    return { id: standard.id, chip };
  });
  chipsEl.append(Object.assign(document.createElement("span"), { textContent: "Kiểm tra theo chuẩn:" }), ...chips.map((item) => item.chip));
  const paintChips = (): void => chips.forEach(({ id, chip }) => chip.setAttribute("aria-pressed", String(id === switcher.currentStandard().id)));
  paintChips();

  switcher.onChange(() => {
    report.refresh();
    quick.refresh();
    header.refresh();
    footer.refresh();
    wizard.refresh();
    extras.refresh();
    settings.refresh();
    settings.applyDefault();
    paintChips();
    const current = switcher.currentStandard().id;
    if (pendingStandardTab === current || standardIds.includes(tabs.active())) activate(current);
    pendingStandardTab = null;
  });

  watchControls(document.body);
  setStatus("Sẵn sàng.");
  void switcher.start().then(() => settings.applyDefault());
}
