import { insertSampleParagraph } from "../office/insertSample";
import { isDocumentStandardSupported, readDocumentSnapshot, writeDocumentNoAsk, writeDocumentStandard } from "../office/documentStandard";
import { browserStore } from "../settings/storage";
import type { Issue } from "../checker/types";
import { STANDARDS } from "../standards";
import type { StandardProfile } from "../standards/types";
import { watchControls } from "../ui/a11y";
import { appVersion, mountEnvironmentPanel } from "../ui/EnvironmentPanel";
import { mountFixAllButton } from "../ui/FixAllButton";
import { mountTemplateMaker } from "../ui/TemplateMaker";
import { mountFooterForm } from "../ui/FooterForm";
import { mountFormatPanel } from "../ui/FormatPanel";
import { mountHeaderForm } from "../ui/HeaderForm";
import { mountHeaderRebuild } from "../ui/HeaderRebuild";
import { mountInsertPanel } from "../ui/InsertPanel";
import { mountNewDocWizard } from "../ui/NewDocWizard";
import { mountPageSetupPanel } from "../ui/PageSetupPanel";
import { mountQuickActions } from "../ui/QuickActions";
import { mountQuickOptions } from "../ui/QuickOptions";
import { mountReportPanel } from "../ui/ReportPanel";
import { mountSettingsPanel } from "../ui/SettingsPanel";
import { addEntry, entryFromDocument } from "../settings/soDangKy";
import { mountStandardSwitcher } from "../ui/StandardSwitcher";
import { mountTabs, tabButtonId } from "../ui/Tabs";
import { loadActiveSection, loadActiveTab, saveActiveSection, saveActiveTab } from "../ui/tabsLogic";

/** Biểu tượng của các tab kiểu ribbon (nét vẽ SVG, khung 24×24). */
const ICONS = {
  home: ["M13 2 4 14h7l-1 8 9-12h-7z"],
  selection: ["M4 7V4h3", "M17 4h3v3", "M20 17v3h-3", "M7 20H4v-3", "M8 10h8", "M8 14h5"],
  sparkle: ["M12 3v4", "M12 17v4", "M3 12h4", "M17 12h4", "m6 6 2 2", "m16 16 2 2", "m6 18 2-2", "m16 8 2-2"],
  page: ["M5 3h14v18H5z", "M8 6h8v12H8z"],
  pageNumber: ["M5 3h14v18H5z", "M10 7h4", "M8 12h8", "M8 16h8"],
  styles: ["M4 19 9 5l5 14", "M6.5 13h5", "M16 9h4", "M16 14h4", "M16 19h4"],
  template: ["M4 4h16v4H4z", "M4 12h7v8H4z", "M15 12h5", "M15 16h5", "M15 20h5"],
  list: ["M9 6h11", "M9 12h11", "M9 18h11", "m3 6 1 1 2-2", "m3 12 1 1 2-2", "m3 18 1 1 2-2"],
  compose: ["M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z", "M14 3v5h5", "M12 11v6", "M9 14h6"],
  blocks: ["M4 4h16v6H4z", "M4 14h7v6H4z", "M15 14h5v6h-5z"],
  format: ["M4 6h16", "M4 10h16", "M4 14h16", "M4 18h10"],
  check: ["M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0z", "m8 12 3 3 5-6"],
  agency: ["M4 21V8l8-5 8 5v13", "M9 21v-6h6v6", "M3 21h18"],
} as const;

interface Section {
  id: string;
  label: string;
  content: HTMLElement;
}

interface RibbonTab {
  id: string;
  label: string;
  icon: readonly string[];
  sections: Section[];
}

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
 * Dựng toàn bộ giao diện TTA-Format: chuẩn văn bản (Nghị định 30 hay Hướng dẫn 05) chọn bằng nút gạt ở đầu trang; năm
 * tab kiểu ribbon (Nhanh, Soạn, Khối, Định dạng, Kiểm tra, Cơ quan), mỗi tab có các mục con, mỗi lần hiện một mục; tab
 * Nhanh gom các việc hay làm, mỗi việc một lần bấm.
 */
export function mountApp(options: { check?: (profile: StandardProfile) => Promise<Issue[]> } = {}): void {
  const store = browserStore();
  const noticeEl = need("notice");
  need("app-version").textContent = `v${appVersion()}`;

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
  const rebuildEl = element("div", "header-rebuild");
  const settingsEl = element("div", "settings-panel");
  const environmentEl = element("div", "environment");
  const sampleEl = element("div", "sample");

  mountEnvironmentPanel(environmentEl);
  const sampleButton = element("button", "btn-sample", { type: "button", textContent: "Chèn văn bản thử nghiệm" });
  sampleButton.addEventListener("click", () => void onInsertSample());
  sampleEl.append(sampleButton);

  // Tab "Nhanh": việc hay làm nhất, một lần bấm (đi tới đúng mục rồi bấm nút ở đó).
  const homeEl = element("div", "quick-actions");
  mountQuickActions(homeEl, [
    {
      title: "Chạy ngay",
      actions: [
        { id: "format-selection", label: "Định dạng phần tô đen", hint: "Bôi đen đoạn cần sửa rồi bấm", icon: ICONS.selection, runs: true, onClick: () => go("format", "body", "fb-selection") },
        { id: "check", label: "Kiểm tra văn bản", hint: "Liệt kê lỗi theo nhóm", icon: ICONS.check, runs: true, onClick: () => go("check", "report", "rp-check") },
        { id: "quick-fix", label: "Sửa nhanh", hint: "Chuẩn hoá cả văn bản, hoàn tác được", icon: ICONS.sparkle, runs: true, onClick: () => go("check", "report", "rp-quick") },
        { id: "page", label: "Thiết lập trang", hint: "Khổ A4, lề theo chuẩn", icon: ICONS.page, runs: true, onClick: () => go("format", "page", "ps-apply") },
        { id: "page-number", label: "Đánh số trang", hint: "Giữa lề trên, ẩn trang 1", icon: ICONS.pageNumber, runs: true, onClick: () => go("format", "page", "ps-page-number") },
        { id: "styles", label: "Cài bộ style", hint: "Style của chuẩn vào văn bản", icon: ICONS.styles, runs: true, onClick: () => go("format", "page", "ps-install-styles") },
        { id: "fix-header", label: "Sửa khối đầu cũ", hint: "Đọc khối đầu sai, dựng lại đúng quy định", icon: ICONS.blocks, runs: true, onClick: () => go("blocks", "fix", "hr-read") },
      ],
    },
    {
      title: "Mở nhanh",
      actions: [
        { id: "new", label: "Tạo văn bản mới", hint: "Chọn loại, điền rồi tạo", icon: ICONS.compose, runs: false, onClick: () => go("compose", "new") },
        { id: "template", label: "Tạo mẫu", hint: "Biến văn bản thành mẫu", icon: ICONS.template, runs: false, onClick: () => go("compose", "template") },
        { id: "head", label: "Khối đầu", hint: "Cơ quan, số, ngày tháng", icon: ICONS.blocks, runs: false, onClick: () => go("blocks", "head") },
        { id: "foot", label: "Khối cuối", hint: "Nơi nhận, người ký", icon: ICONS.blocks, runs: false, onClick: () => go("blocks", "foot") },
        { id: "fixall", label: "Chỉnh toàn bộ", hint: "Nhiều bước, có xem trước", icon: ICONS.list, runs: false, onClick: () => go("check", "fixall") },
        { id: "agency", label: "Hồ sơ cơ quan", hint: "Điền sẵn khối đầu, khối cuối", icon: ICONS.agency, runs: false, onClick: () => go("agency", "profiles") },
      ],
    },
  ]);

  // Sáu tab theo trình tự công việc: nhanh, soạn, khối đầu và cuối, định dạng, kiểm tra, cơ quan. Mỗi tab chia mục con
  // (chỉ hiện một mục mỗi lần) để hạn chế cuộn trang.
  const ribbon: RibbonTab[] = [
    { id: "home", label: "Nhanh", icon: ICONS.home, sections: [{ id: "actions", label: "Việc hay làm", content: homeEl }] },
    {
      id: "compose",
      label: "Soạn",
      icon: ICONS.compose,
      sections: [
        { id: "new", label: "Tạo mới", content: wizardEl },
        { id: "template", label: "Tạo mẫu", content: templateEl },
        { id: "extras", label: "Chèn thêm", content: extrasEl },
      ],
    },
    {
      id: "blocks",
      label: "Khối",
      icon: ICONS.blocks,
      sections: [
        { id: "head", label: "Khối đầu", content: headerEl },
        { id: "fix", label: "Sửa khối đầu", content: rebuildEl },
        { id: "foot", label: "Khối cuối", content: footerEl },
      ],
    },
    {
      id: "format",
      label: "Định dạng",
      icon: ICONS.format,
      sections: [
        { id: "body", label: "Nội dung", content: formatEl },
        { id: "page", label: "Trang, style", content: pageEl },
        { id: "options", label: "Tuỳ chọn", content: quickEl },
      ],
    },
    {
      id: "check",
      label: "Kiểm tra",
      icon: ICONS.check,
      sections: [
        { id: "report", label: "Kiểm tra lỗi", content: reportEl },
        { id: "fixall", label: "Chỉnh toàn bộ", content: fixAllEl },
      ],
    },
    {
      id: "agency",
      label: "Cơ quan",
      icon: ICONS.agency,
      sections: [
        { id: "profiles", label: "Hồ sơ", content: settingsEl },
        { id: "environment", label: "Môi trường", content: environmentEl },
        { id: "sample", label: "Thử nghiệm", content: sampleEl },
      ],
    },
  ];

  const panels = need("panels");
  const sectionShowers = new Map<string, (id: string) => void>();
  for (const tab of ribbon) {
    const panel = element("div", `panel-${tab.id}`, { className: "panel" });
    panel.setAttribute("role", "tabpanel");
    panel.setAttribute("aria-labelledby", tabButtonId(tab.id));
    // Tab chỉ có một mục thì không cần thanh mục con.
    const nav = element("div", `subnav-${tab.id}`, { className: "subnav-bar", hidden: tab.sections.length < 2 });
    const sectionEls = tab.sections.map((section) => {
      const el = element("section", `sec-${tab.id}-${section.id}`, { className: "sec" });
      el.setAttribute("role", "tabpanel");
      el.setAttribute("aria-labelledby", tabButtonId(`${tab.id}-${section.id}`));
      el.append(section.content);
      return el;
    });
    panel.append(nav, ...sectionEls);
    panels.append(panel);
    const ids = tab.sections.map((section) => section.id);
    const initial = loadActiveSection(store, tab.id, ids);
    const subTabs = mountTabs(
      nav,
      tab.sections.map((section) => ({ id: `${tab.id}-${section.id}`, label: section.label, panelId: `sec-${tab.id}-${section.id}` })),
      `${tab.id}-${initial}`,
      (id) => showSection(id.slice(tab.id.length + 1)),
      { className: "subnav", label: `Mục của tab ${tab.label}` },
    );
    function showSection(id: string): void {
      sectionEls.forEach((el, index) => (el.hidden = ids[index] !== id));
      subTabs.setActive(`${tab.id}-${id}`);
      saveActiveSection(store, tab.id, id);
    }
    showSection(initial);
    sectionShowers.set(tab.id, showSection);
  }

  /** Đi tới mục `sectionId` của tab `tabId`; có `clickId` thì bấm luôn nút đó để chạy ngay. */
  function go(tabId: string, sectionId: string, clickId?: string): void {
    activate(tabId);
    sectionShowers.get(tabId)?.(sectionId);
    window.scrollTo(0, 0);
    if (clickId) document.getElementById(clickId)?.click();
  }

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
    check: options.check,
  });
  mountFixAllButton(fixAllEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs, behavior: quick.behavior });
  mountTemplateMaker(templateEl, { currentStandard: switcher.currentStandard });
  const rebuild = mountHeaderRebuild(rebuildEl, { currentStandard: switcher.currentStandard, values: quick.values });
  mountPageSetupPanel(pageEl, { currentStandard: switcher.currentStandard, values: quick.values, styleSpecs: quick.styleSpecs });
  mountFormatPanel(formatEl, { currentStandard: switcher.currentStandard, behavior: quick.behavior, styleSpecs: quick.styleSpecs });

  // Thanh tab kiểu ribbon.
  const tabIds = ribbon.map((tab) => tab.id);
  const tabs = mountTabs(
    need("tabs"),
    ribbon.map((tab) => ({ id: tab.id, label: tab.label, panelId: `panel-${tab.id}`, icon: tab.icon })),
    loadActiveTab(store, tabIds),
    activate,
    { className: "tabs ribbon-tabs" },
  );
  function activate(id: string): void {
    for (const tab of ribbon) need(`panel-${tab.id}`).hidden = tab.id !== id;
    tabs.setActive(id);
    saveActiveTab(store, id);
  }
  activate(tabs.active());

  // Chuẩn văn bản: nút gạt ở đầu trang, dùng chung cho mọi tab.
  const standardButtons = STANDARDS.map((standard) => {
    // Nhãn gọn "NĐ 30", "HD 05" lấy trong tên chuẩn; tên đầy đủ ở chú thích.
    const short = /\(([^)]+)\)/.exec(standard.name)?.[1] ?? standard.tabLabel;
    const button = element("button", `std-${standard.id}`, { type: "button", textContent: short, title: `${standard.tabLabel}: ${standard.name}` });
    button.addEventListener("click", () => switcher.choose(standard.id));
    return { id: standard.id, button };
  });
  need("standard-switch").append(...standardButtons.map((item) => item.button));
  const paintStandard = (): void => standardButtons.forEach(({ id, button }) => button.setAttribute("aria-pressed", String(id === switcher.currentStandard().id)));
  paintStandard();

  switcher.onChange(() => {
    rebuild.refresh();
    report.refresh();
    quick.refresh();
    header.refresh();
    footer.refresh();
    wizard.refresh();
    extras.refresh();
    settings.refresh();
    settings.applyDefault();
    paintStandard();
  });

  watchControls(document.body);
  setStatus("Sẵn sàng.");
  void switcher.start().then(() => settings.applyDefault());
}
