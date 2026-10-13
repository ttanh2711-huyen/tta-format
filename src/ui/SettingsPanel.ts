import { exportConfig, importConfig } from "../settings/exportImport";
import { defaultProfile, deleteProfile, listProfiles, saveProfile, setDefaultProfile, type AgencyProfile } from "../settings/profile";
import { listCounters, removeCounter, setLastNumber } from "../settings/soVanBan";
import type { KeyValueStore } from "../settings/storage";
import type { StandardProfile } from "../standards/types";

export interface SettingsPanelDeps {
  currentStandard: () => StandardProfile;
  store: KeyValueStore;
  /** Điền hồ sơ vào Khối đầu và Khối cuối. */
  useProfile: (profile: AgencyProfile) => void;
  /** Gọi sau khi nhập cấu hình để các bảng khác đọc lại. */
  onImported: () => void;
}

function create<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: Array<Node | string>): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

const lines = (text: string): string[] => text.split(/\r?\n/).map((line) => line.trim()).filter((line) => line !== "");

/**
 * Bảng "Hồ sơ cơ quan và cấu hình" (T22): danh sách hồ sơ của chuẩn đang chọn (thêm, sửa, xoá, đặt mặc định, dùng để điền
 * sẵn Khối đầu và Khối cuối) và xuất, nhập file cấu hình để chuyển sang máy khác.
 */
export function mountSettingsPanel(container: HTMLElement, deps: SettingsPanelDeps): { refresh: () => void; applyDefault: () => void } {
  let editing: { id?: string; draft: Draft } | null = null;
  const body = create("div");
  const message = create("p", { id: "sp-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Hồ sơ cơ quan và cấu hình" }), body, message);

  interface Draft {
    label: string;
    parent: string;
    issuer: string;
    place: string;
    abbrs: string;
    position: string;
    signer: string;
  }

  const emptyDraft = (): Draft => ({ label: "", parent: "", issuer: "", place: "", abbrs: "", position: "", signer: "" });
  const draftOf = (p: AgencyProfile): Draft => ({ label: p.label, parent: p.parentLines.join("\n"), issuer: p.issuerLines.join("\n"), place: p.place, abbrs: p.abbrs, position: p.positionLines.join("\n"), signer: p.signerName });

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  const button = (text: string, onClick: () => void, props: Partial<HTMLButtonElement> = {}): HTMLButtonElement => {
    const element = create("button", { type: "button", textContent: text, ...props });
    element.addEventListener("click", onClick);
    return element;
  };

  function applyDefault(): void {
    const profile = defaultProfile(deps.store, deps.currentStandard());
    if (profile) deps.useProfile(profile);
  }

  function editor(profile: StandardProfile): HTMLElement {
    const draft = editing!.draft;
    const field = (id: string, caption: string, placeholder: string, key: keyof Draft, multiline = false): HTMLElement => {
      const el = multiline ? create("textarea", { id, rows: 2, value: draft[key], placeholder }) : create("input", { type: "text", id, value: draft[key], placeholder });
      el.addEventListener("input", () => {
        draft[key] = el.value;
      });
      return create("label", { className: "fld" }, create("span", { textContent: caption }), el);
    };
    const party = profile.headerKind === "dang";
    return create(
      "div",
      { className: "hf-box", id: "sp-editor" },
      field("sp-label", "Tên gợi nhớ", "Ví dụ: Chi cục Thủy lợi", "label"),
      field("sp-parent", party ? "Cơ quan cấp trên (tối đa 2 dòng)" : "Cơ quan chủ quản (tối đa 3 dòng)", "Mỗi dòng một cơ quan", "parent", true),
      field("sp-issuer", party ? "Cơ quan ban hành (tối đa 3 dòng)" : "Cơ quan ban hành (một dòng)", "Tên cơ quan ban hành", "issuer", true),
      field("sp-place", "Địa danh", "Ví dụ: Cần Thơ", "place"),
      field("sp-abbrs", "Chữ viết tắt cơ quan (nối bằng gạch nối)", "Ví dụ: CCTL-ĐĐPCTT", "abbrs"),
      field("sp-position", "Chức vụ người ký thường dùng (tối đa 3 dòng)", "Mỗi dòng một chức vụ", "position", true),
      field("sp-signer", "Họ tên người ký thường dùng", "Họ và tên", "signer"),
      button("Lưu hồ sơ", () => {
        const result = saveProfile(deps.store, profile, {
          id: editing?.id,
          standardId: profile.id,
          label: draft.label,
          parentLines: lines(draft.parent),
          issuerLines: lines(draft.issuer),
          place: draft.place,
          abbrs: draft.abbrs,
          positionLines: lines(draft.position),
          signerName: draft.signer,
        });
        if (!result.ok) {
          show("error", result.reasons.join(" "));
          return;
        }
        editing = null;
        render();
        show("ok", `Đã lưu hồ sơ "${result.profile.label}".`);
      }, { id: "sp-save" }),
      button("Huỷ", () => {
        editing = null;
        render();
      }),
    );
  }

  /** Sổ số văn bản (T26): xem và sửa số cuối cùng đã dùng của từng loại, từng năm. */
  function numberBook(standard: StandardProfile): HTMLElement {
    const counters = listCounters(deps.store, standard);
    const list = create("ul", { id: "sp-book" });
    for (const counter of counters) {
      const input = create("input", { type: "text", value: String(counter.last), inputMode: "numeric" });
      input.style.width = "60px";
      input.setAttribute("aria-label", `Số cuối đã dùng của ${counter.typeKey === "-" ? "công văn" : counter.typeKey} năm ${counter.year}`);
      list.append(
        create(
          "li",
          {},
          `${counter.typeKey === "-" ? "Công văn" : counter.typeKey} năm ${counter.year}: số cuối `,
          input,
          button("Lưu", () => {
            const ok = setLastNumber(deps.store, standard, counter.year, counter.typeKey === "-" ? null : counter.typeKey, Number(input.value));
            show(ok ? "ok" : "error", ok ? "Đã sửa số cuối." : "Số phải là số nguyên từ 1 trở lên.");
            if (ok) render();
          }),
          button("Xoá", () => {
            removeCounter(deps.store, standard, counter.year, counter.typeKey === "-" ? null : counter.typeKey);
            render();
          }),
        ),
      );
    }
    return create(
      "div",
      {},
      create("p", { className: "note", textContent: "Add-in ghi số văn bản mỗi lần tạo hoặc chèn khối đầu để gợi ý số kế tiếp. Số đánh liên tục từ 1 trong năm cho từng loại; sửa ở đây nếu số thực tế khác." }),
      counters.length === 0 ? create("p", { className: "note", textContent: "Chưa có số nào được ghi." }) : list,
    );
  }

  function transfer(): HTMLElement {
    const area = create("textarea", { id: "sp-transfer", rows: 6, placeholder: "Dán nội dung file cấu hình vào đây để nhập, hoặc bấm Xuất để xem nội dung." });
    const exportButton = button("Xuất cấu hình", () => {
      const text = exportConfig(deps.store);
      area.value = text;
      area.select();
      void navigator.clipboard?.writeText(text).catch(() => undefined);
      link.href = URL.createObjectURL(new Blob([text], { type: "application/json" }));
      link.hidden = false;
      show("ok", "Đã xuất cấu hình; nội dung nằm trong ô bên dưới và đã được chép vào bộ nhớ tạm nếu trình duyệt cho phép.");
    }, { id: "sp-export" });
    const link = create("a", { id: "sp-download", textContent: "Tải về .json", hidden: true, download: "nd30-addin-cau-hinh.json" });
    const importButton = button("Nhập cấu hình", () => {
      const report = importConfig(deps.store, area.value);
      if (!report.ok) {
        show("error", report.error ?? "Không nhập được.");
        return;
      }
      deps.onImported();
      render();
      const skipped = report.skipped.length > 0 ? ` Bỏ qua ${report.skipped.length} mục: ${report.skipped.join(" ")}` : "";
      show(report.skipped.length > 0 ? "warning" : "ok", `Đã nhập ${report.profiles} hồ sơ, ${report.presets} preset, ${report.defaults} mặc định, ${report.agencyEntries} cơ quan.${skipped}`);
    }, { id: "sp-import" });
    return create("div", {}, create("p", { className: "note", textContent: "Xuất để chuyển hồ sơ cơ quan, preset và danh sách cơ quan sang máy khác; nhập sẽ gộp vào cấu hình hiện có (hồ sơ trùng mã được thay)." }), exportButton, " ", link, area, importButton);
  }

  function render(): void {
    const standard = deps.currentStandard();
    body.replaceChildren();
    const profiles = listProfiles(deps.store, standard);
    body.append(create("p", { className: "note", textContent: `Hồ sơ cơ quan của chuẩn ${standard.name}. Hồ sơ mặc định được điền sẵn vào Khối đầu và Khối cuối khi mở add-in; "Dùng" điền ngay.` }));
    const list = create("ul", { id: "sp-list" });
    for (const profile of profiles) {
      list.append(
        create(
          "li",
          {},
          create("strong", { textContent: `${profile.label}${profile.isDefault ? " (mặc định)" : ""}` }),
          create("div", { className: "note", textContent: [...profile.parentLines, ...profile.issuerLines].join(" · ") }),
          button("Dùng", () => {
            deps.useProfile(profile);
            show("ok", `Đã điền "${profile.label}" vào Khối đầu và Khối cuối.`);
          }),
          button("Đặt mặc định", () => {
            setDefaultProfile(deps.store, profile.id);
            render();
          }, { disabled: profile.isDefault === true }),
          button("Sửa", () => {
            editing = { id: profile.id, draft: draftOf(profile) };
            render();
          }),
          button("Xoá", () => {
            deleteProfile(deps.store, profile.id);
            render();
            show("ok", `Đã xoá hồ sơ "${profile.label}".`);
          }),
        ),
      );
    }
    body.append(profiles.length === 0 ? create("p", { className: "note", textContent: "Chưa có hồ sơ nào." }) : list);
    if (editing) body.append(editor(standard));
    else {
      body.append(button("Thêm hồ sơ", () => {
        editing = { draft: emptyDraft() };
        render();
      }, { id: "sp-add" }));
    }
    body.append(create("details", { id: "sp-book-group" }, create("summary", { textContent: "Sổ số văn bản" }), numberBook(standard)));
    body.append(create("details", { id: "sp-transfer-group" }, create("summary", { textContent: "Xuất, nhập cấu hình" }), transfer()));
  }

  render();
  return {
    refresh: () => {
      editing = null;
      show("", "");
      render();
    },
    applyDefault,
  };
}
