import { applyPageOptions, applyStyleOptions, type ApplyOutcome } from "../office/applyOptions";
import { parseMm } from "../office/pageSetupPlan";
import { buildStyleSpecs, type StyleSpec } from "../ooxml/styleSpecs";
import {
  behaviorOf,
  buildRegistry,
  resetGroup,
  resetOption,
  setOption,
  toStyleOptions,
  initialState,
  type FormatRegistry,
  type FormatState,
  type FormatValues,
  type GroupId,
  type OptionDef,
  type OptionValue,
} from "../settings/formatOptions";
import {
  applyPreset,
  clearDefaults,
  deletePreset,
  duplicatePreset,
  listPresets,
  loadDefaults,
  renamePreset,
  saveDefaults,
  savePreset,
  type Preset,
} from "../settings/presets";
import { browserStore, type KeyValueStore } from "../settings/storage";
import type { StandardProfile } from "../standards/types";

export interface QuickOptionsHandle {
  values(): FormatValues;
  behavior(): ReturnType<typeof behaviorOf>;
  /** Bộ style theo tuỳ chọn hiện hành. */
  styleSpecs(): StyleSpec[];
  /** Dựng lại theo chuẩn đang chọn (gọi khi người dùng đổi chuẩn). */
  refresh(): void;
}

export interface QuickOptionsDeps {
  currentStandard: () => StandardProfile;
  store?: KeyValueStore;
}

function create<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> = {},
  ...children: Array<Node | string>
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  Object.assign(element, props);
  element.append(...children);
  return element;
}

function button(text: string, onClick: () => void, props: Partial<HTMLButtonElement> = {}): HTMLButtonElement {
  const element = create("button", { type: "button", textContent: text, ...props });
  element.addEventListener("click", onClick);
  return element;
}

function showValue(value: OptionValue, def: OptionDef): string {
  const choice = def.choices.find((c) => c.value === value);
  if (choice) return choice.label;
  if (typeof value === "boolean") return value ? "Bật" : "Tắt";
  return `${String(value).replace(".", ",")}${def.unit ? ` ${def.unit}` : ""}`;
}

/** Bảng "Tuỳ chọn nhanh" (T6A): sinh từ registry của chuẩn đang chọn. */
export function mountQuickOptions(container: HTMLElement, deps: QuickOptionsDeps): QuickOptionsHandle {
  const store = deps.store ?? browserStore();
  let profile = deps.currentStandard();
  let registry: FormatRegistry = buildRegistry(profile);
  let state: FormatState = initialState(registry);
  let pending: { id: string; value: OptionValue; reason: string } | null = null;
  let presetName = "";
  let deleteArmed: string | null = null;
  const openGroups = new Set<GroupId>(["font", "paragraph"]);
  let queue: Promise<void> = Promise.resolve();

  const body = create("div", { id: "qo-body" });
  const message = create("p", { id: "qo-message" });
  message.setAttribute("role", "status");
  container.append(create("h2", { textContent: "Tuỳ chọn nhanh" }), body, message);

  function show(kind: "ok" | "warning" | "error" | "", text: string): void {
    message.className = kind;
    message.textContent = text;
  }

  function showOutcome(outcome: ApplyOutcome): void {
    show(outcome.kind, outcome.text);
  }

  function enqueue(task: () => Promise<ApplyOutcome | undefined>): void {
    queue = queue.then(async () => {
      try {
        const outcome = await task();
        if (outcome) showOutcome(outcome);
      } catch (error) {
        show("error", `Không áp dụng được: ${error instanceof Error ? error.message : String(error)}`);
      }
    });
  }

  function applyLive(groups: ReadonlySet<GroupId> | "all"): void {
    const values = state.values;
    enqueue(async () => {
      const outcomes: ApplyOutcome[] = [];
      const wants = (group: GroupId) => groups === "all" || groups.has(group);
      if (wants("page")) outcomes.push(await applyPageOptions(values));
      if (wants("font") || wants("paragraph") || wants("table")) outcomes.push(await applyStyleOptions(profile, values));
      const worst = outcomes.find((o) => o.kind === "error") ?? outcomes.find((o) => o.kind === "warning") ?? outcomes[0];
      return worst ? { kind: worst.kind, text: outcomes.map((o) => o.text).join(" ") } : undefined;
    });
  }

  function commit(def: OptionDef, value: OptionValue, confirmed = false): void {
    const result = setOption(registry, state, def.id, value, { confirmed });
    if (result.status === "rejected") {
      show("error", result.reason);
      return;
    }
    if (result.status === "needsConfirm") {
      pending = { id: def.id, value, reason: result.reason };
      render();
      return;
    }
    state = result.state;
    pending = null;
    show("", "");
    render();
    if (def.effect === "now") {
      applyLive(new Set<GroupId>([def.group]));
    } else if (def.effect === "onCreate") {
      show("ok", "Đã ghi nhận, sẽ dùng khi tạo văn bản hoặc chèn khối.");
    }
  }

  function optionRow(def: OptionDef): HTMLElement {
    const value = state.values[def.id];
    const row = create("div", { className: "qo-row" });
    const isException = state.exceptions.includes(def.id);
    const title = create("div", { className: "qo-title" }, `${def.label}: `, create("strong", { textContent: showValue(value, def) }));
    if (isException) title.append(create("span", { className: "qo-exception", textContent: " (ngoại lệ chủ ý)" }));
    if (def.effect === "onCreate") title.append(create("span", { className: "note", textContent: " · áp dụng khi tạo văn bản" }));
    row.append(title);

    if (def.kind === "boolean") {
      const box = create("input", { type: "checkbox", checked: value === true, id: `qo-${def.id}` });
      box.addEventListener("change", () => commit(def, box.checked));
      row.append(create("label", {}, box, " Bật"));
    } else {
      const choices = create("div", { className: "qo-choices" });
      for (const choice of def.choices) {
        const chosen = choice.value === value;
        const choiceButton = button(choice.label, () => commit(def, choice.value), { id: `qo-${def.id}-${String(choice.value)}` });
        choiceButton.setAttribute("aria-pressed", String(chosen));
        if (choice.exception) choiceButton.title = choice.exception;
        choices.append(choiceButton);
      }
      row.append(choices);
    }

    if (def.custom) {
      const input = create("input", { type: "text", inputMode: "decimal", id: `qo-${def.id}-custom`, placeholder: "Tuỳ chỉnh…" });
      input.setAttribute("aria-label", `${def.label}: nhập giá trị riêng${def.unit ? ` (${def.unit})` : ""}`);
      const submit = () => {
        const parsed = parseMm(input.value);
        if (Number.isNaN(parsed)) {
          show("error", `"${def.label}": hãy nhập một số, ví dụ 13 hoặc 1,5.`);
          return;
        }
        commit(def, parsed);
      };
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") submit();
      });
      row.append(create("div", { className: "qo-custom" }, input, def.unit ? ` ${def.unit} ` : " ", button("Áp dụng", submit)));
    }

    if (value !== def.default || isException) {
      row.append(button("Về mặc định", () => resetAndApply((s) => resetOption(registry, s, def.id), new Set<GroupId>([def.group])), { className: "qo-small" }));
    }
    if (pending && pending.id === def.id) {
      const box = create("div", { className: "qo-confirm", id: "qo-confirm" });
      box.setAttribute("role", "alert");
      box.append(
        create("p", { textContent: `${pending.reason} Áp dụng giá trị ${showValue(pending.value, def)} thì mục này được ghi là ngoại lệ chủ ý.` }),
        button("Vẫn áp dụng", () => {
          const ask = pending;
          if (ask) commit(def, ask.value, true);
        }),
        button("Huỷ", () => {
          pending = null;
          render();
        }),
      );
      row.append(box);
    }
    return row;
  }

  function resetAndApply(change: (s: FormatState) => FormatState, groups: ReadonlySet<GroupId> | "all"): void {
    state = change(state);
    pending = null;
    render();
    applyLive(groups);
  }

  function groupSection(group: { id: GroupId; label: string }): HTMLElement | null {
    const options = registry.options.filter((o) => o.group === group.id);
    if (options.length === 0) return null;
    const details = create("details", { open: openGroups.has(group.id) });
    details.addEventListener("toggle", () => {
      if (details.open) openGroups.add(group.id);
      else openGroups.delete(group.id);
    });
    details.append(create("summary", { textContent: group.label }));
    for (const def of options) details.append(optionRow(def));
    details.append(button("Đặt lại nhóm về mặc định của chuẩn", () => resetAndApply((s) => resetGroup(registry, s, group.id), new Set<GroupId>([group.id])), { className: "qo-small" }));
    return details;
  }

  function applyPresetNow(preset: Preset): void {
    const result = applyPreset(registry, preset);
    state = result.state;
    pending = null;
    render();
    applyLive("all");
    if (result.ignored.length > 0) show("warning", `Đã áp preset "${preset.name}". ${result.ignored.join(" ")}`);
  }

  function presetsSection(): HTMLElement {
    const details = create("details");
    details.append(create("summary", { textContent: "Preset (bộ tuỳ chọn đặt tên)" }));
    const nameInput = create("input", { type: "text", id: "qo-preset-name", value: presetName, placeholder: "Tên preset" });
    nameInput.addEventListener("input", () => {
      presetName = nameInput.value;
    });
    details.append(
      create("p", { className: "note", textContent: `Preset thuộc chuẩn ${profile.name}. Nhập tên rồi bấm "Lưu preset mới"; "Đổi tên" dùng tên trong ô.` }),
      create("div", {}, nameInput, " ", button("Lưu preset mới", () => {
        const saved = savePreset(store, profile, presetName, state);
        if (saved.ok) {
          presetName = "";
          show("ok", `Đã lưu preset "${saved.preset.name}".`);
          render();
        } else {
          show("error", saved.reason);
        }
      })),
    );
    for (const preset of listPresets(store, profile)) {
      const row = create("div", { className: "qo-preset" }, create("strong", { textContent: preset.name }), " ");
      row.append(button("Áp dụng", () => applyPresetNow(preset)));
      row.append(button("Nhân bản", () => {
        const copy = duplicatePreset(store, profile, preset.id);
        show(copy.ok ? "ok" : "error", copy.ok ? `Đã nhân bản thành "${copy.preset.name}".` : copy.reason);
        render();
      }));
      if (!preset.builtIn) {
        row.append(button("Đổi tên", () => {
          const renamed = renamePreset(store, preset.id, presetName);
          if (renamed.ok) presetName = "";
          show(renamed.ok ? "ok" : "error", renamed.ok ? "Đã đổi tên." : renamed.reason);
          render();
        }));
        row.append(button(deleteArmed === preset.id ? "Bấm lại để xoá" : "Xoá", () => {
          if (deleteArmed !== preset.id) {
            deleteArmed = preset.id;
            render();
            return;
          }
          deleteArmed = null;
          const removed = deletePreset(store, preset.id);
          show(removed.ok ? "ok" : "error", removed.ok ? `Đã xoá preset "${preset.name}".` : removed.reason);
          render();
        }));
      }
      details.append(row);
    }
    return details;
  }

  function render(): void {
    body.replaceChildren();
    body.append(
      create("p", { className: "note", textContent: `Chuẩn: ${profile.name}. Giá trị ngoài quy định chỉ áp dụng sau khi bạn xác nhận và được ghi là ngoại lệ chủ ý.` }),
      create(
        "div",
        { className: "qo-toolbar" },
        button("Lưu làm mặc định", () => {
          saveDefaults(store, profile.id, state);
          show("ok", "Đã lưu giá trị hiện tại làm mặc định cho chuẩn này trên máy này.");
        }),
        button("Đặt lại tất cả", () => resetAndApply(() => initialState(registry), "all")),
        button("Xoá mặc định đã lưu", () => {
          clearDefaults(store, profile.id);
          show("ok", "Đã xoá mặc định đã lưu; lần mở sau dùng mặc định của chuẩn.");
        }),
      ),
    );
    for (const group of registry.groups) {
      const section = groupSection(group);
      if (section) body.append(section);
    }
    body.append(presetsSection());
  }

  function load(): void {
    profile = deps.currentStandard();
    registry = buildRegistry(profile);
    const loaded = loadDefaults(store, registry, profile.id);
    state = loaded.state;
    pending = null;
    deleteArmed = null;
    render();
    show(loaded.ignored.length > 0 ? "warning" : "", loaded.ignored.join(" "));
  }

  load();

  return {
    values: () => state.values,
    behavior: () => behaviorOf(state.values),
    styleSpecs: () => buildStyleSpecs(profile, toStyleOptions(profile, state.values)),
    refresh: load,
  };
}
