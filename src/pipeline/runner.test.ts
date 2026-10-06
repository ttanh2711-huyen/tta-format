import {
  buildPreview,
  planPipeline,
  runPipeline,
  StepRegistry,
  summarizeReport,
  type PipelineStep,
  type StepPlan,
  type StepResult,
} from "./runner";

interface Ctx {
  log: string[];
}

function step(id: string, order: number, options: { plan?: Partial<StepPlan> | Error; run?: StepResult | Error } = {}): PipelineStep<Ctx> {
  return {
    id,
    label: `Bước ${id}`,
    order,
    async plan(ctx) {
      ctx.log.push(`plan:${id}`);
      if (options.plan instanceof Error) throw options.plan;
      return { willRun: true, lines: [`làm ${id}`], ...options.plan };
    },
    async run(ctx) {
      ctx.log.push(`run:${id}`);
      if (options.run instanceof Error) throw options.run;
      return options.run ?? { status: "done", kind: "ok", text: `xong ${id}` };
    },
  };
}

describe("StepRegistry", () => {
  it("trả các bước theo thứ tự order, bước đăng ký sau không cần sửa khung", () => {
    const registry = new StepRegistry<Ctx>();
    registry.register(step("thân", 30));
    registry.register(step("trang", 10));
    registry.register(step("style", 20));
    registry.register(step("số trang", 40));
    expect(registry.steps().map((s) => s.id)).toEqual(["trang", "style", "thân", "số trang"]);
  });

  it("cùng order thì giữ thứ tự đăng ký", () => {
    const registry = new StepRegistry<Ctx>();
    registry.register(step("b", 10));
    registry.register(step("a", 10));
    expect(registry.steps().map((s) => s.id)).toEqual(["b", "a"]);
  });

  it("đăng ký trùng mã bước là lỗi", () => {
    const registry = new StepRegistry<Ctx>();
    registry.register(step("a", 1));
    expect(() => registry.register(step("a", 2))).toThrow(/a/);
  });
});

describe("planPipeline", () => {
  it("lập kế hoạch từng bước theo thứ tự, không chạy bước nào", async () => {
    const ctx: Ctx = { log: [] };
    const planned = await planPipeline([step("a", 1), step("b", 2, { plan: { willRun: false, lines: [], skipReason: "đã đúng" } })], ctx);
    expect(ctx.log).toEqual(["plan:a", "plan:b"]);
    expect(planned).toEqual([
      { id: "a", label: "Bước a", willRun: true, lines: ["làm a"] },
      { id: "b", label: "Bước b", willRun: false, lines: [], skipReason: "đã đúng" },
    ]);
  });

  it("bước lập kế hoạch bị lỗi thì không chạy bước đó và ghi lý do, không làm hỏng các bước khác", async () => {
    const ctx: Ctx = { log: [] };
    const planned = await planPipeline([step("a", 1, { plan: new Error("không đọc được") }), step("b", 2)], ctx);
    expect(planned[0]).toMatchObject({ id: "a", willRun: false, skipReason: expect.stringContaining("không đọc được") });
    expect(planned[1].willRun).toBe(true);
  });
});

describe("runPipeline", () => {
  it("chạy lần lượt các bước có kế hoạch chạy, bước còn lại bị bỏ qua kèm lý do", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1), step("b", 2), step("c", 3)];
    const planned = await planPipeline(steps, ctx);
    planned[1] = { ...planned[1], willRun: false, skipReason: "đã đúng" };
    ctx.log.length = 0;
    const report = await runPipeline(steps, ctx, planned);
    expect(ctx.log).toEqual(["run:a", "run:c"]);
    expect(report.steps.map((s) => [s.id, s.state])).toEqual([["a", "done"], ["b", "skipped"], ["c", "done"]]);
    expect(report.steps[1].text).toBe("đã đúng");
    expect(report).toMatchObject({ ok: true, changed: true });
  });

  it("bước lỗi thì dừng, các bước trước giữ kết quả, các bước sau đánh dấu chưa chạy", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1), step("b", 2, { run: new Error("Word từ chối") }), step("c", 3)];
    const report = await runPipeline(steps, ctx, await planPipeline(steps, ctx));
    expect(report.steps.map((s) => [s.id, s.state])).toEqual([["a", "done"], ["b", "failed"], ["c", "notRun"]]);
    expect(report.steps[1].text).toContain("Word từ chối");
    expect(report).toMatchObject({ ok: false, failedStepId: "b", changed: true });
    expect(ctx.log).not.toContain("run:c");
  });

  it("bước tự báo bỏ qua khi chạy (không có gì để làm) được ghi là bỏ qua", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1, { run: { status: "skipped", kind: "ok", text: "không có đoạn nào" } })];
    const report = await runPipeline(steps, ctx, await planPipeline(steps, ctx));
    expect(report.steps[0].state).toBe("skipped");
    expect(report.changed).toBe(false);
  });

  it("chạy lại khi mọi bước đều đã đúng thì không chạy gì và không đổi gì", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1, { plan: { willRun: false, lines: [], skipReason: "đã đúng" } }), step("b", 2, { plan: { willRun: false, lines: [], skipReason: "đã đúng" } })];
    const report = await runPipeline(steps, ctx, await planPipeline(steps, ctx));
    expect(ctx.log.filter((l) => l.startsWith("run:"))).toEqual([]);
    expect(report).toMatchObject({ ok: true, changed: false });
  });

  it("báo tiến độ trước và sau mỗi bước", async () => {
    const ctx: Ctx = { log: [] };
    const events: string[] = [];
    const steps = [step("a", 1), step("b", 2)];
    await runPipeline(steps, ctx, await planPipeline(steps, ctx), (e) => events.push(`${e.phase}:${e.stepId}`));
    expect(events).toEqual(["start:a", "end:a", "start:b", "end:b"]);
  });

  it("không có kế hoạch truyền vào thì tự lập kế hoạch", async () => {
    const ctx: Ctx = { log: [] };
    const report = await runPipeline([step("a", 1)], ctx);
    expect(report.steps[0].state).toBe("done");
  });
});

describe("buildPreview", () => {
  it("nêu chuẩn, liệt kê bước sẽ chạy và bước bỏ qua", async () => {
    const ctx: Ctx = { log: [] };
    const planned = await planPipeline([step("a", 1), step("b", 2, { plan: { willRun: false, lines: [], skipReason: "đã đúng" } })], ctx);
    const preview = buildPreview("Đảng (HD 05)", planned);
    expect(preview.heading).toContain("Đảng (HD 05)");
    expect(preview.items).toEqual([
      { label: "Bước a", willRun: true, lines: ["làm a"] },
      { label: "Bước b", willRun: false, lines: ["Bỏ qua: đã đúng"] },
    ]);
    expect(preview.canRun).toBe(true);
    expect(preview.nothingToDo).toBe(false);
  });

  it("mọi bước đều bỏ qua thì báo không có gì để chỉnh và không cho chạy", async () => {
    const ctx: Ctx = { log: [] };
    const planned = await planPipeline([step("a", 1, { plan: { willRun: false, lines: [], skipReason: "đã đúng" } })], ctx);
    const preview = buildPreview("Chính quyền (NĐ 30)", planned);
    expect(preview).toMatchObject({ canRun: false, nothingToDo: true });
  });
});

describe("summarizeReport", () => {
  it("thành công: từng bước một dòng, mức ok", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1), step("b", 2)];
    const summary = summarizeReport(await runPipeline(steps, ctx));
    expect(summary.kind).toBe("ok");
    expect(summary.lines).toEqual(["Bước a: xong a", "Bước b: xong b"]);
  });

  it("có bước lỗi: mức error, nêu bước lỗi và các bước chưa chạy", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1), step("b", 2, { run: new Error("hỏng") }), step("c", 3)];
    const summary = summarizeReport(await runPipeline(steps, ctx));
    expect(summary.kind).toBe("error");
    expect(summary.lines[1]).toBe("Bước b: LỖI: hỏng");
    expect(summary.lines[2]).toBe("Bước c: chưa chạy");
  });

  it("có bước cảnh báo: mức warning", async () => {
    const ctx: Ctx = { log: [] };
    const steps = [step("a", 1, { run: { status: "done", kind: "warning", text: "chưa đặt được Exactly" } })];
    expect(summarizeReport(await runPipeline(steps, ctx)).kind).toBe("warning");
  });
});
