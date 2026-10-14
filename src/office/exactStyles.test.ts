import { buildStyleSpecs } from "../ooxml/styleSpecs";
import { HD05 } from "../standards/hd05/profile";
import { ND30 } from "../standards/nd30/profile";
import { planExactRepair } from "./exactStyles";
import { readbackFromSpec, type StyleReadback } from "./styleVerify";

const hd = buildStyleSpecs(HD05);
const noiDung = hd.find((s) => s.name === "HD05_NoiDung")!;
const missing: StyleReadback = { ...readbackFromSpec(noiDung), exists: false };
const atLeast: StyleReadback = { ...readbackFromSpec(noiDung), lineRule: "atLeast" };
const maps = (rb: StyleReadback) => new Map(hd.map((s) => [s.name, s.name === noiDung.name ? rb : readbackFromSpec(s)]));

describe("kế hoạch sửa giãn dòng Exactly của style", () => {
  it("chuẩn Chính quyền (giãn dòng bội số) không có gì phải dựng lại", () => {
    const nd = buildStyleSpecs(ND30);
    const plan = planExactRepair(nd, new Map(nd.map((s) => [s.name, readbackFromSpec(s)])), new Set(), false);
    expect(plan).toEqual({ rebuild: [], blocked: [] });
  });

  it("style Exactly đã đúng thì giữ nguyên", () => {
    expect(planExactRepair(hd, maps(readbackFromSpec(noiDung)), new Set(), false)).toEqual({ rebuild: [], blocked: [] });
  });

  it("style chưa có thì dựng mới bằng OOXML, kể cả khi đang có đoạn trỏ tới tên đó", () => {
    expect(planExactRepair(hd, maps(missing), new Set(["HD05NoiDung"]), false).rebuild.map((s) => s.name)).toEqual(["HD05_NoiDung"]);
  });

  it("style đang sai quy tắc và chưa ai dùng thì xoá rồi dựng lại", () => {
    expect(planExactRepair(hd, maps(atLeast), new Set(), false)).toEqual({ rebuild: [noiDung], blocked: [] });
  });

  it("style sai quy tắc nhưng đang có đoạn dùng thì giữ nguyên và báo bị chặn, trừ khi thân sắp được thay", () => {
    const used = new Set(["HD05NoiDung"]);
    expect(planExactRepair(hd, maps(atLeast), used, false)).toEqual({ rebuild: [], blocked: [noiDung] });
    expect(planExactRepair(hd, maps(atLeast), used, true)).toEqual({ rebuild: [noiDung], blocked: [] });
  });
});
