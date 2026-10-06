import { probesForRunAll } from "./selection";

describe("probesForRunAll", () => {
  const probes = [
    { id: "p02" },
    { id: "p12b", risky: true },
    { id: "p13", risky: false },
  ];

  it("bỏ các phép thử đánh dấu rủi ro khỏi lượt chạy tất cả, giữ nguyên thứ tự", () => {
    expect(probesForRunAll(probes).map((p) => p.id)).toEqual(["p02", "p13"]);
  });

  it("không đổi danh sách gốc", () => {
    probesForRunAll(probes);
    expect(probes).toHaveLength(3);
  });
});
