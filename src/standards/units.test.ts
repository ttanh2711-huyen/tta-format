import { mmToPt, ptToMm } from "./units";

describe("đổi đơn vị mm và pt", () => {
  it("25,4 mm bằng 72 pt (1 inch)", () => {
    expect(mmToPt(25.4)).toBeCloseTo(72, 6);
    expect(ptToMm(72)).toBeCloseTo(25.4, 6);
  });

  it("lề trái 30 mm bằng khoảng 85,04 pt", () => {
    expect(mmToPt(30)).toBeCloseTo(85.0394, 3);
  });

  it("đổi qua lại không làm lệch giá trị", () => {
    expect(ptToMm(mmToPt(20))).toBeCloseTo(20, 9);
  });
});
