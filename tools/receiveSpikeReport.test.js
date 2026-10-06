const { EventEmitter } = require("node:events");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { createSpikeReportHandler, MAX_REPORT_BYTES } = require("./receiveSpikeReport");

function fakeResponse() {
  const res = { statusCode: 200, body: undefined };
  res.status = (code) => {
    res.statusCode = code;
    return res;
  };
  res.json = (body) => {
    res.body = body;
    return res;
  };
  return res;
}

function send(handler, payload) {
  const req = new EventEmitter();
  req.destroyed = false;
  req.destroy = () => {
    req.destroyed = true;
  };
  const res = fakeResponse();
  handler(req, res);
  const chunks = Array.isArray(payload) ? payload : [payload];
  for (const chunk of chunks) {
    if (req.destroyed) break;
    req.emit("data", Buffer.from(chunk, "utf8"));
  }
  if (!req.destroyed) req.emit("end");
  return { req, res };
}

describe("createSpikeReportHandler", () => {
  let docsDir;

  beforeEach(() => {
    docsDir = fs.mkdtempSync(path.join(os.tmpdir(), "spike-"));
  });

  afterEach(() => {
    fs.rmSync(docsDir, { recursive: true, force: true });
  });

  it("ghi docs/spike-results.json và .md khi báo cáo hợp lệ, giữ nguyên chữ có dấu", () => {
    const handler = createSpikeReportHandler(docsDir);
    const report = { generatedAt: "2026-10-06", results: [{ id: "p02", detail: "Lề trái đúng" }] };
    const { res } = send(handler, JSON.stringify({ report, markdown: "# Kết quả\n" }));
    expect(res.statusCode).toBe(200);
    expect(res.body).toEqual({ ok: true });
    expect(JSON.parse(fs.readFileSync(path.join(docsDir, "spike-results.json"), "utf8"))).toEqual(report);
    expect(fs.readFileSync(path.join(docsDir, "spike-results.md"), "utf8")).toBe("# Kết quả\n");
  });

  it("nhận được báo cáo gửi thành nhiều mảnh", () => {
    const handler = createSpikeReportHandler(docsDir);
    const whole = JSON.stringify({ report: { a: 1 }, markdown: "ok" });
    const { res } = send(handler, [whole.slice(0, 10), whole.slice(10)]);
    expect(res.statusCode).toBe(200);
  });

  it("từ chối JSON hỏng và không ghi file", () => {
    const handler = createSpikeReportHandler(docsDir);
    const { res } = send(handler, "{không phải json");
    expect(res.statusCode).toBe(400);
    expect(fs.existsSync(path.join(docsDir, "spike-results.json"))).toBe(false);
  });

  it("từ chối dữ liệu thiếu markdown hoặc report", () => {
    const handler = createSpikeReportHandler(docsDir);
    expect(send(handler, JSON.stringify({ report: {} })).res.statusCode).toBe(400);
    expect(send(handler, JSON.stringify({ markdown: "x" })).res.statusCode).toBe(400);
    expect(send(handler, JSON.stringify({ report: null, markdown: "x" })).res.statusCode).toBe(400);
    expect(fs.readdirSync(docsDir)).toEqual([]);
  });

  it("từ chối dữ liệu quá lớn và ngắt kết nối", () => {
    const handler = createSpikeReportHandler(docsDir);
    const { req, res } = send(handler, "x".repeat(MAX_REPORT_BYTES + 1));
    expect(res.statusCode).toBe(413);
    expect(req.destroyed).toBe(true);
    expect(fs.readdirSync(docsDir)).toEqual([]);
  });

  it("chỉ ghi vào hai file cố định trong thư mục docs, không theo đường dẫn do client gửi", () => {
    const handler = createSpikeReportHandler(docsDir);
    send(handler, JSON.stringify({ report: {}, markdown: "m", path: "../../evil.txt", file: "../evil" }));
    expect(fs.readdirSync(docsDir).sort()).toEqual(["spike-results.json", "spike-results.md"]);
  });
});
