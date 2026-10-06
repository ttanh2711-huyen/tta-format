const fs = require("node:fs");
const path = require("node:path");

const MAX_REPORT_BYTES = 1_000_000;

/**
 * Chỉ dùng khi phát triển: nhận báo cáo spike T3 từ task pane (cùng nguồn https://localhost)
 * và ghi vào hai file cố định trong `docsDir`. Không theo đường dẫn nào do client gửi.
 */
function createSpikeReportHandler(docsDir) {
  return (req, res) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_REPORT_BYTES) {
        res.status(413).json({ ok: false });
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => {
      if (size > MAX_REPORT_BYTES) {
        return;
      }
      try {
        const { report, markdown } = JSON.parse(Buffer.concat(chunks).toString("utf8"));
        if (typeof markdown !== "string" || typeof report !== "object" || report === null) {
          res.status(400).json({ ok: false });
          return;
        }
        fs.mkdirSync(docsDir, { recursive: true });
        fs.writeFileSync(path.join(docsDir, "spike-results.json"), `${JSON.stringify(report, null, 2)}\n`);
        fs.writeFileSync(path.join(docsDir, "spike-results.md"), markdown);
        res.json({ ok: true });
      } catch {
        res.status(400).json({ ok: false });
      }
    });
  };
}

/** Dùng làm `devServer.setupMiddlewares` của webpack. */
function receiveSpikeReport(docsDir) {
  return (middlewares, devServer) => {
    devServer.app.post("/spike-report", createSpikeReportHandler(docsDir));
    return middlewares;
  };
}

module.exports = { createSpikeReportHandler, receiveSpikeReport, MAX_REPORT_BYTES };
