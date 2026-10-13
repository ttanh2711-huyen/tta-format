const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

/** Chỉ dùng khi phát triển: nhận kết quả tự kiểm thử trong Word và ghi vào thư mục tạm của máy, không ghi vào repo. */
function receiveSelftest() {
  return (middlewares, devServer) => {
    devServer.app.post("/selftest-report", (req, res) => {
      const chunks = [];
      req.on("data", (chunk) => chunks.push(chunk));
      req.on("end", () => {
        fs.writeFileSync(path.join(os.tmpdir(), "tta-selftest.json"), Buffer.concat(chunks));
        res.json({ ok: true });
      });
    });
    return middlewares;
  };
}

module.exports = { receiveSelftest };
