const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const { version } = require("./package.json");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { receiveSpikeReport } = require("./tools/receiveSpikeReport");

const OFFICE_SCRIPT = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>';
const styles = () => fs.readFileSync(path.resolve(__dirname, "src/taskpane/taskpane.css"), "utf8");

const port = Number(process.env.npm_package_config_dev_server_port) || 3000;

async function getHttpsOptions() {
  // Chỉ nạp khi chạy dev server: lần đầu có thể hiện hộp thoại cài chứng chỉ tin cậy của Windows.
  const devCerts = require("office-addin-dev-certs");
  const { ca, key, cert } = await devCerts.getHttpsServerOptions();
  return { ca, key, cert };
}

module.exports = async (env) => {
  const serving = Boolean(env && env.WEBPACK_SERVE);
  // Bản phát hành không gồm trang thăm dò API (spike) dùng khi phát triển.
  const release = Boolean(env && env.release);
  return {
    entry: release ? { taskpane: "./src/taskpane/taskpane.ts" } : { taskpane: "./src/taskpane/taskpane.ts", spike: "./src/spike/spike.ts", preview: "./src/taskpane/preview.ts" },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".js"] },
    // Add-in chạy từ máy hoặc trang tĩnh, không cần cảnh báo dung lượng gói của web công cộng.
    performance: { hints: false },
    module: {
      rules: [{ test: /\.ts$/, exclude: /node_modules/, use: "ts-loader" }],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        templateParameters: { officeScript: OFFICE_SCRIPT, styles: styles() },
      }),
      ...(release
        ? []
        : [
            // Bản xem trước giao diện ngoài Word (Office giả), chỉ có khi phát triển.
            new HtmlWebpackPlugin({
              filename: "preview.html",
              template: "./src/taskpane/taskpane.html",
              chunks: ["preview"],
              templateParameters: { officeScript: "", styles: styles() },
            }),
            new HtmlWebpackPlugin({
              filename: "spike.html",
              template: "./src/spike/spike.html",
              chunks: ["spike"],
            }),
          ]),
      new CopyWebpackPlugin({ patterns: [{ from: "assets", to: "assets" }] }),
      new webpack.DefinePlugin({ __APP_VERSION__: JSON.stringify(version) }),
    ],
    devServer: serving
      ? {
          headers: { "Access-Control-Allow-Origin": "*" },
          server: { type: "https", options: await getHttpsOptions() },
          setupMiddlewares: receiveSpikeReport(path.resolve(__dirname, "docs")),
          port,
        }
      : undefined,
  };
};
