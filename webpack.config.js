const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const { version } = require("./package.json");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { receiveSelftest } = require("./tools/receiveSelftest");

const OFFICE_SCRIPT = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>';
/**
 * Trang task pane dựng bằng cách thay hai chỗ trống trong taskpane.html (`<%= officeScript %>`, `<%= styles %>`) ngay tại đây,
 * không qua bộ nạp mẫu của html-webpack-plugin: hai trang (taskpane, preview) dùng chung một mẫu với tham số khác nhau nên
 * bộ nạp mẫu báo "officeScript is not defined" khi chạy dev server.
 */
const taskpaneHtml = (officeScript) => () =>
  fs
    .readFileSync(path.resolve(__dirname, "src/taskpane/taskpane.html"), "utf8")
    .replace("<%= officeScript %>", () => officeScript)
    .replace("<%= styles %>", () => fs.readFileSync(path.resolve(__dirname, "src/taskpane/taskpane.css"), "utf8"));

const port = Number(process.env.npm_package_config_dev_server_port) || 3000;

async function getHttpsOptions() {
  // Chỉ nạp khi chạy dev server: lần đầu có thể hiện hộp thoại cài chứng chỉ tin cậy của Windows.
  const devCerts = require("office-addin-dev-certs");
  const { ca, key, cert } = await devCerts.getHttpsServerOptions();
  return { ca, key, cert };
}

module.exports = async (env) => {
  const serving = Boolean(env && env.WEBPACK_SERVE);
  // Bản phát hành không gồm các trang chỉ dùng khi phát triển (xem trước giao diện, tự kiểm thử trong Word).
  const release = Boolean(env && env.release);
  return {
    entry: release ? { taskpane: "./src/taskpane/taskpane.ts" } : { taskpane: "./src/taskpane/taskpane.ts", preview: "./src/taskpane/preview.ts", selftest: "./src/selftest/selftest.ts" },
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
        templateContent: taskpaneHtml(OFFICE_SCRIPT),
        chunks: ["taskpane"],
      }),
      ...(release
        ? []
        : [
            // Bản xem trước giao diện ngoài Word (Office giả), chỉ có khi phát triển.
            new HtmlWebpackPlugin({
              filename: "preview.html",
              templateContent: taskpaneHtml(""),
              chunks: ["preview"],
            }),
            new HtmlWebpackPlugin({
              filename: "selftest.html",
              template: "./src/selftest/selftest.html",
              chunks: ["selftest"],
            }),
          ]),
      new CopyWebpackPlugin({ patterns: [{ from: "assets", to: "assets" }] }),
      new webpack.DefinePlugin({ __APP_VERSION__: JSON.stringify(version) }),
    ],
    devServer: serving
      ? {
          headers: { "Access-Control-Allow-Origin": "*" },
          server: { type: "https", options: await getHttpsOptions() },
          setupMiddlewares: receiveSelftest(),
          port,
        }
      : undefined,
  };
};
