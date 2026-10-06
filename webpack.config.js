const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { receiveSpikeReport } = require("./tools/receiveSpikeReport");

const port = Number(process.env.npm_package_config_dev_server_port) || 3000;

async function getHttpsOptions() {
  // Chỉ nạp khi chạy dev server: lần đầu có thể hiện hộp thoại cài chứng chỉ tin cậy của Windows.
  const devCerts = require("office-addin-dev-certs");
  const { ca, key, cert } = await devCerts.getHttpsServerOptions();
  return { ca, key, cert };
}

module.exports = async (env) => {
  const serving = Boolean(env && env.WEBPACK_SERVE);
  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      spike: "./src/spike/spike.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [{ test: /\.ts$/, exclude: /node_modules/, use: "ts-loader" }],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "spike.html",
        template: "./src/spike/spike.html",
        chunks: ["spike"],
      }),
      new CopyWebpackPlugin({ patterns: [{ from: "assets", to: "assets" }] }),
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
