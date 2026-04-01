/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const urlDev = "https://localhost:3000";

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  // 仅在开发模式下尝试加载 HTTPS 证书
  let devServerOptions = {};
  if (dev) {
    try {
      const devCerts = require("office-addin-dev-certs");
      devServerOptions = await devCerts.getHttpsServerOptions();
    } catch {
      console.warn("⚠️  Could not load dev certs, using default HTTPS config");
      devServerOptions = {};
    }
  }

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json"],
      alias: {
        "@": path.resolve(__dirname, "src"),
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/thumbnails",
            to: "assets/thumbnails",
            noErrorOnMissing: true,
          },
          {
            from: "assets/hd",
            to: "assets/hd",
            noErrorOnMissing: true,
          },
          {
            from: "assets/source.pptx",
            to: "assets/source.pptx",
            noErrorOnMissing: true,
          },
          {
            from: "assets/slides",
            to: "assets/slides",
            noErrorOnMissing: true,
          },
          {
            from: "src/data",
            to: "data",
            noErrorOnMissing: true,
          },
          // Service Worker — 离线缓存支持
          {
            from: "src/service-worker.js",
            to: "service-worker.js",
          },
          // 图标资源（manifest.xml 中引用）
          {
            from: "assets/icon-16.png",
            to: "assets/icon-16.png",
            noErrorOnMissing: true,
          },
          {
            from: "assets/icon-32.png",
            to: "assets/icon-32.png",
            noErrorOnMissing: true,
          },
          {
            from: "assets/icon-64.png",
            to: "assets/icon-64.png",
            noErrorOnMissing: true,
          },
          {
            from: "assets/icon-80.png",
            to: "assets/icon-80.png",
            noErrorOnMissing: true,
          },
          // bundle.pptx（素材源文件）
          {
            from: "assets/bundle.pptx",
            to: "assets/bundle.pptx",
            noErrorOnMissing: true,
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: devServerOptions,
      },
      port: 3000,
      static: [
        {
          directory: path.resolve(__dirname, "assets"),
          publicPath: "/assets",
        },
        {
          directory: path.resolve(__dirname, "src/data"),
          publicPath: "/data",
        },
      ],
      // bundle.pptx 约 2.6MB
      client: {
        overlay: true,
      },
      hot: true,
    },
  };
};
