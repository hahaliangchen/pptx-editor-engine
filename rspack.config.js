const path = require("path");
const rspack = require("@rspack/core");

module.exports = {
  entry: "./src/demo-entry.ts",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "index.js",
    library: {
      name: "PptEditor",
      type: "umd",
      export: "default",
    },
    globalObject: "globalThis",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".js", ".wasm"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: [/node_modules/],
        use: {
          loader: "builtin:swc-loader",
          options: {
            jsc: {
              parser: {
                syntax: "typescript",
              },
            },
          },
        },
        type: "javascript/auto",
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
        type: "javascript/auto",
      },
    ],
  },
  plugins: [
    new rspack.HtmlRspackPlugin({
      template: "./index.html",
    }),
  ],
  experiments: {
    asyncWebAssembly: true,
  },
  devServer: {
    port: process.env.PORT ? Number(process.env.PORT) : 3000,
    hot: true,
  },
};
