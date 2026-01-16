/* eslint-disable no-undef */

const path = require("path");
const webpack = require("webpack");
const dotenv = require("dotenv");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

dotenv.config();

function requireEnv(name, fallback) {
  const value = process.env[name];
  if (value) {
    return value;
  }
  if (fallback) {
    return fallback;
  }
  throw new Error(`Missing required environment variable: ${name}`);
}

const urlProd = requireEnv("APP_BASE_URL").replace(/\/+$/, "");
const urlDev = requireEnv("APP_BASE_URL_DEV", urlProd).replace(/\/+$/, "");
const defineEnv = {
  "process.env.AAD_CLIENT_ID": JSON.stringify(requireEnv("AAD_CLIENT_ID")),
  "process.env.AAD_AUTHORITY": JSON.stringify(requireEnv("AAD_AUTHORITY")),
  "process.env.GRAPH_BASE_URL": JSON.stringify(requireEnv("GRAPH_BASE_URL")),
  "process.env.GRAPH_SCOPES": JSON.stringify(requireEnv("GRAPH_SCOPES")),
  "process.env.APP_BASE_URL": JSON.stringify(requireEnv("APP_BASE_URL")),
  "process.env.APP_BASE_URL_DEV": JSON.stringify(requireEnv("APP_BASE_URL_DEV", urlProd)),
};

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      createEventDialog: "./src/dialogs/create-event.js",
      authDialog: "./src/dialogs/auth.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin(defineEnv),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "public/index.html",
            to: "index.html",
          },
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              }
              return content.toString().replace(new RegExp(escapeRegExp(urlDev), "g"), urlProd);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "create-event.html",
        template: "./src/dialogs/create-event.html",
        chunks: ["polyfill", "createEventDialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/dialogs/auth.html",
        chunks: ["polyfill", "authDialog"],
      }),
    ],
    devServer: {
      allowedHosts: "all",
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
