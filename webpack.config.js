/* eslint-disable no-undef */

const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";

// Each production target maps to its own manifest source file (kept as
// separate manifest*.xml files rather than templating one file, since
// AppDomains/WebApplicationInfo differ per Azure AD app registration).
const targets = {
  freshminds: {
    urlProd: "https://researchcloud.app/msoffice-addin/",
    manifestFile: "manifest.xml",
  },
  verityone: {
    urlProd: "https://d2yq86j6lz79ng.cloudfront.net/msoffice-addin/",
    manifestFile: "manifest.verityone.xml",
  },
};

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const target = targets[env.target] ? env.target : "freshminds";
  const { urlProd, manifestFile } = targets[target];
  // In dev mode, ADDIN_ENV picks which backend (Cognito/API) the dev-server
  // talks to — independent of `target`, which only drives the prod URL/manifest
  // substitution above. Defaults to "development" (personal account) unless
  // an explicit --env target is passed, e.g. `--env target=verityone`.
  const addinEnv = dev ? env.target || "development" : target;
  const config = {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
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
      new webpack.DefinePlugin({
        "process.env.ADDIN_ENV": JSON.stringify(addinEnv),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "callback.html",
        template: "./src/taskpane/callback.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "ms-callback.html",
        template: "./src/taskpane/ms-callback.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: dev ? "manifest.xml" : manifestFile,
            to: "manifest.xml",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
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
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
