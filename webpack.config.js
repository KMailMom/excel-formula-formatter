/* eslint-disable no-undef */
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

// Development URL (webpack-dev-server). Do not change.
const urlDev = "https://localhost:3000/";
// GitHub repository ("owner/name"). In GitHub Actions it comes from GITHUB_REPOSITORY,
// locally from the `origin` remote, so forks build with their own URLs automatically.
function githubRepo() {
  if (process.env.GITHUB_REPOSITORY) return process.env.GITHUB_REPOSITORY;
  try {
    const remote = execSync("git remote get-url origin", { encoding: "utf8" }).trim();
    const m = /github\.com[/:]([^/]+\/[^/]+?)(?:\.git)?$/.exec(remote);
    if (m) return m[1];
  } catch (e) {
    // no git remote; fall through
  }
  throw new Error("Cannot determine the GitHub repository. Set GITHUB_REPOSITORY=owner/name.");
}

// Production URL: the GitHub Pages site of the repository. Must end with a slash.
function prodUrls() {
  const [owner, name] = githubRepo().split("/");
  return {
    base: `https://${owner.toLowerCase()}.github.io/${name}/`,
    repo: `https://github.com/${owner}/${name}`,
  };
}

// Add-in Id, read from the manifest so the install scripts always match it.
const addinId = /<Id>([^<]+)<\/Id>/.exec(fs.readFileSync(path.join(__dirname, "manifest.xml"), "utf8"))[1];

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

// Wraps a PowerShell script in a .cmd file that runs it on double-click (CRLF line endings).
const cmdWrapper = fs.readFileSync(path.join(__dirname, "install", "cmd-wrapper.txt"), "utf8");
function toCmd(psScript, title) {
  return cmdWrapper
    .replace("__TITLE__", title)
    .replace("__PS_SCRIPT__", () => psScript)
    .replace(/\r?\n/g, "\r\n");
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const prod = dev ? null : prodUrls();
  const baseUrl = dev ? urlDev : prod.base;
  const repoUrl = dev ? "https://github.com/" : prod.repo;
  // Fills the placeholders in the landing page and install scripts.
  const fillPlaceholders = (content) =>
    content.toString().replace(/__BASE_URL__/g, baseUrl).replace(/__ADDIN_ID__/g, addinId).replace(/__REPO_URL__/g, repoUrl);
  const config = {
    devtool: "source-map",
    entry: {
      // Shared runtime: the task pane bundle also contains the command handlers.
      taskpane: "./src/taskpane/taskpane.ts",
      // Browser version of the formatter (online.html), same formatter code.
      web: "./src/web/web.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "online.html",
        template: "./src/web/online.html",
        chunks: ["web"],
        templateParameters: { repoUrl },
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", globOptions: { ignore: ["**/*.svg", "**/*.mjs"] } },
          { from: "src/dialog/dialog.html", to: "dialog.html" },
          { from: "shortcuts.json", to: "shortcuts.json" },
          // Landing, privacy and support pages, plus the one-line installers.
          { from: "site", to: ".", transform: fillPlaceholders },
          { from: "docs/images", to: "images" },
          { from: "install", to: ".", transform: fillPlaceholders, globOptions: { ignore: ["**/cmd-wrapper.txt"] } },
          // Double-clickable Windows versions of the PowerShell scripts (same code, wrapped).
          { from: "install/install.ps1", to: "install.cmd", transform: (c) => toCmd(fillPlaceholders(c), "Formula Formatter for Excel - installer") },
          { from: "install/uninstall.ps1", to: "uninstall.cmd", transform: (c) => toCmd(fillPlaceholders(c), "Formula Formatter for Excel - uninstaller") },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), baseUrl);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: env && env.WEBPACK_BUILD ? {} : await getHttpsOptions(),
      },
      port: 3000,
    },
  };
  return config;
};
