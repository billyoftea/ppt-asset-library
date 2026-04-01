/**
 * 生成生产环境 manifest.xml
 * 
 * 用法:
 *   GITHUB_PAGES_URL=https://your-username.github.io/ppt-asset-library node scripts/generate-manifest.js
 */

const fs = require("fs");
const path = require("path");

const GITHUB_PAGES_URL = process.env.GITHUB_PAGES_URL;

if (!GITHUB_PAGES_URL) {
  console.error("❌ 请设置 GITHUB_PAGES_URL 环境变量");
  console.error('   export GITHUB_PAGES_URL="https://your-username.github.io/ppt-asset-library"');
  process.exit(1);
}

// 去掉末尾斜杠
const baseUrl = GITHUB_PAGES_URL.replace(/\/+$/, "");

const templatePath = path.resolve(__dirname, "..", "manifest-production.xml");
const outputPath = path.resolve(__dirname, "..", "dist", "manifest.xml");
const rootOutputPath = path.resolve(__dirname, "..", "manifest-gh-pages.xml");

// 读取模板
let content = fs.readFileSync(templatePath, "utf-8");

// 替换占位符
content = content.replace(/\{\{GITHUB_PAGES_URL\}\}/g, baseUrl);

// 确保 dist 目录存在
const distDir = path.dirname(outputPath);
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

// 写入 dist
fs.writeFileSync(outputPath, content, "utf-8");
console.log(`✅ manifest.xml → ${outputPath}`);

// 写入项目根目录
fs.writeFileSync(rootOutputPath, content, "utf-8");
console.log(`✅ manifest-gh-pages.xml → ${rootOutputPath}`);
console.log(`\n📋 GitHub Pages URL: ${baseUrl}`);
console.log("📋 把 manifest-gh-pages.xml 发给同事即可使用！");
