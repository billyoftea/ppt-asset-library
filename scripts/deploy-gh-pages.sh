#!/bin/bash
#
# 部署到 GitHub Pages 的脚本
#
# 使用方法:
#   1. 先设置你的 GitHub Pages URL:
#      export GITHUB_PAGES_URL="https://your-username.github.io/ppt-asset-library"
#   2. 运行此脚本:
#      bash scripts/deploy-gh-pages.sh
#
# 前提条件:
#   - 已安装 git
#   - 已关联 GitHub 远程仓库 (git remote add origin ...)
#   - 已在 GitHub 仓库 Settings → Pages 中启用 GitHub Pages (选择 gh-pages 分支)
#

set -e

# 颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${GREEN}🚀 开始部署到 GitHub Pages...${NC}"

# 检查 GITHUB_PAGES_URL 环境变量
if [ -z "$GITHUB_PAGES_URL" ]; then
  echo -e "${RED}❌ 请先设置 GITHUB_PAGES_URL 环境变量${NC}"
  echo "   export GITHUB_PAGES_URL=\"https://your-username.github.io/ppt-asset-library\""
  exit 1
fi

# 去掉末尾斜杠
GITHUB_PAGES_URL="${GITHUB_PAGES_URL%/}"

echo -e "${YELLOW}📦 GitHub Pages URL: ${GITHUB_PAGES_URL}${NC}"

# 1. 构建生产版本
echo -e "\n${GREEN}[1/5] 构建生产版本...${NC}"
npm run build

# 2. 生成 manifest.xml (替换占位符) 并放入 dist 根目录
echo -e "\n${GREEN}[2/5] 生成生产环境 manifest.xml...${NC}"
sed "s|{{GITHUB_PAGES_URL}}|${GITHUB_PAGES_URL}|g" manifest-production.xml > dist/manifest.xml
echo "   ✅ manifest.xml 已生成到 dist/（将随 GitHub Pages 一起部署）"

# 同时复制一份到项目根目录方便分发
cp dist/manifest.xml manifest-gh-pages.xml
echo "   ✅ manifest-gh-pages.xml 已复制到项目根目录"

# 3. 验证 manifest.xml 可通过 URL 访问
echo -e "\n${GREEN}[3/5] 验证部署结构...${NC}"
if [ -f "dist/manifest.xml" ] && [ -f "dist/taskpane.html" ]; then
  echo "   ✅ dist/manifest.xml   — 同事在信任中心填入 ${GITHUB_PAGES_URL}/ 即可发现插件"
  echo "   ✅ dist/taskpane.html  — 插件 UI 页面"
else
  echo -e "${RED}   ❌ 部署文件不完整，请检查构建${NC}"
  exit 1
fi

# 4. 部署 dist 到 gh-pages 分支
echo -e "\n${GREEN}[4/5] 推送 dist/ 到 gh-pages 分支...${NC}"

# 检查是否安装了 gh-pages 工具
if ! npx --no-install gh-pages --version &>/dev/null; then
  echo "   📥 安装 gh-pages 工具..."
  npm install --save-dev gh-pages
fi

npx gh-pages -d dist

echo -e "\n${GREEN}[5/5] ✅ 部署完成！${NC}"
echo ""
echo "============================================================"
echo -e "${GREEN}🎉 插件已部署到: ${GITHUB_PAGES_URL}${NC}"
echo "============================================================"
echo ""
echo "📋 同事安装方法（超简单，无需下载任何文件）："
echo ""
echo "   1. 打开 PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置"
echo "   2. 左侧选「受信任的加载项目录」"
echo "   3. 在「目录 URL」中填入："
echo ""
echo "      ${GITHUB_PAGES_URL}/"
echo ""
echo "   4. 点击「添加目录」→ 勾选「显示在菜单中」→ 确定 → 确定"
echo "   5. 重启 PowerPoint → 插入 → 获取加载项 → 共享文件夹"
echo "   6. 找到「PPT Asset Library」→ 点击添加 → 完成！🎉"
echo ""
echo "   ⚡ 第一次打开需要联网（自动缓存），之后离线也能用"
echo ""
