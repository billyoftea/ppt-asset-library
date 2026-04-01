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
echo -e "\n${GREEN}[1/4] 构建生产版本...${NC}"
npm run build

# 2. 生成 manifest.xml (替换占位符)
echo -e "\n${GREEN}[2/4] 生成生产环境 manifest.xml...${NC}"
sed "s|{{GITHUB_PAGES_URL}}|${GITHUB_PAGES_URL}|g" manifest-production.xml > dist/manifest.xml
echo "   ✅ manifest.xml 已生成到 dist/"

# 同时复制一份到项目根目录方便分发
cp dist/manifest.xml manifest-gh-pages.xml
echo "   ✅ manifest-gh-pages.xml 已复制到项目根目录（可直接发给同事）"

# 3. 部署 dist 到 gh-pages 分支
echo -e "\n${GREEN}[3/4] 推送 dist/ 到 gh-pages 分支...${NC}"

# 检查是否安装了 gh-pages 工具
if ! npx --no-install gh-pages --version &>/dev/null; then
  echo "   📥 安装 gh-pages 工具..."
  npm install --save-dev gh-pages
fi

npx gh-pages -d dist

echo -e "\n${GREEN}[4/4] ✅ 部署完成！${NC}"
echo ""
echo "================================================="
echo -e "${GREEN}🎉 插件已部署到: ${GITHUB_PAGES_URL}${NC}"
echo "================================================="
echo ""
echo "📋 接下来："
echo "   1. 把 manifest-gh-pages.xml 发给同事"
echo "   2. 同事在 PowerPoint 中: 插入 → 获取加载项 → 上传我的加载项 → 选择 manifest-gh-pages.xml"
echo "   3. 第一次打开插件需要联网（自动缓存所有资源）"
echo "   4. 之后即使离线也能正常使用 ✅"
echo ""
