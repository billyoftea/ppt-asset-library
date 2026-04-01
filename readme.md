# PPT 素材库 Office Add-in 插件

## 项目背景

将 PPT 中积累的图形素材库（500+ 页）以 **Office Add-in 任务窗格插件** 的形式封装，方便在制作 PPT 时通过侧边栏快速浏览、搜索并一键插入素材。

---

## 素材库概览

当前素材来源：`Graphics, elements, and icon's – 500+.pptx`（~66MB，565 张幻灯片，5300+ 媒体文件）

### 素材分类（从源文件提取）

| 分类 | 数量 | 说明 |
|------|------|------|
| **Circular Diagrams** | ~59 页 | 饼图、环形图、圆形流程图等 |
| **Weeble（人形图标）** | ~68 页 | 各种人形/角色插画 |
| **Process Charts** | ~21 页 | 流程图、步骤图 |
| **Maps** | ~27 页 | 世界地图、美国地图、区域地图、地图联系人等 |
| **Org Charts** | ~20 页 | 组织架构图、层级图 |
| **Icons** | ~36 页 | 通用图标、社交媒体图标、实心/线框图标等 |
| **3-Dimensional** | ~15 页 | 3D 立体图形 |
| **Arrows** | ~10 页 | 各类箭头和方向指示 |
| **Infographics** | ~19 页 | 信息图、数据可视化排版 |
| **Timelines** | ~10 页 | 时间轴、里程碑、阶段图 |
| **Other Shapes** | ~17 页 | 六边形、齿轮、雷达图、蜘蛛图等 |
| **Tables** | ~6 页 | 表格样式模板 |
| **Puzzles** | ~6 页 | 拼图类图形 |
| **Quotes / Bubbles** | ~2 页 | 引用框、对话气泡 |
| **Newspapers** | ~2 页 | 报纸风格排版 |
| **其他** | ~50+ 页 | 漏斗、路线图、平衡图、评分卡等 |

---

## 功能需求

### MVP（v1.0）

- [ ] **侧边栏任务窗格（Task Pane）**：在 PowerPoint 中打开一个侧边面板
- [ ] **素材分类浏览**：按上述分类展示缩略图网格，支持分类折叠/展开
- [ ] **搜索功能**：按关键词搜索素材（匹配分类名、标签）
- [ ] **预览**：点击缩略图查看大图预览
- [ ] **一键插入**：点击插入按钮，将选中的素材（图形/形状组）插入到当前幻灯片

### v1.1（增强）

- [ ] **收藏夹**：标记常用素材，快速访问
- [ ] **最近使用**：记录最近插入的素材
- [ ] **多选批量插入**：支持同时选择多个素材
- [ ] **拖拽插入**：支持从面板拖拽到幻灯片
- [ ] **自定义素材**：支持用户上传/导入自己的素材

### v2.0（进阶）

- [ ] **素材编辑**：插入后支持快速换色、缩放等
- [ ] **模板组合**：预设常用素材组合，一键插入整个版式
- [ ] **团队共享**：通过云端同步素材库，团队成员共享
- [ ] **AI 推荐**：根据当前幻灯片内容智能推荐合适的素材

---

## 技术方案

### 技术栈

| 层次 | 技术选型 | 说明 |
|------|----------|------|
| **Add-in 框架** | Office Add-in (Web-based) | 使用 Yeoman Generator (`yo office`) 脚手架 |
| **清单文件** | `manifest.xml` | 定义插件元信息、权限、任务窗格入口 |
| **前端框架** | React + TypeScript | 构建任务窗格 UI |
| **UI 组件库** | Fluent UI React v9 | 与 Office 风格保持一致 |
| **Office API** | Office.js (PowerPoint API) | 操作幻灯片、插入内容 |
| **构建工具** | Webpack / Vite | 打包前端资源 |
| **素材存储** | 本地静态资源 / CDN | 缩略图 + 原始矢量数据 |

### 项目结构（预期）

```
add-in_test/
├── manifest.xml                  # Office Add-in 清单文件
├── package.json
├── webpack.config.js
├── src/
│   ├── taskpane/
│   │   ├── index.tsx             # 任务窗格入口
│   │   ├── App.tsx               # 主应用组件
│   │   ├── components/
│   │   │   ├── CategoryList.tsx  # 分类列表
│   │   │   ├── AssetGrid.tsx     # 素材缩略图网格
│   │   │   ├── AssetPreview.tsx  # 大图预览弹窗
│   │   │   ├── SearchBar.tsx     # 搜索栏
│   │   │   └── Favorites.tsx     # 收藏夹
│   │   └── styles/
│   │       └── taskpane.css
│   ├── services/
│   │   ├── officeApi.ts          # Office.js 封装（插入图形等）
│   │   └── assetManager.ts       # 素材数据管理（索引、搜索）
│   └── data/
│       └── assetIndex.json       # 素材索引（分类、标签、缩略图路径）
├── assets/
│   ├── thumbnails/               # 素材缩略图（PNG, 150x150）
│   └── vectors/                  # 原始矢量素材（SVG / EMF / 形状XML）
├── scripts/
│   ├── extract_assets.py         # 从源PPTX提取素材为独立文件
│   ├── generate_thumbnails.py    # 生成缩略图
│   └── build_index.py            # 生成素材索引JSON
└── README.md
```

### 核心 API 使用

```typescript
// 插入图片到当前幻灯片
async function insertImage(base64Image: string) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.addImage(base64Image);
    await context.sync();
  });
}

// 插入预定义形状组（从XML还原）
async function insertShapeGroup(shapeXml: string) {
  // 通过 Office.js 的 OOXML 插入功能还原形状
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    // 使用 insertSlidesFromBase64 或 setSelectedDataAsync
    await context.sync();
  });
}
```

---

## 素材预处理流程

将源 PPTX 中的 565 页素材转换为插件可用的格式：

```
源PPTX ──> 解包(unzip) ──> 分析slide XML ──> 提取形状/图片
                                    │
                                    ├──> 生成缩略图 (PNG 150px)
                                    ├──> 导出矢量素材 (SVG / Shape XML)
                                    ├──> 提取分类标签 (from slide text)
                                    └──> 构建索引 JSON (assetIndex.json)
```

### 预处理脚本说明

| 脚本 | 功能 |
|------|------|
| `extract_assets.py` | 遍历每张幻灯片，将图形组/形状/图片提取为独立文件 |
| `generate_thumbnails.py` | 将每个素材渲染为 150×150 PNG 缩略图 |
| `build_index.py` | 根据幻灯片文本和结构，生成分类/标签/路径索引 |

---

## 开发计划

### Phase 1：素材预处理（1-2 天） ✅ 已完成

1. ✅ 编写脚本从源 PPTX 解包并分析 slide XML（561 张素材页）
2. ✅ 提取每张幻灯片的主要图形元素为独立文件（含形状 XML 和媒体引用）
3. ✅ 生成缩略图（562 张，含媒体提取和占位图）
4. ✅ 构建素材索引 JSON（14 个分类，317KB 完整索引 + 177KB 精简版）

### Phase 2：插件框架搭建（1-2 天） ✅ 已完成

1. ✅ 创建 PowerPoint 任务窗格项目（React + TypeScript + Webpack）
2. ✅ 配置 `manifest.xml`（含 Ribbon 按钮、任务窗格入口、HTTPS 开发证书）
3. ✅ 搭建 React 组件结构（App / SearchBar / CategoryList / AssetGrid / AssetPreview）
4. ✅ 集成 Fluent UI React v9 组件库 + Office.js API 封装

### Phase 3：核心功能开发（3-5 天） ✅ 已完成

1. ✅ 实现分类浏览 UI（缩略图网格 + 分类折叠/展开 + 分类排序 + Tooltip 提示）
2. ✅ 实现搜索功能（防抖搜索 + 多词 AND 匹配 + 搜索缓存预构建 + 结果计数 + Enter/Escape 快捷键）
3. ✅ 实现大图预览（Dialog 弹窗 + 前后导航 + 键盘左右箭头 + 计数器）
4. ✅ 对接 Office.js 实现一键插入（shapes.addImage 首选 + setSelectedDataAsync 回退 + 环境检测 + 超时处理）
5. ✅ 处理不同素材类型的插入逻辑（图片插入 + 整页幻灯片插入 + PPTX base64 加载 + 多路径图片加载）

### Phase 4：测试和优化（1-2 天） ✅ 已完成

1. ✅ 在 PowerPoint Desktop 上侧载测试（macOS，`npm run sideload` 直接启动 PPT + 插件）
2. ✅ 素材加载性能优化（懒加载缩略图 + 无限滚动分页 + 搜索缓存预构建）
3. ✅ 高清插入图（960×720）替代缩略图（200×150），插入质量提升 ~22 倍
4. ✅ 整页矢量插入（`insertSlidesFromBase64`）保留原生可编辑性
5. ✅ 生成 Ribbon 图标（16/32/64/80px），manifest 验证通过
6. ✅ HTTPS 证书配置完成，所有资源端点可用

---

## 运行环境要求

- **Node.js** ≥ 18
- **Python** ≥ 3.10（预处理脚本）
- **PowerPoint**：Microsoft 365 桌面版 或 PowerPoint Online
- **浏览器**：Edge / Chrome（用于调试任务窗格）

---

## 快速开始

```bash
# 1. 安装 Yeoman 和 Office 脚手架
npm install -g yo generator-office

# 2. 生成项目（或从本仓库直接开发）
yo office --projectType taskpane --name "PPT Asset Library" --host powerpoint --ts true

# 3. 安装依赖
npm install

# 4. 运行预处理脚本（提取素材）
python scripts/extract_assets.py "Graphics, elements, and icon's – 500+.pptx"

# 5. 启动开发服务器
npm start

# 6. 在 PowerPoint 中侧载插件进行测试
# 参考：https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins
```

---

## 🚀 部署到 GitHub Pages

本插件托管在 GitHub Pages 上，内置 **Service Worker 离线缓存**，第一次联网使用后离线也能正常工作。

### 发布/更新素材

```bash
# 设置 GitHub Pages URL
export GITHUB_PAGES_URL="https://billyoftea.github.io/ppt-asset-library"

# 一键构建并部署
npm run deploy
```

部署脚本会自动构建、生成 manifest、推送到 `gh-pages` 分支。

---

## 📥 安装指南

插件内容托管在 GitHub Pages，**manifest.xml 是唯一需要下载的文件**。

### 下载 manifest.xml

浏览器打开以下地址，右键保存文件：

```
https://billyoftea.github.io/ppt-asset-library/manifest.xml
```

---

### Windows 安装

1. 在「文档」目录下新建一个文件夹，例如：

   ```
   C:\Users\你的用户名\Documents\OfficeAddins
   ```

2. 将下载的 `manifest.xml` 放进这个文件夹

3. 打开 PowerPoint → **文件** → **选项** → **信任中心** → **信任中心设置**

4. 左侧选择 **「受信任的加载项目录」**

5. 在 **「目录 URL」** 中填入刚才的文件夹路径：

   ```
   C:\Users\你的用户名\Documents\OfficeAddins
   ```

6. 点击 **「添加目录」** → 勾选 ✅ **「显示在菜单中」** → 确定 → 确定

7. **完全关闭 PowerPoint 并重新打开**（必须重启！）

8. 打开任意 PPT → **插入** → **获取加载项** → **共享文件夹** 选项卡

9. 找到 **「PPT Asset Library」** → 点击 **「添加」** 🎉

---

### Mac 安装

1. 打开 **访达（Finder）**

2. 按 **Command + Shift + G**（前往文件夹）

3. 粘贴以下路径，点击「前往」：

   ```
   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   ```

4. 如果 `wef` 文件夹不存在，就**新建一个**叫 `wef` 的文件夹

5. 将下载的 `manifest.xml` 放进 `wef` 文件夹

6. **Command + Q** 完全退出 PowerPoint

7. 重新打开 PowerPoint，打开任意 PPT

8. **插入** → **加载项** 下拉 ▼ → **我的加载项** → **开发人员加载项**

9. 找到 **「PPT Asset Library」** 🎉

---

### 素材更新

- 素材由管理员更新并发布到 GitHub Pages
- **你不需要重新下载 manifest.xml**，插件内容会自动从服务器加载最新版本
- 如果需要强制刷新，点击插件标题栏右侧的 🔄 **刷新按钮**即可
- 首次使用需联网，之后支持离线使用（自动缓存）

---

### 离线缓存机制

| 项目 | 说明 |
|------|------|
| **技术方案** | Service Worker + Cache API |
| **缓存策略** | Cache First + 后台静默更新 |
| **预缓存内容** | taskpane.html、JS、assetIndex.json、slideIds.json、bundle.pptx |
| **运行时缓存** | 所有缩略图在首次加载时自动缓存 |
| **手动刷新** | 点击插件内 🔄 按钮，强制清除并重新下载所有资源 |

---

## 参考资料

- [Office Add-in 官方文档](https://learn.microsoft.com/office/dev/add-ins/)
- [PowerPoint Add-in Tutorial](https://learn.microsoft.com/office/dev/add-ins/tutorials/powerpoint-tutorial-yo)
- [Office.js PowerPoint API](https://learn.microsoft.com/javascript/api/powerpoint)
- [Fluent UI React](https://react.fluentui.dev/)
- [Office Add-in 示例代码](https://officedev.github.io/Office-Add-in-samples/)
