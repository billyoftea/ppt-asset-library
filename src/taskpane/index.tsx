import React from "react";
import { createRoot } from "react-dom/client";
import {
  FluentProvider,
  webLightTheme,
  createLightTheme,
  createDarkTheme,
  BrandVariants,
  Theme,
} from "@fluentui/react-components";
import App from "./App";
import "./styles/taskpane.css";

/* global Office */

/**
 * Tencent Blue 腾讯蓝品牌色阶（基于 TDesign 色彩体系）
 * 用于 Fluent UI v9 的 BrandVariants（10~160 共 16 级）
 */
const tencentBlue: BrandVariants = {
  10:  "#001433",
  20:  "#001D4D",
  30:  "#002566",
  40:  "#002A7C",
  50:  "#003099",
  60:  "#0034B5",
  70:  "#0041CC",
  80:  "#0052D9",   // ← 主色 Tencent Blue
  90:  "#4787EB",
  100: "#699EF0",
  110: "#88B2F4",
  120: "#A6C6F7",
  130: "#BDD2FB",
  140: "#D4E3FC",
  150: "#E8EFFF",
  160: "#ECF2FE",
};

const tencentLightTheme: Theme = {
  ...createLightTheme(tencentBlue),
};

// 覆盖字体为腾讯体
tencentLightTheme.fontFamilyBase =
  '"TencentSans", "腾讯体", -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", "Helvetica Neue", Helvetica, Arial, sans-serif';

// 等待 Office.js 初始化完成后再渲染 React 应用
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("Office.js ready - PowerPoint host detected");
  }

  const rootElement = document.getElementById("root");
  if (rootElement) {
    const root = createRoot(rootElement);
    root.render(
      <React.StrictMode>
        <FluentProvider theme={tencentLightTheme}>
          <App />
        </FluentProvider>
      </React.StrictMode>
    );
  }
});
