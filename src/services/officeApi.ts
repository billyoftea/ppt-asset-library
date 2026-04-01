/**
 * officeApi.ts
 * Office.js API 封装 — 从 bundle.pptx 插入原生可编辑幻灯片页面
 *
 * 策略：加载 bundle.pptx（包含第 103~191 页）→ insertSlidesFromBase64（指定 sourceSlideIds）
 *       → 完全保留原始页面的所有形状、文字、图片、动画等特性，100% 可编辑
 */

import { Asset } from "../types";

/* global PowerPoint, Office */

// 标记是否运行在 Office 环境中
let officeReady = false;

if (typeof Office !== "undefined") {
  try {
    Office.onReady(() => {
      officeReady = true;
    });
  } catch {
    // Office.js not available
  }
}

// bundle.pptx 的 base64 缓存（~2.6MB，只需加载一次）
let cachedBundleBase64: string | null = null;

// slideIds 映射缓存：{ "103": 1006, "104": 798, ... }
let cachedSlideIds: Record<string, number> | null = null;

/**
 * 检测当前是否在 Office 宿主环境中运行
 */
export function isOfficeEnvironment(): boolean {
  return officeReady || (typeof Office !== "undefined" && typeof PowerPoint !== "undefined");
}

/**
 * 加载 bundle.pptx 为 base64（带缓存）
 */
async function loadBundlePptxBase64(): Promise<string> {
  if (cachedBundleBase64) return cachedBundleBase64;

  console.log("Loading bundle.pptx...");
  const urls = ["/assets/bundle.pptx", "assets/bundle.pptx"];

  for (const url of urls) {
    try {
      const response = await fetch(url);
      if (!response.ok) continue;

      const blob = await response.blob();
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          const dataUri = reader.result as string;
          const b64 = dataUri.includes(",") ? dataUri.split(",")[1] : dataUri;
          resolve(b64);
        };
        reader.onerror = () => reject(new Error("Failed to read bundle.pptx"));
        reader.readAsDataURL(blob);
      });

      cachedBundleBase64 = base64;
      console.log("✅ bundle.pptx loaded and cached");
      return base64;
    } catch {
      continue;
    }
  }

  throw new Error("Could not load bundle.pptx");
}

/**
 * 加载 slideIds 映射表
 */
async function loadSlideIds(): Promise<Record<string, number>> {
  if (cachedSlideIds) return cachedSlideIds;

  const paths = ["./data/slideIds.json", "/data/slideIds.json", "../data/slideIds.json"];

  for (const path of paths) {
    try {
      const response = await fetch(path);
      if (response.ok) {
        cachedSlideIds = await response.json();
        console.log("✅ Loaded slideIds mapping");
        return cachedSlideIds!;
      }
    } catch {
      continue;
    }
  }

  throw new Error("Could not load slideIds.json");
}

/**
 * 预加载 bundle.pptx 和 slideIds（可在 App 初始化时调用，加速首次插入）
 */
export async function preloadBundle(): Promise<void> {
  await Promise.all([loadBundlePptxBase64(), loadSlideIds()]);
}

/**
 * 插入一页原生 PPT 幻灯片
 * 从 bundle.pptx 中按 slideId 提取指定页面，插入到当前选中幻灯片之后
 * 完全保留原始页面的所有特征（形状、文字、图片、动画、样式）
 *
 * @param asset 要插入的素材（使用 slideNumber 定位页面）
 */
export async function insertSlide(asset: Asset): Promise<void> {
  if (!isOfficeEnvironment()) {
    throw new Error("PowerPoint API not available. Please open this add-in in PowerPoint.");
  }

  // 并行加载 bundle.pptx 和 slideIds
  const [pptxBase64, slideIds] = await Promise.all([loadBundlePptxBase64(), loadSlideIds()]);

  // 查找该页在 bundle.pptx 中的 slideId
  const slideNum = String(asset.slideNumber);
  const sourceSlideId = slideIds[slideNum];

  if (!sourceSlideId) {
    throw new Error(`Slide ID not found for slide ${asset.slideNumber}`);
  }

  await PowerPoint.run(async (context) => {
    const options: PowerPoint.InsertSlideOptions = {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      sourceSlideIds: [String(sourceSlideId)],
    };

    // 尝试在当前选中幻灯片后面插入
    try {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();
      if (selectedSlides.items.length > 0) {
        options.targetSlideId = selectedSlides.items[0].id;
      }
    } catch {
      // 不设置 targetSlideId，默认插入到末尾
    }

    context.presentation.insertSlidesFromBase64(pptxBase64, options);
    await context.sync();
  });
}
