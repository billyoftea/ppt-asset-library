/**
 * assetManager.ts
 * 素材数据管理 - 加载索引、搜索素材、分类管理
 * v3.0 - 页面级素材（每页一个原生 PPT 幻灯片）
 */

import { AssetIndex, Asset, AssetRaw, Category, normalizeAsset } from "../types";

let cachedIndex: AssetIndex | null = null;

// 搜索索引缓存 (预构建的小写化搜索文本)
const searchCache = new Map<string, string>();

/**
 * 加载素材索引 JSON
 * 支持多路径尝试和缓存
 */
export async function loadAssetIndex(): Promise<AssetIndex> {
  if (cachedIndex) {
    return cachedIndex;
  }

  // 尝试多个路径加载索引文件（开发 vs 打包环境路径差异）
  const paths = [
    "./data/assetIndex.json",
    "/data/assetIndex.json",
    "../data/assetIndex.json",
    "./assets/data/assetIndex.json",
  ];

  let lastError: Error | null = null;

  for (const path of paths) {
    try {
      const response = await fetch(path);
      if (response.ok) {
        const rawData = await response.json();

        // 验证数据结构完整性
        if (!rawData.assets || !Array.isArray(rawData.assets)) {
          throw new Error("Invalid asset index: missing assets array");
        }
        if (!rawData.categories || !Array.isArray(rawData.categories)) {
          throw new Error("Invalid asset index: missing categories array");
        }

        // 规范化素材数据（从精简字段名转为完整字段名）
        const assets: Asset[] = rawData.assets.map((raw: AssetRaw) =>
          normalizeAsset(raw)
        );

        const data: AssetIndex = {
          version: rawData.version || "2.0",
          totalAssets: rawData.totalAssets || assets.length,
          categories: rawData.categories,
          assets,
        };

        // 预构建搜索缓存
        buildSearchCache(data);

        cachedIndex = data;
        console.log(
          `✅ Loaded asset index from ${path}: ${data.totalAssets} assets in ${data.categories.length} categories`
        );
        return data;
      }
    } catch (err) {
      lastError = err as Error;
      continue;
    }
  }

  throw new Error(
    `Failed to load asset index: ${lastError?.message || "All paths failed"}`
  );
}

/**
 * 预构建搜索文本缓存
 * 将每个素材的可搜索文本预先拼接并小写化，避免搜索时重复计算
 */
function buildSearchCache(index: AssetIndex): void {
  searchCache.clear();
  for (const asset of index.assets) {
    const searchableText = [
      asset.title,
      asset.category,
      asset.subcategory,
      ...asset.tags,
      ...asset.keywords,
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();
    searchCache.set(asset.id, searchableText);
  }
}

/**
 * 搜索素材
 * @param index 素材索引
 * @param query 搜索关键词（支持多词 AND 搜索）
 * @param category 可选的分类筛选
 * @returns 匹配的素材列表
 */
export function searchAssets(
  index: AssetIndex,
  query: string,
  category?: string
): Asset[] {
  let results = index.assets;

  // 按分类筛选
  if (category) {
    results = results.filter((a) => a.category === category);
  }

  // 按关键词搜索
  if (query.trim()) {
    const queryLower = query.toLowerCase().trim();
    const queryWords = queryLower.split(/\s+/).filter(Boolean);

    results = results.filter((asset) => {
      // 使用预缓存的搜索文本
      const searchableText = searchCache.get(asset.id) || "";
      return queryWords.every((word) => searchableText.includes(word));
    });
  }

  return results;
}

/**
 * 获取指定分类下的素材
 */
export function getAssetsByCategory(
  index: AssetIndex,
  categoryName: string
): Asset[] {
  return index.assets.filter((a) => a.category === categoryName);
}

/**
 * 按 ID 获取素材
 */
export function getAssetById(
  index: AssetIndex,
  assetId: string
): Asset | undefined {
  return index.assets.find((a) => a.id === assetId);
}

/**
 * 获取所有分类及其素材数量
 */
export function getCategorySummary(index: AssetIndex): Category[] {
  return index.categories.map((cat) => ({
    ...cat,
    count: index.assets.filter((a) => a.category === cat.name).length,
  }));
}

/**
 * 获取推荐/相关素材（同分类下的其他素材）
 */
export function getRelatedAssets(
  index: AssetIndex,
  asset: Asset,
  limit: number = 6
): Asset[] {
  return index.assets
    .filter((a) => a.id !== asset.id && a.category === asset.category)
    .slice(0, limit);
}

/**
 * 清除缓存（用于重新加载）
 */
export function clearCache(): void {
  cachedIndex = null;
  searchCache.clear();
}
