/**
 * 素材库核心类型定义
 * v3.0 - 页面级素材（每页一个原生 PPT 幻灯片）
 */

/** 单个素材（精简字段名，对应 assetIndex.json） */
export interface AssetRaw {
  id: string;
  sn: number;           // slideNumber
  t: string;            // title
  cat: string;          // category
  sub?: string;         // subcategory
  tags?: string[];
  kw?: string[];        // keywords
  thumb: string;        // thumbnailPath
  type: "slide";
  img: boolean;         // hasImage (real image thumbnail)
}

/** 规范化后的素材对象（前端使用） */
export interface Asset {
  id: string;
  slideNumber: number;
  title: string;
  category: string;
  subcategory: string;
  tags: string[];
  keywords: string[];
  thumbnailPath: string;
  type: "slide";
  hasImage: boolean;
}

/** 素材分类 */
export interface Category {
  id: string;
  name: string;
  count: number;
  icon?: string;
}

/** 素材索引（assetIndex.json 的结构） */
export interface AssetIndex {
  version: string;
  totalAssets: number;
  categories: Category[];
  assets: Asset[];
}

/** 将 JSON 中的精简素材转为规范化对象 */
export function normalizeAsset(raw: AssetRaw): Asset {
  return {
    id: raw.id,
    slideNumber: raw.sn,
    title: raw.t,
    category: raw.cat,
    subcategory: raw.sub ?? "",
    tags: raw.tags ?? [],
    keywords: raw.kw ?? [],
    thumbnailPath: raw.thumb,
    type: "slide",
    hasImage: raw.img,
  };
}
