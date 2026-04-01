#!/usr/bin/env python3
"""
build_index.py
根据提取的素材元数据构建最终的素材索引 JSON (assetIndex.json)。
这个索引文件将被 Office Add-in 前端直接使用。

用法: python scripts/build_index.py
"""

import json
import os
import re
from pathlib import Path
from collections import OrderedDict

PROJECT_ROOT = Path(__file__).parent.parent
METADATA_DIR = PROJECT_ROOT / "assets" / "metadata"
THUMBNAILS_DIR = PROJECT_ROOT / "assets" / "thumbnails"
OUTPUT_DIR = PROJECT_ROOT / "src" / "data"


def normalize_category(raw_category):
    """规范化分类名称"""
    mapping = {
        'Arrows': 'Arrows',
        'Circular Diagrams': 'Circular Diagrams',
        'Tables': 'Tables',
        'Other Shapes': 'Other Shapes',
        'Puzzle Pieces': 'Puzzles',
        'Timelines': 'Timelines',
        'Process Charts': 'Process Charts',
        'Maps': 'Maps',
        'Icons': 'Icons',
        'Infographics': 'Infographics',
        'Org Charts': 'Org Charts',
        'Weeble': 'Weeble',
        '3-Dimensional': '3D Shapes',
        'Quotes & Bubbles': 'Quotes & Bubbles',
        'Newspapers': 'Newspapers',
        'Logos & Themes': 'Logos & Themes',
        'Other': 'Other',
        'Uncategorized': 'Other',
    }
    return mapping.get(raw_category, raw_category)


def generate_search_keywords(asset):
    """生成搜索关键词列表"""
    keywords = set()

    # 从标题提取
    if asset.get('title'):
        for word in re.split(r'[\s,\-–—/&]+', asset['title']):
            word = word.strip().lower()
            if word and len(word) > 1 and not word.isdigit():
                keywords.add(word)

    # 从子分类提取
    if asset.get('subcategory'):
        for word in re.split(r'[\s,\-–—/&]+', asset['subcategory']):
            word = word.strip().lower()
            if word and len(word) > 1 and not word.isdigit():
                keywords.add(word)

    # 从标签提取
    for tag in asset.get('tags', []):
        for word in re.split(r'[\s,\-–—/&]+', tag):
            word = word.strip().lower()
            if word and len(word) > 1 and not word.isdigit():
                keywords.add(word)

    # 添加分类名
    for word in re.split(r'[\s,\-–—/&]+', asset.get('category', '')):
        word = word.strip().lower()
        if word and len(word) > 1:
            keywords.add(word)

    # 过滤常见无用词
    stop_words = {'the', 'and', 'for', 'with', 'from', 'text', 'slide',
                  'mauris', 'lorem', 'ipsum', 'xxx', 'xxxxxx', 'xxxxxxxx',
                  'insert', 'date', 'here', 'title', 'presentation',
                  'pwc', 'header', 'name', 'surname'}
    keywords -= stop_words

    return sorted(list(keywords))


def build_index():
    """构建最终的素材索引"""
    # 读取原始元数据
    raw_path = METADATA_DIR / "assets_raw.json"
    if not raw_path.exists():
        print("❌ assets_raw.json not found. Run extract_assets.py first.")
        return

    with open(raw_path, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)

    # 检查哪些缩略图存在
    existing_thumbnails = set()
    if THUMBNAILS_DIR.exists():
        for f in THUMBNAILS_DIR.iterdir():
            if f.suffix == '.png':
                existing_thumbnails.add(f.name)

    # 构建索引
    categories = OrderedDict()
    assets_list = []
    category_order = [
        'Circular Diagrams', 'Process Charts', 'Arrows', 'Timelines',
        'Infographics', 'Maps', 'Org Charts', 'Icons', 'Weeble',
        '3D Shapes', 'Tables', 'Other Shapes', 'Puzzles',
        'Quotes & Bubbles', 'Newspapers', 'Logos & Themes', 'Other'
    ]

    # 初始化分类
    for cat in category_order:
        categories[cat] = {
            'id': cat.lower().replace(' ', '-').replace('&', 'and'),
            'name': cat,
            'count': 0,
            'icon': get_category_icon(cat),
        }

    for raw_asset in raw_data.get('assets', []):
        category = normalize_category(raw_asset.get('category', 'Other'))
        slide_num = raw_asset['slideNumber']

        # 检查缩略图
        thumb_file = f"slide_{slide_num}.png"
        has_thumbnail = thumb_file in existing_thumbnails

        asset = {
            'id': raw_asset['id'],
            'slideNumber': slide_num,
            'title': raw_asset.get('title', f'Slide {slide_num}'),
            'category': category,
            'subcategory': raw_asset.get('subcategory', ''),
            'tags': raw_asset.get('tags', []),
            'keywords': generate_search_keywords(raw_asset),
            'thumbnailPath': f'assets/thumbnails/{thumb_file}' if has_thumbnail else '',
            'vectorPath': f'assets/vectors/{raw_asset["slideFile"]}',
            'mediaCount': raw_asset.get('mediaCount', 0),
            'shapeCount': raw_asset.get('shapeCount', 0),
        }

        assets_list.append(asset)

        # 更新分类计数
        if category in categories:
            categories[category]['count'] += 1
        else:
            categories[category] = {
                'id': category.lower().replace(' ', '-').replace('&', 'and'),
                'name': category,
                'count': 1,
                'icon': '📦',
            }

    # 移除空分类
    categories = {k: v for k, v in categories.items() if v['count'] > 0}

    # 构建最终索引
    index = {
        'version': '1.0.0',
        'totalAssets': len(assets_list),
        'categories': list(categories.values()),
        'assets': assets_list,
    }

    # 保存索引
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = OUTPUT_DIR / "assetIndex.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

    # 同时保存一份精简版（不含 keywords，减小文件体积）
    lite_assets = []
    for a in assets_list:
        lite = {k: v for k, v in a.items() if k != 'keywords'}
        lite_assets.append(lite)

    lite_index = {
        'version': '1.0.0',
        'totalAssets': len(lite_assets),
        'categories': list(categories.values()),
        'assets': lite_assets,
    }

    lite_path = OUTPUT_DIR / "assetIndex.lite.json"
    with open(lite_path, 'w', encoding='utf-8') as f:
        json.dump(lite_index, f, ensure_ascii=False)

    print(f"\n📊 Index statistics:")
    print(f"  Total assets: {len(assets_list)}")
    print(f"  Categories: {len(categories)}")
    print(f"  With thumbnails: {sum(1 for a in assets_list if a['thumbnailPath'])}")
    print(f"\n  Category breakdown:")
    for cat in categories.values():
        print(f"    {cat['icon']} {cat['name']}: {cat['count']} assets")

    print(f"\n💾 Full index: {output_path} ({os.path.getsize(output_path) / 1024:.1f} KB)")
    print(f"💾 Lite index: {lite_path} ({os.path.getsize(lite_path) / 1024:.1f} KB)")

    return index


def get_category_icon(category):
    """获取分类图标 emoji"""
    icons = {
        'Circular Diagrams': '🔴',
        'Process Charts': '📊',
        'Arrows': '➡️',
        'Timelines': '📅',
        'Infographics': '📈',
        'Maps': '🗺️',
        'Org Charts': '🏢',
        'Icons': '🎨',
        'Weeble': '🧑',
        '3D Shapes': '🧊',
        'Tables': '📋',
        'Other Shapes': '🔷',
        'Puzzles': '🧩',
        'Quotes & Bubbles': '💬',
        'Newspapers': '📰',
        'Logos & Themes': '🏷️',
        'Other': '📦',
    }
    return icons.get(category, '📦')


if __name__ == '__main__':
    print("=" * 60)
    print("📑 Asset Index Builder")
    print("=" * 60)

    index = build_index()
    if index:
        print("\n✨ Index build complete!")
