#!/usr/bin/env python3
"""
extract_assets.py
从解包后的 PPTX 中提取每张幻灯片的图形素材为独立文件。
- 将每张幻灯片的形状 XML 提取到 assets/vectors/ 目录
- 将关联的媒体文件（图片）复制到对应目录
- 输出素材元数据用于后续索引构建

用法: python scripts/extract_assets.py [unpacked_dir]
"""

import xml.etree.ElementTree as ET
import os
import sys
import re
import json
import shutil
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
UNPACKED_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else PROJECT_ROOT / "unpacked"
SLIDES_DIR = UNPACKED_DIR / "ppt" / "slides"
RELS_DIR = UNPACKED_DIR / "ppt" / "slides" / "_rels"
MEDIA_DIR = UNPACKED_DIR / "ppt" / "media"
OUTPUT_VECTORS = PROJECT_ROOT / "assets" / "vectors"
OUTPUT_METADATA = PROJECT_ROOT / "assets" / "metadata"

# XML namespaces
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# 分类关键词映射 - 基于对源 PPT 内容的分析
CATEGORY_KEYWORDS = {
    'Arrows': ['arrow', 'directional', 'arros'],
    'Circular Diagrams': ['circular', 'circle', 'venn', 'harvey ball', 'pie', 'donut',
                          'cycle', 'speedometer', 'quadrant', 'radar', 'sections'],
    'Tables': ['table'],
    'Other Shapes': ['other shape', 'diamond', 'pyramid', 'hexagon', 'puzzle', 'gear',
                     'spider', 'balance', 'funnel', 'scorecard'],
    'Puzzle Pieces': ['puzzle', 'jigsaw'],
    'Timelines': ['timeline', 'bridge diagram', 'waterfall', 'milestone', 'roadmap',
                  'gantt'],
    'Process Charts': ['process', 'flow', 'maze', 'diagram', 'graph', 'chevron',
                       'staircase', 'steps'],
    'Maps': ['map'],
    'Icons': ['icon', 'clip art', 'silhouette', 'social media icon'],
    'Infographics': ['infographic'],
    'Org Charts': ['org', 'organizational', 'hierarchy', 'hierachy'],
    'Weeble': ['weeble', 'weeblemania'],
    '3-Dimensional': ['3-dimensional', '3d', 'cube', 'building block'],
    'Quotes & Bubbles': ['quote', 'speech bubble', 'bubble'],
    'Newspapers': ['newspaper', 'clipping'],
    'Logos & Themes': ['logo', 'theme', 'pwc'],
}

# 跳过的幻灯片（封面、目录、色板等非素材页）
SKIP_SLIDES = {1, 2, 3, 4}


def extract_text_from_slide(slide_path):
    """提取幻灯片中的所有文本"""
    tree = ET.parse(slide_path)
    root = tree.getroot()
    texts = []
    for t in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text and t.text.strip():
            texts.append(t.text.strip())
    return texts


def classify_slide(texts):
    """根据幻灯片文本内容进行分类"""
    if not texts:
        return 'Uncategorized', []

    # 合并前几个文本作为分类依据
    combined = ' '.join(texts[:5]).lower()

    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in combined:
                return category, texts[:3]

    return 'Other', texts[:3]


def get_slide_rels(slide_file):
    """获取幻灯片的关系文件中引用的媒体"""
    rels_file = RELS_DIR / f"{slide_file}.rels"
    media_refs = []
    if rels_file.exists():
        tree = ET.parse(rels_file)
        root = tree.getroot()
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            target = rel.get('Target', '')
            rel_type = rel.get('Type', '')
            if 'image' in rel_type or 'media' in target:
                # 获取实际的媒体文件名
                media_name = os.path.basename(target)
                media_refs.append({
                    'rId': rel.get('Id'),
                    'target': target,
                    'media_file': media_name,
                    'type': rel_type.split('/')[-1]
                })
    return media_refs


def count_shapes(slide_path):
    """统计幻灯片中的形状数量和类型"""
    tree = ET.parse(slide_path)
    root = tree.getroot()

    shape_count = 0
    shape_types = set()

    # 统计各种形状元素
    for sp in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
        shape_count += 1
        shape_types.add('shape')

    for grp in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp'):
        shape_count += 1
        shape_types.add('group')

    for pic in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}pic'):
        shape_count += 1
        shape_types.add('picture')

    for cxn in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}cxnSp'):
        shape_count += 1
        shape_types.add('connector')

    # Also check drawingml namespace shapes
    for sp in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}sp'):
        shape_count += 1

    return shape_count, list(shape_types)


def extract_slide_xml(slide_path, output_path):
    """将幻灯片 XML 复制到输出目录"""
    shutil.copy2(slide_path, output_path)


def process_all_slides():
    """处理所有幻灯片，提取素材信息"""
    slide_files = sorted(
        [f for f in os.listdir(SLIDES_DIR) if f.endswith('.xml')],
        key=lambda x: int(re.search(r'(\d+)', x).group())
    )

    assets = []
    category_counts = {}
    current_category = 'Uncategorized'

    print(f"Processing {len(slide_files)} slides...")

    for i, sf in enumerate(slide_files):
        slide_num = int(re.search(r'(\d+)', sf).group())

        if slide_num in SKIP_SLIDES:
            continue

        slide_path = SLIDES_DIR / sf
        texts = extract_text_from_slide(slide_path)

        # 分类
        category, tags = classify_slide(texts)
        if category != 'Other' and category != 'Uncategorized':
            current_category = category

        # 获取媒体引用
        media_refs = get_slide_rels(sf)

        # 统计形状
        shape_count, shape_types = count_shapes(slide_path)

        # 生成子分类/标签
        subcategory = ''
        if len(texts) >= 2:
            # 第二个文本通常是子分类
            sub = texts[1] if texts[1] not in [str(slide_num)] else ''
            if sub and not sub.isdigit():
                subcategory = sub

        # 提取幻灯片的标题
        title = texts[0] if texts else f'Slide {slide_num}'

        # 构建素材记录
        asset = {
            'id': f'slide_{slide_num}',
            'slideNumber': slide_num,
            'slideFile': sf,
            'title': title,
            'category': current_category if category in ('Other', 'Uncategorized') else category,
            'subcategory': subcategory,
            'tags': [t for t in texts[:5] if t and not t.isdigit() and len(t) > 1],
            'mediaFiles': [m['media_file'] for m in media_refs],
            'mediaCount': len(media_refs),
            'shapeCount': shape_count,
            'shapeTypes': shape_types,
            'thumbnailPath': f'thumbnails/slide_{slide_num}.png',
            'vectorPath': f'vectors/{sf}',
        }

        assets.append(asset)

        # 统计分类数量
        cat = asset['category']
        category_counts[cat] = category_counts.get(cat, 0) + 1

        # 复制幻灯片 XML 到 vectors 目录
        os.makedirs(OUTPUT_VECTORS, exist_ok=True)
        extract_slide_xml(slide_path, OUTPUT_VECTORS / sf)

        # 复制关联的 rels 文件
        rels_src = RELS_DIR / f"{sf}.rels"
        if rels_src.exists():
            rels_dest = OUTPUT_VECTORS / "_rels"
            os.makedirs(rels_dest, exist_ok=True)
            shutil.copy2(rels_src, rels_dest / f"{sf}.rels")

    # 输出统计信息
    print(f"\n✅ Processed {len(assets)} asset slides")
    print(f"\n📊 Category distribution:")
    for cat, count in sorted(category_counts.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count} slides")

    return assets, category_counts


def copy_media_files():
    """复制所有媒体文件到素材目录"""
    media_dest = OUTPUT_VECTORS / "media"
    if MEDIA_DIR.exists():
        print(f"\n📁 Copying media files...")
        # 创建符号链接而不是复制（节省空间）
        if media_dest.exists():
            if media_dest.is_symlink():
                os.remove(media_dest)
            else:
                shutil.rmtree(media_dest)
        os.symlink(MEDIA_DIR.resolve(), media_dest)
        media_count = len(list(MEDIA_DIR.iterdir()))
        print(f"  Linked {media_count} media files")


def save_metadata(assets, category_counts):
    """保存素材元数据"""
    os.makedirs(OUTPUT_METADATA, exist_ok=True)

    # 保存完整的素材列表
    metadata_path = OUTPUT_METADATA / "assets_raw.json"
    with open(metadata_path, 'w', encoding='utf-8') as f:
        json.dump({
            'totalAssets': len(assets),
            'categories': category_counts,
            'assets': assets
        }, f, ensure_ascii=False, indent=2)
    print(f"\n💾 Saved metadata to {metadata_path}")

    return metadata_path


if __name__ == '__main__':
    print("=" * 60)
    print("📦 PPT Asset Extractor")
    print("=" * 60)

    if not SLIDES_DIR.exists():
        print(f"❌ Slides directory not found: {SLIDES_DIR}")
        print(f"   Please unpack the PPTX first.")
        sys.exit(1)

    assets, category_counts = process_all_slides()
    copy_media_files()
    save_metadata(assets, category_counts)

    print("\n✨ Extraction complete!")
