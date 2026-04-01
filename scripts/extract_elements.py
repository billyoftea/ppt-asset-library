#!/usr/bin/env python3
"""
extract_elements.py
以元素为单位提取素材：将每张幻灯片中的独立形状组(grpSp)和图片(pic)拆分为单独素材。

策略：
- Icons/Weeble 类：每个 pic 或 grpSp 是一个独立素材 → 逐个提取
- Charts/Diagrams 类：如果页面有多个 grpSp，每个 grpSp 是一个独立素材
- 如果页面只有 1 个有意义的顶层组，则整页作为一个素材
- 纯文本框(Title/Slide Number)不算素材

对每个素材：
1. 提取关联的图片资源生成缩略图
2. 对纯矢量形状组，用 Pillow 绘制简单的示意缩略图
3. 构建新的 assetIndex.json
"""

import xml.etree.ElementTree as ET
import os
import sys
import json
import re
import shutil
import hashlib
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

PROJECT_ROOT = Path(__file__).parent.parent
UNPACKED_DIR = PROJECT_ROOT / "unpacked"
SLIDES_DIR = UNPACKED_DIR / "ppt" / "slides"
RELS_DIR = SLIDES_DIR / "_rels"
MEDIA_DIR = UNPACKED_DIR / "ppt" / "media"
ASSETS_DIR = PROJECT_ROOT / "assets"
THUMBNAILS_DIR = ASSETS_DIR / "thumbnails"
VECTORS_DIR = ASSETS_DIR / "vectors"
DATA_DIR = PROJECT_ROOT / "src" / "data"

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

THUMB_SIZE = (300, 225)

# ─── 分类配置 ─────────────────────────────────
# 哪些分类应该以"元素"为单位拆分
ELEMENT_LEVEL_CATEGORIES = {
    'Icons', 'Weeble', 'Logos & Themes', 'Quotes & Bubbles', 'Newspapers'
}

# 哪些分类应该以"形状组"为单位拆分（每个 grpSp 是一个素材）
GROUP_LEVEL_CATEGORIES = {
    'Circular Diagrams', 'Arrows', '3D Shapes', 'Tables',
    'Timelines', 'Other Shapes', 'Maps', 'Process Charts', 'Org Charts'
}


def load_slide_rels(slide_num):
    """加载幻灯片的关系文件，返回 rId -> media_file 的映射"""
    rels_file = RELS_DIR / f"slide{slide_num}.xml.rels"
    if not rels_file.exists():
        return {}

    tree = ET.parse(rels_file)
    root = tree.getroot()
    rels = {}
    for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid = rel.get('Id', '')
        target = rel.get('Target', '')
        rel_type = rel.get('Type', '')
        if 'image' in rel_type:
            media_file = os.path.basename(target)
            rels[rid] = media_file
    return rels


def get_element_name(elem):
    """获取元素名称"""
    for child in elem.iter():
        if child.tag.endswith('cNvPr'):
            return child.get('name', '')
    return ''


def get_element_text(elem):
    """提取元素内的所有文本"""
    texts = []
    for t in elem.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text and t.text.strip():
            texts.append(t.text.strip())
    return texts


def get_element_images(elem, rels_map):
    """获取元素引用的所有图片文件路径"""
    images = []
    for blip in elem.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
        rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
        if rid in rels_map:
            media_file = rels_map[rid]
            media_path = MEDIA_DIR / media_file
            if media_path.exists() and media_path.suffix.lower() in ('.png', '.jpg', '.jpeg', '.gif'):
                images.append(media_path)
    return images


def get_element_bounds(elem):
    """获取元素的位置和尺寸 (x, y, cx, cy) in EMU"""
    xfrm = None
    # 直接子元素的 xfrm
    for child in elem:
        tag = child.tag.split('}')[-1]
        if tag in ('spPr', 'grpSpPr'):
            xfrm_elem = child.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
            if xfrm_elem is not None:
                xfrm = xfrm_elem
                break

    if xfrm is None:
        return None

    off = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off')
    ext = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')

    if off is None or ext is None:
        return None

    try:
        x = int(off.get('x', 0))
        y = int(off.get('y', 0))
        cx = int(ext.get('cx', 0))
        cy = int(ext.get('cy', 0))
        return (x, y, cx, cy)
    except (ValueError, TypeError):
        return None


def is_meaningful_element(elem):
    """判断元素是否是有意义的素材（排除标题、页码等）"""
    name = get_element_name(elem).lower()
    tag = elem.tag.split('}')[-1]

    # 跳过标题和页码占位符
    skip_patterns = ['title', 'slide number', 'placeholder', 'date', 'footer',
                     'スライド番号', '正方形/長方形']
    for p in skip_patterns:
        if p in name:
            return False

    # 形状组和图片总是有意义的
    if tag in ('grpSp', 'pic'):
        return True

    # 单独的形状：检查是否有实质内容
    if tag == 'sp':
        bounds = get_element_bounds(elem)
        if bounds:
            _, _, cx, cy = bounds
            # 太小的形状忽略（< 0.5cm x 0.5cm）
            if cx < 180000 and cy < 180000:
                return False
        # 纯文本框如果很小也忽略
        texts = get_element_text(elem)
        if not texts and name.startswith('TextBox'):
            return False
        return True

    return tag in ('cxnSp',)  # 连接线通常是辅助的，但保留


def generate_image_thumbnail(image_path, output_path):
    """从图片文件生成缩略图"""
    try:
        img = Image.open(image_path).convert('RGBA')
        # 白色背景
        canvas = Image.new('RGBA', THUMB_SIZE, (255, 255, 255, 255))
        img.thumbnail((THUMB_SIZE[0] - 20, THUMB_SIZE[1] - 20), Image.LANCZOS)
        x = (THUMB_SIZE[0] - img.width) // 2
        y = (THUMB_SIZE[1] - img.height) // 2
        canvas.paste(img, (x, y), img)
        canvas.convert('RGB').save(output_path, 'PNG')
        return True
    except Exception as e:
        return False


def generate_composite_thumbnail(image_paths, output_path):
    """组合多张图片生成缩略图"""
    loaded = []
    for p in image_paths:
        try:
            img = Image.open(p).convert('RGBA')
            loaded.append((img, os.path.getsize(p)))
        except:
            pass

    if not loaded:
        return False

    # 按大小排序，取前4张
    loaded.sort(key=lambda x: x[1], reverse=True)
    images = [img for img, _ in loaded[:4]]

    canvas = Image.new('RGBA', THUMB_SIZE, (255, 255, 255, 255))

    if len(images) == 1:
        img = images[0]
        img.thumbnail((THUMB_SIZE[0] - 20, THUMB_SIZE[1] - 20), Image.LANCZOS)
        x = (THUMB_SIZE[0] - img.width) // 2
        y = (THUMB_SIZE[1] - img.height) // 2
        canvas.paste(img, (x, y), img)
    elif len(images) == 2:
        cell_w = THUMB_SIZE[0] // 2 - 15
        cell_h = THUMB_SIZE[1] - 20
        for i, img in enumerate(images):
            img.thumbnail((cell_w, cell_h), Image.LANCZOS)
            x = 10 + i * (THUMB_SIZE[0] // 2) + (cell_w - img.width) // 2
            y = (THUMB_SIZE[1] - img.height) // 2
            canvas.paste(img, (x, y), img)
    elif len(images) == 3:
        # 上面1个大的，下面2个小的
        images[0].thumbnail((THUMB_SIZE[0] - 20, THUMB_SIZE[1] // 2 - 15), Image.LANCZOS)
        x = (THUMB_SIZE[0] - images[0].width) // 2
        canvas.paste(images[0], (x, 5), images[0])
        cell_w = THUMB_SIZE[0] // 2 - 15
        cell_h = THUMB_SIZE[1] // 2 - 15
        for i, img in enumerate(images[1:3]):
            img.thumbnail((cell_w, cell_h), Image.LANCZOS)
            x = 10 + i * (THUMB_SIZE[0] // 2) + (cell_w - img.width) // 2
            y = THUMB_SIZE[1] // 2 + 5 + (cell_h - img.height) // 2
            canvas.paste(img, (x, y), img)
    else:
        # 2x2 网格
        cell_w = THUMB_SIZE[0] // 2 - 15
        cell_h = THUMB_SIZE[1] // 2 - 15
        for i, img in enumerate(images[:4]):
            img.thumbnail((cell_w, cell_h), Image.LANCZOS)
            col = i % 2
            row = i // 2
            x = 10 + col * (THUMB_SIZE[0] // 2) + (cell_w - img.width) // 2
            y = 5 + row * (THUMB_SIZE[1] // 2) + (cell_h - img.height) // 2
            canvas.paste(img, (x, y), img)

    canvas.convert('RGB').save(output_path, 'PNG')
    return True


def generate_vector_placeholder(output_path, category, label=""):
    """为纯矢量素材生成占位缩略图"""
    colors = {
        'Arrows': (52, 152, 219),
        'Circular Diagrams': (231, 76, 60),
        'Tables': (46, 204, 113),
        'Other Shapes': (155, 89, 182),
        'Timelines': (26, 188, 156),
        'Process Charts': (230, 126, 34),
        'Maps': (52, 73, 94),
        'Icons': (22, 160, 133),
        'Org Charts': (39, 174, 96),
        'Weeble': (211, 84, 0),
        '3D Shapes': (44, 62, 80),
        'Quotes & Bubbles': (192, 57, 43),
        'Newspapers': (127, 140, 141),
        'Logos & Themes': (241, 196, 15),
    }
    color = colors.get(category, (149, 165, 166))

    img = Image.new('RGB', THUMB_SIZE, (245, 245, 250))
    draw = ImageDraw.Draw(img)

    # 网格背景
    for x in range(0, THUMB_SIZE[0], 25):
        draw.line([(x, 0), (x, THUMB_SIZE[1])], fill=(235, 235, 240), width=1)
    for y in range(0, THUMB_SIZE[1], 25):
        draw.line([(0, y), (THUMB_SIZE[0], y)], fill=(235, 235, 240), width=1)

    # 中心放一个带颜色的圆角矩形
    r, g, b = color
    light_color = (r + (255 - r) // 2, g + (255 - g) // 2, b + (255 - b) // 2)
    draw.rounded_rectangle([50, 30, 250, 170], radius=12, fill=light_color,
                           outline=color, width=2)

    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 12)
        font_small = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 10)
    except:
        font = ImageFont.load_default()
        font_small = font

    # 分类名
    bbox = draw.textbbox((0, 0), category, font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((THUMB_SIZE[0] - tw) // 2, 180), category, fill=(120, 130, 150), font=font)

    # 标签
    if label:
        short_label = label[:30] + ('...' if len(label) > 30 else '')
        bbox2 = draw.textbbox((0, 0), short_label, font=font_small)
        tw2 = bbox2[2] - bbox2[0]
        draw.text(((THUMB_SIZE[0] - tw2) // 2, 90), short_label, fill=color, font=font_small)

    img.save(output_path, 'PNG')


def extract_slide_assets(slide_num, category, subcategory, rels_map):
    """
    从一张幻灯片中提取所有独立素材。
    返回素材列表，每个素材包含 id, thumbnail, tags 等信息。
    """
    slide_path = SLIDES_DIR / f"slide{slide_num}.xml"
    if not slide_path.exists():
        return []

    tree = ET.parse(slide_path)
    root = tree.getroot()
    sp_tree = root.find('.//p:cSld/p:spTree', NS)
    if sp_tree is None:
        return []

    # 收集顶层有意义的元素
    meaningful_elements = []
    for child in sp_tree:
        tag = child.tag.split('}')[-1]
        if tag in ('sp', 'grpSp', 'pic', 'cxnSp') and is_meaningful_element(child):
            meaningful_elements.append(child)

    # 对于只有 0-1 个有意义元素的页面，整页作为一个素材
    if len(meaningful_elements) <= 1:
        return extract_whole_slide(slide_num, category, subcategory, rels_map, meaningful_elements)

    # 对于 ELEMENT_LEVEL 分类，每个 pic/grpSp 都是独立素材
    if category in ELEMENT_LEVEL_CATEGORIES:
        return extract_individual_elements(slide_num, category, subcategory, rels_map, meaningful_elements)

    # 对于 GROUP_LEVEL 分类
    if category in GROUP_LEVEL_CATEGORIES:
        # 如果有多个 grpSp，每个 grpSp 是一个素材
        groups = [e for e in meaningful_elements if e.tag.split('}')[-1] == 'grpSp']
        if len(groups) > 1:
            return extract_individual_elements(slide_num, category, subcategory, rels_map, meaningful_elements)
        else:
            # 只有1个或0个 group，整页作为一个素材
            return extract_whole_slide(slide_num, category, subcategory, rels_map, meaningful_elements)

    # 默认：整页一个素材
    return extract_whole_slide(slide_num, category, subcategory, rels_map, meaningful_elements)


def extract_whole_slide(slide_num, category, subcategory, rels_map, elements):
    """整张幻灯片作为一个素材"""
    asset_id = f"slide_{slide_num}"

    # 收集所有图片
    all_images = []
    for elem in elements:
        all_images.extend(get_element_images(elem, rels_map))

    # 收集文本做标签
    all_texts = []
    for elem in elements:
        all_texts.extend(get_element_text(elem))

    # 生成缩略图
    thumb_path = THUMBNAILS_DIR / f"{asset_id}.png"
    if all_images:
        if len(all_images) == 1:
            generate_image_thumbnail(all_images[0], thumb_path)
        else:
            generate_composite_thumbnail(all_images, thumb_path)
    else:
        label = get_element_name(elements[0]) if elements else ""
        generate_vector_placeholder(thumb_path, category, label)

    keywords = list(set(
        [w.lower() for t in all_texts for w in re.findall(r'[a-zA-Z]+', t) if len(w) > 2]
    ))[:10]

    return [{
        'id': asset_id,
        'slideNumber': slide_num,
        'elementIndex': 0,
        'title': all_texts[0] if all_texts else f"{category} - Slide {slide_num}",
        'category': category,
        'subcategory': subcategory,
        'tags': list(set(all_texts[:5])),
        'keywords': keywords,
        'thumbnailPath': f"assets/thumbnails/{asset_id}.png",
        'vectorPath': f"assets/vectors/slide{slide_num}.xml",
        'type': 'slide',
        'hasImage': len(all_images) > 0,
    }]


def extract_individual_elements(slide_num, category, subcategory, rels_map, elements):
    """将每个有意义的元素作为独立素材提取"""
    assets = []
    elem_idx = 0

    for elem in elements:
        tag = elem.tag.split('}')[-1]
        name = get_element_name(elem)
        texts = get_element_text(elem)
        images = get_element_images(elem, rels_map)

        # 连接线不单独提取
        if tag == 'cxnSp':
            continue

        # 纯文本框（没有图片且是 sp 类型）如果内容太少则跳过
        if tag == 'sp' and not images:
            if not texts or all(len(t) <= 3 for t in texts):
                continue

        elem_idx += 1
        asset_id = f"slide_{slide_num}_e{elem_idx}"

        # 生成缩略图
        thumb_path = THUMBNAILS_DIR / f"{asset_id}.png"
        if images:
            if len(images) == 1:
                generate_image_thumbnail(images[0], thumb_path)
            else:
                generate_composite_thumbnail(images, thumb_path)
        else:
            generate_vector_placeholder(thumb_path, category, name)

        title = name if name else (texts[0] if texts else f"{category} Element {elem_idx}")
        keywords = list(set(
            [w.lower() for t in texts for w in re.findall(r'[a-zA-Z]+', t) if len(w) > 2]
        ))[:10]

        assets.append({
            'id': asset_id,
            'slideNumber': slide_num,
            'elementIndex': elem_idx,
            'title': title,
            'category': category,
            'subcategory': subcategory,
            'tags': list(set(texts[:5])) + [name] if name else list(set(texts[:5])),
            'keywords': keywords,
            'thumbnailPath': f"assets/thumbnails/{asset_id}.png",
            'vectorPath': f"assets/vectors/slide{slide_num}.xml",
            'type': 'element',
            'elementTag': tag,
            'hasImage': len(images) > 0,
        })

    # 如果拆分后只有0个元素，回退到整页模式
    if not assets:
        return extract_whole_slide(slide_num, category, subcategory, rels_map, elements)

    return assets


def main():
    print("=" * 60)
    print("🔧 Element-Level Asset Extractor")
    print("=" * 60)

    # 清理旧缩略图
    if THUMBNAILS_DIR.exists():
        print("\n🗑️  Cleaning old thumbnails...")
        for f in THUMBNAILS_DIR.glob('*.png'):
            f.unlink()

    os.makedirs(THUMBNAILS_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)

    # 加载现有索引获取分类信息
    old_index_path = DATA_DIR / "assetIndex.json"
    if not old_index_path.exists():
        print("❌ assetIndex.json not found! Run build_index.py first.")
        sys.exit(1)

    with open(old_index_path, 'r') as f:
        old_data = json.load(f)

    # 建立 slideNumber -> category/subcategory 的映射
    slide_info = {}
    for asset in old_data.get('assets', []):
        sn = asset['slideNumber']
        slide_info[sn] = {
            'category': asset.get('category', 'Uncategorized'),
            'subcategory': asset.get('subcategory', ''),
        }

    all_assets = []
    category_counts = {}
    stats = {'slide_level': 0, 'element_level': 0, 'total_with_image': 0}

    total_slides = len(slide_info)
    processed = 0

    for snum in sorted(slide_info.keys()):
        info = slide_info[snum]
        category = info['category']
        subcategory = info['subcategory']

        # 加载关系映射
        rels_map = load_slide_rels(snum)

        # 提取素材
        assets = extract_slide_assets(snum, category, subcategory, rels_map)

        for a in assets:
            if a['type'] == 'slide':
                stats['slide_level'] += 1
            else:
                stats['element_level'] += 1
            if a.get('hasImage'):
                stats['total_with_image'] += 1

        all_assets.extend(assets)
        category_counts[category] = category_counts.get(category, 0) + len(assets)

        processed += 1
        if processed % 50 == 0:
            print(f"  Progress: {processed}/{total_slides} slides, {len(all_assets)} assets extracted")

    # 构建新索引
    categories = []
    for cat_name in sorted(category_counts.keys()):
        cat_id = re.sub(r'[^a-z0-9]+', '-', cat_name.lower()).strip('-')
        categories.append({
            'id': cat_id,
            'name': cat_name,
            'count': category_counts[cat_name],
        })

    new_index = {
        'version': '2.0',
        'totalAssets': len(all_assets),
        'categories': categories,
        'assets': all_assets,
    }

    # 保存
    output_path = DATA_DIR / "assetIndex.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(new_index, f, ensure_ascii=False, indent=2)

    print(f"\n{'=' * 60}")
    print(f"✅ Extraction complete!")
    print(f"   Total assets: {len(all_assets)}")
    print(f"   Slide-level assets: {stats['slide_level']}")
    print(f"   Element-level assets: {stats['element_level']}")
    print(f"   Assets with real images: {stats['total_with_image']}")
    print(f"\n📊 Category breakdown:")
    for cat, count in sorted(category_counts.items(), key=lambda x: -x[1]):
        print(f"   {cat}: {count}")
    print(f"\n📁 Index saved to: {output_path}")
    print(f"   Index size: {os.path.getsize(output_path) // 1024}KB")


if __name__ == '__main__':
    main()
