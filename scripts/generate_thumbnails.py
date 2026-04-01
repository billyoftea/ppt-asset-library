#!/usr/bin/env python3
"""
generate_thumbnails.py
从解包后的 PPTX 幻灯片生成 PNG 缩略图。

方法: 使用 LibreOffice 将 PPTX 转为 PDF，再用 pdftoppm 将每页转为 PNG。
若 LibreOffice 不可用，则从幻灯片 XML 中提取关联图片作为缩略图。

用法: python scripts/generate_thumbnails.py [pptx_file]
"""

import os
import sys
import subprocess
import json
import shutil
import re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
PPTX_FILE = Path(sys.argv[1]) if len(sys.argv) > 1 else None
ASSETS_DIR = PROJECT_ROOT / "assets"
THUMBNAILS_DIR = ASSETS_DIR / "thumbnails"
METADATA_DIR = ASSETS_DIR / "metadata"
UNPACKED_DIR = PROJECT_ROOT / "unpacked"
MEDIA_DIR = UNPACKED_DIR / "ppt" / "media"
RELS_DIR = UNPACKED_DIR / "ppt" / "slides" / "_rels"


def find_pptx():
    """自动查找项目中的 PPTX 文件"""
    if PPTX_FILE and PPTX_FILE.exists():
        return PPTX_FILE
    for f in PROJECT_ROOT.iterdir():
        if f.suffix == '.pptx' and not f.name.startswith('~'):
            return f
    return None


def generate_via_libreoffice(pptx_path):
    """使用 LibreOffice + pdftoppm 生成缩略图"""
    import tempfile
    tmp_dir = tempfile.mkdtemp()

    try:
        # 检查 soffice 是否可用
        soffice_paths = [
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            'soffice',
            '/usr/bin/soffice',
        ]

        soffice = None
        for sp in soffice_paths:
            try:
                subprocess.run([sp, '--version'], capture_output=True, timeout=10)
                soffice = sp
                break
            except (FileNotFoundError, subprocess.TimeoutExpired):
                continue

        if not soffice:
            return False

        print("  Using LibreOffice to convert PPTX → PDF...")
        result = subprocess.run(
            [soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir,
             str(pptx_path)],
            capture_output=True, text=True, timeout=300
        )

        if result.returncode != 0:
            print(f"  LibreOffice error: {result.stderr}")
            return False

        # 找到生成的 PDF
        pdf_files = list(Path(tmp_dir).glob('*.pdf'))
        if not pdf_files:
            print("  No PDF generated")
            return False

        pdf_path = pdf_files[0]

        # 检查 pdftoppm 是否可用
        try:
            subprocess.run(['pdftoppm', '-v'], capture_output=True, timeout=5)
        except FileNotFoundError:
            print("  pdftoppm not found, trying alternative...")
            # 尝试使用 ImageMagick 的 convert
            try:
                subprocess.run(['magick', '-version'], capture_output=True, timeout=5)
                return generate_from_pdf_magick(pdf_path)
            except FileNotFoundError:
                return False

        print("  Converting PDF pages to PNG thumbnails...")
        os.makedirs(THUMBNAILS_DIR, exist_ok=True)

        result = subprocess.run(
            ['pdftoppm', '-png', '-r', '150', '-scale-to', '300',
             str(pdf_path), str(THUMBNAILS_DIR / 'slide')],
            capture_output=True, text=True, timeout=600
        )

        if result.returncode == 0:
            # 重命名文件: slide-01.png → slide_1.png 
            for f in THUMBNAILS_DIR.glob('slide-*.png'):
                num = int(re.search(r'slide-(\d+)', f.name).group(1))
                new_name = f"slide_{num}.png"
                f.rename(THUMBNAILS_DIR / new_name)
            return True

        return False

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def generate_from_pdf_magick(pdf_path):
    """使用 ImageMagick 从 PDF 生成缩略图"""
    os.makedirs(THUMBNAILS_DIR, exist_ok=True)
    result = subprocess.run(
        ['magick', str(pdf_path), '-resize', '300x225',
         '-quality', '90', str(THUMBNAILS_DIR / 'slide_%d.png')],
        capture_output=True, text=True, timeout=600
    )
    return result.returncode == 0


def generate_from_media_fallback():
    """
    备选方案：从幻灯片关系文件中找到关联的第一张图片作为缩略图。
    这不如 LibreOffice 方案完美，但无需额外依赖。
    """
    import xml.etree.ElementTree as ET

    print("  Using media extraction fallback for thumbnails...")
    os.makedirs(THUMBNAILS_DIR, exist_ok=True)

    slides_dir = UNPACKED_DIR / "ppt" / "slides"
    slide_files = sorted(
        [f for f in os.listdir(slides_dir) if f.endswith('.xml')],
        key=lambda x: int(re.search(r'(\d+)', x).group())
    )

    generated = 0
    try:
        from PIL import Image
        has_pil = True
    except ImportError:
        has_pil = False

    for sf in slide_files:
        slide_num = int(re.search(r'(\d+)', sf).group())
        rels_file = RELS_DIR / f"{sf}.rels"

        if not rels_file.exists():
            continue

        tree = ET.parse(rels_file)
        root = tree.getroot()

        # 找到第一个图片引用
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            target = rel.get('Target', '')
            rel_type = rel.get('Type', '')
            if 'image' in rel_type:
                media_file = os.path.basename(target)
                media_path = MEDIA_DIR / media_file
                thumb_path = THUMBNAILS_DIR / f"slide_{slide_num}.png"

                if media_path.exists() and not thumb_path.exists():
                    if has_pil:
                        try:
                            img = Image.open(media_path)
                            img.thumbnail((300, 225), Image.LANCZOS)
                            img.save(thumb_path, 'PNG')
                            generated += 1
                        except Exception:
                            # 直接复制
                            shutil.copy2(media_path, thumb_path)
                            generated += 1
                    else:
                        shutil.copy2(media_path, thumb_path)
                        generated += 1
                    break  # 只取第一张图

    return generated


def generate_placeholder_thumbnails():
    """
    为没有缩略图的幻灯片生成占位图。
    使用 PIL 生成带分类名称的纯色图片。
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        has_pil = True
    except ImportError:
        has_pil = False
        return 0

    metadata_path = METADATA_DIR / "assets_raw.json"
    if not metadata_path.exists():
        return 0

    with open(metadata_path, 'r') as f:
        data = json.load(f)

    generated = 0
    # 分类颜色映射
    colors = {
        'Arrows': '#3498db',
        'Circular Diagrams': '#e74c3c',
        'Tables': '#2ecc71',
        'Other Shapes': '#9b59b6',
        'Puzzle Pieces': '#f39c12',
        'Timelines': '#1abc9c',
        'Process Charts': '#e67e22',
        'Maps': '#34495e',
        'Icons': '#16a085',
        'Infographics': '#8e44ad',
        'Org Charts': '#27ae60',
        'Weeble': '#d35400',
        '3-Dimensional': '#2c3e50',
        'Quotes & Bubbles': '#c0392b',
        'Newspapers': '#7f8c8d',
        'Logos & Themes': '#f1c40f',
        'Other': '#95a5a6',
        'Uncategorized': '#bdc3c7',
    }

    for asset in data.get('assets', []):
        thumb_path = THUMBNAILS_DIR / f"slide_{asset['slideNumber']}.png"
        if not thumb_path.exists():
            color = colors.get(asset['category'], '#95a5a6')
            img = Image.new('RGB', (300, 225), color)
            draw = ImageDraw.Draw(img)

            # 绘制分类标签
            text = asset['category']
            try:
                font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 16)
            except (OSError, IOError):
                font = ImageFont.load_default()

            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
            x = (300 - text_w) // 2
            y = (225 - text_h) // 2 - 15
            draw.text((x, y), text, fill='white', font=font)

            # 绘制幻灯片编号
            num_text = f"Slide {asset['slideNumber']}"
            bbox2 = draw.textbbox((0, 0), num_text, font=font)
            nw = bbox2[2] - bbox2[0]
            draw.text(((300 - nw) // 2, y + text_h + 10), num_text, fill='white', font=font)

            img.save(thumb_path, 'PNG')
            generated += 1

    return generated


def main():
    print("=" * 60)
    print("🖼️  Thumbnail Generator")
    print("=" * 60)

    os.makedirs(THUMBNAILS_DIR, exist_ok=True)

    # 方案1: 尝试 LibreOffice
    pptx_path = find_pptx()
    if pptx_path:
        print(f"\n📄 Found PPTX: {pptx_path.name}")
        print("  Attempting LibreOffice conversion...")
        if generate_via_libreoffice(pptx_path):
            count = len(list(THUMBNAILS_DIR.glob('slide_*.png')))
            print(f"\n✅ Generated {count} thumbnails via LibreOffice")
            return

    # 方案2: 从媒体文件提取
    print("\n⚠️  LibreOffice not available, using media extraction fallback")
    count = generate_from_media_fallback()
    print(f"  Generated {count} thumbnails from media files")

    # 方案3: 生成占位图
    print("\n🎨 Generating placeholder thumbnails for remaining slides...")
    placeholder_count = generate_placeholder_thumbnails()
    print(f"  Generated {placeholder_count} placeholder thumbnails")

    total = len(list(THUMBNAILS_DIR.glob('slide_*.png')))
    print(f"\n✅ Total thumbnails: {total}")


if __name__ == '__main__':
    main()
