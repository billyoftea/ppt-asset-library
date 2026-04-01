#!/usr/bin/env python3
"""
Apply Tencent Brand Theme to bundle.pptx
=========================================
- 修改所有 theme XML 的颜色方案为腾讯蓝 (TDesign 色彩体系)
- 修改所有 theme XML 的字体方案为腾讯体 (TencentSans)
- 批量替换幻灯片中硬编码的 PwC 品牌色
- 批量替换幻灯片中硬编码的字体名称
"""

import os
import re
import glob

UNPACKED_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "unpacked_bundle")

# ============================================================
# 1. 腾讯品牌色定义 (TDesign 标准)
# ============================================================
TENCENT_COLOR_SCHEME = {
    "dk1": "181818",      # 主文字色 (深色1)
    "lt1": "FFFFFF",      # 浅色1 (白色)
    "dk2": "0052D9",      # 深色2 = 腾讯蓝主色
    "lt2": "ECF2FE",      # 浅色2 = 腾讯蓝最浅
    "accent1": "0052D9",  # 强调色1 = 腾讯蓝
    "accent2": "0034B5",  # 强调色2 = 腾讯蓝 Hover
    "accent3": "2BA471",  # 强调色3 = TDesign 成功绿
    "accent4": "D54941",  # 强调色4 = TDesign 错误红
    "accent5": "E37318",  # 强调色5 = TDesign 警告橙
    "accent6": "4787EB",  # 强调色6 = 腾讯蓝90级
    "hlink": "0052D9",    # 超链接 = 腾讯蓝
    "folHlink": "0034B5", # 已访问链接
}

# ============================================================
# 2. PwC 品牌色 → 腾讯品牌色 映射 (用于替换幻灯片中硬编码的颜色)
# ============================================================
COLOR_REPLACEMENTS = {
    # PwC Orange 主色 → 腾讯蓝
    "DC6900": "0052D9",
    "dc6900": "0052D9",
    # PwC 金色 → 腾讯蓝 hover
    "FFB600": "0034B5",
    "ffb600": "0034B5",
    # PwC 深褐 → 腾讯蓝 active
    "602320": "002A7C",
    # PwC 玫瑰粉 → 腾讯蓝 90
    "E27588": "4787EB",
    "e27588": "4787EB",
    # PwC 深红 → TDesign 红
    "A32020": "D54941",
    "a32020": "D54941",
    # PwC 红 → TDesign 红 hover
    "E0301E": "C9514B",
    "e0301e": "C9514B",
    # PwC 灰金 → 腾讯蓝浅色
    "968C6D": "0052D9",
    # PwC 灰金浅色 → 腾讯蓝最浅
    "D5D1C5": "D4E3FC",
    "d5d1c5": "D4E3FC",
}

# ============================================================
# 3. 字体替换映射
# ============================================================
FONT_REPLACEMENTS = {
    # 标题字体
    "Georgia": "TencentSans W7",
    # 正文字体
    "Arial": "TencentSans W3",
}

# 东亚字体设置
EA_FONT_MAJOR = "腾讯体 W7"  # 标题
EA_FONT_MINOR = "腾讯体 W3"  # 正文


def replace_color_scheme_in_theme(xml_content: str) -> str:
    """替换 theme XML 中的 <a:clrScheme> 整块"""
    
    # 构建新的颜色方案 XML
    new_clr_items = []
    for tag, color in TENCENT_COLOR_SCHEME.items():
        new_clr_items.append(f'      <a:{tag}>\n        <a:srgbClr val="{color}"/>\n      </a:{tag}>')
    
    new_clr_scheme = '    <a:clrScheme name="Tencent Blue">\n' + \
                     '\n'.join(new_clr_items) + \
                     '\n    </a:clrScheme>'
    
    # 用正则替换整个 clrScheme 块
    pattern = r'<a:clrScheme\b[^>]*>.*?</a:clrScheme>'
    result = re.sub(pattern, new_clr_scheme, xml_content, flags=re.DOTALL)
    return result


def replace_font_scheme_in_theme(xml_content: str) -> str:
    """替换 theme XML 中的 <a:fontScheme> 块"""
    
    new_font_scheme = '''    <a:fontScheme name="Tencent">
      <a:majorFont>
        <a:latin typeface="TencentSans W7"/>
        <a:ea typeface="腾讯体 W7"/>
        <a:cs typeface="TencentSans W7"/>
        <a:font script="Hans" typeface="腾讯体 W7"/>
        <a:font script="Hant" typeface="腾讯体 W7"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="TencentSans W3"/>
        <a:ea typeface="腾讯体 W3"/>
        <a:cs typeface="TencentSans W3"/>
        <a:font script="Hans" typeface="腾讯体 W3"/>
        <a:font script="Hant" typeface="腾讯体 W3"/>
      </a:minorFont>
    </a:fontScheme>'''
    
    pattern = r'<a:fontScheme\b[^>]*>.*?</a:fontScheme>'
    result = re.sub(pattern, new_font_scheme, xml_content, flags=re.DOTALL)
    return result


def replace_theme_name(xml_content: str) -> str:
    """替换主题名称"""
    result = re.sub(
        r'<a:theme\b([^>]*)\bname="[^"]*"',
        r'<a:theme\1name="Tencent Blue Theme"',
        xml_content
    )
    return result


def replace_object_defaults_fonts(xml_content: str) -> str:
    """替换 objectDefaults 中硬编码的字体"""
    for old_font, new_font in FONT_REPLACEMENTS.items():
        xml_content = xml_content.replace(
            f'typeface="{old_font}"',
            f'typeface="{new_font}"'
        )
    return xml_content


def replace_hardcoded_colors_in_slide(xml_content: str) -> str:
    """替换幻灯片 XML 中硬编码的 PwC 品牌色"""
    for old_color, new_color in COLOR_REPLACEMENTS.items():
        # 替换 srgbClr val="XXXXXX"
        xml_content = xml_content.replace(
            f'val="{old_color}"',
            f'val="{new_color}"'
        )
    return xml_content


def replace_hardcoded_fonts_in_slide(xml_content: str) -> str:
    """替换幻灯片 XML 中硬编码的字体名称"""
    for old_font, new_font in FONT_REPLACEMENTS.items():
        xml_content = xml_content.replace(
            f'typeface="{old_font}"',
            f'typeface="{new_font}"'
        )
        # 也处理带 pitchFamily 的情况
        xml_content = xml_content.replace(
            f'typeface="{old_font}" pitchFamily=',
            f'typeface="{new_font}" pitchFamily='
        )
    return xml_content


def process_theme_files():
    """处理所有主题文件"""
    theme_dir = os.path.join(UNPACKED_DIR, "ppt", "theme")
    modified = 0
    
    for theme_file in sorted(glob.glob(os.path.join(theme_dir, "theme*.xml"))):
        filename = os.path.basename(theme_file)
        with open(theme_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original = content
        content = replace_color_scheme_in_theme(content)
        content = replace_font_scheme_in_theme(content)
        content = replace_theme_name(content)
        content = replace_object_defaults_fonts(content)
        
        if content != original:
            with open(theme_file, 'w', encoding='utf-8') as f:
                f.write(content)
            modified += 1
            print(f"  ✅ {filename} — 颜色+字体已更新")
        else:
            print(f"  ⏭️  {filename} — 无需修改")
    
    # 处理 themeOverride
    for override_file in sorted(glob.glob(os.path.join(theme_dir, "themeOverride*.xml"))):
        filename = os.path.basename(override_file)
        with open(override_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original = content
        content = replace_color_scheme_in_theme(content)
        content = replace_font_scheme_in_theme(content)
        content = replace_object_defaults_fonts(content)
        
        if content != original:
            with open(override_file, 'w', encoding='utf-8') as f:
                f.write(content)
            modified += 1
            print(f"  ✅ {filename} — 颜色+字体已更新")
        else:
            print(f"  ⏭️  {filename} — 无需修改")
    
    return modified


def process_slide_files():
    """处理所有幻灯片文件"""
    slides_dir = os.path.join(UNPACKED_DIR, "ppt", "slides")
    color_modified = 0
    font_modified = 0
    
    for slide_file in sorted(glob.glob(os.path.join(slides_dir, "slide*.xml"))):
        filename = os.path.basename(slide_file)
        with open(slide_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original = content
        
        # 替换硬编码颜色
        content_after_color = replace_hardcoded_colors_in_slide(content)
        if content_after_color != content:
            color_modified += 1
        
        # 替换硬编码字体
        content_after_font = replace_hardcoded_fonts_in_slide(content_after_color)
        if content_after_font != content_after_color:
            font_modified += 1
        
        if content_after_font != original:
            with open(slide_file, 'w', encoding='utf-8') as f:
                f.write(content_after_font)
    
    return color_modified, font_modified


def process_master_and_layout_files():
    """处理母版和版式文件"""
    modified = 0
    
    # 母版
    masters_dir = os.path.join(UNPACKED_DIR, "ppt", "slideMasters")
    for master_file in sorted(glob.glob(os.path.join(masters_dir, "slideMaster*.xml"))):
        filename = os.path.basename(master_file)
        with open(master_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original = content
        content = replace_hardcoded_colors_in_slide(content)
        content = replace_hardcoded_fonts_in_slide(content)
        
        if content != original:
            with open(master_file, 'w', encoding='utf-8') as f:
                f.write(content)
            modified += 1
    
    # 版式
    layouts_dir = os.path.join(UNPACKED_DIR, "ppt", "slideLayouts")
    if os.path.exists(layouts_dir):
        for layout_file in sorted(glob.glob(os.path.join(layouts_dir, "slideLayout*.xml"))):
            with open(layout_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original = content
            content = replace_hardcoded_colors_in_slide(content)
            content = replace_hardcoded_fonts_in_slide(content)
            
            if content != original:
                with open(layout_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                modified += 1
    
    return modified


def main():
    print("=" * 60)
    print("  Applying Tencent Brand Theme to bundle.pptx")
    print("  腾讯蓝 #0052D9 | TencentSans 腾讯体")
    print("=" * 60)
    
    print("\n📦 Step 1: 修改主题文件 (theme XML)")
    theme_count = process_theme_files()
    print(f"   → {theme_count} 个主题文件已更新")
    
    print("\n🎨 Step 2: 替换幻灯片中硬编码的颜色和字体")
    color_count, font_count = process_slide_files()
    print(f"   → {color_count} 张幻灯片的颜色已替换")
    print(f"   → {font_count} 张幻灯片的字体已替换")
    
    print("\n📐 Step 3: 替换母版和版式中的颜色和字体")
    master_count = process_master_and_layout_files()
    print(f"   → {master_count} 个母版/版式文件已更新")
    
    print("\n" + "=" * 60)
    print("  ✅ 腾讯品牌主题应用完成！")
    print("  下一步：运行 pack 脚本重新打包为 bundle.pptx")
    print("=" * 60)


if __name__ == "__main__":
    main()
