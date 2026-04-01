#!/usr/bin/env python3
"""
Replace PwC branding with FiT branding in bundle.pptx XML files:
1. Replace "PwC" text with "FiT"
2. Replace PwC red/orange brand colors with Tencent blue palette
"""

import os
import re
import glob

UNPACKED_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'unpacked_bundle')

# ============================================================
# Part 1: PwC -> FiT text replacement
# ============================================================
def replace_pwc_text():
    """Replace all PwC text occurrences with FiT"""
    count = 0
    files_modified = 0
    
    xml_files = glob.glob(os.path.join(UNPACKED_DIR, 'ppt', '**', '*.xml'), recursive=True)
    
    for fpath in xml_files:
        with open(fpath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Count occurrences
        occurrences = content.count('PwC')
        if occurrences == 0:
            continue
        
        new_content = content.replace('PwC', 'FiT')
        
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        count += occurrences
        files_modified += 1
        print(f"  ✅ {os.path.relpath(fpath, UNPACKED_DIR)}: {occurrences} 处 PwC -> FiT")
    
    return count, files_modified

# ============================================================
# Part 2: Red/Orange color replacement
# ============================================================

# Tencent Design Token blue palette (from dark to light)
TENCENT_BLUE = {
    'primary':     '0052D9',  # 主蓝
    'dark':        '0034B5',  # 深蓝
    'darker':      '001F97',  # 更深蓝
    'light':       '366EF4',  # 亮蓝
    'lighter':     '618DFF',  # 更亮蓝
    'soft':        '96BBF8',  # 柔蓝
    'softer':      'BBD3FB',  # 更柔蓝
    'pale':        'D1E3FF',  # 浅蓝
    'palest':      'ECF2FE',  # 最浅蓝
    'accent':      '4787EB',  # 辅助蓝
    'teal':        '2BA471',  # 腾讯绿（辅助色）
}

# Color mapping: PwC red/orange -> Tencent blue palette
# Strategy: map by brightness/saturation level to maintain visual hierarchy
COLOR_MAP = {
    # === PwC 经典橙色 -> 腾讯主蓝 ===
    'EB8C00': '0052D9',  # PwC 橙 -> 主蓝 (209次，最多)
    'EB8C03': '0052D9',  # PwC 橙变体
    'BF8800': '0034B5',  # 暗橙 -> 深蓝
    'A54E00': '001F97',  # 深橙 -> 更深蓝
    'E88C14': '366EF4',  # 橙色 -> 亮蓝
    'EE9C34': '4787EB',  # 浅橙 -> 辅助蓝
    'E98C23': '366EF4',  # 橙色 -> 亮蓝
    'FFC227': '96BBF8',  # 亮黄橙 -> 柔蓝
    'FFC021': '96BBF8',  # 亮黄橙变体
    'FCC000': '96BBF8',  # 金黄 -> 柔蓝
    'FDB515': 'BBD3FB',  # 金黄 -> 更柔蓝
    'FFA451': 'BBD3FB',  # 浅橙 -> 更柔蓝
    'FCB770': 'D1E3FF',  # 淡橙 -> 浅蓝
    'FBB36A': 'D1E3FF',  # 淡橙变体
    'F8A26C': 'BBD3FB',  # 浅橙 -> 更柔蓝
    'F6B67F': 'D1E3FF',  # 淡橙 -> 浅蓝
    'F89C63': 'BBD3FB',  # 浅橙 -> 更柔蓝

    # === PwC 橙红 -> 主蓝/亮蓝 ===
    'EC6408': '366EF4',  # 橙红 -> 亮蓝
    'E37318': '366EF4',  # 橙红 -> 亮蓝 (主题中的)
    'F36F24': '618DFF',  # 橙红 -> 更亮蓝
    'DB6A27': '4787EB',  # 暗橙 -> 辅助蓝
    'F48F17': '618DFF',  # 橙 -> 更亮蓝
    'EB660B': '366EF4',  # 橙红 -> 亮蓝
    'EA8804': '4787EB',  # 橙色 -> 辅助蓝
    'E36A00': '366EF4',  # 橙色

    # === PwC 深红 -> 深蓝/主蓝 ===
    'AD2624': '001F97',  # 暗红 -> 更深蓝
    'A10000': '001F97',  # 深红 -> 更深蓝
    'A32020': '001F97',  # 深红 -> 更深蓝
    'A32022': '001F97',  # 深红变体
    'A21B1B': '001F97',  # 深红变体
    '9A1702': '001F97',  # 深红变体
    'B91E2B': '0034B5',  # 暗红 -> 深蓝
    'BB2741': '0034B5',  # 暗红 -> 深蓝
    'C00000': '0034B5',  # 深红 -> 深蓝
    'C01000': '0034B5',  # 深红变体
    'C93C00': '0034B5',  # 深红变体
    'C22303': '0034B5',  # 深红变体
    'C42303': '0034B5',  # 深红变体
    'C43450': '0034B5',  # 暗红紫 -> 深蓝

    # === PwC 亮红 -> 主蓝/亮蓝 ===
    'D54941': '0052D9',  # 腾讯红(主题色) -> 主蓝 ★ 核心映射
    'D62E1C': '0052D9',  # 红 -> 主蓝
    'D74021': '0052D9',  # 红 -> 主蓝
    'D13A0D': '0052D9',  # 红 -> 主蓝
    'D1390D': '0052D9',  # 红变体
    'D61400': '0052D9',  # 红 -> 主蓝
    'CD2F12': '0034B5',  # 红 -> 深蓝
    'DF3326': '0052D9',  # 红 -> 主蓝
    'DB536A': '366EF4',  # 粉红 -> 亮蓝
    'DB546A': '366EF4',  # 粉红变体
    'DB4D56': '366EF4',  # 粉红 -> 亮蓝

    # === 亮红/粉红 -> 亮蓝/柔蓝 ===
    'E0301E': '0052D9',  # 红 -> 主蓝
    'E02504': '0052D9',  # 红 -> 主蓝
    'E04C00': '0052D9',  # 红橙 -> 主蓝
    'E40428': '0052D9',  # 亮红 -> 主蓝
    'E64135': '366EF4',  # 亮红 -> 亮蓝
    'E93409': '0052D9',  # 红 -> 主蓝
    'ED1E37': '0052D9',  # 红 -> 主蓝
    'EE2E3C': '0052D9',  # 红 -> 主蓝

    # === 粉红系 -> 柔蓝系 ===
    'E27083': '618DFF',  # 粉红 -> 更亮蓝
    'C66766': '4787EB',  # 暗粉 -> 辅助蓝
    'D48C8C': '96BBF8',  # 淡粉 -> 柔蓝
    'E669A2': '618DFF',  # 玫红 -> 更亮蓝
    'E998A6': 'BBD3FB',  # 浅粉 -> 更柔蓝
    'F08371': '618DFF',  # 珊瑚粉 -> 更亮蓝
    'F47C6B': '618DFF',  # 珊瑚 -> 更亮蓝
    'F47F6D': '618DFF',  # 珊瑚变体
    'EF4364': '366EF4',  # 玫红 -> 亮蓝
    'F48D92': 'BBD3FB',  # 浅粉 -> 更柔蓝
    'F4888D': 'BBD3FB',  # 浅粉变体
    'C9514B': '4787EB',  # 暗珊瑚 -> 辅助蓝
    'C55852': '4787EB',  # 暗珊瑚变体
    'DA9089': '96BBF8',  # 淡粉 -> 柔蓝
    'AB5E55': '0034B5',  # 暗棕红 -> 深蓝
    'A75750': '0034B5',  # 暗棕红变体

    # === 纯红/火红 -> 主蓝 ===
    'FB3D32': '0052D9',  # 火红 -> 主蓝
    'FF0000': '0052D9',  # 纯红 -> 主蓝
}

def replace_colors():
    """Replace PwC red/orange colors with Tencent blue palette"""
    total_replacements = 0
    files_modified = 0
    
    xml_files = glob.glob(os.path.join(UNPACKED_DIR, 'ppt', '**', '*.xml'), recursive=True)
    
    for fpath in xml_files:
        with open(fpath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        file_replacements = 0
        new_content = content
        
        for old_color, new_color in COLOR_MAP.items():
            # 替换大写和小写形式
            for variant in [old_color.upper(), old_color.lower()]:
                count = new_content.count(f'"{variant}"')
                if count > 0:
                    new_content = new_content.replace(f'"{variant}"', f'"{new_color}"')
                    file_replacements += count
        
        if file_replacements > 0:
            with open(fpath, 'w', encoding='utf-8') as f:
                f.write(new_content)
            total_replacements += file_replacements
            files_modified += 1
            print(f"  🎨 {os.path.relpath(fpath, UNPACKED_DIR)}: {file_replacements} 处颜色替换")
    
    return total_replacements, files_modified


def main():
    print("=" * 60)
    print("🔄 PwC -> FiT 品牌替换")
    print("=" * 60)
    
    # Step 1: Text replacement
    print("\n📝 Part 1: 替换 PwC 文字 -> FiT")
    print("-" * 40)
    text_count, text_files = replace_pwc_text()
    print(f"\n  总计: {text_count} 处文字替换，{text_files} 个文件")
    
    # Step 2: Color replacement
    print(f"\n🎨 Part 2: 替换红色/橙色系颜色 -> 腾讯蓝色系")
    print(f"  颜色映射: {len(COLOR_MAP)} 种颜色")
    print("-" * 40)
    color_count, color_files = replace_colors()
    print(f"\n  总计: {color_count} 处颜色替换，{color_files} 个文件")
    
    # Summary
    print("\n" + "=" * 60)
    print(f"✅ 全部完成!")
    print(f"  📝 文字: {text_count} 处 PwC -> FiT ({text_files} 个文件)")
    print(f"  🎨 颜色: {color_count} 处红/橙 -> 蓝 ({color_files} 个文件)")
    print("=" * 60)


if __name__ == '__main__':
    main()
