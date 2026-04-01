#!/usr/bin/env python3
"""
Remove all italic attributes (i="1") from PPTX XML files.
Targets: <a:rPr ... i="1" ...>, <a:defRPr ... i="1" ...>, <a:endParaRPr ... i="1" ...>
"""
import os
import re
import glob

UNPACKED_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "unpacked_bundle")

# Pattern: match i="1" as a standalone attribute in XML tags
# This handles: i="1" preceded by a space, potentially followed by space or >
ITALIC_ATTR_PATTERN = re.compile(r'\s+i="1"')

def remove_italic_from_file(filepath):
    """Remove i="1" attributes from a single XML file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Count occurrences before
    count = len(ITALIC_ATTR_PATTERN.findall(content))
    if count == 0:
        return 0
    
    # Remove i="1" attribute (the space before it gets consumed too)
    new_content = ITALIC_ATTR_PATTERN.sub('', content)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    return count

def main():
    total_files = 0
    total_replacements = 0
    
    # Process all XML files in ppt/ directory
    search_dirs = [
        os.path.join(UNPACKED_DIR, "ppt", "slides"),
        os.path.join(UNPACKED_DIR, "ppt", "slideMasters"),
        os.path.join(UNPACKED_DIR, "ppt", "slideLayouts"),
        os.path.join(UNPACKED_DIR, "ppt", "theme"),
    ]
    
    for search_dir in search_dirs:
        if not os.path.exists(search_dir):
            continue
        for filepath in sorted(glob.glob(os.path.join(search_dir, "*.xml"))):
            count = remove_italic_from_file(filepath)
            if count > 0:
                total_files += 1
                total_replacements += count
                print(f"  ✓ {os.path.relpath(filepath, UNPACKED_DIR)}: 移除 {count} 处斜体")
    
    print(f"\n✅ 完成！共修改 {total_files} 个文件，移除 {total_replacements} 处斜体属性 (i=\"1\")")

if __name__ == "__main__":
    main()
