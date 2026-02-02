#!/usr/bin/env python3
"""
Fix template to add proper placeholders for the PPT Migrator.

This script:
1. Removes warning text boxes from slide 1
2. Adds <p:ph type="ctrTitle"/> to title text boxes (containing "Titel")
3. Adds <p:ph type="subTitle"/> to subtitle text boxes (containing "Subtitel")
"""

import zipfile
import os
import shutil
import re
from xml.etree import ElementTree as ET

# Paths
ORIGINAL = '/home/user/ppt-transformer/template/Template_Transitionsphase_OESL.pptx'
OUTPUT = '/home/user/ppt-transformer/template/Template_Transitionsphase_OESL_fixed.pptx'
TEMP_DIR = '/tmp/pptx_fix_v2'

# XML Namespaces used in PPTX
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

def cleanup():
    """Remove temp directory if exists"""
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)

def extract_pptx():
    """Extract the original PPTX"""
    cleanup()
    with zipfile.ZipFile(ORIGINAL, 'r') as z:
        z.extractall(TEMP_DIR)
    print(f"Extracted to {TEMP_DIR}")

def add_placeholder_to_textbox(slide_content, title_text, placeholder_type):
    """
    Add a <p:ph> element to text boxes containing the specified text.

    Args:
        slide_content: XML string of the slide
        title_text: Text to search for (e.g., "Titel" or "Subtitel")
        placeholder_type: The placeholder type to add (e.g., "ctrTitle", "subTitle")

    Returns:
        Modified XML string
    """
    # Find shapes with the specified text
    # Pattern: <p:sp>...<a:t>Titel</a:t>...</p:sp>

    # We need to find <p:sp> elements that contain the text
    sp_pattern = r'<p:sp\b[^>]*>[\s\S]*?</p:sp>'

    def process_shape(match):
        shape = match.group(0)

        # Check if this shape contains the target text
        text_pattern = rf'<a:t>{re.escape(title_text)}</a:t>'
        if not re.search(text_pattern, shape):
            return shape  # Not our target, return unchanged

        # Check if it's a textbox (has txBox="1")
        if 'txBox="1"' not in shape:
            return shape  # Not a textbox, return unchanged

        # Check if it already has a placeholder
        if '<p:ph ' in shape or '<p:ph/' in shape:
            return shape  # Already has placeholder, return unchanged

        # Check if we already added this placeholder type
        if f'type="{placeholder_type}"' in shape:
            return shape  # Already has this placeholder type

        # Now add the placeholder to <p:nvPr>
        # Case 1: <p:nvPr/> - empty nvPr
        if '<p:nvPr/>' in shape:
            shape = shape.replace(
                '<p:nvPr/>',
                f'<p:nvPr><p:ph type="{placeholder_type}"/></p:nvPr>'
            )
            print(f"  Added {placeholder_type} placeholder (empty nvPr)")
            return shape

        # Case 2: <p:nvPr>...</p:nvPr> - nvPr with content
        nvpr_pattern = r'<p:nvPr>([\s\S]*?)</p:nvPr>'
        nvpr_match = re.search(nvpr_pattern, shape)
        if nvpr_match:
            inner = nvpr_match.group(1)
            new_inner = f'<p:ph type="{placeholder_type}"/>{inner}'
            shape = shape.replace(
                f'<p:nvPr>{inner}</p:nvPr>',
                f'<p:nvPr>{new_inner}</p:nvPr>'
            )
            print(f"  Added {placeholder_type} placeholder (with content)")
            return shape

        return shape

    return re.sub(sp_pattern, process_shape, slide_content)

def remove_warning_boxes(slide_content):
    """Remove warning text boxes from slide 1"""
    # Pattern for warning box 1: "Textfeld 7" containing "PLEASE DO NOT CHANGE"
    pattern1 = r'<p:sp><p:nvSpPr><p:cNvPr id="8" name="Textfeld 7">[\s\S]*?</p:sp>'
    # Pattern for warning box 2: "Textfeld 8" containing "Save a copy"
    pattern2 = r'<p:sp><p:nvSpPr><p:cNvPr id="9" name="Textfeld 8">[\s\S]*?</p:sp>'

    count = 0
    if re.search(pattern1, slide_content):
        slide_content = re.sub(pattern1, '', slide_content)
        count += 1
    if re.search(pattern2, slide_content):
        slide_content = re.sub(pattern2, '', slide_content)
        count += 1

    if count > 0:
        print(f"  Removed {count} warning box(es)")

    return slide_content

def process_slides():
    """Process all slides to add placeholders"""
    slides_dir = os.path.join(TEMP_DIR, 'ppt', 'slides')

    for filename in sorted(os.listdir(slides_dir)):
        if not filename.endswith('.xml') or filename.startswith('_'):
            continue

        filepath = os.path.join(slides_dir, filename)
        print(f"\nProcessing {filename}...")

        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        original_content = content

        # For slide1.xml, remove warning boxes first
        if filename == 'slide1.xml':
            content = remove_warning_boxes(content)

        # Add title placeholder to text boxes with "Titel" text
        content = add_placeholder_to_textbox(content, 'Titel', 'ctrTitle')

        # Add subtitle placeholder to text boxes with "Subtitel" text
        content = add_placeholder_to_textbox(content, 'Subtitel', 'subTitle')

        # Write back if changed
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"  Saved changes")
        else:
            print(f"  No changes needed")

def repack_pptx():
    """Repack the modified files into a PPTX"""
    # Remove old output if exists
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)

    # Create new PPTX (ZIP file)
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(TEMP_DIR):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, TEMP_DIR)
                zout.write(file_path, arcname)

    print(f"\nCreated: {OUTPUT}")
    print(f"Size: {os.path.getsize(OUTPUT)} bytes")

def main():
    print("=" * 60)
    print("PPTX Template Placeholder Fixer")
    print("=" * 60)

    print("\n1. Extracting original template...")
    extract_pptx()

    print("\n2. Processing slides...")
    process_slides()

    print("\n3. Repacking template...")
    repack_pptx()

    print("\n4. Cleaning up...")
    cleanup()

    print("\n" + "=" * 60)
    print("Done! Fixed template saved to:")
    print(OUTPUT)
    print("=" * 60)

if __name__ == '__main__':
    main()
