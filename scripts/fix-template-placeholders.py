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


def add_body_placeholder(slide_content):
    """
    Add a body placeholder to a slide that has title but no body placeholder.
    Used for slides 13 and 14 which are title-only layouts.

    Returns:
        Modified XML string
    """
    # Check if slide already has a body placeholder
    if '<p:ph type="body"' in slide_content:
        print("  Already has body placeholder")
        return slide_content

    # Check if slide has a title placeholder (required for this to make sense)
    if '<p:ph type="title"' not in slide_content:
        print("  No title placeholder found, skipping body placeholder")
        return slide_content

    # Body placeholder shape to add
    # Position: below the title, covering most of the slide
    # Dimensions based on template: x=457200 (same as title), y=1200000, width=17373600, height=7800000
    body_placeholder = '''<p:sp><p:nvSpPr><p:cNvPr id="100" name="Content Placeholder 100"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1200000"/><a:ext cx="17373600" cy="7800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>'''

    # Insert before </p:spTree>
    if '</p:spTree>' in slide_content:
        slide_content = slide_content.replace(
            '</p:spTree>',
            body_placeholder + '</p:spTree>'
        )
        print("  Added body placeholder")

    return slide_content


def add_two_column_body_placeholders(slide_content):
    """
    Add two body placeholders for a two-column layout (slide 15).
    Left column and right column body placeholders.

    Returns:
        Modified XML string
    """
    # Check if slide already has body placeholders
    if '<p:ph type="body"' in slide_content:
        print("  Already has body placeholder(s)")
        return slide_content

    # Two-column body placeholders
    # Left column: x=457200, y=2000000, width=8200000, height=7000000
    # Right column: x=9600000, y=2000000, width=8200000, height=7000000
    left_body = '''<p:sp><p:nvSpPr><p:cNvPr id="101" name="Content Placeholder Left"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="2000000"/><a:ext cx="8200000" cy="7000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>'''

    right_body = '''<p:sp><p:nvSpPr><p:cNvPr id="102" name="Content Placeholder Right"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="9600000" y="2000000"/><a:ext cx="8200000" cy="7000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>'''

    # Insert before </p:spTree>
    if '</p:spTree>' in slide_content:
        slide_content = slide_content.replace(
            '</p:spTree>',
            left_body + right_body + '</p:spTree>'
        )
        print("  Added two-column body placeholders")

    return slide_content

def create_slide_16():
    """
    Create a new slide 16 with header (title), subheader (subTitle), and body placeholders.
    This is a commonly needed layout for content slides.
    """
    slides_dir = os.path.join(TEMP_DIR, 'ppt', 'slides')
    rels_dir = os.path.join(slides_dir, '_rels')

    # Check if slide16.xml already exists
    slide16_path = os.path.join(slides_dir, 'slide16.xml')
    if os.path.exists(slide16_path):
        print("  slide16.xml already exists, skipping creation")
        return

    # Slide 16 XML with title, subtitle, and body placeholders
    # Using same layout reference as slide 13/14 (slideLayout4)
    slide16_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="17373600" cy="700000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE"/><a:t>Header</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1000000"/><a:ext cx="17373600" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE"/><a:t>Subheader</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Content 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600000"/><a:ext cx="17373600" cy="7400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE"/><a:t>Original Equipment Solutions </a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{DFC702AA-6566-462A-94F4-D9A629CE43FC}" type="slidenum"><a:rPr lang="de-DE" smtClean="0"/><a:t>16</a:t></a:fld><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''

    # Write slide16.xml
    with open(slide16_path, 'w', encoding='utf-8') as f:
        f.write(slide16_content)
    print("  Created slide16.xml")

    # Create slide16.xml.rels (relationship to layout)
    slide16_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout5.xml"/></Relationships>'''

    rels_path = os.path.join(rels_dir, 'slide16.xml.rels')
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(slide16_rels)
    print("  Created slide16.xml.rels")

    # Update presentation.xml to include slide16
    pres_path = os.path.join(TEMP_DIR, 'ppt', 'presentation.xml')
    with open(pres_path, 'r', encoding='utf-8') as f:
        pres_content = f.read()

    # Find the highest rId for slides and add slide16
    # Add slide reference to sldIdLst
    if '</p:sldIdLst>' in pres_content:
        # Find max slide id
        import re as regex
        slide_ids = regex.findall(r'<p:sldId id="(\d+)"', pres_content)
        max_id = max(int(sid) for sid in slide_ids) if slide_ids else 256
        new_id = max_id + 1

        # Find max rId
        rids = regex.findall(r'rId(\d+)', pres_content)
        max_rid = max(int(rid) for rid in rids) if rids else 10
        new_rid = max_rid + 1

        # Add slide reference
        pres_content = pres_content.replace(
            '</p:sldIdLst>',
            f'<p:sldId id="{new_id}" r:id="rId{new_rid}"/></p:sldIdLst>'
        )
        print(f"  Added slide16 to presentation.xml (id={new_id}, rId{new_rid})")

        with open(pres_path, 'w', encoding='utf-8') as f:
            f.write(pres_content)

    # Update presentation.xml.rels to add relationship to slide16
    pres_rels_path = os.path.join(TEMP_DIR, 'ppt', '_rels', 'presentation.xml.rels')
    with open(pres_rels_path, 'r', encoding='utf-8') as f:
        pres_rels_content = f.read()

    # Add relationship for slide16
    if '</Relationships>' in pres_rels_content:
        pres_rels_content = pres_rels_content.replace(
            '</Relationships>',
            f'<Relationship Id="rId{new_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide16.xml"/></Relationships>'
        )

        with open(pres_rels_path, 'w', encoding='utf-8') as f:
            f.write(pres_rels_content)
        print("  Added slide16 relationship to presentation.xml.rels")

    # Update [Content_Types].xml to include slide16
    content_types_path = os.path.join(TEMP_DIR, '[Content_Types].xml')
    with open(content_types_path, 'r', encoding='utf-8') as f:
        ct_content = f.read()

    # Add override for slide16
    if '/ppt/slides/slide16.xml' not in ct_content:
        ct_content = ct_content.replace(
            '</Types>',
            '<Override PartName="/ppt/slides/slide16.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
        )

        with open(content_types_path, 'w', encoding='utf-8') as f:
            f.write(ct_content)
        print("  Added slide16 to [Content_Types].xml")


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

        # DISABLED - placeholder additions may cause corruption
        # # Add title placeholder to text boxes with "Titel" text
        # content = add_placeholder_to_textbox(content, 'Titel', 'ctrTitle')

        # # Add subtitle placeholder to text boxes with "Subtitel" text
        # content = add_placeholder_to_textbox(content, 'Subtitel', 'subTitle')

        # Disabled - body placeholders cause corruption
        # # For slides 13 and 14, add body placeholder (most used template slides)
        # if filename in ['slide13.xml', 'slide14.xml']:
        #     content = add_body_placeholder(content)

        # # For slide 15, add two-column body placeholders
        # if filename == 'slide15.xml':
        #     content = add_two_column_body_placeholders(content)

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

    # Skip slide 16 creation - causes PPTX corruption
    # print("\n3. Creating slide 16 (header/subheader/body)...")
    # create_slide_16()

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
