#!/usr/bin/env python3
"""
Script to find and delete empty pages from DOCX files.
Works by extracting DOCX (ZIP), parsing document.xml, and removing empty pages.
No external libraries required - uses only standard Python.

Detects empty pages as:
- Consecutive empty paragraphs (created by Enter key)
- Section breaks followed by empty content
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
import tempfile
from pathlib import Path
import re

# DOCX XML namespace
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def extract_docx(docx_path, extract_dir):
    """Extract DOCX file to a directory."""
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)


def repack_docx(extract_dir, output_path):
    """Repack extracted files back into a DOCX."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zipf.write(file_path, arcname)


def get_paragraph_text(para):
    """Get all text content from a paragraph."""
    text = ""
    for t in para.findall('.//w:t', NS):
        if t.text:
            text += t.text
    return text.strip()


def is_empty_paragraph(para):
    """Check if paragraph is empty (no text content)."""
    return len(get_paragraph_text(para)) == 0


def has_page_break(para):
    """Check if paragraph contains a page break."""
    # Check for <w:br w:type="page"/>
    for br in para.findall('.//w:br', NS):
        br_type = br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
        if br_type == 'page':
            return True
    return False


def has_section_break(para):
    """Check if paragraph contains a section break (which creates a new page)."""
    sect = para.find('.//w:sectPr', NS)
    if sect is not None:
        sect_type = sect.find('.//w:type', NS)
        if sect_type is not None:
            type_val = sect_type.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            return type_val in ['nextPage', 'oddPage', 'evenPage']
    return False


def find_empty_page_regions(document_xml_path, min_empty_paragraphs=15):
    """
    Find regions of consecutive empty paragraphs that likely form empty pages.
    
    Args:
        document_xml_path: Path to document.xml
        min_empty_paragraphs: Minimum consecutive empty paragraphs to consider as "empty page"
    
    Returns:
        List of empty page regions with their paragraph indices
    """
    tree = ET.parse(document_xml_path)
    root = tree.getroot()
    
    body = root.find('.//w:body', NS)
    if body is None:
        print("Could not find document body.")
        return [], tree, root, body
    
    paragraphs = list(body)
    empty_regions = []
    
    # Track consecutive empty paragraphs
    current_empty_start = None
    current_empty_count = 0
    current_empty_indices = []
    
    for i, para in enumerate(paragraphs):
        tag_local = para.tag.split('}')[-1] if '}' in para.tag else para.tag
        
        if tag_local == 'p':  # It's a paragraph
            if is_empty_paragraph(para) and not has_section_break(para):
                # Empty paragraph (and not containing a section break we want to keep)
                if current_empty_start is None:
                    current_empty_start = i
                current_empty_count += 1
                current_empty_indices.append(i)
            else:
                # Non-empty paragraph - check if we had a run of empties
                if current_empty_count >= min_empty_paragraphs:
                    empty_regions.append({
                        'start_index': current_empty_start,
                        'end_index': current_empty_indices[-1],
                        'count': current_empty_count,
                        'indices': current_empty_indices.copy()
                    })
                # Reset
                current_empty_start = None
                current_empty_count = 0
                current_empty_indices = []
        else:
            # Not a paragraph - reset counter
            if current_empty_count >= min_empty_paragraphs:
                empty_regions.append({
                    'start_index': current_empty_start,
                    'end_index': current_empty_indices[-1],
                    'count': current_empty_count,
                    'indices': current_empty_indices.copy()
                })
            current_empty_start = None
            current_empty_count = 0
            current_empty_indices = []
    
    # Check if document ends with empty paragraphs
    if current_empty_count >= min_empty_paragraphs:
        empty_regions.append({
            'start_index': current_empty_start,
            'end_index': current_empty_indices[-1],
            'count': current_empty_count,
            'indices': current_empty_indices.copy()
        })
    
    return empty_regions, tree, root, body, paragraphs


def delete_empty_regions(regions_to_delete, body, paragraphs, tree, document_xml_path):
    """Delete selected empty page regions from the document."""
    # Collect all indices to delete
    indices_to_delete = set()
    for region in regions_to_delete:
        indices_to_delete.update(region['indices'])
    
    # Remove elements (in reverse order to maintain indices)
    for i in sorted(indices_to_delete, reverse=True):
        elem = paragraphs[i]
        try:
            body.remove(elem)
        except ValueError:
            pass
    
    # Re-register namespaces to preserve them
    namespaces_to_register = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'o': 'urn:schemas-microsoft-com:office:office',
        'v': 'urn:schemas-microsoft-com:vml',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    }
    for prefix, uri in namespaces_to_register.items():
        ET.register_namespace(prefix, uri)
    
    # Save the modified document
    tree.write(document_xml_path, xml_declaration=True, encoding='UTF-8')


def process_docx(docx_path, min_empty=15):
    """Main function to process a single DOCX file."""
    docx_path = Path(docx_path)
    
    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}")
        return
    
    if not docx_path.suffix.lower() == '.docx':
        print(f"Error: Not a DOCX file: {docx_path}")
        return
    
    print(f"\n{'='*60}")
    print(f"Processing: {docx_path.name}")
    print('='*60)
    
    # Create temp directory for extraction
    with tempfile.TemporaryDirectory() as temp_dir:
        extract_dir = Path(temp_dir) / 'extracted'
        extract_dir.mkdir()
        
        # Extract DOCX
        print("\n📦 Extracting DOCX...")
        extract_docx(docx_path, extract_dir)
        
        # Find document.xml
        document_xml = extract_dir / 'word' / 'document.xml'
        if not document_xml.exists():
            print("Error: document.xml not found in DOCX")
            return
        
        # Find empty page regions
        print(f"🔍 Scanning for empty page regions ({min_empty}+ consecutive empty paragraphs)...")
        result = find_empty_page_regions(document_xml, min_empty)
        empty_regions, tree, root, body, paragraphs = result
        
        if not empty_regions:
            print("\n✅ No empty page regions found!")
            print(f"   (Looking for {min_empty}+ consecutive empty paragraphs)")
            return
        
        # Display empty regions
        print(f"\n⚠️  Found {len(empty_regions)} empty page region(s):")
        for i, region in enumerate(empty_regions, 1):
            print(f"  {i}. {region['count']} empty paragraphs (elements {region['start_index']}-{region['end_index']})")
        
        # Ask for user confirmation
        print("\nOptions:")
        print("  [a] Delete ALL empty page regions")
        print("  [s] Select regions to delete")
        print("  [n] Cancel - don't delete anything")
        
        choice = input("\nYour choice (a/s/n): ").strip().lower()
        
        regions_to_delete = []
        
        if choice == 'a':
            confirm = input("\n⚠️  Delete all empty page regions? Type 'yes' to confirm: ").strip().lower()
            if confirm == 'yes':
                regions_to_delete = empty_regions
        elif choice == 's':
            print("\nFor each region, type 'y' to delete or 'n' to keep:")
            for i, region in enumerate(empty_regions, 1):
                response = input(f"  Delete region {i} ({region['count']} empty paragraphs)? (y/n): ").strip().lower()
                if response == 'y':
                    regions_to_delete.append(region)
        else:
            print("Operation cancelled.")
            return
        
        if not regions_to_delete:
            print("No regions selected for deletion.")
            return
        
        # Count total paragraphs to remove
        total_to_remove = sum(r['count'] for r in regions_to_delete)
        
        # Delete the empty regions
        print(f"\n🗑️  Deleting {total_to_remove} empty paragraphs from {len(regions_to_delete)} region(s)...")
        delete_empty_regions(regions_to_delete, body, paragraphs, tree, document_xml)
        
        # Create backup and save
        backup_path = docx_path.with_suffix('.backup.docx')
        shutil.copy(docx_path, backup_path)
        print(f"📄 Backup created: {backup_path.name}")
        
        # Repack DOCX
        repack_docx(extract_dir, docx_path)
        print(f"✅ Saved: {docx_path.name}")
        
        print(f"\n🎉 Successfully removed {total_to_remove} empty paragraphs!")


def main():
    """Main entry point."""
    print("=" * 60)
    print("      DOCX Empty Page Finder & Remover")
    print("      (No external libraries required)")
    print("=" * 60)
    
    # Get DOCX file path from user
    docx_input = input("\nEnter path to DOCX file: ").strip()
    
    if not docx_input:
        print("No file specified.")
        return
    
    # Handle quoted paths
    docx_input = docx_input.strip('"\'')
    
    # Ask for sensitivity
    print("\nHow many consecutive empty paragraphs should be considered an 'empty page'?")
    print("(A typical page has ~25-30 lines, so 15-20 is a good minimum)")
    min_input = input("Minimum empty paragraphs [default: 15]: ").strip()
    
    try:
        min_empty = int(min_input) if min_input else 15
    except ValueError:
        min_empty = 15
    
    process_docx(docx_input, min_empty)
    
    print("\nDone!")


if __name__ == "__main__":
    main()
