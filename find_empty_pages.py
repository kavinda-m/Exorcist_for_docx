#!/usr/bin/env python3
"""
Script to find and delete empty pages from DOCX files.
Works by extracting DOCX (ZIP), parsing document.xml, and removing empty pages.
No external libraries required - uses only standard Python.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
import tempfile
from pathlib import Path

# DOCX XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# Register namespaces to preserve them when writing
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


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


def find_empty_pages(document_xml_path):
    """
    Find empty pages in document.xml.
    Empty pages are identified by page breaks followed by no content before the next break.
    Returns list of (page_number, element_info) for empty pages.
    """
    tree = ET.parse(document_xml_path)
    root = tree.getroot()
    
    body = root.find('.//w:body', NAMESPACES)
    if body is None:
        print("Could not find document body.")
        return [], tree, root
    
    empty_pages = []
    page_num = 1
    current_page_content = []
    current_page_elements = []
    
    # Iterate through all children of body
    for elem in list(body):
        # Check for page break
        page_break = elem.find('.//w:br[@w:type="page"]', NAMESPACES)
        section_break = elem.find('.//w:sectPr', NAMESPACES)
        
        if page_break is not None or section_break is not None:
            # Check if current page is empty
            text_content = ''.join(current_page_content).strip()
            if not text_content and current_page_elements:
                empty_pages.append({
                    'page_num': page_num,
                    'elements': current_page_elements.copy(),
                    'type': 'page_break' if page_break is not None else 'section_break'
                })
            
            page_num += 1
            current_page_content = []
            current_page_elements = []
        else:
            # Collect text content
            for text_elem in elem.findall('.//w:t', NAMESPACES):
                if text_elem.text:
                    current_page_content.append(text_elem.text)
            current_page_elements.append(elem)
    
    # Check last page
    text_content = ''.join(current_page_content).strip()
    if not text_content and current_page_elements:
        empty_pages.append({
            'page_num': page_num,
            'elements': current_page_elements.copy(),
            'type': 'end_of_document'
        })
    
    return empty_pages, tree, body


def delete_empty_pages(empty_pages, body, tree, document_xml_path):
    """Delete selected empty pages from the document."""
    elements_to_remove = []
    
    for page_info in empty_pages:
        elements_to_remove.extend(page_info['elements'])
    
    for elem in elements_to_remove:
        try:
            body.remove(elem)
        except ValueError:
            pass  # Element already removed or not found
    
    # Save the modified document
    tree.write(document_xml_path, xml_declaration=True, encoding='UTF-8')


def process_docx(docx_path):
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
        
        # Find empty pages
        print("🔍 Scanning for empty pages...")
        empty_pages, tree, body = find_empty_pages(document_xml)
        
        if not empty_pages:
            print("\n✅ No empty pages found!")
            return
        
        # Display empty pages
        print(f"\n⚠️  Found {len(empty_pages)} empty page(s):")
        for i, page in enumerate(empty_pages, 1):
            print(f"  {i}. Page {page['page_num']} ({page['type']})")
        
        # Ask for user confirmation
        print("\nOptions:")
        print("  [a] Delete ALL empty pages")
        print("  [s] Select pages to delete")
        print("  [n] Cancel - don't delete anything")
        
        choice = input("\nYour choice (a/s/n): ").strip().lower()
        
        pages_to_delete = []
        
        if choice == 'a':
            confirm = input("\n⚠️  Delete all empty pages? Type 'yes' to confirm: ").strip().lower()
            if confirm == 'yes':
                pages_to_delete = empty_pages
        elif choice == 's':
            print("\nFor each page, type 'y' to delete or 'n' to keep:")
            for page in empty_pages:
                response = input(f"  Delete page {page['page_num']}? (y/n): ").strip().lower()
                if response == 'y':
                    pages_to_delete.append(page)
        else:
            print("Operation cancelled.")
            return
        
        if not pages_to_delete:
            print("No pages selected for deletion.")
            return
        
        # Delete the empty pages
        print(f"\n🗑️  Deleting {len(pages_to_delete)} page(s)...")
        delete_empty_pages(pages_to_delete, body, tree, document_xml)
        
        # Create backup and save
        backup_path = docx_path.with_suffix('.backup.docx')
        shutil.copy(docx_path, backup_path)
        print(f"📄 Backup created: {backup_path.name}")
        
        # Repack DOCX
        repack_docx(extract_dir, docx_path)
        print(f"✅ Saved: {docx_path.name}")
        
        print(f"\n🎉 Successfully removed {len(pages_to_delete)} empty page(s)!")


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
    
    process_docx(docx_input)
    
    print("\nDone!")


if __name__ == "__main__":
    main()
