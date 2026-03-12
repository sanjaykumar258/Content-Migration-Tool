import zipfile
import lxml.etree as ET

docx_path = r"d:\CONTENT MIGRATION TASK\docx_migration_test_file.docx"
with zipfile.ZipFile(docx_path) as z:
    content = z.read("word/document.xml")
    root = ET.fromstring(content)
    # Print a small part of the XML near where an image might be
    # Or just search for image-related tags
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    drawings = root.xpath(".//w:drawing", namespaces=namespaces)
    picts = root.xpath(".//w:pict", namespaces=namespaces)
    
    print(f"Found {len(drawings)} w:drawing elements")
    print(f"Found {len(picts)} w:pict elements")
    
    if drawings:
        # Check first drawing for blip
        blips = drawings[0].xpath(".//a:blip", namespaces=namespaces)
        print(f"First drawing has {len(blips)} a:blip elements")
        if blips:
            print(f"First blip r:embed: {blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')}")
