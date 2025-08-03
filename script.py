import tempfile
import zipfile
import os
import xml.etree.ElementTree as ET

docx_path = "hello_world.docx"
PARAGRAPH_CRITERIA= {
    'min_sentences':2,
    'min_words':15
}

def _extract_docx(docx_path: str, extract_to: str):
    """Extract docx file contents to directory"""
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def _create_docx(source_dir: str, output_path: str):
    """Create docx file from directory contents"""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)        

def _create_empty_paragraph():
    """Create a standard empty paragraph element"""
    # Create paragraph element
    para = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    # Add basic paragraph properties
    pPr = ET.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    
    # Add spacing properties for consistent formatting
    spacing = ET.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '276')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
    
    # Add empty run
    run = ET.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    
    return para


def _count_sentences(text: str) -> int:
    """Count sentences in text, handling abbreviations properly"""
    import re
    
    # Get abbreviations list
    abbreviations = [
      "PT", "CV", "UD", "Tbk", "Ltd", "Inc", "Corp",
      "Dr", "dr", "Prof", "Ir", "Drs", "Dra", "ST", "SE", "SH", "MM", "M.Si", "M.Kom", "M.Pd", "S.Kom", "S.E", "S.H",
      "Hj", "H", "KH",
      "No", "Nomor", "Tel", "Telp", "Fax", "Hp",
      "Jl", "Jln", "Gg",
      "Kec", "Kel", "Kab", "Prov", "RT", "RW",
      "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"
    ]
    
    # Create a pattern for abbreviations with periods
    abbrev_pattern = r'\b(' + '|'.join(re.escape(abbr) for abbr in abbreviations) + r')\.'
    
    # Replace abbreviations with placeholders to avoid false sentence endings
    text_processed = re.sub(abbrev_pattern, r'\1<ABBREV>', text, flags=re.IGNORECASE)
    
    # Count actual sentence endings
    # Look for periods, exclamation marks, or question marks that are:
    # 1. Followed by whitespace and uppercase letter (new sentence)
    # 2. At the end of the text
    sentence_pattern = r'[.!?]+(?:\s+[A-Z]|$)'
    sentences = re.findall(sentence_pattern, text_processed)
    
    # If no sentence endings found but text exists, count as 1 sentence
    sentence_count = len(sentences)
    if sentence_count == 0 and len(text.strip()) > 0:
        sentence_count = 1
        
    return sentence_count        

def _get_text_content(element):
    text_elements = element.findall('.//w:t', namespaces)
    text_content = ''.join([t.text or '' for t in text_elements])
    return text_content

def _is_list_item(element):
    """Check if element is part of a list"""
    numPr = element.find('.//w:numPr', namespaces)
    return numPr is not None   


def _is_last_list_item(element, next_element):
    """Check if current element is the last item in its list/sublist"""
    # Current must be a list item
    current_numPr = element.find('.//w:numPr', namespaces)
    if current_numPr is None:
        return False

    # Get current list level
    current_ilvl = current_numPr.find('.//w:ilvl', namespaces)
    current_level = 0
    if current_ilvl is not None:
        current_level = int(current_ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))
    
    # Get current list ID
    current_numId = current_numPr.find('.//w:numId', namespaces)
    current_list_id = None
    if current_numId is not None:
        current_list_id = current_numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')

    # If no next element, it's the last
    if next_element is None:
        return True

    # Check next element
    next_numPr = next_element.find('.//w:numPr', namespaces)

    # If next is not a list item, current is last in list
    if next_numPr is None:
        return True

    # Get next list properties
    next_ilvl = next_numPr.find('.//w:ilvl', namespaces)
    next_level = 0
    if next_ilvl is not None:
        next_level = int(next_ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))

    next_numId = next_numPr.find('.//w:numId', namespaces)
    next_list_id = None
    if next_numId is not None:
        next_list_id = next_numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')

    # Different list ID = current is last of its list
    if current_list_id != next_list_id:
        return True

    # If next level is less than current (going back to parent level)
    # then current is last at its level
    if next_level < current_level:
        return True

    return False 


def _cleanup_excess_empty_paragraphs(body, all_elements, current_index):
    """
    Ensure only one empty paragraph exists after the current element.
    Removes excess empty paragraphs if more than one found.
    
    Args:
        body: The document body element
        all_elements: List of all elements in body
        current_index: Index of current element we're checking after
    """
    # Start checking from the next element
    i = current_index + 1
    empty_count = 0
    elements_to_remove = []

    # Count consecutive empty paragraphs
    while i < len(all_elements):
        elem = all_elements[i]

        # Only check paragraph elements
        if not elem.tag.endswith('p'):
            break

        # Check if it's empty
        text_content = _get_text_content(elem)
        if text_content.strip() == '':
            empty_count += 1
            # Keep first empty, mark others for removal
            if empty_count > 1:
                elements_to_remove.append(elem)
            i += 1
        else:
            # Stop at first non-empty paragraph
            break

    # Remove excess empty paragraphs from body
    for elem in elements_to_remove:
        body.remove(elem)
        # Also remove from all_elements list to keep it in sync
        if elem in all_elements:
            all_elements.remove(elem)

    return len(elements_to_remove)  # Return number of removed elements    


def _count_list_items(element, all_elements, current_index):
    """
    Count total number of items in the current list (including all levels).
    
    Args:
        element: Current element that should be a list item
        all_elements: List of all elements in body
        current_index: Index of current element
        
    Returns:
        int: Total count of items in this list
    """
    # Get current list properties
    current_numPr = element.find('.//w:numPr', namespaces)
    if current_numPr is None:
        return 0

    # Get current list ID
    current_numId = current_numPr.find('.//w:numId', namespaces)
    if current_numId is None:
        return 0

    current_list_id = current_numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')

    # Count items with same list ID
    count = 0

    # Search backward from current position
    for i in range(current_index, -1, -1):
        elem = all_elements[i]
        elem_numPr = elem.find('.//w:numPr', namespaces)
        if elem_numPr:
            elem_numId = elem_numPr.find('.//w:numId', namespaces)
            if elem_numId:
                elem_list_id = elem_numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if elem_list_id == current_list_id:
                    count += 1
                else:
                    break  # Different list, stop counting
        else:
            # Not a list item, check if we've started counting
            if count > 0:
                break  # We've left the list

    # Search forward from current position + 1
    for i in range(current_index + 1, len(all_elements)):
        elem = all_elements[i]
        elem_numPr = elem.find('.//w:numPr', namespaces)
        if elem_numPr:
            elem_numId = elem_numPr.find('.//w:numId', namespaces)
            if elem_numId:
                elem_list_id = elem_numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if elem_list_id == current_list_id:
                    count += 1
                else:
                    break  # Different list, stop counting
        else:
            break  # Not a list item, we've left the list

    return count    

with tempfile.TemporaryDirectory() as temp_dir:
    # Extract docx contents
    _extract_docx(docx_path, temp_dir)
    
    # Load and process document.xml
    doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
    tree = ET.parse(doc_xml_path)
    namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)       
    root = tree.getroot()
    body = root.find('.//w:body',namespaces)
    all_elements = list(body)

    i = 0
    while i < len(all_elements):  # Use while loop since list length may change
        element = all_elements[i]
        text_content = _get_text_content(element)
        sentence_count = _count_sentences(text_content)

        next_element = all_elements[i+1] if i+1 < len(all_elements) else None

        # Check if it's a paragraph OR last list item
        is_paragraph = (sentence_count >= PARAGRAPH_CRITERIA['min_sentences'] or
                        len(text_content.split()) >= PARAGRAPH_CRITERIA['min_words'])
        is_last_of_list = _is_last_list_item(element, next_element)


        if (is_paragraph) and not _is_list_item(element):
            # Check what comes next
            if i+1 < len(all_elements):
                next_text = _get_text_content(all_elements[i+1])

                if next_text.strip() == '':
                    # There's already at least one empty paragraph
                    # Clean up any excess (keep only 1)
                    removed = _cleanup_excess_empty_paragraphs(body, all_elements, i)
                    print(f"Found empty after: {text_content[:50]}... (removed {removed} excess)")
                else:
                    # No empty paragraph exists, add one
                    new_empty = _create_empty_paragraph()
                    body.insert(i+1, new_empty)
                    all_elements.insert(i+1, new_empty)  # Keep list in sync
                    print(f"Added empty after: {text_content[:50]}...")
        elif (is_last_of_list and not text_content.isupper()):
            list_item_count = _count_list_items(element, all_elements, i)
            
            # Only add spacing for lists with more than 3 items
            if list_item_count > 3: 
                # Check what comes next
                if i+1 < len(all_elements):
                    next_text = _get_text_content(all_elements[i+1])

                    if next_text.strip() == '':
                        # There's already at least one empty paragraph
                        # Clean up any excess (keep only 1)
                        removed = _cleanup_excess_empty_paragraphs(body, all_elements, i)
                        print(f"Found empty after: {text_content[:50]}... (removed {removed} excess)")
                    else:
                        # No empty paragraph exists, add one
                        new_empty = _create_empty_paragraph()
                        body.insert(i+1, new_empty)
                        all_elements.insert(i+1, new_empty)  # Keep list in sync
                        print(f"Added empty after: {text_content[:50]}...")            
        i += 1

    # Save the modified document
    tree.write(doc_xml_path, encoding='utf-8', xml_declaration=True)
    _create_docx(temp_dir, 'output5.docx')
    



    
