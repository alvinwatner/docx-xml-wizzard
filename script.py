import tempfile
import zipfile
import os
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from typing import List, Dict

docx_path = "hello_world.docx"
PARAGRAPH_CRITERIA= {
    'min_sentences':2,
    'min_words':15
}

PAGE_CONTENT_HEIGHT = 727.2  # A4 page content height in points
CONTENT_WIDTH = 447.9  # A4 content width in points

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

def _ensure_single_empty_paragraph_after(body: ET.Element, all_elements: list[ET.Element], current_index: int, text_content: str):
      """
      Ensure exactly one empty paragraph exists after the current element.
      Either adds one if missing or removes excess if more than one exists.
      
      Args:
          body: The document body element
          all_elements: List of all elements in body
          current_index: Index of current element
          text_content: Text content of current element (for logging)
      
      Returns:
          str: Action taken ('added', 'cleaned', or 'none')
      """
      # Check what comes next
      if current_index + 1 < len(all_elements):
          next_text = _get_text_content(all_elements[current_index + 1])

          if next_text.strip() == '':
              # There's already at least one empty paragraph
              # Clean up any excess (keep only 1)
              removed = _cleanup_excess_empty_paragraphs(body, all_elements, current_index)
              print(f"Found empty after: {text_content[:50]}... (removed {removed} excess)")
              return 'cleaned'
          else:
              # No empty paragraph exists, add one
              new_empty = _create_empty_paragraph()
              body.insert(current_index + 1, new_empty)
              all_elements.insert(current_index + 1, new_empty)  # Keep list in sync
              print(f"Added empty after: {text_content[:50]}...")
              return 'added'
      return 'none'    

def _get_list_level(element):
    """Get the list level of an element (0-based)"""
    numPr = element.find('.//w:numPr', namespaces)
    if numPr is None:
        return None
    
    ilvl = numPr.find('.//w:ilvl', namespaces)
    if ilvl is not None:
        return int(ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))
    return 0

def _is_heading(element):
    """
    Determine if an element is a heading.
    For now, we check if it's a level-0 list item with uppercase text
    (like "1. IDENTIFIKASI STATUS PENILAI")
    """
    text = _get_text_content(element).strip()
    
    if not text:
        return False
    
    # Check if it's a numbered list item at level 0
    if _is_list_item(element):
        level = _get_list_level(element)
        if level == 0:
            # Remove numbering and check if remaining text has uppercase words
            # Split and skip the first part (usually the number)
            words = text.split()
            if len(words) > 1:
                # Check if significant words are uppercase
                significant_text = ' '.join(words[1:])
                # If more than 50% of alphabetic characters are uppercase, consider it a heading
                upper_count = sum(1 for c in significant_text if c.isupper())
                alpha_count = sum(1 for c in significant_text if c.isalpha())
                if alpha_count > 0 and upper_count / alpha_count > 0.5:
                    return True
    
    return False


def _estimate_element_height(element) -> float:
    """
    Estimate the height of an element in points.
    Based on our document analysis work.
    """
    text = _get_text_content(element)
    
    # Default font measurements
    font_size = 11.0  # Default font size in points
    line_height = font_size * 1.5  # Typical line spacing (1.5x)
    
    if not text.strip():
        # Empty paragraph - single line height
        return line_height
    
    # Text height calculation using our formula
    char_width = font_size * 0.6  # Average character width
    chars_per_line = CONTENT_WIDTH / char_width
    
    # Calculate number of lines
    num_lines = max(1, len(text) / chars_per_line)
    height = num_lines * line_height
    
    # Add some padding for paragraph spacing
    height += 6  # Small padding between paragraphs
    
    return height

def _calculate_page_positions(all_elements: List[Element]) -> Dict[int, int]:
    """
    Calculate which page each element falls on.
    Returns a dictionary mapping element index to page number.
    """
    cumulative_height = 0.0
    current_page = 1
    element_pages = {}
    
    for i, element in enumerate(all_elements):
        elem_height = _estimate_element_height(element)
        
        # Check if this element would overflow to next page
        if cumulative_height + elem_height > PAGE_CONTENT_HEIGHT:
            # Start new page
            current_page += 1
            cumulative_height = elem_height
        else:
            cumulative_height += elem_height
        
        element_pages[i] = current_page
        
        # Debug output
        text_preview = _get_text_content(element)[:50]
        print(f"Element {i}: Page {current_page}, Height: {elem_height:.1f}, Total: {cumulative_height:.1f} - {text_preview}...")
    
    return element_pages    

def _insert_line_break_after(body: Element, all_elements: List[Element], index: int):
    """Insert an empty paragraph (line break) after the element at the given index"""
    if index < 0 or index >= len(all_elements):
        return
    
    empty_para = _create_empty_paragraph()
    # Insert after the specified element
    insert_position = all_elements[index]
    body_list = list(body)
    insert_idx = body_list.index(insert_position) + 1
    body.insert(insert_idx, empty_para)
    all_elements.insert(index + 1, empty_para)
    
    return True


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
    groups = []
    current_list_count = 0
    while i < len(all_elements):  # Use while loop since list length may change
        element = all_elements[i]
        text_content = _get_text_content(element)
        sentence_count = _count_sentences(text_content)

        next_element = all_elements[i+1] if i+1 < len(all_elements) else None
                
        # Track list items as we encounter them
        if (_is_list_item(element)):
            current_list_count += 1

        # Apply formatting rules
        is_paragraph = (sentence_count >= PARAGRAPH_CRITERIA['min_sentences'] or
                        len(text_content.split()) >= PARAGRAPH_CRITERIA['min_words'])
        
        ## below is the logic to group elements that we considered need to be on the same page
        if is_paragraph:                                    
            groups.append([element])
        if _is_heading(element):
            next_element_text_content = _get_text_content(next_element)
            next_sentence_count = _count_sentences(next_element_text_content)
            is_next_text_content_is_paragraph = (next_sentence_count >= PARAGRAPH_CRITERIA['min_sentences'] or
                        len(next_element_text_content.split()) >= PARAGRAPH_CRITERIA['min_words'])
            if is_next_text_content_is_paragraph:
                groups.append([element, next_element])
        
        ### This is basic rule for adding empty paragraph after a paragraph
        ### Temporarily commented out since we are focusing on the groups
        # is_last_of_list = _is_last_list_item(element, next_element)

        # if (is_paragraph) and not _is_list_item(element):
        #     _ensure_single_empty_paragraph_after(body, all_elements, i, text_content)           
        # elif (is_last_of_list and not text_content.isupper()):
        #     # Only add spacing after lists with more than 3 items
        #     if current_list_count > 3: 
        #         _ensure_single_empty_paragraph_after(body, all_elements, i, text_content)
        
        # if _is_last_list_item(element, next_element): 
        #     current_list_count = 0

        i += 1

    # print_groups(groups)

    print(f"\nIdentified {len(groups)} groups that should stay together")

    # Step 2: Calculate page positions for all elements
    print("\nCalculating page positions...")
    element_pages = _calculate_page_positions(all_elements)
  
    # Step 3: Check if any groups are split across pages and fix them
    print("\nChecking for split groups...")
    total_line_breaks_added = 0
    
    for group_idx, group in enumerate(groups):
        # Get indices of elements in this group
        group_indices = []
        for elem in group:
            try:
                idx = all_elements.index(elem)
                group_indices.append(idx)
            except ValueError:
                continue
        
        if not group_indices:
            continue
        
        # Check if group spans multiple pages
        pages = [element_pages.get(idx, 1) for idx in group_indices]
        unique_pages = set(pages)
        
        if len(unique_pages) > 1:
            group_text = _get_text_content(group[0])[:50]
            print(f"\nGroup {group_idx} spans pages {unique_pages}: {group_text}...")
            
            # Find the element before this group
            first_idx = min(group_indices)
            if first_idx > 0:
                # Keep adding line breaks until the group stays together
                line_breaks_added = 0
                max_attempts = 10  # Prevent infinite loop
                
                while line_breaks_added < max_attempts:
                    # Add a line break before the group
                    insert_after_idx = first_idx - 1 + line_breaks_added
                    _insert_line_break_after(body, all_elements, insert_after_idx)
                    line_breaks_added += 1
                    total_line_breaks_added += 1
                    
                    # Recalculate page positions after adding line break
                    element_pages = _calculate_page_positions(all_elements)
                    
                    # Update group indices since we added elements
                    group_indices = []
                    for elem in group:
                        try:
                            idx = all_elements.index(elem)
                            group_indices.append(idx)
                        except ValueError:
                            continue
                    
                    # Check if group is now on the same page
                    pages = [element_pages.get(idx, 1) for idx in group_indices]
                    unique_pages = set(pages)
                    
                    if len(unique_pages) == 1:
                        print(f"  Fixed! Added {line_breaks_added} line breaks. Group now on page {pages[0]}")
                        break
                    else:
                        print(f"  Still split after {line_breaks_added} line breaks, continuing...")
                
                if line_breaks_added >= max_attempts:
                    print(f"  Warning: Could not fix group after {max_attempts} attempts")
    
    print(f"\nTotal line breaks added: {total_line_breaks_added}")  

    '''
    1. after all groups are collected, we will iterate over all the elements
    2. we will calculate the height of the content of each element and accumulate
    3. if it reach the end of the page, we check the grouped element. If the grouped element separated 
    into different pages, we will keep adding line break, until the grouped element get into the same page
    '''

    ## TODO: (1) check if the accumulated height is accurate
    ## (2) watch closely how it insert the line break/empty paragraph
    ## (3) finalize, add more types of groups
    ## (4) holiday to japan

    # Save the modified document
    tree.write(doc_xml_path, encoding='utf-8', xml_declaration=True)
    _create_docx(temp_dir, 'output8.docx')

    


    
