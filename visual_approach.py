import tempfile
import zipfile
import os
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from typing import List, Dict, Tuple

# ============================================================================
# CONSTANTS
# ============================================================================

# Input document path
docx_path = "template_variables.docx"

# Group patterns that must stay on the same page (G in the algorithm)
GROUP_PATTERNS = {
    'heading+paragraph': ['heading', 'paragraph'],
    'paragraph': ['paragraph'],
    'heading+list': ['heading', 'list']
}

# Paragraph detection criteria
PARAGRAPH_CRITERIA = {
    'min_sentences': 2,
    'min_words': 15
}

# Page dimensions (A4)
PAGE_CONTENT_HEIGHT = 727.2  # Points
CONTENT_WIDTH = 447.9  # Points

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

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

def _get_text_content(element, namespaces):
    """Get text content from an element"""
    text_elements = element.findall('.//w:t', namespaces)
    text_content = ''.join([t.text or '' for t in text_elements])
    return text_content

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
    sentence_pattern = r'[.!?]+(?:\s+[A-Z]|$)'
    sentences = re.findall(sentence_pattern, text_processed)

    # If no sentence endings found but text exists, count as 1 sentence
    sentence_count = len(sentences)
    if sentence_count == 0 and len(text.strip()) > 0:
        sentence_count = 1

    return sentence_count

def _is_list_item(element, namespaces):
    """Check if element is part of a list"""
    numPr = element.find('.//w:numPr', namespaces)
    return numPr is not None

def _get_list_level(element, namespaces):
    """Get the list level of an element (0-based)"""
    numPr = element.find('.//w:numPr', namespaces)
    if numPr is None:
        return None

    ilvl = numPr.find('.//w:ilvl', namespaces)
    if ilvl is not None:
        return int(ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0'))
    return 0

def _is_heading(element, namespaces):
    """
    Determine if an element is a heading.
    Check if it's a level-0 list item with uppercase text
    """
    text = _get_text_content(element, namespaces).strip()

    if not text:
        return False

    # Check if it's a numbered list item at level 0
    if _is_list_item(element, namespaces):
        level = _get_list_level(element, namespaces)
        if level == 0:
            # Remove numbering and check if remaining text has uppercase words
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

def _is_paragraph(element, namespaces):
    """Check if element qualifies as a paragraph based on criteria"""
    text = _get_text_content(element, namespaces)
    sentence_count = _count_sentences(text)

    return (sentence_count >= PARAGRAPH_CRITERIA['min_sentences'] or
            len(text.split()) >= PARAGRAPH_CRITERIA['min_words'])

def _get_element_type(element, namespaces):
    """
    Determine the type of an element for pattern matching.
    Returns: 'heading', 'paragraph', 'list', or None
    """
    if _is_heading(element, namespaces):
        return 'heading'
    elif _is_paragraph(element, namespaces):
        return 'paragraph'
    elif _is_list_item(element, namespaces):
        return 'list'
    return None

def _match_pattern(window_types: List[str]) -> Tuple[bool, str]:
    """
    Check if a window of element types matches any pattern in G.

    Uses strict matching - the window must exactly match the pattern.
    None values are NOT filtered out.

    Args:
        window_types: List of element types in the window

    Returns:
        Tuple of (matches, pattern_name)
    """
    # Strict matching - no filtering of None values
    # The window must exactly match the pattern

    # Check against each pattern in G
    for pattern_name, pattern_types in GROUP_PATTERNS.items():
        if window_types == pattern_types:
            return True, pattern_name

    return False, None

# ============================================================================
# PHASE 1: GROUP EXTRACTION
# ============================================================================

def extract_groups(all_elements: List[Element], namespaces: dict) -> List[Dict]:
    """
    Extract all groups from document using sliding window approach.

    This implements lines 7-17 of Algorithm 1:
    - For each position i in document
    - Try all window sizes from L-1 down to 0
    - Check if window matches any pattern in G
    - If matched, record the group

    Args:
        all_elements: List of all XML elements from document body
        namespaces: XML namespaces dict

    Returns:
        List of groups P, where each group is:
        {
            'type': pattern name (e.g., 'heading+paragraph'),
            'group_index': position in P list,
            'doc_indices': [i, i+1, ..., i+j] indices in document,
            'elements': [elem1, elem2, ...] actual XML elements
        }
    """
    P = []  # Extracted groups
    n = len(all_elements)

    # L = max group size (longest pattern)
    L = max(len(pattern) for pattern in GROUP_PATTERNS.values())

    print("\n" + "="*70)
    print("PHASE 1: GROUP EXTRACTION")
    print("="*70)
    print(f"Total elements: {n}")
    print(f"Maximum group size (L): {L}")
    print(f"Group patterns: {list(GROUP_PATTERNS.keys())}")
    print()

    # For each position i in document (line 7)
    for i in range(n):
        # Try all window sizes from L-1 down to 0 (line 8)
        for j in range(L - 1, -1, -1):
            # Check if window fits in document
            if i + j >= n:
                continue

            # Build temp_group (lines 9-12)
            temp_group = []
            temp_indices = []
            temp_types = []

            for k in range(j + 1):  # j+1 because range is exclusive
                elem = all_elements[i + k]
                elem_type = _get_element_type(elem, namespaces)

                temp_group.append(elem)
                temp_indices.append(i + k)
                temp_types.append(elem_type)

            # Check if pattern matches any in G (line 13)
            matches, pattern_name = _match_pattern(temp_types)

            if matches:
                # Record the group (line 14-15)
                group = {
                    'type': pattern_name,
                    'group_index': len(P),
                    'doc_indices': temp_indices,
                    'elements': temp_group
                }
                P.append(group)

                # Debug output
                print(f"Group {len(P)-1}: {pattern_name}")
                print(f"  Doc indices: {temp_indices}")
                print(f"  ")
                # print(f"  First element: \"{first_text}...\"")
                print()

    print(f"\nTotal groups extracted: {len(P)}")
    print("="*70)

    return P

# ============================================================================
# MAIN EXECUTION
# ============================================================================

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

    # Register namespaces for output
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = tree.getroot()
    body = root.find('.//w:body', namespaces)
    all_elements = list(body)

    print(f"\n{'='*70}")
    print(f"DOCUMENT: {docx_path}")
    print(f"{'='*70}")
    print(f"Total XML elements in document: {len(all_elements)}")

    # PHASE 1: Extract all groups
    P = extract_groups(all_elements, namespaces)

    # Print detailed summary
    print("\n" + "="*70)
    print("GROUP EXTRACTION SUMMARY")
    print("="*70)

    pattern_counts = {}
    for group in P:
        pattern = group['type']
        pattern_counts[pattern] = pattern_counts.get(pattern, 0) + 1

    for pattern, count in pattern_counts.items():
        print(f"{pattern}: {count} groups")

    print(f"\nTotal groups: {len(P)}")
    print("="*70)

    # Print first 10 groups for inspection
    print("\nFirst 10 groups (detailed):")
    print("-"*70)
    for i, group in enumerate(P[:10]):
        print(f"\nGroup {i}:")
        print(f"  Type: {group['type']}")
        print(f"  Doc indices: {group['doc_indices']}")
        for idx, elem in enumerate(group['elements']):
            text = _get_text_content(elem, namespaces)[:60]
            print(f"    Element {idx}: \"{text}...\"")
