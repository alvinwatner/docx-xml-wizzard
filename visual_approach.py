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
docx_path = "template_variables_messy_multi_paragraph.docx"

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
# PDF GENERATION AND EXTRACTION
# ============================================================================

def generate_pdf(docx_path: str, output_pdf: str) -> str:
    """
    Generate PDF from DOCX using Google Docs API.

    Args:
        docx_path: Path to input DOCX file
        output_pdf: Path for output PDF file

    Returns:
        Path to generated PDF
    """
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload

    SCOPES = ["https://www.googleapis.com/auth/drive.file"]

    print("\n" + "="*70)
    print("GENERATING PDF VIA GOOGLE DOCS API")
    print("="*70)

    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    service = build("drive", "v3", credentials=creds)

    # Upload DOCX and convert to Google Docs
    file_metadata = {
        "name": os.path.basename(docx_path),
        "mimeType": "application/vnd.google-apps.document"
    }
    media = MediaFileUpload(
        docx_path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True
    )

    print("Uploading DOCX to Google Drive...")
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()
    file_id = uploaded.get("id")
    print(f"Uploaded (File ID: {file_id})")

    # Export as PDF
    print("Exporting to PDF...")
    pdf_request = service.files().export_media(
        fileId=file_id,
        mimeType="application/pdf"
    )
    pdf_data = pdf_request.execute()

    with open(output_pdf, "wb") as pdf_file:
        pdf_file.write(pdf_data)
    print(f"PDF saved: {output_pdf}")

    # Delete temporary Google Doc
    service.files().delete(fileId=file_id).execute()
    print("Temporary Google Doc deleted")
    print("="*70)

    return output_pdf

def extract_pdf_data(pdf_path: str) -> Dict:
    """
    Extract text blocks with coordinates from PDF using PyMuPDF.

    Args:
        pdf_path: Path to PDF file

    Returns:
        Dict with structure:
        {
            'pages': [
                {
                    'page_num': 1,
                    'width': 596.0,
                    'height': 842.0,
                    'blocks': [
                        {'text': '...', 'bbox': {'x0': ..., 'y0': ..., 'x1': ..., 'y1': ...}},
                        ...
                    ],
                    'all_text': '...'  # Combined text for this page
                },
                ...
            ]
        }
    """
    import fitz  # PyMuPDF

    print("\n" + "="*70)
    print("EXTRACTING PDF DATA")
    print("="*70)
    print(f"PDF: {pdf_path}")

    doc = fitz.open(pdf_path)
    pages_data = []

    for page_num, page in enumerate(doc, start=1):
        print(f"\nProcessing page {page_num}...")

        page_data = {
            'page_num': page_num,
            'width': page.rect.width,
            'height': page.rect.height,
            'blocks': [],
            'all_text': ''
        }

        # Extract text with detailed positioning
        text_dict = page.get_text("dict")
        all_texts = []

        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:  # Text block
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        bbox = span.get("bbox", [0, 0, 0, 0])

                        if text.strip():
                            page_data["blocks"].append({
                                "text": text,
                                "bbox": {
                                    "x0": bbox[0],
                                    "y0": bbox[1],
                                    "x1": bbox[2],
                                    "y1": bbox[3]
                                }
                            })
                            all_texts.append(text)

        page_data["all_text"] = " ".join(all_texts)
        pages_data.append(page_data)

        print(f"  Extracted {len(page_data['blocks'])} text blocks")
        print(f"  Text length: {len(page_data['all_text'])} chars")

    doc.close()

    print(f"\nTotal pages extracted: {len(pages_data)}")
    print("="*70)

    return {'pages': pages_data}

def get_overlapped_words(A, B):
    """
    Find the longest sequence of overlapping words between two sentences.
    
    Args:
        A (str): First sentence
        B (str): Second sentence
    
    Returns:
        list: List of overlapped words (longest sequence found)
    """
    # Normalize sentences: lowercase and split into words
    words_A = A.lower().split()
    words_B = B.lower().split()
    
    # Track the longest overlap found
    max_overlap_length = 0
    max_overlap_words = []
    
    # Iterate through all possible starting positions in A
    for i in range(len(words_A)):
        # Iterate through all possible starting positions in B
        for j in range(len(words_B)):
            # Count consecutive matching words
            overlap_count = 0
            while (i + overlap_count < len(words_A) and 
                   j + overlap_count < len(words_B) and 
                   words_A[i + overlap_count] == words_B[j + overlap_count]):
                overlap_count += 1
            
            # Update max if we found a longer overlap
            if overlap_count > max_overlap_length:
                max_overlap_length = overlap_count
                max_overlap_words = words_A[i:i + overlap_count]
    
    return max_overlap_words

# ============================================================================
# PHASE 2: SPLIT DETECTION
# ============================================================================

def _find_exact_match(elem_text_normalized: str, pages: List[Dict]) -> int:
    """
    Find first page containing exact element text.
    
    Returns:
        Page number if found, None otherwise
    """
    import re
    for page in pages:
        page_text_normalized = re.sub(r'\s+', ' ', page['all_text'].strip().lower())
        if elem_text_normalized in page_text_normalized:
            return page['page_num']
    return None


def _detect_partial_split(elem_text_normalized: str, pages: List[Dict], min_word_overlap: int = 6) -> List[int]:
    """
    Detect if element is split across multiple pages using word overlap.
    
    Strategy:
    - For each page, find longest overlapping word sequence
    - If overlap > threshold, this might be first chunk of split element
    - Continue to next page and concatenate chunks
    - If concatenated text matches full element, we found a partial split
    
    Args:
        elem_text_normalized: Normalized element text to search for
        pages: List of PDF pages with text data
        min_word_overlap: Minimum overlapping words to consider a match
        
    Returns:
        List of page numbers if partial split detected, None otherwise
    """
    import re
    
    first_chunk = None
    first_page_num = None
    
    for page in pages:
        page_text_normalized = re.sub(r'\s+', ' ', page['all_text'].strip().lower())
        overlapped_words = get_overlapped_words(elem_text_normalized, page_text_normalized)
        
        if len(overlapped_words) > min_word_overlap:
            if first_chunk is None:
                # Found potential first chunk of split element
                first_chunk = ' '.join(overlapped_words)
                first_page_num = page['page_num']
            else:
                # Found potential continuation on next page
                remaining_chunk = ' '.join(overlapped_words)
                combined_text = first_chunk + ' ' + remaining_chunk
                normalized_combined = re.sub(r'\s+', ' ', combined_text.strip())
                
                # Check if concatenated chunks match full element
                if elem_text_normalized in normalized_combined:
                    return [first_page_num, page['page_num']]
    
    return None


def detect_split_groups(V: Dict, P: List[Dict], all_elements: List[Element], 
                       namespaces: dict, debug: bool = False) -> List[Dict]:
    """
    Detect all groups that are split across pages.

    Strategy:
    - For each group, find which page(s) each element appears on
    - First try exact text matching (sequential, stops at first match)
    - If not found, try partial split detection (element spanning pages)
    - If elements appear on different pages → SPLIT!

    Args:
        V: PDF data with pages from extract_pdf_data()
        P: List of all groups from Phase 1
        all_elements: All DOCX elements
        namespaces: XML namespaces
        debug: Enable debug logging

    Returns:
        W: List of split groups with metadata
    """
    import re

    if debug:
        print("\n" + "="*70)
        print("DETECTING SPLIT GROUPS")
        print("="*70)

    W = []  # Split groups

    for group in P:
        elem_indices = group['doc_indices']
        pages_for_group = []

        if debug:
            print(f"CHECKING group {group['group_index']}: indices {elem_indices}")

        for elem_idx in elem_indices:
            elem = all_elements[elem_idx]
            elem_text = _get_text_content(elem, namespaces)

            # Skip empty elements
            if not elem_text.strip():
                continue

            # Normalize whitespace - replace all sequences with single space
            elem_text_normalized = re.sub(r'\s+', ' ', elem_text.strip().lower())

            # STRATEGY 1: Try exact match first (stops at first occurrence)
            found_page = _find_exact_match(elem_text_normalized, V['pages'])
            
            if found_page:
                pages_for_group.append(found_page)
                if debug:
                    print(f"  Element [{elem_idx}] found on page {found_page} (exact match)")
            else:
                # STRATEGY 2: Try partial split detection (element spanning pages)
                split_pages = _detect_partial_split(elem_text_normalized, V['pages'])
                
                if split_pages:
                    pages_for_group.extend(split_pages)
                    if debug:
                        print(f"  Element [{elem_idx}] spans pages {split_pages} (partial split)")

        # Check if group elements appear on multiple pages
        unique_pages = set(pages_for_group)
        
        if len(unique_pages) > 1:
            W.append({
                **group,
                'split_reason': f'Elements on pages {sorted(unique_pages)}'
            })
            if debug:
                print(f"  → SPLIT DETECTED on pages {sorted(unique_pages)}")

        if debug:
            print()

    if debug:
        print(f"Split groups detected: {len(W)} / {len(P)}")
        print("="*70)

    return W

# ============================================================================
# DOCUMENT CLEANUP
# ============================================================================

def remove_empty_elements(body: Element, all_elements: List[Element], namespaces: dict) -> List[Element]:
    """
    Remove all empty paragraph elements from the document.

    Empty elements (line breaks with no text) can block group detection
    because they appear between heading and paragraph elements.

    Args:
        body: The document body element
        all_elements: List of all elements in the body
        namespaces: XML namespaces

    Returns:
        Cleaned list of elements with empty ones removed
    """
    print("\n" + "="*70)
    print("CLEANING DOCUMENT - REMOVING EMPTY ELEMENTS")
    print("="*70)

    elements_to_remove = []

    for elem in all_elements:
        text = _get_text_content(elem, namespaces)
        if text.strip() == "":  # Empty element
            elements_to_remove.append(elem)

    # Remove from body XML
    for elem in elements_to_remove:
        body.remove(elem)

    # Return cleaned list
    cleaned_elements = [e for e in all_elements if e not in elements_to_remove]

    print(f"Original elements: {len(all_elements)}")
    print(f"Empty elements removed: {len(elements_to_remove)}")
    print(f"Cleaned elements: {len(cleaned_elements)}")
    print("="*70)

    return cleaned_elements

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

    # CLEANUP: Remove empty elements before group extraction
    all_elements = remove_empty_elements(body, all_elements, namespaces)

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

    # ========================================================================
    # PHASE 2: SPLIT DETECTION (Testing f(V, P[i]))
    # ========================================================================

    print("\n" + "="*70)
    print("PHASE 2: SPLIT DETECTION TEST")
    print("="*70)

    # # Generate PDF from DOCX
    pdf_path = "template_variables_visual.pdf"
    generate_pdf(docx_path, pdf_path)

    # Extract PDF data
    V = extract_pdf_data(pdf_path)

    # # Detect split groups
    W = detect_split_groups(V, P, all_elements, namespaces)

    # ========================================================================
    # PRINT W FOR MANUAL VERIFICATION
    # ========================================================================

    print("\n" + "="*70)
    print("SPLIT GROUPS DETECTED (W) - FOR MANUAL VERIFICATION")
    print("="*70)
    print(f"Total split groups: {len(W)}")
    print()

    if len(W) == 0:
        print("No split groups detected! All groups are on the same page.")
    else:
        for i, group in enumerate(W):
            print(f"\n{'='*70}")
            print(f"SPLIT GROUP {i}")
            print(f"{'='*70}")
            print(f"Type: {group['type']}")
            print(f"Group index in P: {group['group_index']}")
            print(f"Doc indices: {group['doc_indices']}")
            print(f"Split reason: {group['split_reason']}")
            print(f"\nElements:")

            for j, elem_idx in enumerate(group['doc_indices']):
                elem = group['elements'][j]
                text = _get_text_content(elem, namespaces)

                # Truncate long text
                if len(text) > 100:
                    text_display = text[:100] + "..."
                else:
                    text_display = text

                print(f"  [{elem_idx}] \"{text_display}\"")

    print("\n" + "="*70)
    print("VERIFICATION INSTRUCTIONS")
    print("="*70)
    print(f"1. Open the PDF: {pdf_path}")
    print("2. For each split group above, check:")
    print("   - Are the elements actually on different pages?")
    print("   - Is any single element cut across pages (partial split)?")
    print("3. Confirm if W is accurate")
    print("="*70)
