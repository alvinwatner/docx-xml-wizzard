# Page Content Estimation Guide

## 1. Estimating Page Content Boundary for a New Document

### What You Need to Find
The page content boundary is the **usable height** on each page where content can be placed, calculated as:
```
Content Height = Page Height - Top Margin - Bottom Margin
```

### Where to Look
Check **`document_content_analyzer.py`** - specifically the `_extract_page_setup()` method which:

1. **Extracts from DOCX XML** (`word/document.xml`):
   - Looks for `w:sectPr` (section properties) 
   - Reads `w:pgSz` for page dimensions
   - Reads `w:pgMar` for margins

2. **Key XML attributes**:
   ```xml
   <w:pgSz w:w="11906" w:h="16838"/>  <!-- Page size in twips -->
   <w:pgMar w:top="1559" w:bottom="737" w:left="1531" w:right="1418"/>  <!-- Margins in twips -->
   ```

3. **Conversion formula**:
   - **Twips to Points**: `value / 20` (since 1 point = 20 twips)
   - **EMUs to Points**: `value / 12700` (for images)
   - **Points to Inches**: `value / 72` (since 72 points = 1 inch)

### Example Output
For standard A4 document:
```json
{
  "page_width": 595.35,      // 8.27 inches
  "page_height": 842.0,       // 11.69 inches  
  "margins": {
    "top": 77.95,            // points
    "bottom": 36.85          // points
  },
  "content_height": 727.2     // 842.0 - 77.95 - 36.85 = 727.2 points
}
```

**Result**: You have **727.2 points** (~10.1 inches) of vertical space per page for content.

---

## 2. Estimating Element Heights and Accumulation

### Height Estimation Logic

The core formula is in `_estimate_element_height()` (from both `document_content_analyzer.py` and `page_break_formatter.py`):

#### For Text Elements
```python
# Base measurements
font_size = 11.0            # Default, or extract from styles
line_height = font_size * 1.5   # Standard 1.5x line spacing
content_width = 447.9        # Usable width after margins

# Character and line calculation
char_width = font_size * 0.6     # Empirical ratio for average character width
chars_per_line = content_width / char_width
num_lines = text_length / chars_per_line

# Final height
element_height = num_lines * line_height + padding
```

**Why 0.6?** Based on typography research, average character width in proportional fonts is approximately 0.6x the font size.

#### For Images
```python
# Extract from drawing properties in XML
# Convert EMUs to points
image_height = emu_value / 12700
```

#### For Tables
```python
# Sum all row heights
table_height = sum(row_heights) + (row_count * padding)
```

#### For Empty Paragraphs
```python
# Single line height
empty_height = line_height  # Usually ~16.5 points for 11pt font
```

### Accumulation Process

The accumulation happens in `_calculate_page_positions()`:

```python
def _calculate_page_positions(all_elements):
    cumulative_height = 0.0
    current_page = 1
    PAGE_CONTENT_HEIGHT = 727.2  # From step 1
    
    for element in all_elements:
        elem_height = _estimate_element_height(element)
        
        # Check if adding this element exceeds page boundary
        if cumulative_height + elem_height > PAGE_CONTENT_HEIGHT:
            # Would overflow - start new page
            current_page += 1
            cumulative_height = elem_height  # Reset to just this element
        else:
            # Fits on current page
            cumulative_height += elem_height
        
        # Record which page this element is on
        element_pages[element_index] = current_page
```

### Key Insights

1. **Page Boundary**: Not the full 727.2 points - use ~75% (550 points) as practical threshold to prevent:
   - Orphan headings (heading alone at page bottom)
   - Split paragraphs (text breaking mid-paragraph)
   - Widows (single line at page top)

2. **Height Estimation Accuracy**: Our formula gives reasonable approximations but actual rendering depends on:
   - Specific font metrics
   - Kerning and letter spacing
   - Word's line breaking algorithm
   - Paragraph spacing settings

3. **Adaptive Adjustments**: Different element types need different thresholds:
   - **Headings**: Need more space after (to keep with content)
   - **Tables**: Need more buffer (to avoid splitting)
   - **Lists**: Can be more flexible

### Practical Usage

```python
# 1. Extract page setup once
page_setup = extract_page_setup('document.docx')
CONTENT_HEIGHT = page_setup['content_height']  # e.g., 727.2

# 2. Process each element
cumulative = 0
for element in document_elements:
    height = estimate_height(element)
    
    if cumulative + height > CONTENT_HEIGHT * 0.75:  # 75% threshold
        # Start new page
        insert_page_break()
        cumulative = height
    else:
        cumulative += height
```

---

## Quick Reference

### Constants for A4
- **Page Height**: 842.0 points (11.69 inches)
- **Content Height**: ~727.2 points (varies by margins)
- **Content Width**: ~447.9 points (varies by margins)
- **Practical Threshold**: ~550 points (75% utilization)

### Conversion Factors
- **1 inch** = 72 points = 1440 twips = 914,400 EMUs
- **1 point** = 20 twips = 12,700 EMUs
- **Character width** ≈ font_size × 0.6
- **Line height** ≈ font_size × 1.5

### Files to Reference
- **`document_content_analyzer.py`**: Full implementation of height estimation and page setup extraction
- **`page_break_predictor.py`**: Adaptive threshold algorithm for smart page breaks
- **`page_break_formatter.py`**: Simple implementation with iterative line break approach