"""Intelligent column detector and heuristic AI scorer.
Automatically locates 'Design Number' and 'Image' columns across any sheet layout,
even with arbitrary column positions, multi-row headers, or missing headers.
"""

import re
from typing import List, Dict, Any, Optional, Tuple
import openpyxl.utils

# Positive keywords for Design Number
DESIGN_STRONG_KEYWORDS = [
    "design number", "design no", "designno", "design #", "designo",
    "item code", "item no", "item#", "style no", "style#", "style number",
    "sku", "product code", "product no", "article no", "article#", "model no", "model number"
]

DESIGN_MODERATE_KEYWORDS = [
    "design", "style", "item", "model", "ref", "reference", "code", "article"
]

# Negative keywords for Design Number (avoid picking weights, counts, serials, costs, etc.)
DESIGN_NEGATIVE_KEYWORDS = [
    "serial", "s.no", "sr no", "id", "gross", "net", "weight", "stone", "count",
    "cost", "price", "quantity", "qty", "total", "colour", "color", "metal",
    "category", "size", "diamond", "precious", "carat", "ct", "gm", "rate"
]

# Positive keywords for Image column
IMAGE_STRONG_KEYWORDS = [
    "image", "photo", "pic", "picture", "img", "photograph", "design image", "product image"
]


def score_design_column(
    header: str,
    cell_values: List[Any]
) -> Tuple[float, str]:
    """Scores a single column for likelihood of being the Design Number column.
    
    Returns:
        (score, reason_note)
    """
    score = 0.0
    reasons = []
    h_lower = str(header or "").lower().strip()
    
    # 1. Header keyword check
    if any(k in h_lower for k in DESIGN_STRONG_KEYWORDS):
        score += 120.0
        reasons.append(f"Header '{header}' strongly matches design number")
    elif any(k in h_lower for k in DESIGN_MODERATE_KEYWORDS):
        score += 50.0
        reasons.append(f"Header '{header}' contains design/item keyword")
        
    # Check negative keywords in header
    for neg in DESIGN_NEGATIVE_KEYWORDS:
        if neg in h_lower:
            score -= 75.0
            reasons.append(f"Header contains exclusion term '{neg}'")
            break

    # 2. Cell values inspection
    valid_vals = [str(v).strip() for v in cell_values if v is not None and str(v).strip()]
    if not valid_vals:
        return max(score, 0.0), "; ".join(reasons)
        
    # Uniqueness check (design numbers are typically distinct per row)
    unique_ratio = len(set(valid_vals)) / len(valid_vals)
    if unique_ratio > 0.8:
        score += 35.0
        reasons.append(f"High uniqueness ({int(unique_ratio*100)}%)")
    elif unique_ratio < 0.3:
        score -= 40.0
        reasons.append(f"Low uniqueness ({int(unique_ratio*100)}% - likely a category or attribute)")

    # Alphanumeric pattern check
    # Typical jewelry design codes: 18259W1BS8WV, 70002QBS8WP, BP49008R-20WD8RP, 26024ULB8WX
    alphanumeric_count = 0
    pure_numeric_count = 0
    short_words_count = 0
    
    for val in valid_vals:
        # Check if pure float or int
        try:
            float(val.replace(",", ""))
            pure_numeric_count += 1
            continue
        except ValueError:
            pass
            
        # Check length
        if len(val) < 3:
            short_words_count += 1
            continue
            
        # Match alphanumeric code pattern
        has_letters = bool(re.search(r'[A-Za-z]', val))
        has_digits = bool(re.search(r'[0-9]', val))
        if has_letters and has_digits and len(val) >= 4:
            alphanumeric_count += 1
        elif has_letters and len(val) >= 5 and re.match(r'^[A-Za-z0-9\-_]+$', val):
            alphanumeric_count += 0.8
            
    num_ratio = pure_numeric_count / len(valid_vals)
    if num_ratio > 0.6:
        score -= 60.0
        reasons.append(f"Mostly numeric values ({int(num_ratio*100)}%)")
        
    code_ratio = alphanumeric_count / len(valid_vals)
    if code_ratio > 0.4:
        score += code_ratio * 70.0
        reasons.append(f"{int(code_ratio*100)}% sample values match alphanumeric style codes")

    return max(score, 0.0), "; ".join(reasons)


class ColumnDetectionResult:
    def __init__(
        self,
        sheet_name: str,
        design_col_idx: int,
        design_col_name: str,
        image_col_idx: int,
        image_col_name: str,
        header_row_idx: int,
        confidence: float,
        reasoning: str,
        all_columns: List[Dict[str, Any]]
    ):
        self.sheet_name = sheet_name
        self.design_col_idx = design_col_idx  # 0-indexed
        self.design_col_name = design_col_name
        self.design_col_letter = openpyxl.utils.get_column_letter(design_col_idx + 1)
        self.image_col_idx = image_col_idx  # 0-indexed
        self.image_col_name = image_col_name
        self.image_col_letter = openpyxl.utils.get_column_letter(image_col_idx + 1)
        self.header_row_idx = header_row_idx  # 1-indexed
        self.confidence = confidence
        self.reasoning = reasoning
        self.all_columns = all_columns

    def to_dict(self) -> Dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "design_col_idx": self.design_col_idx,
            "design_col_letter": self.design_col_letter,
            "design_col_name": self.design_col_name,
            "image_col_idx": self.image_col_idx,
            "image_col_letter": self.image_col_letter,
            "image_col_name": self.image_col_name,
            "header_row_idx": self.header_row_idx,
            "confidence": self.confidence,
            "reasoning": self.reasoning,
            "all_columns": self.all_columns
        }


def detect_sheet_columns(
    sheet_name: str,
    rows_data: List[List[Any]],
    image_anchor_cols: List[int]
) -> ColumnDetectionResult:
    """Intelligently detects Design Number and Image columns for a sheet.
    
    Args:
        sheet_name: Name of the sheet.
        rows_data: List of rows, where each row is a list of cell values.
        image_anchor_cols: List of 0-indexed column indices where images are anchored.
        
    Returns:
        ColumnDetectionResult with best identified columns and fallback options.
    """
    if not rows_data:
        # Empty sheet fallback
        return ColumnDetectionResult(
            sheet_name=sheet_name,
            design_col_idx=0,
            design_col_name="Column A",
            image_col_idx=1,
            image_col_name="Column B",
            header_row_idx=1,
            confidence=0.0,
            reasoning="Sheet contains no data rows.",
            all_columns=[{"index": 0, "letter": "A", "name": "Column A"}]
        )

    # 1. Determine the maximum number of columns
    max_cols = max(len(r) for r in rows_data)
    
    # 2. Find best candidate header row (scan rows 1 to min(15, len(rows_data)))
    header_row_idx = 1
    best_header_score = -1.0
    scan_limit = min(15, len(rows_data))
    
    for r_idx in range(scan_limit):
        row = rows_data[r_idx]
        text_cells = [str(c).strip().lower() for c in row if c is not None and str(c).strip()]
        if not text_cells:
            continue
        
        # Header score: diversity of text + match with design/image keywords
        h_score = len(text_cells) * 2.0
        for text in text_cells:
            if any(k in text for k in DESIGN_STRONG_KEYWORDS):
                h_score += 50.0
            elif any(k in text for k in IMAGE_STRONG_KEYWORDS):
                h_score += 40.0
                
        if h_score > best_header_score:
            best_header_score = h_score
            header_row_idx = r_idx + 1  # 1-indexed

    header_cells = rows_data[header_row_idx - 1] if header_row_idx <= len(rows_data) else []
    
    # Build all_columns list
    all_columns = []
    for c_idx in range(max_cols):
        h_name = ""
        if c_idx < len(header_cells) and header_cells[c_idx] is not None:
            h_name = str(header_cells[c_idx]).strip()
        col_letter = openpyxl.utils.get_column_letter(c_idx + 1)
        display_label = f"Col {col_letter} - {h_name}" if h_name else f"Col {col_letter}"
        all_columns.append({
            "index": c_idx,
            "letter": col_letter,
            "name": h_name or f"Column {col_letter}",
            "display": display_label
        })

    # 3. Detect Image Column
    # Check physical image anchor frequency first
    image_col_idx = -1
    image_col_confidence = 0.5
    image_reason = ""
    
    if image_anchor_cols:
        col_counts: Dict[int, int] = {}
        for col in image_anchor_cols:
            col_counts[col] = col_counts.get(col, 0) + 1
        # Get most frequent anchor column
        best_anchor_col, count = max(col_counts.items(), key=lambda x: x[1])
        image_col_idx = best_anchor_col
        image_col_confidence = min(0.99, count / len(image_anchor_cols))
        col_letter = openpyxl.utils.get_column_letter(image_col_idx + 1)
        image_reason = f"Detected {count}/{len(image_anchor_cols)} images physically anchored in Column {col_letter}"
    else:
        # Fallback to header matching for image
        for c_idx, col_info in enumerate(all_columns):
            name_lower = col_info["name"].lower()
            if any(k in name_lower for k in IMAGE_STRONG_KEYWORDS):
                image_col_idx = c_idx
                image_col_confidence = 0.8
                image_reason = f"Header matches image keyword '{col_info['name']}'"
                break
                
    if image_col_idx < 0:
        # Default to second column or column after header
        image_col_idx = 1 if max_cols > 1 else 0
        image_reason = "No image anchor or header found; defaulted to Column B"

    # 4. Detect Design Number Column
    # Collect sample data rows after header
    sample_rows = rows_data[header_row_idx:header_row_idx + 60]
    if not sample_rows and header_row_idx > 1:
        sample_rows = rows_data[header_row_idx - 1:]

    design_scores = []
    for c_idx in range(max_cols):
        # Don't pick the image column as design column
        if c_idx == image_col_idx and image_col_confidence > 0.8:
            continue
            
        header_text = all_columns[c_idx]["name"] if c_idx < len(all_columns) else ""
        col_values = [r[c_idx] for r in sample_rows if c_idx < len(r)]
        score, reason = score_design_column(header_text, col_values)
        design_scores.append((c_idx, score, reason))

    if design_scores:
        design_scores.sort(key=lambda x: x[1], reverse=True)
        best_col_idx, best_score, best_reason = design_scores[0]
        confidence = min(0.99, max(0.5, best_score / 180.0))
        reasoning = f"Design column: {best_reason}. Image column: {image_reason}."
    else:
        best_col_idx = 0
        confidence = 0.4
        reasoning = "Defaulted to Column A (no confident design number candidate found)."

    design_name = all_columns[best_col_idx]["name"] if best_col_idx < len(all_columns) else f"Column {best_col_idx+1}"
    image_name = all_columns[image_col_idx]["name"] if image_col_idx < len(all_columns) else f"Column {image_col_idx+1}"

    return ColumnDetectionResult(
        sheet_name=sheet_name,
        design_col_idx=best_col_idx,
        design_col_name=design_name,
        image_col_idx=image_col_idx,
        image_col_name=image_name,
        header_row_idx=header_row_idx,
        confidence=round(confidence, 2),
        reasoning=reasoning,
        all_columns=all_columns
    )
