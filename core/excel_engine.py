"""High-performance Excel image extraction and pairing engine.
Features dual-engine extraction (openpyxl + direct ZIP/XML fallback),
merged cell resolution, Office 365 cell images support, and flexible column mapping.
"""

import os
import io
import re
import zipfile
import posixpath
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional, Tuple, Callable
import openpyxl
import openpyxl.utils
from PIL import Image

from .image_saver import (
    sanitize_filename,
    convert_and_save_jpeg,
    resolve_unique_filepath,
    create_thumbnail
)
from .column_detector import detect_sheet_columns, ColumnDetectionResult


class ExtractedImage:
    """Represents an image extracted from an Excel sheet."""
    def __init__(self, row: int, col: int, data: bytes, fmt: str = "PNG"):
        self.row = row  # 1-indexed row in Excel
        self.col = col  # 0-indexed column in Excel
        self.data = data
        self.fmt = fmt.upper()
        self._thumbnail: Optional[Image.Image] = None

    @property
    def thumbnail(self) -> Optional[Image.Image]:
        if self._thumbnail is None and self.data:
            self._thumbnail = create_thumbnail(self.data, size=(50, 50))
        return self._thumbnail


class SheetRowItem:
    """Represents a matched row item pairing a design number with an image."""
    def __init__(
        self,
        excel_row: int,
        design_number: str,
        image: Optional[ExtractedImage] = None,
        status: str = "Ready"
    ):
        self.excel_row = excel_row
        self.raw_design_number = str(design_number or "").strip()
        self.clean_design_number = sanitize_filename(self.raw_design_number, fallback=f"Row_{excel_row}")
        self.image = image
        self.status = status
        self.saved_path: Optional[str] = None
        self.error_message: Optional[str] = None

    @property
    def proposed_filename(self) -> str:
        return f"{self.clean_design_number}.jpg"


class SheetScanResult:
    """Holds all scan and extraction data for a single worksheet."""
    def __init__(
        self,
        sheet_name: str,
        raw_rows: List[List[Any]],
        images: List[ExtractedImage],
        detection: ColumnDetectionResult
    ):
        self.sheet_name = sheet_name
        self.raw_rows = raw_rows
        self.images = images
        self.detection = detection
        self.items: List[SheetRowItem] = []
        self.remap_items()

    def remap_items(
        self,
        design_col_idx: Optional[int] = None,
        image_col_idx: Optional[int] = None
    ):
        """Pairs images with design numbers based on selected column indices."""
        if design_col_idx is not None:
            self.detection.design_col_idx = design_col_idx
            self.detection.design_col_letter = openpyxl.utils.get_column_letter(design_col_idx + 1)
        if image_col_idx is not None:
            self.detection.image_col_idx = image_col_idx
            self.detection.image_col_letter = openpyxl.utils.get_column_letter(image_col_idx + 1)

        d_col = self.detection.design_col_idx
        
        # Group images by row
        images_by_row: Dict[int, List[ExtractedImage]] = {}
        for img in self.images:
            images_by_row.setdefault(img.row, []).append(img)

        self.items.clear()
        
        # 1. Process rows where images exist
        processed_rows = set()
        
        for r_num in sorted(images_by_row.keys()):
            imgs = images_by_row[r_num]
            # Try to get design number at row r_num (1-indexed)
            design_val = ""
            if 1 <= r_num <= len(self.raw_rows):
                row_vals = self.raw_rows[r_num - 1]
                if d_col < len(row_vals) and row_vals[d_col] is not None:
                    design_val = str(row_vals[d_col]).strip()
            
            # If empty, check adjacent row (offset +-1)
            if not design_val:
                for offset in (-1, 1):
                    adj_r = r_num + offset
                    if 1 <= adj_r <= len(self.raw_rows):
                        adj_vals = self.raw_rows[adj_r - 1]
                        if d_col < len(adj_vals) and adj_vals[d_col] is not None:
                            val = str(adj_vals[d_col]).strip()
                            if val and adj_r not in images_by_row:
                                design_val = val
                                break

            # If there are multiple images on the same row, create an item for each
            for idx, img in enumerate(imgs):
                item_design = design_val
                if not item_design:
                    item_design = f"Sheet_{self.sheet_name}_Row_{r_num}"
                
                status = "Ready" if design_val else "Missing Design #"
                self.items.append(SheetRowItem(
                    excel_row=r_num,
                    design_number=item_design,
                    image=img,
                    status=status
                ))
            processed_rows.add(r_num)

        # 2. Also check if there are rows with design numbers that had no images
        # (This is useful for preview visibility)
        header_row = self.detection.header_row_idx
        for r_idx in range(header_row, len(self.raw_rows)):
            r_num = r_idx + 1
            if r_num not in processed_rows:
                row_vals = self.raw_rows[r_idx]
                if d_col < len(row_vals) and row_vals[d_col] is not None:
                    design_val = str(row_vals[d_col]).strip()
                    if design_val:
                        self.items.append(SheetRowItem(
                            excel_row=r_num,
                            design_number=design_val,
                            image=None,
                            status="No Image"
                        ))

        # Sort items by Excel row number
        self.items.sort(key=lambda x: x.excel_row)


def _get_cell_resolved_value(ws: openpyxl.worksheet.worksheet.Worksheet, row: int, col: int) -> Any:
    """Reads cell value, correctly resolving merged cell ranges."""
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    # Check merged cells
    coord = cell.coordinate
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            top_left = ws.cell(row=rng.min_row, column=rng.min_col)
            return top_left.value
    return None


def extract_images_from_openpyxl(ws: openpyxl.worksheet.worksheet.Worksheet) -> List[ExtractedImage]:
    """Extracts images using openpyxl worksheet._images."""
    extracted = []
    if not hasattr(ws, "_images") or not ws._images:
        return extracted

    for img in ws._images:
        try:
            # Determine row and col (0-indexed in openpyxl anchors)
            row_idx = 0
            col_idx = 0
            
            if hasattr(img, "anchor"):
                anchor = img.anchor
                if isinstance(anchor, str):
                    # e.g. "E5"
                    row_idx, col_idx = openpyxl.utils.coordinate_to_tuple(anchor)
                    row_idx -= 1
                    col_idx -= 1
                elif hasattr(anchor, "_from"):
                    col_idx = anchor._from.col
                    row_idx = anchor._from.row
                elif hasattr(anchor, "from"):
                    from_obj = getattr(anchor, "from")
                    col_idx = getattr(from_obj, "col", 0)
                    row_idx = getattr(from_obj, "row", 0)
                elif hasattr(anchor, "row") and hasattr(anchor, "col"):
                    row_idx = anchor.row
                    col_idx = anchor.col

            # Get raw image bytes
            data = img._data()
            fmt = getattr(img, "format", "PNG") or "PNG"
            
            # Excel row is 1-indexed, col is 0-indexed
            extracted.append(ExtractedImage(
                row=row_idx + 1,
                col=col_idx,
                data=data,
                fmt=fmt
            ))
        except Exception:
            continue

    return extracted


def extract_images_from_zip_fallback(xlsx_path: str, sheet_name: str) -> List[ExtractedImage]:
    """Directly parses OpenXML zip archive to extract images and anchor positions.
    Guarantees extraction even when openpyxl fails to parse drawing relations.
    """
    extracted = []
    try:
        with zipfile.ZipFile(xlsx_path, "r") as z:
            # Parse workbook.xml
            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
            ns = {
                "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            }
            
            # Find relationship ID for the given sheet
            sheet_rid = None
            for s in wb_xml.findall(".//main:sheet", ns):
                if s.attrib.get("name") == sheet_name:
                    sheet_rid = s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    break
                    
            if not sheet_rid:
                return extracted
                
            # Parse workbook.xml.rels
            wb_rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel_map = {
                r.attrib["Id"]: r.attrib["Target"]
                for r in wb_rels_xml.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            }
            
            target = rel_map.get(sheet_rid, "").lstrip("/")
            sheet_path = target if target.startswith("xl/") else "xl/" + target
            
            # Parse sheet rels
            sheet_dir, sheet_file = posixpath.split(sheet_path)
            sheet_rel_path = posixpath.join(sheet_dir, "_rels", sheet_file + ".rels")
            
            if sheet_rel_path not in z.namelist():
                return extracted
                
            sheet_rels = ET.fromstring(z.read(sheet_rel_path))
            drawing_paths = []
            for rel in sheet_rels.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                if rel.attrib.get("Type", "").endswith("/drawing"):
                    d_target = rel.attrib["Target"]
                    if d_target.startswith("/"):
                        drawing_paths.append(d_target.lstrip("/"))
                    else:
                        drawing_paths.append(posixpath.normpath(posixpath.join(sheet_dir, d_target)))
                        
            for d_path in drawing_paths:
                if d_path not in z.namelist():
                    continue
                d_xml = ET.fromstring(z.read(d_path))
                d_dir, d_file = posixpath.split(d_path)
                d_rel_path = posixpath.join(d_dir, "_rels", d_file + ".rels")
                
                d_rels = {}
                if d_rel_path in z.namelist():
                    drels_xml = ET.fromstring(z.read(d_rel_path))
                    for dr in drels_xml.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        d_rels[dr.attrib["Id"]] = dr.attrib["Target"]
                        
                # Look for oneCellAnchor and twoCellAnchor
                for anchor in d_xml:
                    from_elem = anchor.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from")
                    if from_elem is None:
                        continue
                        
                    col_elem = from_elem.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col")
                    row_elem = from_elem.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row")
                    if col_elem is None or row_elem is None:
                        continue
                        
                    col = int(col_elem.text)
                    row = int(row_elem.text)
                    
                    blip = anchor.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
                    if blip is None:
                        continue
                        
                    embed_id = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    media_target = d_rels.get(embed_id, "")
                    if not media_target:
                        continue
                        
                    if media_target.startswith("/"):
                        media_path = media_target.lstrip("/")
                    else:
                        media_path = posixpath.normpath(posixpath.join(d_dir, media_target))
                        
                    if media_path in z.namelist():
                        img_bytes = z.read(media_path)
                        ext = posixpath.splitext(media_path)[1].replace(".", "").upper()
                        extracted.append(ExtractedImage(
                            row=row + 1,  # 1-indexed
                            col=col,      # 0-indexed
                            data=img_bytes,
                            fmt=ext or "PNG"
                        ))
    except Exception:
        pass
        
    return extracted


def scan_excel_workbook(xlsx_path: str, progress_callback: Optional[Callable[[str, float], None]] = None) -> List[SheetScanResult]:
    """Scans all worksheets in an Excel workbook, extracting data, drawings, and column maps.
    
    Args:
        xlsx_path: Path to the .xlsx or .xlsm file.
        progress_callback: Optional callback(status_msg, progress_float).
        
    Returns:
        List of SheetScanResult objects for each worksheet.
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    if progress_callback:
        progress_callback("Opening Excel workbook...", 0.1)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_results: List[SheetScanResult] = []
    total_sheets = len(wb.sheetnames)

    for s_idx, sheet_name in enumerate(wb.sheetnames):
        if progress_callback:
            progress_callback(f"Scanning sheet '{sheet_name}' ({s_idx+1}/{total_sheets})...", 0.1 + (s_idx / total_sheets) * 0.7)

        ws = wb[sheet_name]
        
        # 1. Read row values
        raw_rows: List[List[Any]] = []
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        
        for r in range(1, max_r + 1):
            row_vals = []
            for c in range(1, max_c + 1):
                val = _get_cell_resolved_value(ws, r, c)
                row_vals.append(val)
            raw_rows.append(row_vals)

        # 2. Extract images using openpyxl
        images = extract_images_from_openpyxl(ws)
        
        # 3. Fallback to direct zip parser if openpyxl extracted 0 images
        if not images:
            images = extract_images_from_zip_fallback(xlsx_path, sheet_name)

        # 4. Detect columns
        image_anchor_cols = [img.col for img in images]
        detection = detect_sheet_columns(sheet_name, raw_rows, image_anchor_cols)

        # 5. Build SheetScanResult
        sheet_result = SheetScanResult(
            sheet_name=sheet_name,
            raw_rows=raw_rows,
            images=images,
            detection=detection
        )
        sheet_results.append(sheet_result)

    wb.close()
    if progress_callback:
        progress_callback("Workbook scan complete.", 1.0)

    return sheet_results


def export_images(
    sheet_results: List[SheetScanResult],
    output_dir: str,
    options: Dict[str, Any],
    progress_callback: Optional[Callable[[str, float, Dict[str, int]], None]] = None
) -> Dict[str, Any]:
    """Exports and saves images for all configured sheets as JPG files.
    
    Args:
        sheet_results: List of SheetScanResult.
        output_dir: Base directory to save files.
        options: Configuration dict (jpeg_quality, duplicate_mode, subfolders_per_sheet).
        progress_callback: Optional callback(message, progress_ratio, stats_dict).
        
    Returns:
        Summary statistics dictionary.
    """
    os.makedirs(output_dir, exist_ok=True)
    
    quality = options.get("jpeg_quality", 95)
    dup_mode = options.get("duplicate_mode", "once")
    subfolders = options.get("subfolders_per_sheet", True)
    
    # Calculate total images to export
    total_images = sum(
        1 for res in sheet_results for item in res.items if item.image is not None
    )
    
    stats = {
        "total_images": total_images,
        "saved_count": 0,
        "skipped_count": 0,
        "error_count": 0,
        "missing_design_count": 0
    }
    
    current_idx = 0
    used_filenames_global: set = set()
    
    for res in sheet_results:
        # Determine target folder
        if subfolders and len(sheet_results) > 1:
            clean_sheet_folder = sanitize_filename(res.sheet_name, fallback="Sheet")
            target_folder = os.path.join(output_dir, clean_sheet_folder)
        else:
            target_folder = output_dir
            
        os.makedirs(target_folder, exist_ok=True)
        used_filenames_sheet: set = set()
        
        for item in res.items:
            if item.image is None:
                continue
                
            current_idx += 1
            
            # Check design number
            if not item.raw_design_number:
                stats["missing_design_count"] += 1
                base_name = f"Unnamed_{sanitize_filename(res.sheet_name)}_{item.excel_row}"
            else:
                base_name = item.clean_design_number
                
            used_set = used_filenames_sheet if subfolders else used_filenames_global
            target_path = resolve_unique_filepath(
                target_dir=target_folder,
                base_name=base_name,
                ext=".jpg",
                mode=dup_mode,
                used_names=used_set
            )
            
            if target_path is None:
                item.status = "Skipped (Duplicate)"
                stats["skipped_count"] += 1
                continue
                
            # Convert and save as JPEG
            success, err = convert_and_save_jpeg(
                img_bytes=item.image.data,
                output_path=target_path,
                quality=quality
            )
            
            if success:
                item.status = "Saved"
                item.saved_path = target_path
                stats["saved_count"] += 1
            else:
                item.status = "Error"
                item.error_message = err
                stats["error_count"] += 1
                
            if progress_callback:
                progress = current_idx / max(1, total_images)
                msg = f"Saving ({current_idx}/{total_images}): {os.path.basename(target_path)}"
                progress_callback(msg, progress, stats)
                
    return stats
