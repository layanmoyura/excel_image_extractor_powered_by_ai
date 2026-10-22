"""Image processing and saving utility.
Handles converting Excel embedded images to standard high-quality JPG,
sanitizing Windows filenames, handling transparency, and managing duplicate names.
"""

import os
import re
import io
from typing import Tuple, Optional, Set
from PIL import Image, ImageOps

# Windows invalid filename characters: < > : " / \ | ? *
INVALID_WIN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename(name: str, fallback: str = "unnamed") -> str:
    """Sanitizes a string to be a safe, valid Windows filename.
    
    Args:
        name: Proposed name (e.g., design number).
        fallback: Name to use if input is empty or entirely invalid.
        
    Returns:
        A clean filename string without extension.
    """
    if name is None:
        return fallback
    
    # Convert to string and strip surrounding whitespace
    clean = str(name).strip()
    
    # Replace invalid Windows characters with a hyphen
    clean = INVALID_WIN_CHARS.sub("-", clean)
    
    # Replace multiple spaces or hyphens with a single character
    clean = re.sub(r'[\s_]+', '_', clean)
    clean = re.sub(r'-+', '-', clean)
    
    # Windows does not permit trailing dots or spaces
    clean = clean.strip(". ")
    
    # Check for reserved Windows filenames (CON, PRN, AUX, NUL, COM1-9, LPT1-9)
    reserved = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}
    if clean.upper() in reserved:
        clean = f"{clean}_item"
        
    if not clean:
        return fallback
        
    # Cap maximum length for Windows MAX_PATH safety
    if len(clean) > 180:
        clean = clean[:180].rstrip(". ")
        
    return clean


def convert_and_save_jpeg(
    img_bytes: bytes,
    output_path: str,
    quality: int = 95,
    bg_color: Tuple[int, int, int] = (255, 255, 255)
) -> Tuple[bool, Optional[str]]:
    """Converts image bytes of any format (PNG, WEBP, BMP, TIFF, EMF, etc.)
    into a high-quality RGB JPEG file on disk.
    
    Args:
        img_bytes: Raw binary bytes of the image.
        output_path: Full file path where the .jpg should be saved.
        quality: JPEG compression quality (1-100, default 95).
        bg_color: Background color for transparent pixels (default white).
        
    Returns:
        Tuple of (success: bool, error_message: str | None).
    """
    try:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        with Image.open(io.BytesIO(img_bytes)) as pil_img:
            # Fix EXIF orientation if present
            try:
                pil_img = ImageOps.exif_transpose(pil_img)
            except Exception:
                pass
                
            # Handle transparency (RGBA, LA, or Palette with transparency)
            if pil_img.mode in ("RGBA", "LA") or (pil_img.mode == "P" and "transparency" in pil_img.info):
                if pil_img.mode != "RGBA":
                    pil_img = pil_img.convert("RGBA")
                # Create white background canvas
                background = Image.new("RGB", pil_img.size, bg_color)
                # Paste image using its alpha channel as mask
                background.paste(pil_img, mask=pil_img.split()[3])
                rgb_img = background
            elif pil_img.mode == "CMYK":
                rgb_img = pil_img.convert("RGB")
            elif pil_img.mode != "RGB":
                rgb_img = pil_img.convert("RGB")
            else:
                rgb_img = pil_img
                
            # Save as JPEG
            rgb_img.save(output_path, "JPEG", quality=quality, optimize=True)
            return True, None
    except Exception as e:
        return False, str(e)


def resolve_unique_filepath(
    target_dir: str,
    base_name: str,
    ext: str = ".jpg",
    mode: str = "once",
    used_names: Optional[Set[str]] = None
) -> Optional[str]:
    """Generates a target file path respecting duplicate handling rules.
    Saves each unique design number strictly once as <DesignNumber>.jpg without suffixes.
    
    Args:
        target_dir: Destination directory.
        base_name: Sanitized base filename (e.g. '18259W1BS8WV').
        ext: File extension (e.g. '.jpg').
        mode: 'once' (save once, skip duplicate occurrences) or 'overwrite'.
        used_names: Set of filenames already assigned in this batch.
        
    Returns:
        Full file path string, or None if skipped (already saved).
    """
    if used_names is None:
        used_names = set()
        
    target_filename = f"{base_name}{ext}"
    full_path = os.path.join(target_dir, target_filename)
    
    if mode == "overwrite":
        used_names.add(target_filename.lower())
        return full_path
        
    # Default 'once': Save same design number once, skip duplicates (NO suffix _ saves)
    if os.path.exists(full_path) or target_filename.lower() in used_names:
        return None
        
    used_names.add(target_filename.lower())
    return full_path


def create_thumbnail(img_bytes: bytes, size: Tuple[int, int] = (60, 60)) -> Optional[Image.Image]:
    """Creates a small PIL thumbnail for GUI preview."""
    try:
        with Image.open(io.BytesIO(img_bytes)) as pil_img:
            thumb = pil_img.copy()
            thumb.thumbnail(size, Image.Resampling.LANCZOS)
            if thumb.mode in ("RGBA", "LA") or (thumb.mode == "P" and "transparency" in thumb.info):
                bg = Image.new("RGB", thumb.size, (255, 255, 255))
                if thumb.mode != "RGBA":
                    thumb = thumb.convert("RGBA")
                bg.paste(thumb, mask=thumb.split()[3])
                return bg
            return thumb.convert("RGB")
    except Exception:
        return None
