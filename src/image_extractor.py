"""
src/image_extractor.py

Extracts embedded images from an Ericsson TSSR PDF and organizes
them by the Nokia T7 template slot they belong to.

Public API:
    extractor = ImageExtractor("path.pdf", output_dir="assets/temp_images")
    image_map = extractor.extract_all()
    # image_map = {
    #     "vicinity_map":     [".../vicinity_map_p4_0.jpeg", ...],
    #     "proposed_space":   [...],
    #     "room_layout":      [...],
    #     "cable_routing":    [...],
    # }
"""

import os
from typing import Dict, List

# PyMuPDF renamed its package from "fitz" to "pymupdf".
# Try the new name first, fall back to the old one.
try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz  # older PyMuPDF (< 1.24)
    except ImportError as e:
        raise ImportError(
            "PyMuPDF is required. Install it via: pip install PyMuPDF"
        ) from e


class ImageExtractor:
    """
    Extracts images from an Ericsson TSSR PDF page by page.

    Each PDF page number maps to a logical section of the Ericsson TSSR:
      - page 4  → vicinity map + site photos
      - page 7  → proposed Nokia OLT space (reference photos)
      - page 37 → room / equipment layout diagram
      - page 39 → cable routing / equipment installation photos

    Adjust IMAGE_SECTIONS if the Ericsson PDF layout changes.
    """

    IMAGE_SECTIONS: Dict[int, str] = {
        4:  "vicinity_map",
        7:  "proposed_space",
        37: "room_layout",
        39: "cable_routing",
    }

    # Anything smaller than this (in bytes) is probably an icon or logo.
    MIN_SIZE_BYTES: int = 15_000

    def __init__(self, pdf_path: str, output_dir: str = "assets/temp_images"):
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        self.extracted: Dict[str, List[str]] = {}

    # ------------------------------------------------------------------
    # PUBLIC
    # ------------------------------------------------------------------
    def extract_all(self) -> Dict[str, List[str]]:
        """
        Walk the PDF, extract every embedded image ≥ MIN_SIZE_BYTES,
        and organize them by section. Returns a dict mapping
        section-name → list of file paths.
        """
        doc = fitz.open(self.pdf_path)
        try:
            for page_num, section in self.IMAGE_SECTIONS.items():
                if page_num >= len(doc):
                    continue

                page = doc[page_num]
                saved = self._extract_page_images(page, page_num, section)
                if saved:
                    self.extracted.setdefault(section, []).extend(saved)
        finally:
            doc.close()

        return self.extracted

    def cleanup(self):
        """Delete all extracted temp images."""
        for files in self.extracted.values():
            for f in files:
                try:
                    os.remove(f)
                except OSError:
                    pass
        self.extracted = {}

    # ------------------------------------------------------------------
    # INTERNAL
    # ------------------------------------------------------------------
    def _extract_page_images(self, page, page_num: int, section: str) -> List[str]:
        """Extract and save all qualifying images from one page."""
        saved: List[str] = []

        try:
            images = page.get_images(full=True)
        except Exception:
            return saved

        for idx, img_info in enumerate(images):
            try:
                xref = img_info[0]
                base = page.parent.extract_image(xref)
            except Exception:
                continue

            data = base.get("image", b"")
            ext = base.get("ext", "png")

            if len(data) < self.MIN_SIZE_BYTES:
                continue

            fname = f"{section}_p{page_num}_{idx}.{ext}"
            fpath = os.path.join(self.output_dir, fname)

            try:
                with open(fpath, "wb") as f:
                    f.write(data)
                saved.append(fpath)
            except OSError:
                continue

        return saved
