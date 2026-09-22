import fitz  # PyMuPDF
import os
from typing import List, Dict

class ImageExtractor:
    """
    Extracts embedded images from the Ericsson TSSR PDF.
    Categorizes them by page/section for placement in Nokia template.
    """

    # Page numbers based on analysis of the Ericsson PDF
    # Page 4  -> Vicinity Map
    # Page 8  -> C2 Existing Antenna RRU MW Photos
    # Page 9  -> C3 Panoramic View
    # Page 37 -> F1 Site Construction Layout
    # Page 39 -> H1 Existing Equipment Layout
    IMAGE_SECTIONS = {
        4:  "vicinity_map",
        8:  "antenna_photos",
        9:  "panoramic_view",
        37: "site_layout",
        39: "equipment_layout",
    }

    def __init__(self, pdf_path: str, output_dir: str = "assets/temp_images"):
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        self.extracted: Dict[str, List[str]] = {}

    def extract_all(self) -> Dict[str, List[str]]:
        """Loop through pages, extract images, and organize by section."""
        doc = fitz.open(self.pdf_path)

        for page_num, section in self.IMAGE_SECTIONS.items():
            if page_num >= len(doc):
                continue

            page = doc[page_num]
            images = page.get_images(full=True)
            saved_paths = []

            for idx, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                ext = base_image["ext"]

                # Skip tiny images (logos, icons) — keep only real photos
                if len(image_bytes) < 20_000:  # 20 KB threshold
                    continue

                filename = f"{section}_p{page_num}_{idx}.{ext}"
                filepath = os.path.join(self.output_dir, filename)

                with open(filepath, "wb") as f:
                    f.write(image_bytes)

                saved_paths.append(filepath)

            if saved_paths:
                self.extracted[section] = saved_paths

        doc.close()
        return self.extracted

    def get_section_images(self, section: str) -> List[str]:
        return self.extracted.get(section, [])

    def cleanup(self):
        """Remove temp images after DOCX generation."""
        for files in self.extracted.values():
            for f in files:
                if os.path.exists(f):
                    os.remove(f)
        self.extracted = {}
