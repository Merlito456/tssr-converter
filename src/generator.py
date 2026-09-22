from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os

class TSSRGenerator:
    """
    Renders the Nokia template with text + images.
    Requires the template to have {{ }} placeholders.
    """

    def __init__(self, template_path: str):
        self.doc = DocxTemplate(template_path)

    def _prep_images(self, image_paths, width_mm=80):
        """Convert file paths to InlineImage objects."""
        images = []
        for p in image_paths:
            if os.path.exists(p):
                try:
                    images.append(InlineImage(self.doc, p, width=Mm(width_mm)))
                except Exception as e:
                    print(f"[WARN] Could not embed {p}: {e}")
        return images

    def generate(self, context: dict, image_map: dict, output_path: str) -> str:
        """
        Args:
            context:   Jinja2 variables for text fields
            image_map: dict like {"vicinity_map": [path1], "antenna_photos": [path2, path3]}
            output_path: where to save the .docx
        """
        # Add images into the context
        for key, paths in image_map.items():
            context[f"{key}_images"] = self._prep_images(paths)

        self.doc.render(context)
        self.doc.save(output_path)
        return output_path
