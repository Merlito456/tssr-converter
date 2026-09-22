"""
src/generator.py

Orchestrates the filling of the Nokia T7 TSSR template:

    1. Fill text cells (site info, materials, change history)
    2. Fill rectifier tables (numbers + equipment rows)
    3. Replace images
    4. Save the output .docx

Returns a report dict with counts of what was filled, so the
Streamlit UI can display it and the developer can debug.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from src.template_filler import TemplateFiller


class TSSRGenerator:
    """
    Fills a Nokia TSSR template with extracted data.

    Usage:
        gen = TSSRGenerator("templates/nokia_t7_template.docx", verbose=True)
        report = gen.generate(context, image_map, "out.docx")
    """

    # These are the image keys that `TemplateFiller.replace_images`
    # expects (see IMAGE_ORDER in template_filler.py).
    NOKIA_IMAGE_SLOTS = [
        "vicinity_map",
        "site_photo_1",
        "site_photo_2",
        "proposed_space_1",
        "proposed_space_2",
        "trs_diagram",
        "room_layout",
        "cable_routing",
    ]

    def __init__(self, template_path: str, verbose: bool = False):
        self.template_path = template_path
        self.verbose = verbose

        if not Path(template_path).exists():
            raise FileNotFoundError(
                f"Template not found: {template_path}\n"
                f"Upload it via the Streamlit sidebar or commit it to "
                f"templates/ in the repo."
            )

    # ------------------------------------------------------------------
    # LOGGING
    # ------------------------------------------------------------------
    def _log(self, msg: str):
        if self.verbose:
            print(f"[generator] {msg}")

    # ------------------------------------------------------------------
    # MAIN
    # ------------------------------------------------------------------
    def generate(
        self,
        context: Dict[str, Any],
        image_map: Dict[str, List[str]],
        output_path: str,
    ) -> Dict[str, Any]:
        """
        Run the full pipeline. Returns a report dict:

            {
              "text_filled":     True,
              "rectifier_1_ok":  True,
              "rectifier_2_ok":  True,
              "images_used":     6,
              "images_missing":  ["site_photo_2", "trs_diagram"],
              "output_path":     "out.docx",
            }
        """
        self._log(f"Opening template: {self.template_path}")
        filler = TemplateFiller(self.template_path)

        report: Dict[str, Any] = {
            "text_filled": False,
            "rectifier_1_ok": False,
            "rectifier_2_ok": False,
            "images_used": 0,
            "images_missing": [],
            "output_path": output_path,
        }

        # ---------- 1. TEXT ----------
        try:
            filler.fill_text(context)
            report["text_filled"] = True
            self._log("Text cells filled")
        except Exception as e:
            self._log(f"fill_text failed: {e}")
            raise

        # ---------- 2. RECTIFIERS ----------
        try:
            filler.fill_rectifier_1(context)
            report["rectifier_1_ok"] = True
            self._log("Rectifier 1 filled")
        except Exception as e:
            self._log(f"fill_rectifier_1 failed (continuing): {e}")

        try:
            filler.fill_rectifier_2(context)
            report["rectifier_2_ok"] = True
            self._log("Rectifier 2 filled")
        except Exception as e:
            self._log(f"fill_rectifier_2 failed (continuing): {e}")

        # ---------- 3. IMAGES ----------
        flat = self._flatten_images(image_map)
        report["images_used"] = sum(1 for v in flat.values() if v)
        report["images_missing"] = [k for k, v in flat.items() if not v]
        self._log(f"Images: {report['images_used']} used, "
                  f"{len(report['images_missing'])} missing")

        try:
            filler.replace_images(flat)
        except Exception as e:
            self._log(f"replace_images failed (continuing): {e}")

        # ---------- 4. SAVE ----------
        filler.save(output_path)
        self._log(f"Saved: {output_path}")

        return report

    # ------------------------------------------------------------------
    # IMAGE FLATTENING
    # ------------------------------------------------------------------
    def _flatten_images(self, image_map: Dict[str, List[str]]) -> Dict[str, Optional[str]]:
        """
        Map Ericsson PDF image sections → Nokia template slots.

        Tolerant to multiple possible section keys because different
        versions of ImageExtractor have used different names.
        """
        # Try every plausible key for each logical group
        vmp = self._first_present(
            image_map,
            "vicinity_map_and_site_photos",
            "vicinity_map",
            "site_photos",
        )
        prop = self._first_present(
            image_map,
            "proposed_space",
            "proposed_space_photos",
        )
        room = self._first_present(
            image_map,
            "room_layout",
            "room_lay_out",
            "equipment_layout",
        )
        cable = self._first_present(
            image_map,
            "cable_routing",
            "cable_routing_and_layout",
            "cable_layout",
        )

        return {
            "vicinity_map":     self._pick(vmp, 0),
            "site_photo_1":     self._pick(vmp, 1),
            "site_photo_2":     self._pick(vmp, 2),
            "proposed_space_1": self._pick(prop, 0),
            "proposed_space_2": self._pick(prop, 1),
            "trs_diagram":      self._pick(prop, 2),
            "room_layout":      self._pick(room, 0),
            "cable_routing":    self._pick(cable, 0),
        }

    @staticmethod
    def _first_present(d: Dict[str, List[str]], *keys: str) -> List[str]:
        """Return the first non-empty list from the given keys."""
        for k in keys:
            if k in d and d[k]:
                return d[k]
        return []

    @staticmethod
    def _pick(lst: List[str], idx: int) -> Optional[str]:
        """Return the idx-th item, or the last item, or None."""
        if not lst:
            return None
        if idx < len(lst):
            return lst[idx]
        return lst[-1]

    # ------------------------------------------------------------------
    # CONVENIENCE (used by app.py to show a quick summary)
    # ------------------------------------------------------------------
    def describe_plan(self, image_map: Dict[str, List[str]]) -> str:
        """Human-readable summary of what will be filled."""
        flat = self._flatten_images(image_map)
        lines = ["Generation plan:"]
        for slot in self.NOKIA_IMAGE_SLOTS:
            path = flat.get(slot)
            status = "✅" if path else "⬜ (empty)"
            name = Path(path).name if path else "—"
            lines.append(f"  {status}  {slot:20s} {name}")
        return "\n".join(lines)
