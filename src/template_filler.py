"""
src/template_filler.py

Fills the Nokia TSSR - T7 template AS-IS by matching label cells
and existing image shapes. No manual tagging required.

The template structure has been analyzed against:
  - templates/nokia_t7_template.docx
  - MIN449_LEBAK_TSSR (reference output)
  - Ericsson TSSR (input)

Update only the CELL TARGETING MAPS below if the template layout changes.

Compatible with python-docx 1.1+ and 1.2+.
"""

from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional
import re

from docx import Document
from docx.oxml.ns import qn


class TemplateFiller:
    """
    Fills the Nokia T7 TSSR template in place.

    Public API:
        filler = TemplateFiller("templates/nokia_t7_template.docx")
        filler.fill_text(context)                    # page 3, 4, 15, 18
        filler.fill_rectifier_1(context)             # page 9
        filler.fill_rectifier_2(context)             # page 10
        filler.replace_images(image_map)             # pages 5, 7, 13
        filler.save("output.docx")
    """

    # ================================================================
    # CELL TARGETING MAPS
    # ================================================================

    # Page 3 — Revision / Approvals (table index 0)
    REVISION_FIELDS: Dict[str, str] = {
        "site name": "site_name",
        "plaid":     "site_id",
        "region":    "region",
    }

    # Page 4 — Site Information (table index 1, 4-column layout)
    SITE_INFO_FIELDS: Dict[str, str] = {
        "site address":             "site_address",
        "site coordinates":         "_site_coordinates_",   # special
        "site owner":               "site_owner",
        "site contact person":      "contact_person",
        "tco name":                 "tco_name",
        "site class":               "site_class",
        "site type":                "site_type",
        "room access":              "room_access",
        "cabin location":           "cabin_location",
        "flood history":            "flood_history",
        "hauling remarks":          "hauling_remarks",
        "site profile":             "site_profile",
        "site key location":        "site_key_location",
        "fo name":                  "fo_name",
        "site access requirement":  "site_access",
    }

    # Page 15 — Materials table
    MATERIALS_TABLE_KEY: str = "materials needed"
    MATERIALS_ROWS: Dict[int, str] = {
        0: "grounding_length",
        1: "patchcord_length",
        2: "power_cable_length",
        3: "terminal_lugs_qty",
        4: "shrink_tube_qty",
        5: "tie_wrap_qty",
        6: "terminal_log_qty",
        7: "conduit_length",
        8: "dc_breaker_rating",
    }

    # Page 18 — Change history
    CHANGE_HISTORY_ROW: int = 1

    # Rectifier tables (adjust after inspecting the actual template)
    RECTIFIER_1_TABLE_IDX: int = 3
    RECTIFIER_2_TABLE_IDX: int = 4

    # Image ordering in template (top-to-bottom, left-to-right)
    IMAGE_ORDER: List[str] = [
        "vicinity_map",       # page 5, first image
        "site_photo_1",       # page 5, second image
        "site_photo_2",       # page 5, third image
        "proposed_space_1",   # page 7, left
        "proposed_space_2",   # page 7, right
        "trs_diagram",        # page 7, bottom
        "room_layout",        # page 13, top
        "cable_routing",      # page 13, bottom
    ]

    # ================================================================
    # INIT
    # ================================================================

    def __init__(self, template_path: str):
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        self.template_path = template_path
        self.doc = Document(template_path)

    # ================================================================
    # PUBLIC API
    # ================================================================

    def fill_text(self, ctx: Dict[str, Any]) -> None:
        """Fill all text cells across the document."""
        self._fill_revision_table(ctx)
        self._fill_site_info_table(ctx)
        self._fill_site_remarks_paragraph(ctx)
        self._fill_materials_table(ctx)
        self._fill_change_history(ctx)

    def fill_rectifier_1(self, ctx: Dict[str, Any]) -> None:
        """Fill Rectifier 1 table."""
        self._fill_rectifier_block(
            table_idx=self.RECTIFIER_1_TABLE_IDX,
            prefix="r1",
            ctx=ctx,
        )

    def fill_rectifier_2(self, ctx: Dict[str, Any]) -> None:
        """Fill Rectifier 2 table + summary computations."""
        self._fill_rectifier_block(
            table_idx=self.RECTIFIER_2_TABLE_IDX,
            prefix="r2",
            ctx=ctx,
        )
        self._fill_sufficiency_summary(ctx)

    def replace_images(self, image_map: Dict[str, Optional[str]]) -> None:
        """Replace existing image shapes with new ones."""
        self._replace_images_by_order(self.IMAGE_ORDER, image_map)

    def save(self, output_path: str) -> str:
        self.doc.save(output_path)
        return output_path

    # ================================================================
    # INTERNAL — Low-level helpers
    # ================================================================

    def _set_cell(self, cell, value: Any) -> None:
        """Set cell text while preserving formatting of the first run."""
        if cell is None or value is None:
            return
        text = str(value)
        p = cell.paragraphs[0]
        # Remove extra paragraphs
        for extra in cell.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        # Clear existing runs
        for run in p.runs:
            run.text = ""
        # Write new text
        if p.runs:
            p.runs[0].text = text
        else:
            p.add_run(text)

    def _find_cell_by_label(self, table, label: str) -> Optional[Any]:
        """Return the value cell to the right of a label cell, or None."""
        label_lc = label.lower().strip()
        for row in table.rows:
            cells = row.cells
            for idx, cell in enumerate(cells):
                if label_lc in cell.text.strip().lower():
                    if idx + 1 < len(cells):
                        return cells[idx + 1]
        return None

    def _find_all_cells_with_label(self, table, label: str) -> List[Any]:
        """Return ALL value cells whose label matches exactly."""
        label_lc = label.lower().strip()
        matches: List[Any] = []
        for row in table.rows:
            cells = row.cells
            for idx, cell in enumerate(cells):
                if cell.text.strip().lower() == label_lc:
                    if idx + 1 < len(cells):
                        matches.append(cells[idx + 1])
        return matches

    # ================================================================
    # Page 3 — Revision / Approvals
    # ================================================================

    def _fill_revision_table(self, ctx: Dict[str, Any]) -> None:
        if len(self.doc.tables) < 1:
            return
        table = self.doc.tables[0]

        # --- Simple label → value pairs ---
        for label, key in self.REVISION_FIELDS.items():
            cell = self._find_cell_by_label(table, label)
            if cell is not None and key in ctx:
                self._set_cell(cell, ctx[key])

        # --- SITE ADDRESS: value is on the next row ---
        rows_list = list(table.rows)
        for row_idx, row in enumerate(rows_list):
            for cell in row.cells:
                txt_lc = cell.text.strip().lower()
                if txt_lc == "site address" or txt_lc.startswith("site address"):
                    if row_idx + 1 < len(rows_list):
                        next_row = rows_list[row_idx + 1]
                        if len(next_row.cells) > 0:
                            self._set_cell(
                                next_row.cells[0],
                                ctx.get("site_address", ""),
                            )
                    return

    # ================================================================
    # Page 4 — Site Information
    # ================================================================

    def _fill_site_info_table(self, ctx: Dict[str, Any]) -> None:
        if len(self.doc.tables) < 2:
            return
        table = self.doc.tables[1]

        # --- Direct label → value ---
        for label, key in self.SITE_INFO_FIELDS.items():
            if key == "_site_coordinates_":
                continue  # handled below
            cell = self._find_cell_by_label(table, label)
            if cell is not None and key in ctx:
                self._set_cell(cell, ctx[key])

        # --- Site Coordinates: combine lat + lon ---
        cell = self._find_cell_by_label(table, "site coordinates")
        if cell is not None:
            lat = ctx.get("latitude", "N/A")
            lon = ctx.get("longitude", "N/A")
            self._set_cell(cell, f"{lat}, {lon}")

        # --- Mobile # appears twice (contact + FO) ---
        mobile_cells = self._find_all_cells_with_label(table, "mobile #")
        if len(mobile_cells) >= 1:
            self._set_cell(mobile_cells[0], ctx.get("contact_number", ""))
        if len(mobile_cells) >= 2:
            self._set_cell(mobile_cells[1], ctx.get("fo_mobile", ""))

    def _fill_site_remarks_paragraph(self, ctx: Dict[str, Any]) -> None:
        """Page 4 — remarks paragraph below the site info table."""
        remarks_text = ctx.get("site_remarks", "")
        if not remarks_text:
            return
        for p in self.doc.paragraphs:
            if p.text.strip().startswith("Remarks"):
                # Append to the same paragraph
                p.add_run(f"\n{remarks_text}")
                return

    # ================================================================
    # Page 15 — Materials
    # ================================================================

    def _fill_materials_table(self, ctx: Dict[str, Any]) -> None:
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and self.MATERIALS_TABLE_KEY in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None:
            return

        # Pre-format values with their units
        values: Dict[int, str] = {
            0: f"{ctx.get('grounding_length', '')}m",
            1: f"{ctx.get('patchcord_length', '')}m",
            2: str(ctx.get("power_cable_length", "")),
            3: f"{ctx.get('terminal_lugs_qty', '')}pcs",
            4: f"{ctx.get('shrink_tube_qty', '')}pcs",
            5: f"{ctx.get('tie_wrap_qty', '')}pc",
            6: f"{ctx.get('terminal_log_qty', '')}pcs",
            7: f"{ctx.get('conduit_length', '')}m",
            8: f"{ctx.get('dc_breaker_rating', '')}A",
        }

        rows_list = list(target.rows)
        for row_idx, value in values.items():
            if row_idx >= len(rows_list):
                continue
            row = rows_list[row_idx]
            if len(row.cells) >= 3:
                self._set_cell(row.cells[2], value)

    # ================================================================
    # Page 18 — Change history
    # ================================================================

    def _fill_change_history(self, ctx: Dict[str, Any]) -> None:
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and "ver" in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None or len(target.rows) < 2:
            return

        rows_list = list(target.rows)
        if self.CHANGE_HISTORY_ROW >= len(rows_list):
            return
        row = rows_list[self.CHANGE_HISTORY_ROW]

        values = [
            "0.1",
            "Draft",
            ctx.get("change_date", ""),
            ctx.get("author_name", ""),
            ctx.get("owner_name", ""),
            ctx.get("reviewer_name", ""),
            ctx.get("approval_date", ""),
            ctx.get("approver_name", ""),
            ctx.get("approval_date", ""),
            ctx.get("change_description", ""),
        ]
        for idx, val in enumerate(values):
            if idx < len(row.cells):
                self._set_cell(row.cells[idx], val)

    # ================================================================
    # Rectifier Tables
    # ================================================================

    def _fill_rectifier_block(
        self,
        table_idx: int,
        prefix: str,
        ctx: Dict[str, Any],
    ) -> None:
        if table_idx >= len(self.doc.tables):
            return
        table = self.doc.tables[table_idx]

        # --- 1. Patch header numbers in merged cells ---
        header_subs = {
            "No. of Rectifier Module": str(ctx.get(f"{prefix}_modules", "")),
            "Rating": str(ctx.get(f"{prefix}_rating_watts", "")),
            "Voltage": str(ctx.get(f"{prefix}_voltage", "")),
            "# of Batteries in Banks": str(ctx.get(f"{prefix}_battery_banks", "")),
            "Capacity Per Cell": str(ctx.get(f"{prefix}_capacity_per_cell", "")),
            "% BCC": str(ctx.get(f"{prefix}_bcc", "")),
            "Rectifier Module rating": str(ctx.get(f"{prefix}_module_amps", "")),
        }
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text
                for label, new_val in header_subs.items():
                    if label in txt and new_val:
                        new_txt = re.sub(
                            rf"({re.escape(label)}[:\s()AHwattvolts]*?)(-?\d+(?:\.\d+)?)",
                            rf"\g<1>{new_val}",
                            txt,
                            count=1,
                        )
                        if new_txt != txt:
                            self._set_cell(cell, new_txt)

        # --- 2. Replace equipment rows ---
        loads = (
            ctx.get(f"{prefix}_dismantle_loads", [])
            + ctx.get(f"{prefix}_proposed_loads", [])
            + ctx.get(f"{prefix}_retain_loads", [])
        )
        if not loads:
            return

        # Find equipment rows (uppercase code + 5+ cols)
        rows_list = list(table.rows)
        data_row_indices: List[int] = []
        for i, row in enumerate(rows_list):
            cells = row.cells
            if len(cells) >= 5:
                first = cells[0].text.strip()
                if (first and any(c.isupper() for c in first)
                        and len(first) < 20
                        and not first.lower().startswith(
                            ("equipment", "decom load", "proposed")
                        )):
                    data_row_indices.append(i)

        if not data_row_indices:
            return

        # Fill the first N equipment rows
        for i, load in enumerate(loads):
            if i >= len(data_row_indices):
                break
            row = rows_list[data_row_indices[i]]
            cells = row.cells
            self._set_cell(cells[0], load.get("name", ""))
            if len(cells) > 1:
                self._set_cell(cells[1], str(load.get("watts", "")))
            if len(cells) > 2:
                self._set_cell(cells[2], str(load.get("qty", "")))
            if len(cells) > 3:
                self._set_cell(cells[3], str(load.get("total_watts", "")))
            if len(cells) > 4:
                self._set_cell(cells[4], str(load.get("amps", "")))
            if len(cells) > 5:
                self._set_cell(cells[5], load.get("remark", ""))

    def _fill_sufficiency_summary(self, ctx: Dict[str, Any]) -> None:
        """Patch hardcoded summary numbers in paragraphs."""
        subs = {
            # Rectifier 1
            "= 107.12 Amps": f"= {ctx.get('total_full_load', '')} Amps",
            "= 283.02 Amps": f"= {ctx.get('existing_capacity', '')} Amps",
            "= 600 AH":      f"= {ctx.get('battery_capacity', '')} AH",
            "= 146 Amperes": f"= {ctx.get('available_capacity', '')} Amperes",
            "= 38%":         f"= {ctx.get('percent_util_tlc', '')}%",
            "= 48%":         f"= {ctx.get('percent_util_tlc_bcc', '')}%",
            "5.60  HR":      f"{ctx.get('but_hours', '')} HR",
            "5.60 Hr":       f"{ctx.get('but_hours', '')} Hr",
            # Rectifier 2
            "= 119.36 Amps": f"= {ctx.get('r2_total_full_load', '')} Amps",
            "= 226.84 Amps": f"= {ctx.get('r2_existing_capacity', '')} Amps",
            "= 77 Amperes":  f"= {ctx.get('r2_available_capacity', '')} Amperes",
            "= 53%":         f"= {ctx.get('r2_percent_util_tlc', '')}%",
            "= 66%":         f"= {ctx.get('r2_percent_util_tlc_bcc', '')}%",
            "5.03  Hr":      f"{ctx.get('r2_but_hours', '')} Hr",
            "5.03 Hr":       f"{ctx.get('r2_but_hours', '')} Hr",
        }
        for p in self.doc.paragraphs:
            for old, new in subs.items():
                if old in p.text:
                    for run in p.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)

    # ================================================================
    # Images
    # ================================================================

    def _replace_images_by_order(
        self,
        order: List[str],
        image_map: Dict[str, Optional[str]],
    ) -> None:
        drawings = self._collect_all_drawings()

        for idx, key in enumerate(order):
            if idx >= len(drawings):
                break
            path = image_map.get(key)
            if not path or not Path(path).exists():
                continue
            self._swap_image(drawings[idx], path)

    def _collect_all_drawings(self) -> List[Any]:
        """Return all <w:drawing> elements in document order."""
        drawings: List[Any] = []
        for p in self.doc.paragraphs:
            drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        for tbl in self.doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        return drawings

    def _swap_image(self, drawing_el, new_image_path: str) -> None:
        """Replace the image bytes referenced by a <w:drawing> element."""
        blip = drawing_el.find(".//" + qn("a:blip"))
        if blip is None:
            return
        r_id = blip.get(qn("r:embed"))
        if r_id is None:
            return

        part = self.doc.part.related_parts.get(r_id)
        if part is None:
            return

        with open(new_image_path, "rb") as f:
            new_bytes = f.read()

        # Replace the raw blob
        part._blob = new_bytes

        # Best-effort content-type update based on extension
        ext = Path(new_image_path).suffix.lower().lstrip(".")
        ct_map = {
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "png": "image/png",
            "gif": "image/gif",
            "bmp": "image/bmp",
        }
        if ext in ct_map:
            try:
                part._content_type = ct_map[ext]
            except Exception:
                pass
