"""
src/template_filler.py

Fills the Nokia TSSR - T7 template AS-IS by matching label cells
and existing image shapes. No manual tagging required.

The template structure has been analyzed against:
  - templates/nokia_t7_template.docx
  - MIN449_LEBAK_TSSR (reference output)
  - Ericsson TSSR (input)

Update only the CELL_MAPS dict if the template layout changes.
"""

from docx import Document
from docx.shared import Mm
from docx.oxml.ns import qn
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional
import re


class TemplateFiller:

    # ================================================================
    # CELL TARGETING MAPS
    # ================================================================

    # Page 3 — Revision/Approvals table (table index 0)
    # Format: label_text (lowercase, partial match) → context_key
    REVISION_FIELDS = {
        "site name": "site_name",
        "plaid":     "site_id",
        "region":    "region",
    }

    # Page 4 — Site Information table (table index 1)
    # All in 4-column layout: label | value | label | value
    SITE_INFO_FIELDS = {
        "site address":           "site_address",
        "site coordinates":       "site_coordinates",  # custom combined
        "site owner":             "site_owner",
        "site contact person":    "contact_person",
        "tco name":               "tco_name",
        "site class":             "site_class",
        "site type":              "site_type",
        "room access":            "room_access",
        "cabin location":         "cabin_location",
        "flood history":          "flood_history",
        "hauling remarks":        "hauling_remarks",
        "site profile":           "site_profile",
        "site key location":      "site_key_location",
        "fo name":                "fo_name",
        "site access requirement": "site_access",
    }

    # Page 15 — Materials table
    MATERIALS_TABLE_KEY = "materials needed"
    MATERIALS_ROWS = {
        0: "grounding_length",     # 10m
        1: "patchcord_length",     # 3m
        2: "power_cable_length",   # undetermined
        3: "terminal_lugs_qty",    # 4pcs
        4: "shrink_tube_qty",      # 6pcs
        5: "tie_wrap_qty",         # 1pc
        6: "terminal_log_qty",     # 6pcs
        7: "conduit_length",       # 20m
        8: "dc_breaker_rating",    # 16A
    }

    # Page 18 — Change history
    CHANGE_HISTORY_ROW = 1  # first data row (0.1)

    # Rectifier table indices (adjust based on inspection)
    RECTIFIER_1_TABLE_IDX = 3
    RECTIFIER_2_TABLE_IDX = 4

    # Image ordering in template (top-to-bottom, left-to-right)
    IMAGE_ORDER = [
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

    def __init__(self, template_path: str):
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        self.doc = Document(template_path)
        self.template_path = template_path

    # ================================================================
    # PUBLIC API
    # ================================================================

    def fill_text(self, ctx: Dict[str, Any]):
        """Fill all text cells in the document."""
        self._fill_revision_table(ctx)
        self._fill_site_info_table(ctx)
        self._fill_site_remarks_paragraph(ctx)
        self._fill_materials_table(ctx)
        self._fill_change_history(ctx)

    def fill_rectifier_1(self, ctx: Dict[str, Any]):
        """Fill Rectifier 1 table + summary."""
        self._fill_rectifier_block(
            table_idx=self.RECTIFIER_1_TABLE_IDX,
            prefix="r1",
            ctx=ctx,
        )

    def fill_rectifier_2(self, ctx: Dict[str, Any]):
        """Fill Rectifier 2 table + summary."""
        self._fill_rectifier_block(
            table_idx=self.RECTIFIER_2_TABLE_IDX,
            prefix="r2",
            ctx=ctx,
        )
        self._fill_sufficiency_summary(ctx)

    def replace_images(self, image_map: Dict[str, Optional[str]]):
        """Replace existing image shapes with new ones."""
        self._replace_images_by_order(self.IMAGE_ORDER, image_map)

    def save(self, output_path: str):
        self.doc.save(output_path)
        return output_path

    # ================================================================
    # INTERNAL — Helpers
    # ================================================================

    def _set_cell(self, cell, value: str):
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
        """
        Find the value cell to the right of a label cell.
        Returns the cell immediately to the right of the matched label.
        """
        label_lc = label.lower().strip()
        for row in table.rows:
            cells = row.cells
            for idx, cell in enumerate(cells):
                if label_lc in cell.text.strip().lower():
                    # Return the next non-duplicate cell
                    if idx + 1 < len(cells):
                        return cells[idx + 1]
        return None

    def _find_all_cells_with_label(self, table, label: str) -> List[Any]:
        """Find ALL cells matching a label (for repeated labels like 'Mobile #')."""
        label_lc = label.lower().strip()
        matches = []
        for row in table.rows:
            cells = row.cells
            for idx, cell in enumerate(cells):
                if cell.text.strip().lower() == label_lc:
                    if idx + 1 < len(cells):
                        matches.append(cells[idx + 1])
        return matches

    # ================================================================
    # INTERNAL — Page 3: Revision/Approvals
    # ================================================================

    def _fill_revision_table(self, ctx: Dict[str, Any]):
        if len(self.doc.tables) < 1:
            return
        table = self.doc.tables[0]

        # Fill simple label→value pairs
        for label, key in self.REVISION_FIELDS.items():
            cell = self._find_cell_by_label(table, label)
            if cell is not None and key in ctx:
                self._set_cell(cell, ctx[key])

        # Site Address spans differently — look for the standalone cell
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                if "site address" in cell.text.lower():
                    # The value is on the next row, first cell
                    row_idx = table.rows.index(row)
                    if row_idx + 1 < len(table.rows):
                        next_row = table.rows[row_idx + 1]
                        if len(next_row.cells) > 0:
                            self._set_cell(next_row.cells[0], ctx.get("site_address", ""))

    # ================================================================
    # INTERNAL — Page 4: Site Information
    # ================================================================

    def _fill_site_info_table(self, ctx: Dict[str, Any]):
        if len(self.doc.tables) < 2:
            return
        table = self.doc.tables[1]

        # Direct label-based fill
        for label, key in self.SITE_INFO_FIELDS.items():
            cell = self._find_cell_by_label(table, label)
            if cell is not None and key in ctx:
                self._set_cell(cell, ctx[key])

        # Coordinates — combined
        cell = self._find_cell_by_label(table, "site coordinates")
        if cell is not None:
            lat = ctx.get("latitude", "N/A")
            lon = ctx.get("longitude", "N/A")
            self._set_cell(cell, f"{lat}, {lon}")

        # Mobile # — two occurrences (contact + FO)
        mobile_cells = self._find_all_cells_with_label(table, "mobile #")
        if len(mobile_cells) >= 1:
            self._set_cell(mobile_cells[0], ctx.get("contact_number", ""))
        if len(mobile_cells) >= 2:
            self._set_cell(mobile_cells[1], ctx.get("fo_mobile", ""))

    def _fill_site_remarks_paragraph(self, ctx: Dict[str, Any]):
        """Page 4 remarks paragraph at the bottom."""
        remarks_text = ctx.get("site_remarks", "")
        if not remarks_text:
            return
        for p in self.doc.paragraphs:
            if p.text.strip().startswith("Remarks"):
                # Add text after the label
                p.add_run(f"\n{remarks_text}")
                break

    # ================================================================
    # INTERNAL — Page 15: Materials
    # ================================================================

    def _fill_materials_table(self, ctx: Dict[str, Any]):
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and self.MATERIALS_TABLE_KEY in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None:
            return

        # Format values
        values = {
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

        for row_idx, value in values.items():
            if row_idx < len(target.rows):
                row = target.rows[row_idx]
                if len(row.cells) >= 3:
                    self._set_cell(row.cells[2], value)

    # ================================================================
    # INTERNAL — Page 18: Change History
    # ================================================================

    def _fill_change_history(self, ctx: Dict[str, Any]):
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and "ver" in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None or len(target.rows) < 2:
            return

        row = target.rows[self.CHANGE_HISTORY_ROW]
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
    # INTERNAL — Rectifier Tables
    # ================================================================

    def _fill_rectifier_block(self, table_idx: int, prefix: str, ctx: Dict[str, Any]):
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

        # --- 2. Replace equipment data rows ---
        loads = (
            ctx.get(f"{prefix}_dismantle_loads", [])
            + ctx.get(f"{prefix}_proposed_loads", [])
            + ctx.get(f"{prefix}_retain_loads", [])
        )
        if not loads:
            return

        # Find equipment rows (uppercase code + 5+ cols)
        data_row_indices = []
        for i, row in enumerate(table.rows):
            cells = row.cells
            if len(cells) >= 5:
                first = cells[0].text.strip()
                if first and any(c.isupper() for c in first) and len(first) < 20:
                    if not first.lower().startswith(("equipment", "decom load", "proposed")):
                        data_row_indices.append(i)

        if not data_row_indices:
            return

        # Populate the first N rows with load data
        for i, load in enumerate(loads):
            if i >= len(data_row_indices):
                break
            row = table.rows[data_row_indices[i]]
            cells = row.cells
            self._set_cell(cells[0], load["name"])
            if len(cells) > 1:
                self._set_cell(cells[1], str(load["watts"]))
            if len(cells) > 2:
                self._set_cell(cells[2], str(load["qty"]))
            if len(cells) > 3:
                self._set_cell(cells[3], str(load["total_watts"]))
            if len(cells) > 4:
                self._set_cell(cells[4], str(load["amps"]))
            if len(cells) > 5:
                self._set_cell(cells[5], load["remark"])

    def _fill_sufficiency_summary(self, ctx: Dict[str, Any]):
        """Patch the hardcoded numbers in the summary paragraphs."""
        subs = {
            "= 107.12 Amps": f"= {ctx.get('total_full_load', '')} Amps",
            "= 283.02 Amps": f"= {ctx.get('existing_capacity', '')} Amps",
            "= 600 AH": f"= {ctx.get('battery_capacity', '')} AH",
            "= 146 Amperes": f"= {ctx.get('available_capacity', '')} Amperes",
            "= 38%": f"= {ctx.get('percent_util_tlc', '')}%",
            "= 48%": f"= {ctx.get('percent_util_tlc_bcc', '')}%",
            "5.60  HR": f"{ctx.get('but_hours', '')} HR",
            "5.60 Hr": f"{ctx.get('but_hours', '')} Hr",
            # Rectifier 2
            "= 119.36 Amps": f"= {ctx.get('r2_total_full_load', '')} Amps",
            "= 226.84 Amps": f"= {ctx.get('r2_existing_capacity', '')} Amps",
            "= 77 Amperes": f"= {ctx.get('r2_available_capacity', '')} Amperes",
            "= 53%": f"= {ctx.get('r2_percent_util_tlc', '')}%",
            "= 66%": f"= {ctx.get('r2_percent_util_tlc_bcc', '')}%",
            "5.03  Hr": f"{ctx.get('r2_but_hours', '')} Hr",
            "5.03 Hr": f"{ctx.get('r2_but_hours', '')} Hr",
        }
        for p in self.doc.paragraphs:
            for old, new in subs.items():
                if old in p.text:
                    for run in p.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)

    # ================================================================
    # INTERNAL — Images
    # ================================================================

    def _replace_images_by_order(self, order: List[str], image_map: Dict[str, Optional[str]]):
        """Replace images sequentially in document order."""
        drawings = self._collect_all_drawings()

        for idx, key in enumerate(order):
            if idx >= len(drawings):
                break
            path = image_map.get(key)
            if not path or not Path(path).exists():
                continue
            self._swap_image(drawings[idx], path)

    def _collect_all_drawings(self):
        """Return all <w:drawing> elements in document order."""
        drawings = []
        for p in self.doc.paragraphs:
            drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        for tbl in self.doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        return drawings

    def _swap_image(self, drawing_el, new_image_path: str):
        """Replace image bytes referenced by a <w:drawing> element."""
        blip = drawing_el.find(".//" + qn("a:blip"))
        if blip is None:
            return
        r_id = blip.get(qn("r:embed"))
        if r_id is None:
            return

        part = self.doc.part.related_parts.get(r_id)
        if part
