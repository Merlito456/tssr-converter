"""
template_filler.py

Opens the Nokia T7 template UNCHANGED, fills in values by locating
labeled cells and existing image placeholders, and saves the result.

No manual template tagging required.
"""

from docx import Document
from docx.shared import Mm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy
from pathlib import Path
from typing import Any, Optional
import io
import os


class TemplateFiller:
    # ---------------------------------------------------------------
    # CONFIGURATION — Cell coordinates for the Nokia T7 template
    # If Nokia updates the template layout, adjust these maps.
    # ---------------------------------------------------------------

    # Table index in doc.tables → (row, col) of the cell to FILL
    # Label cell is at (row, col - 1) for right-side fills,
    # or (row - 1, col) for below-label fills.
    SITE_INFO_FIELDS = {
        # (row, label_col, fill_col) for the site information table
        # Adjust after inspecting: print_table_map(doc)
        "Site Address":        (1, 0, 1),
        "Site Coordinates":    (1, 2, 3),
        "Site owner":          (2, 0, 1),
        "Site Contact person": (3, 0, 1),
        "Mobile #":            (3, 2, 3),
        "TCO Name":            (5, 0, 1),
        "Site Class":          (5, 2, 3),
        "Site Type":           (6, 0, 1),
        "Room Access":         (7, 0, 1),
        "Cabin Location":      (7, 2, 3),
        "Flood History":       (8, 0, 1),
        "Hauling Remarks":     (8, 2, 3),
        "Site profile":        (9, 0, 1),
        "Site key Location":   (9, 2, 3),
        "FO Name":             (12, 0, 1),
        "Mobile # (FO)":       (12, 2, 3),
        "Site Access Requirement": (13, 0, 1),
    }

    # Materials table (page 15): map label → row index in that table
    MATERIALS_ROWS = {
        "Proposed length of grounding": 0,
        "Proposed length of patchcord":  1,
        "Proposed length of Power Cable": 2,
        "Terminal lugs":                 3,
        "Shrinkable tube":               4,
        "Tie wrap":                      5,
        "Terminal Log":                  6,
        "Liquid-tight Flexible Steel Conduit": 7,
        "DC breaker":                    8,
    }

    def __init__(self, template_path: str):
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        self.doc = Document(template_path)
        self.template_path = template_path

    # ===============================================================
    # PUBLIC API
    # ===============================================================

    def fill_text(self, context: dict):
        """Fill all simple text cells by matching labels in the tables."""
        self._fill_revision_table(context)
        self._fill_site_info_table(context)
        self._fill_materials_table(context)
        self._fill_change_history(context)
        self._fill_remarks_paragraph(context)

    def fill_rectifier_1(self, ctx: dict):
        """Fill Rectifier 1 table (typically doc.tables[3])."""
        self._fill_rectifier_block(table_idx=3, prefix="r1", ctx=ctx)

    def fill_rectifier_2(self, ctx: dict):
        """Fill Rectifier 2 table (typically doc.tables[4])."""
        self._fill_rectifier_block(table_idx=4, prefix="r2", ctx=ctx)
        self._fill_sufficiency_computation(ctx)

    def replace_images(self, image_map: dict):
        """
        Replace existing images in the doc with new ones.
        image_map = {
            "vicinity_map":    "/path/to/map.jpg",
            "site_photo_1":    "/path/to/photo1.jpg",
            "site_photo_2":    "/path/to/photo2.jpg",
            "proposed_space_1": "...",
            "proposed_space_2": "...",
            "trs_diagram":     "...",
            "room_layout":     "...",
            "cable_routing":   "...",
        }
        """
        # Order of images in the document (top-to-bottom, left-to-right):
        # The template has a known sequence of images we need to swap.
        # Adjust the ORDER tuple if your template has more/fewer.
        IMAGE_ORDER = [
            "vicinity_map",      # page 5, first shape
            "site_photo_1",      # page 5, second shape (left)
            "site_photo_2",      # page 5, third shape (right)
            "proposed_space_1",  # page 7, first shape
            "proposed_space_2",  # page 7, second shape
            "trs_diagram",       # page 7, third shape
            "room_layout",       # page 13, first shape
            "cable_routing",     # page 13, second shape
        ]
        self._replace_images_by_order(IMAGE_ORDER, image_map)

    def save(self, output_path: str):
        self.doc.save(output_path)
        return output_path

    # ===============================================================
    # INTERNAL — Text cells
    # ===============================================================

    def _set_cell(self, cell, value):
        """Clear and set a table cell's text, preserving formatting."""
        if cell is None:
            return
        # Preserve the first paragraph's style
        p = cell.paragraphs[0]
        # Remove extra paragraphs
        for extra in cell.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        # Clear runs
        for run in p.runs:
            run.text = ""
        if p.runs:
            p.runs[0].text = str(value)
        else:
            p.add_run(str(value))

    def _find_cell_by_label(self, table, label: str) -> Optional[Any]:
        """Find the value cell to the right of a label."""
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                if label.lower() in cell.text.strip().lower():
                    # Return the next cell in the row
                    if idx + 1 < len(row.cells):
                        return row.cells[idx + 1]
        return None

    def _fill_revision_table(self, ctx):
        """Table 0 — Revision / Approvals (page 3)."""
        table = self.doc.tables[0]
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                t = cell.text.strip().lower()
                if "site name" in t and idx + 1 < len(row.cells):
                    self._set_cell(row.cells[idx + 1], ctx["site_name"])
                elif "plaid" in t and idx + 1 < len(row.cells):
                    self._set_cell(row.cells[idx + 1], ctx["site_id"])
                elif "region" in t and idx + 1 < len(row.cells):
                    self._set_cell(row.cells[idx + 1], ctx["region"])
                elif "site address" in t and idx + 1 < len(row.cells):
                    self._set_cell(row.cells[idx + 1], ctx["site_address"])

    def _fill_site_info_table(self, ctx):
        """Table 1 — Site Information (page 4)."""
        if len(self.doc.tables) < 2:
            return
        table = self.doc.tables[1]

        # Direct label matching per field
        label_map = {
            "Site Address":           ctx["site_address"],
            "Site Coordinates":       f"{ctx['latitude']}, {ctx['longitude']}",
            "Site owner":             ctx["site_owner"],
            "Site Contact person":    ctx["contact_person"],
            "TCO Name":               ctx["tco_name"],
            "Site Class":             ctx["site_class"],
            "Site Type":              ctx["site_type"],
            "Room Access":            ctx["room_access"],
            "Cabin Location":         ctx["cabin_location"],
            "Flood History":          ctx["flood_history"],
            "Hauling Remarks":        ctx["hauling_remarks"],
            "Site profile":           ctx["site_profile"],
            "Site key Location":      ctx["site_key_location"],
            "FO Name":                ctx["fo_name"],
            "Site Access Requirement": ctx["site_access"],
        }
        for label, value in label_map.items():
            cell = self._find_cell_by_label(table, label)
            if cell is not None:
                self._set_cell(cell, value)

        # Mobile # — two occurrences (contact + FO). Fill sequentially.
        mobile_cells = []
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                if cell.text.strip().lower() == "mobile #" and idx + 1 < len(row.cells):
                    mobile_cells.append(row.cells[idx + 1])
        if len(mobile_cells) >= 1:
            self._set_cell(mobile_cells[0], ctx["contact_number"])
        if len(mobile_cells) >= 2:
            self._set_cell(mobile_cells[1], ctx["fo_mobile"])

    def _fill_materials_table(self, ctx):
        """Table for page 15 — Materials / Inventory."""
        # Search for the table containing "Materials needed"
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and "materials needed" in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None:
            return

        # Row index → value mapping (from template order)
        row_values = {
            0: f"{ctx['grounding_length']}m",
            1: f"{ctx['patchcord_length']}m",
            2: ctx["power_cable_length"],
            3: f"{ctx['terminal_lugs_qty']}pcs",
            4: f"{ctx['shrink_tube_qty']}pcs",
            5: f"{ctx['tie_wrap_qty']}pc",
            6: f"{ctx['terminal_log_qty']}pcs",
            7: f"{ctx['conduit_length']}m",
            8: f"{ctx['dc_breaker_rating']}A",
        }
        for r_idx, value in row_values.items():
            if r_idx < len(target.rows):
                row = target.rows[r_idx]
                if len(row.cells) >= 3:
                    self._set_cell(row.cells[2], value)

    def _fill_change_history(self, ctx):
        """Page 18 — Change history table."""
        target = None
        for tbl in self.doc.tables:
            if tbl.rows and "ver" in tbl.rows[0].cells[0].text.lower():
                target = tbl
                break
        if target is None or len(target.rows) < 2:
            return

        row = target.rows[1]  # first data row (0.1)
        values = [
            "0.1", "Draft", ctx["change_date"], ctx["author_name"],
            ctx["owner_name"], ctx["reviewer_name"], ctx["review_date"],
            ctx["approver_name"], ctx["approval_date"], ctx["change_description"],
        ]
        for idx, val in enumerate(values):
            if idx < len(row.cells):
                self._set_cell(row.cells[idx], val)

    def _fill_remarks_paragraph(self, ctx):
        """Page 4 — Remarks paragraph below the site info table."""
        for p in self.doc.paragraphs:
            if "Remarks :" in p.text or "Remarks:" in p.text:
                # Find the NEXT paragraph and inject
                # python-docx doesn't allow "insert after" trivially,
                # so append text to this paragraph instead.
                p.add_run(f"\n{ctx['site_remarks']}")
                break

    # ===============================================================
    # INTERNAL — Rectifier tables
    # ===============================================================

    def _fill_rectifier_block(self, table_idx: int, prefix: str, ctx: dict):
        """
        Rectifier tables contain:
          - Header block with labels like 'No. of Rectifier Module:5'
          - Data block with equipment rows
        We:
          1. Replace numbers embedded in text ('...Module:5' → '...Module:6')
          2. Inject equipment rows by cloning the existing data row
        """
        if table_idx >= len(self.doc.tables):
            return
        table = self.doc.tables[table_idx]

        # --- 1. Patch the header numbers ---
        header_subs = {
            "No. of Rectifier Module": str(ctx.get(f"{prefix}_modules", "")),
            "Rating(watt s)":          str(ctx.get(f"{prefix}_rating_watts", "")),
            "Voltage(V)":              str(ctx.get(f"{prefix}_voltage", "")),
            "# of Batteries in Banks": str(ctx.get(f"{prefix}_battery_banks", "")),
            "Capacity Per Cell(AH)":   str(ctx.get(f"{prefix}_capacity_per_cell", "")),
            "% BCC":                   str(ctx.get(f"{prefix}_bcc", "")),
            "Rectifier Module rating(amps)": str(ctx.get(f"{prefix}_module_amps", "")),
        }
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text
                for label, new_val in header_subs.items():
                    if label in txt and new_val:
                        # Replace the number that follows the label
                        import re
                        new_txt = re.sub(
                            rf"({re.escape(label)}[:\s]*)[\d.]+",
                            rf"\g<1>{new_val}",
                            txt
                        )
                        if new_txt != txt:
                            self._set_cell(cell, new_txt)

        # --- 2. Inject equipment rows ---
        # Locate the LAST data row (the one with the deepest content)
        # and clone it for each equipment entry.
        loads = (
            ctx.get(f"{prefix}_dismantle_loads", [])
            + ctx.get(f"{prefix}_proposed_loads", [])
            + ctx.get(f"{prefix}_retain_loads", [])
        )
        if not loads:
            return

        # Find data rows — heuristic: rows whose first cell has an equipment code
        # (all uppercase, no spaces) and at least 5 columns
        data_rows = []
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 5:
                first = cells[0].text.strip()
                # Equipment codes: AMOD, ABIO, AHEGC, RADIO 6626, etc.
                if first and not first.lower().startswith(("equipment", "decom load", "proposed")):
                    if any(c.isupper() for c in first) or first.startswith("RADIO"):
                        data_rows.append(row)

        if not data_rows:
            return

        template_row = data_rows[-1]._tr  # clone the last one
        parent = template_row.getparent()
        insert_after = template_row

        for item in loads:
            new_tr = deepcopy(template_row)
            # Insert after the previous one
            insert_after.addnext(new_tr)
            insert_after = new_tr

            # Now populate the new row
            from docx.table import _Cell
            new_row_cells = [new_tr.findall(qn("w:tc"))]
            # Easier: rebuild via python-docx table API
        # After injection, re-read the table to fill values
        self._populate_equipment_rows(table, loads)

    def _populate_equipment_rows(self, table, loads):
        """Fill the (now existing) equipment data rows with load values."""
        # Collect equipment-coded rows in order
        data_rows = []
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 5:
                first = cells[0].text.strip()
                if first and any(c.isupper() for c in first):
                    data_rows.append(row)

        # Pair rows with load dicts (up to whichever is smaller)
        for row, load in zip(data_rows, loads):
            cells = row.cells
            self._set_cell(cells[0], load["name"])
            self._set_cell(cells[1], str(load["watts"]))
            self._set_cell(cells[2], str(load["qty"]))
            self._set_cell(cells[3], str(load["total_watts"]))
            self._set_cell(cells[4], str(load["amps"]))
            if len(cells) >= 6:
                self._set_cell(cells[5], load["remark"])

    def _fill_sufficiency_computation(self, ctx):
        """Page 10 — Replace hardcoded numbers in the computation text."""
        subs = {
            "600 AH": f"{ctx['battery_capacity']} AH",
            "77 Amperes": f"{ctx['available_capacity']} Amperes",
            "53%": f"{ctx['r2_percent_util_tlc']}%",
            "66%": f"{ctx['r2_percent_util_tlc_bcc']}%",
            "5.03 Hr": f"{ctx['r2_but_hours']} Hr",
            "5.03  Hr": f"{ctx['r2_but_hours']} Hr",
        }
        for p in self.doc.paragraphs:
            for old, new in subs.items():
                if old in p.text:
                    for run in p.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)

    # ===============================================================
    # INTERNAL — Images
    # ===============================================================

    def _replace_images_by_order(self, order, image_map):
        """
        Replace images in the document sequentially.
        Uses the drawing XML to swap the embedded image bytes.
        """
        # Collect all inline drawings in document order
        drawings = self._collect_all_drawings()

        for idx, key in enumerate(order):
            if idx >= len(drawings):
                break
            path = image_map.get(key)
            if not path or not Path(path).exists():
                continue
            self._swap_image(drawings[idx], path)

    def _collect_all_drawings(self):
        """Return a list of <w:drawing> elements in document order."""
        drawings = []
        # Paragraphs
        for p in self.doc.paragraphs:
            drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        # Tables
        for tbl in self.doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        drawings.extend(p._element.findall(".//" + qn("w:drawing")))
        return drawings

    def _swap_image(self, drawing_el, new_image_path: str):
        """Replace the image bytes referenced by a <w:drawing> element."""
        # Find the relationship ID
        blip = drawing_el.find(".//" + qn("a:blip"))
        if blip is None:
            return
        embed_attr = qn("r:embed")
        r_id = blip.get(embed_attr)
        if r_id is None:
            return

        # Get the part (from the docx package)
        part = self.doc.part.related_parts.get(r_id)
        if part is None:
            return

        # Replace the blob
        with open(new_image_path, "rb") as f:
            new_bytes = f.read()

        # The part's _blob attribute is what docx writes when saving
        part._blob = new_bytes
        # Update content type based on extension
        ext = Path(new_image_path).suffix.lower().lstrip(".")
        ct_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
                  "png": "image/png", "gif": "image/gif",
                  "bmp": "image/bmp"}
        if ext in ct_map:
            try:
                part._content_type = ct_map[ext]
            except Exception:
                pass
