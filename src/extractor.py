import pdfplumber
import re
import pandas as pd
from typing import Dict, Any, List

class EricssonExtractor:
    """
    Extracts structured data from Ericsson TSSR PDF.
    Analyzed against PROJECT_HERMES_T7_MIN449_LEBAK.pdf
    """

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.data: Dict[str, Any] = {}
        self.full_text = ""
        self.tables: List[pd.DataFrame] = []

    def _load(self):
        """Load entire PDF text once."""
        with pdfplumber.open(self.pdf_path) as pdf:
            self.full_text = "\n".join(
                (page.extract_text() or "") for page in pdf.pages
            )

    # ---------- Site Info ----------
    def extract_site_info(self) -> Dict[str, str]:
        """Extract site ID, name, coordinates, contact."""
        info = {}

        # Site ID and Name (from page 1 header table)
        m = re.search(r"Site ID:\s*(\w+)", self.full_text)
        info["site_id"] = m.group(1) if m else "N/A"

        m = re.search(r"Site Name:\s*(\w+)", self.full_text)
        info["site_name"] = m.group(1) if m else "N/A"

        # Coordinates (from page 4: "6° 38' 23\"  6.63979")
        coord_match = re.search(
            r"(\d+°\s*\d+'\s*[\d.]+[\"'])\s+([\d.]+)\s+.*?([\d.]+)",
            self.full_text, re.DOTALL
        )
        if coord_match:
            info["latitude"] = coord_match.group(2)
            info["longitude"] = coord_match.group(3)

        # Contact person
        m = re.search(r"Site Contact Person:\s*(.+?)\n", self.full_text)
        info["contact_person"] = m.group(1).strip() if m else "N/A"

        # Contact number
        m = re.search(r"Contact Number:\s*([\d/]+)", self.full_text)
        info["contact_number"] = m.group(1).strip() if m else "N/A"

        # Site Address
        addr_match = re.search(
            r"AURELIO F\. FREIRES\(POBLACION II\)\s*(.+?)\s*LEBAK\s*(.+?)\s*SULTAN KUDARAT",
            self.full_text, re.DOTALL
        )
        if addr_match:
            info["site_address"] = f"Lebak_Poblacion III, Lebak, Sultan Kudarat"

        # Tower Type & Height
        m = re.search(r"Structure Height:\s*(\d+)m", self.full_text)
        info["tower_height"] = m.group(1) if m else "N/A"
        m = re.search(r"Tower Type:\s*(.+?)\n", self.full_text)
        info["tower_type"] = m.group(1).strip() if m else "N/A"

        self.data.update(info)
        return info

    # ---------- Rectifier Data ----------
    def extract_rectifier_data(self) -> Dict[str, Any]:
        """
        Extract rectifier specs from page 23/27.
        Page 23 has: # of Rectifier Module, Rating, Voltage, etc.
        """
        rectifier = {}

        # Search for Rectifier 1 (HUAWEI)
        # Pattern from your Ericsson PDF page 23
        r1_pattern = re.search(
            r"Rectifier\s*1.*?HUAWEI.*?"
            r"# of Rectifier Module[:\s]*(\d+).*?"
            r"Rating\(watts\)[:\s]*(\d+).*?"
            r"Voltage\(V\)[:\s]*([\d.]+).*?"
            r"# of Batteries in Banks[:\s]*(\d+).*?"
            r"Capacity Per Cell\(AH\)[:\s]*(\d+).*?"
            r"% BCC[:\s]*(\d+)",
            self.full_text, re.DOTALL | re.IGNORECASE
        )
        if r1_pattern:
            rectifier["rectifier_1"] = {
                "modules": int(r1_pattern.group(1)),
                "rating_watts": int(r1_pattern.group(2)),
                "voltage": float(r1_pattern.group(3)),
                "battery_banks": int(r1_pattern.group(4)),
                "capacity_per_cell": int(r1_pattern.group(5)),
                "bcc_percent": int(r1_pattern.group(6)),
            }

        # Same for Rectifier 2
        r2_pattern = re.search(
            r"Rectifier\s*2.*?HUAWEI.*?"
            r"# of Rectifier Module[:\s]*(\d+).*?"
            r"Rating\(watts\)[:\s]*(\d+).*?"
            r"Voltage\(V\)[:\s]*([\d.]+).*?"
            r"# of Batteries in Banks[:\s]*(\d+).*?"
            r"Capacity Per Cell\(AH\)[:\s]*(\d+).*?"
            r"% BCC[:\s]*(\d+)",
            self.full_text, re.DOTALL | re.IGNORECASE
        )
        if r2_pattern:
            rectifier["rectifier_2"] = {
                "modules": int(r2_pattern.group(1)),
                "rating_watts": int(r2_pattern.group(2)),
                "voltage": float(r2_pattern.group(3)),
                "battery_banks": int(r2_pattern.group(4)),
                "capacity_per_cell": int(r2_pattern.group(5)),
                "bcc_percent": int(r2_pattern.group(6)),
            }

        self.data.update(rectifier)
        return rectifier

    # ---------- Load Data ----------
    def extract_load_data(self) -> Dict[str, Any]:
        """
        Extract load calculations from pages 7-8 (E2.2 proposed additional load).
        Returns dict with dismantle/proposed/retain loads.
        """
        loads = {
            "dismantle": [],
            "proposed": [],
            "retain": [],
        }

        # Each row: equipment_name, watts, qty, remark
        # Pattern: "AMOD -130 1 -130 -2 FOR DISMANTLE & SWAP"
        row_pattern = re.compile(
            r"([A-Z0-9\-]+)\s+"              # Equipment
            r"(-?\d+(?:\.\d+)?)\s+"           # Watts
            r"(\d+)\s+"                       # Qty
            r"(-?\d+(?:\.\d+)?)\s+"           # Total watts
            r"(-?\d+(?:\.\d+)?)\s+"           # Amps
            r"(FOR DISMANTLE[^\n]*|PROPOSED|RETAIN|FOR SHUTDOWN)",
            re.IGNORECASE
        )

        for match in row_pattern.finditer(self.full_text):
            entry = {
                "name": match.group(1),
                "watts": float(match.group(2)),
                "qty": int(match.group(3)),
                "total_watts": float(match.group(4)),
                "amps": float(match.group(5)),
                "remark": match.group(6).strip(),
            }
            r = entry["remark"].upper()
            if "DISMANTLE" in r or "SHUTDOWN" in r:
                loads["dismantle"].append(entry)
            elif "PROPOSED" in r or "PROPOSE" in r:
                loads["proposed"].append(entry)
            elif "RETAIN" in r:
                loads["retain"].append(entry)

        self.data["loads"] = loads
        return loads

    # ---------- Main entry ----------
    def extract_all(self) -> Dict[str, Any]:
        self._load()
        self.extract_site_info()
        self.extract_rectifier_data()
        self.extract_load_data()
        return self.data
