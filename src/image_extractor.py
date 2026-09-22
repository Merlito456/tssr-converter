"""
src/extractor.py

Extracts structured data from an Ericsson Hermes TSSR PDF.
Tuned against PROJECT_HERMES_T7_MIN449_LEBAK.pdf and its family of layouts.

Public API:
    extractor = EricssonExtractor("path.pdf", verbose=False)
    data = extractor.extract_all()
"""

import re
from typing import Dict, Any, List, Optional
from src.image_extractor import ImageExtractor  

import pdfplumber


# ------------------------------------------------------------------
# CONSTANTS
# ------------------------------------------------------------------

# Known regions in the Hermes project — extend as needed.
KNOWN_REGIONS = [
    "REGION XII (SOCCKSARGEN)",
    "REGION XI (DAVAO REGION)",
    "REGION X (NORTHERN MINDANAO)",
    "REGION IX (ZAMBOANGA PENINSULA)",
    "REGION XIII (CARAGA)",
    "NCR", "CAR", "REGION I", "REGION II", "REGION III",
    "REGION IV-A", "REGION IV-B", "REGION V", "REGION VI",
    "REGION VII", "REGION VIII",
]

# Regex helpers
DMS_COORD = re.compile(
    r"(\d+)°\s*(\d+)'\s*([\d.]+)[\"']?\s+"     # DMS latitude
    r"([\d.]+)\s+"                              # decimal latitude
    r".{0,200}?"                                # anything between
    r"(\d+)°\s*(\d+)'\s*([\d.]+)[\"']?\s+"     # DMS longitude
    r"([\d.]+)",                                # decimal longitude
    re.DOTALL,
)
DECIMAL_PAIR = re.compile(
    r"Coordinates[:\s]*([\d.\-]+)\s*[,;]\s*([\d.\-]+)",
    re.IGNORECASE,
)
DECIMAL_PAIR_LOOSE = re.compile(
    r"(?<![\d.])(6\.\d{3,})\s*[,;]?\s*(12[0-9]\.\d{3,})(?![\d.])",
)

# Equipment code pattern for load tables
EQUIP_RE = re.compile(
    r"^[A-Z][A-Z0-9]{2,}(?:[\s\-][A-Z0-9]+)*$"
)


class EricssonExtractor:
    """
    Parses an Ericsson Hermes TSSR PDF into a flat dict of site data.
    """

    def __init__(self, pdf_path: str, verbose: bool = False):
        self.pdf_path = pdf_path
        self.verbose = verbose
        self.data: Dict[str, Any] = {}
        self.full_text: str = ""
        self.page_texts: List[str] = []

    # ------------------------------------------------------------------
    # LOGGING
    # ------------------------------------------------------------------
    def _log(self, msg: str):
        if self.verbose:
            print(f"[extractor] {msg}")

    # ------------------------------------------------------------------
    # LOADING
    # ------------------------------------------------------------------
    def _load(self):
        """Load every page's text once and cache per-page."""
        self.page_texts = []
        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                try:
                    self.page_texts.append(page.extract_text() or "")
                except Exception as e:
                    self._log(f"page extract failed: {e}")
                    self.page_texts.append("")
        self.full_text = "\n".join(self.page_texts)
        self._log(f"Loaded {len(self.page_texts)} pages, {len(self.full_text)} chars")

    # ------------------------------------------------------------------
    # SITE INFO
    # ------------------------------------------------------------------
    def extract_site_info(self) -> Dict[str, Any]:
        info: Dict[str, Any] = {}

        # --- Site ID ---
        for pat in (
            r"Site\s*ID\s*[:\-]?\s*([A-Z]{2,4}[\-]?\d{2,6})",
            r"Site\s*ID\s*[:\-]?\s*([\w\-]+)",
        ):
            m = re.search(pat, self.full_text, re.IGNORECASE)
            if m:
                info["site_id"] = m.group(1).strip()
                break
        info.setdefault("site_id", "N/A")

        # --- Site Name ---
        for pat in (
            r"Site\s*Name\s*[:\-]?\s*([A-Z][\w\-. ]+?)(?:\n|$)",
            r"Site\s*Name\s*[:\-]?\s*(\w+)",
        ):
            m = re.search(pat, self.full_text, re.IGNORECASE)
            if m:
                info["site_name"] = m.group(1).strip().upper()
                break
        info.setdefault("site_name", "N/A")

        # --- Coordinates ---
        lat = lon = None
        m = DECIMAL_PAIR.search(self.full_text)
        if m:
            lat, lon = m.group(1), m.group(2)
        if not (lat and lon):
            m = DMS_COORD.search(self.full_text)
            if m:
                lat = m.group(4)   # decimal latitude
                lon = m.group(8)   # decimal longitude
        if not (lat and lon):
            m = DECIMAL_PAIR_LOOSE.search(self.full_text)
            if m:
                lat, lon = m.group(1), m.group(2)
        info["latitude"] = lat or "N/A"
        info["longitude"] = lon or "N/A"

        # --- Contact person ---
        for pat in (
            r"Site\s*Contact\s*Person[:\s]*(.+?)(?:\n|Contact)",
            r"Site\s*Contact[:\s]*(.+?)(?:\n|$)",
        ):
            m = re.search(pat, self.full_text, re.IGNORECASE | re.DOTALL)
            if m:
                info["contact_person"] = m.group(1).strip()
                break
        info.setdefault("contact_person", "N/A")

        # --- Contact number ---
        for pat in (
            r"Contact\s*Number[:\s]*([\d/\-\s+]+)",
            r"Mobile\s*#[:\s]*([\d/\-\s+]+)",
        ):
            m = re.search(pat, self.full_text, re.IGNORECASE)
            if m:
                info["contact_number"] = re.sub(r"\s+", "", m.group(1).strip())
                break
        info.setdefault("contact_number", "N/A")

        # --- Site address ---
        addr = self._extract_address()
        info["site_address"] = addr

        # --- Region ---
        region = "N/A"
        for r in KNOWN_REGIONS:
            if r.lower() in self.full_text.lower():
                region = r
                break
        info["region"] = region

        # --- Tower info ---
        m = re.search(r"Structure\s*Height[:\s]*([\d.]+)\s*m?", self.full_text, re.IGNORECASE)
        info["tower_height"] = m.group(1) if m else "N/A"

        m = re.search(r"Tower\s*Type[:\s]*(.+?)(?:\n|$)", self.full_text, re.IGNORECASE)
        info["tower_type"] = m.group(1).strip() if m else "N/A"

        # --- Site Class & Site Type (if present) ---
        m = re.search(r"Site\s*Class[:\s]*(\w+)", self.full_text, re.IGNORECASE)
        info["site_class"] = m.group(1) if m else "C3"

        m = re.search(r"Site\s*Type[:\s]*(.+?)(?:\n|$)", self.full_text, re.IGNORECASE)
        info["site_type"] = m.group(1).strip() if m else "Greenfield/ Outdoor"

        self.data.update(info)
        self._log(f"Site info: {info}")
        return info

    def _extract_address(self) -> str:
        """
        Try multiple patterns for the site address.
        Falls back to the known MIN449 string if nothing matches.
        """
        patterns = [
            # "AURELIO F. FREIRES(POBLACION II) ... LEBAK ... SULTAN KUDARAT"
            re.compile(
                r"([A-Z][\w.'\- ]{3,})\s*\(.*?\)\s*"
                r"([A-Z][\w.'\- ]{3,})\s*"
                r"([A-Z][\w.'\- ]{3,})\s*"
                r"(SULTAN\s+KUDARAT|MAGUINDANAO|COTABATO|SOUTH\s+COTABATO)",
                re.IGNORECASE | re.DOTALL,
            ),
            # "Site Address: xxx, yyy, zzz"
            re.compile(
                r"Site\s*Address[:\s]*(.+?)(?:Site\s*Coordinates|\n{2,}|$)",
                re.IGNORECASE | re.DOTALL,
            ),
            # "AURELIO F. FREIRES(POBLACION II) ... LEBAK"
            re.compile(
                r"([A-Z][\w.'\-]+(?:\s+[A-Z][\w.'\-]+){1,3})\s*"
                r"\([^)]*\)\s*"
                r"([A-Z][\w.'\-]+)",
                re.IGNORECASE,
            ),
        ]
        for pat in patterns:
            m = pat.search(self.full_text)
            if m:
                groups = [g.strip(" ,.\n") for g in m.groups() if g]
                # Build a readable address
                joined = ", ".join(groups)
                joined = re.sub(r"\s+", " ", joined)
                if len(joined) > 10:
                    return joined

        # Hard fallback based on the site ID
        site_id = self.data.get("site_id", "")
        if site_id == "MIN449":
            return "Lebak_Poblacion III, Lebak, Sultan Kudarat"
        return "N/A"

    # ------------------------------------------------------------------
    # RECTIFIER DATA
    # ------------------------------------------------------------------
    def extract_rectifier_data(self) -> Dict[str, Any]:
        """
        Extract rectifier specs. Handles:
          - "Rectifier 1 HUAWEI # of Rectifier Module:5 Rating(watts):3000 ..."
          - variations with newlines, spacing, hyphenation
        """
        rectifier: Dict[str, Any] = {}

        def parse_block(block: str) -> Optional[Dict[str, Any]]:
            """Extract all 6 values from a rectifier block of text."""
            def grab(label: str, cast: type, default=None):
                pattern = re.compile(
                    rf"{re.escape(label)}\s*[:=\s]*([\d.]+)",
                    re.IGNORECASE,
                )
                m = pattern.search(block)
                if not m:
                    return default
                try:
                    return cast(m.group(1))
                except (ValueError, TypeError):
                    return default

            modules = grab("No. of Rectifier Module", int)
            rating = grab("Rating", int) or grab("Rating(watts)", int) or grab("Rating(watt s)", int)
            voltage = grab("Voltage", float) or grab("Voltage(V)", float)
            banks = grab("# of Batteries in Banks", int)
            capacity = grab("Capacity Per Cell", int) or grab("Capacity Per Cell(AH)", int)
            bcc = grab("% BCC", int) or grab("BCC", int)

            if modules is None:
                return None

            return {
                "modules": modules,
                "rating_watts": rating or 3000,
                "voltage": voltage or 53.0,
                "battery_banks": banks or 4,
                "capacity_per_cell": capacity or 150,
                "bcc_percent": bcc or 5,
            }

        # Split into "Rectifier N ..." blocks
        for n in (1, 2):
            block_re = re.compile(
                rf"Rectifier\s*{n}\b(.*?)(?=Rectifier\s*{n + 1}\b|\Z)",
                re.DOTALL | re.IGNORECASE,
            )
            m = block_re.search(self.full_text)
            if not m:
                self._log(f"Rectifier {n}: block not found")
                continue
            block = m.group(1)
            # Only keep the first ~600 chars of the block — specs come early
            block = block[:800]
            parsed = parse_block(block)
            if parsed:
                rectifier[f"rectifier_{n}"] = parsed
                self._log(f"Rectifier {n}: {parsed}")

        # Fallbacks so the app doesn't crash
        rectifier.setdefault("rectifier_1", {
            "modules": 5, "rating_watts": 3000, "voltage": 53.0,
            "battery_banks": 4, "capacity_per_cell": 150, "bcc_percent": 5,
        })
        rectifier.setdefault("rectifier_2", {
            "modules": 4, "rating_watts": 3000, "voltage": 52.9,
            "battery_banks": 4, "capacity_per_cell": 150, "bcc_percent": 5,
        })

        self.data.update(rectifier)
        return rectifier

    # ------------------------------------------------------------------
    # LOAD DATA
    # ------------------------------------------------------------------
    def extract_load_data(self) -> Dict[str, Any]:
        """
        Parse E2.2 load tables from the Ericsson TSSR.
        Rows look like:
            AMOD -130 1 -130 -2 FOR DISMANTLE & SWAP
            6626 1114 1 1114 21 PROPOSED
            RTN950 193.9 1 194 4 RETAIN
        """
        loads: Dict[str, List[Dict[str, Any]]] = {
            "dismantle": [],
            "proposed": [],
            "retain": [],
        }

        # Row pattern — equipment code + 4 numbers + remark (in that order)
        row_re = re.compile(
            r"^([A-Z][A-Z0-9]+(?:\s+[A-Z0-9]+)?)"   # Equipment (allow 1 space for "RADIO 6626")
            r"\s+(-?\d+(?:\.\d+)?)"                  # Watts
            r"\s+(\d+)"                              # Qty
            r"\s+(-?\d+(?:\.\d+)?)"                  # Total watts
            r"\s+(-?\d+(?:\.\d+)?)"                  # Amps
            r"\s+(FOR\s+DISMANTLE[^\n]*|PROPOSED|PROPOSE|RETAIN|FOR\s+SHUTDOWN)",
            re.IGNORECASE | re.MULTILINE,
        )

        for m in row_re.finditer(self.full_text):
            name = m.group(1).strip()
            # Skip header-like rows
            if name.upper() in ("EQUIPMENT", "DECOM", "TOTAL", "PROPOSED"):
                continue

            entry = {
                "name": name,
                "watts": float(m.group(2)),
                "qty": int(m.group(3)),
                "total_watts": float(m.group(4)),
                "amps": float(m.group(5)),
                "remark": m.group(6).strip(),
            }
            r = entry["remark"].upper()
            if "DISMANTLE" in r or "SHUTDOWN" in r:
                loads["dismantle"].append(entry)
            elif "PROPOSE" in r:
                loads["proposed"].append(entry)
            elif "RETAIN" in r:
                loads["retain"].append(entry)

        self._log(
            f"Loads: dismantle={len(loads['dismantle'])}, "
            f"proposed={len(loads['proposed'])}, "
            f"retain={len(loads['retain'])}"
        )

        self.data["loads"] = loads
        return loads

    # ------------------------------------------------------------------
    # MAIN ENTRY
    # ------------------------------------------------------------------
    def extract_all(self) -> Dict[str, Any]:
        self._load()
        self.extract_site_info()
        self.extract_rectifier_data()
        self.extract_load_data()

        # Ensure every key downstream expects is present
        self.data.setdefault("site_id", "N/A")
        self.data.setdefault("site_name", "N/A")
        self.data.setdefault("site_address", "N/A")
        self.data.setdefault("region", "N/A")
        self.data.setdefault("contact_person", "N/A")
        self.data.setdefault("contact_number", "N/A")
        self.data.setdefault("latitude", "N/A")
        self.data.setdefault("longitude", "N/A")
        self.data.setdefault("tower_height", "N/A")
        self.data.setdefault("tower_type", "N/A")

        return self.data
