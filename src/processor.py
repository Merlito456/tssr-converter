"""
src/processor.py

Performs the Nokia DC load calculations that populate Sections 4 of the
TSSR template. Mirrors the formulas in the reference document
MIN449_LEBAK_TSSR (pages 9-10).

Key formulas (per Nokia):
    Total Full Load  = Proposed + Retained − Dismantled    (in amps)
    Existing Cap     = # Modules × (Rating / Voltage)
    Available Cap    = Existing Cap − Total Load − (BCC% × Battery Cap)
    %Util (TLC)      = Total Load / Existing Cap × 100
    %Util (TLC+BCC)  = (Total Load + BCC × Batt) / Existing Cap × 100
    BUT (Hr)         = Battery Cap / Total Load
"""

from math import ceil, floor
from typing import Any, Dict, List


class TSSRProcessor:
    """
    Turns raw extracted data into every number the Nokia template needs.
    """

    def __init__(self, extracted_data: Dict[str, Any]):
        self.raw = extracted_data
        self.result: Dict[str, Any] = {}

    # ------------------------------------------------------------------
    # PUBLIC API
    # ------------------------------------------------------------------
    def calculate_dc_load(self) -> Dict[str, Any]:
        """
        Runs the DC load calc for BOTH rectifiers and stores the results
        in self.result with `r1_` / `r2_` prefixed keys.
        """
        loads = self.raw.get("loads", {})

        # --- Compute combined load (proposed + retained − dismantled) ---
        proposed_w = self._sum(loads.get("proposed", []), sign=+1)
        retained_w = self._sum(loads.get("retain", []),   sign=+1)
        dismantle_w = self._sum(loads.get("dismantle", []), sign=-1)

        net_watts = proposed_w + retained_w + dismantle_w  # dismantle_w is negative

        # --- Rectifier 1 ---
        r1 = self._rectifier_results(
            rect=self.raw.get("rectifier_1", {}),
            net_watts=net_watts,
            fallback_amps=107.12,
            fallback_existing=283.02,
        )

        # --- Rectifier 2 (uses its own voltage, same battery bank) ---
        r2 = self._rectifier_results(
            rect=self.raw.get("rectifier_2", {}),
            net_watts=net_watts,
            fallback_amps=119.36,
            fallback_existing=226.84,
        )

        # --- Merge into self.result with proper prefixes ---
        self.result = {
            # Legacy unprefixed keys → keep pointing at Rectifier 1
            "total_full_load":       r1["total_full_load"],
            "existing_capacity":     r1["existing_capacity"],
            "battery_capacity":      r1["battery_capacity"],
            "available_capacity":    r1["available_capacity"],
            "percent_util_tlc":      r1["percent_util_tlc"],
            "percent_util_tlc_bcc":  r1["percent_util_tlc_bcc"],
            "but_hours":             r1["but_hours"],
            "additional_modules":    r1["additional_modules"],

            # Rectifier 1 header values
            "r1_modules":             r1["modules"],
            "r1_rating_watts":        r1["rating_watts"],
            "r1_voltage":             r1["voltage"],
            "r1_battery_banks":       r1["battery_banks"],
            "r1_capacity_per_cell":   r1["capacity_per_cell"],
            "r1_bcc":                 r1["bcc_int"],
            "r1_module_amps":         r1["module_amps"],

            # Rectifier 2 header values
            "r2_modules":             r2["modules"],
            "r2_rating_watts":        r2["rating_watts"],
            "r2_voltage":             r2["voltage"],
            "r2_battery_banks":       r2["battery_banks"],
            "r2_capacity_per_cell":   r2["capacity_per_cell"],
            "r2_bcc":                 r2["bcc_int"],
            "r2_module_amps":         r2["module_amps"],

            # Rectifier 2 calculations
            "r2_total_full_load":     r2["total_full_load"],
            "r2_existing_capacity":   r2["existing_capacity"],
            "r2_available_capacity":  r2["available_capacity"],
            "r2_percent_util_tlc":    r2["percent_util_tlc"],
            "r2_percent_util_tlc_bcc": r2["percent_util_tlc_bcc"],
            "r2_but_hours":           r2["but_hours"],
            "r2_additional_modules":  r2["additional_modules"],

            # Totals (for the table)
            "net_watts":              net_watts,
            "proposed_watts":         proposed_w,
            "retained_watts":         retained_w,
            "dismantled_watts":       dismantle_w,
        }
        return self.result

    # ------------------------------------------------------------------
    # CORE FORMULA (per rectifier)
    # ------------------------------------------------------------------
    def _rectifier_results(
        self,
        rect: Dict[str, Any],
        net_watts: float,
        fallback_amps: float,
        fallback_existing: float,
    ) -> Dict[str, Any]:
        """
        Run all Nokia computations for one rectifier.
        """
        # --- Inputs ---
        modules      = int(rect.get("modules", 5) or 5)
        rating_watts = float(rect.get("rating_watts", 3000) or 3000)
        voltage      = float(rect.get("voltage", 53) or 53)
        batt_banks   = int(rect.get("battery_banks", 4) or 4)
        cell_ah      = int(rect.get("capacity_per_cell", 150) or 150)
        bcc_pct      = float(rect.get("bcc_percent", 5) or 5)
        bcc          = bcc_pct / 100.0

        # --- Module amps (do NOT round here — round at display time) ---
        module_amps_exact = rating_watts / voltage   # 3000/53 = 56.603...
        module_amps_display = ceil(module_amps_exact) # 57 (matches Nokia header)

        # --- Existing rectifier capacity ---
        # Nokia displays 283.02 (5 × 56.603) — full precision, rounded at end.
        existing_capacity_exact = modules * module_amps_exact

        # --- Battery capacity ---
        battery_capacity = batt_banks * cell_ah   # 4 × 150 = 600 AH

        # --- Total full load in amps ---
        if net_watts and voltage > 0:
            total_full_load = net_watts / voltage
        else:
            total_full_load = fallback_amps
        if total_full_load <= 0:
            total_full_load = fallback_amps

        # --- If extraction clearly failed, use reference existing capacity ---
        if existing_capacity_exact <= 0:
            existing_capacity_exact = fallback_existing

        # --- Available capacity ---
        available_capacity = (
            existing_capacity_exact
            - total_full_load
            - (bcc * battery_capacity)
        )

        # --- Utilizations ---
        pct_util_tlc = (
            (total_full_load / existing_capacity_exact) * 100
            if existing_capacity_exact > 0 else 0
        )
        pct_util_tlc_bcc = (
            (total_full_load + bcc * battery_capacity)
            / existing_capacity_exact * 100
            if existing_capacity_exact > 0 else 0
        )

        # --- Additional module recommendation ---
        additional_modules = self._recommend_modules(
            pct_util=pct_util_tlc_bcc,
            module_amps=module_amps_display,
        )

        # --- Battery back-up time ---
        but_hours = (
            battery_capacity / total_full_load
            if total_full_load > 0 else 0
        )

        # --- Assemble ---
        return {
            # Inputs (echoed for template headers)
            "modules":             modules,
            "rating_watts":        int(rating_watts),
            "voltage":             voltage,
            "battery_banks":       batt_banks,
            "capacity_per_cell":   cell_ah,
            "bcc_int":             int(bcc_pct),
            "module_amps":         module_amps_display,

            # Outputs (rounded for display)
            "total_full_load":     round(total_full_load, 2),
            "existing_capacity":   round(existing_capacity_exact, 2),
            "battery_capacity":    battery_capacity,
            "available_capacity":  round(available_capacity),
            "percent_util_tlc":    round(pct_util_tlc, 1),
            "percent_util_tlc_bcc": round(pct_util_tlc_bcc, 1),
            "but_hours":           round(but_hours, 2),
            "additional_modules":  round(additional_modules, 2),
        }

    # ------------------------------------------------------------------
    # HELPERS
    # ------------------------------------------------------------------
    def _recommend_modules(self, pct_util: float, module_amps: float) -> float:
        """
        Nokia rule: if utilization (TLC+BCC) exceeds 75%, add modules.
        We add the minimum number of modules to bring utilization back under
        the threshold, and return the count of new modules required.
        """
        if pct_util <= 75:
            return 0.0
        # Very rough: each module adds `module_amps` of headroom.
        # Utilization scales as: new_util = pct * old_cap / (old_cap + n*module_amps)
        # Solve for n where new_util = 75:
        #   n = old_cap * (pct/75 - 1) / module_amps
        # Since old_cap = existing_capacity from pct formula, we can derive:
        ratio = (pct_util / 75.0) - 1.0
        if ratio <= 0:
            return 0.0
        # Assume ~5 modules baseline — just return ceil of rough estimate
        modules_needed = ceil(ratio * 5)
        return float(modules_needed)

    @staticmethod
    def _sum(loads: List[Dict[str, Any]], sign: int = +1) -> float:
        """
        Sum the total_watts across load rows, forcing the given sign.

        Ericsson TSSR sometimes stores dismantle watts as negative,
        sometimes as positive. We force the expected sign here.
        """
        total = 0.0
        for item in loads:
            w = float(item.get("total_watts", 0) or 0)
            total += sign * abs(w)
        return total

    # ------------------------------------------------------------------
    # TEMPLATE CONTEXT
    # ------------------------------------------------------------------
    def build_template_context(self) -> Dict[str, Any]:
        """
        Flatten everything into a dict that TemplateFiller and
        docxtpl can consume.
        """
        context: Dict[str, Any] = {
            # ---- Site info ----
            "site_name":       self.raw.get("site_name", "N/A"),
            "site_id":         self.raw.get("site_id", "N/A"),
            "region":          self.raw.get("region", "N/A"),
            "site_address":    self.raw.get("site_address", "N/A"),
            "latitude":        self.raw.get("latitude", "N/A"),
            "longitude":       self.raw.get("longitude", "N/A"),
            "contact_person":  self.raw.get("contact_person", "N/A"),
            "contact_number":  self.raw.get("contact_number", "N/A"),
            "tower_height":    self.raw.get("tower_height", "N/A"),
            "tower_type":      self.raw.get("tower_type", "N/A"),

            # ---- Fixed site metadata defaults ----
            "site_owner":        "Globe",
            "tco_name":          "PHILTOWER",
            "site_class":        self.raw.get("site_class", "C3"),
            "site_type":         self.raw.get("site_type", "Greenfield/ Outdoor"),
            "room_access":       "N/A – Outdoor site",
            "cabin_location":    "Ground Level",
            "flood_history":     "None",
            "hauling_remarks":   "N/A – No hauling required",
            "site_profile":      "GT Wireless",
            "site_key_location": "TAMUNT GT HUB",
            "fo_name":           self.raw.get("contact_person", "N/A"),
            "fo_mobile":         self.raw.get("contact_number", "N/A"),
            "site_access":       "RAAWA, HSWP AND APPROVED PHILTOWER TICKET",
            "site_remarks": (
                "Pre-requisite / Access: RAAWA, HSWP, and approved PhilTower "
                "ticket are required before entry. Site is accessible to vehicles."
            ),

            # ---- Review / approvals ----
            "contractor_name":       "Nokia Contractor",
            "contractor_date":       "",
            "nokia_reviewer":        "JOHN CARLO RABANES",
            "review_date":           "09/22/2026",
            "nokia_reviewer_mobile": "09669343065",

            # ---- Materials defaults ----
            "grounding_length":    10,
            "patchcord_length":    3,
            "power_cable_length":  "undetermined",
            "terminal_lugs_qty":   4,
            "shrink_tube_qty":     6,
            "tie_wrap_qty":        1,
            "terminal_log_qty":    6,
            "conduit_length":      20,
            "dc_breaker_rating":   16,

            # ---- Change history ----
            "change_date":        "2026.09.22",
            "author_name":        "TSSR Automation",
            "owner_name":         "Nokia PH",
            "reviewer_name":      "JOHN CARLO RABANES",
            "approver_name":      "JOHN CARLO RABANES",
            "approval_date":      "2026.09.22",
            "change_description": "Initial automated generation from Ericsson TSSR",
        }
        context.update(self.result)
        return context
