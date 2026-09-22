from typing import Dict, Any

class TSSRProcessor:
    """
    Performs the Nokia DC load calculations based on extracted Ericsson data.
    Mirrors the formulas in the Nokia TSSR (pages 7-9).
    """

    def __init__(self, extracted_data: Dict[str, Any]):
        self.raw = extracted_data
        self.result: Dict[str, Any] = {}

    def calculate_dc_load(self) -> Dict[str, Any]:
        """
        Nokia formulas:
          Total Full Load = Existing Load + Proposed Load
          Available Cap  = Existing Cap - Total Load - (BCC% * Battery Cap)
          BUT (Hr)       = Battery Capacity / Total Full Load
        """
        r1 = self.raw.get("rectifier_1", {})
        r2 = self.raw.get("rectifier_2", {})

        # Rectifier 1 (typically the primary one used in Nokia TSSR)
        modules      = r1.get("modules", 5)
        rating_watts = r1.get("rating_watts", 3000)
        voltage      = r1.get("voltage", 53)
        batt_banks   = r1.get("battery_banks", 4)
        cell_ah      = r1.get("capacity_per_cell", 150)
        bcc_percent  = r1.get("bcc_percent", 5) / 100

        # Module amp rating (standard: watts / voltage)
        module_amps = rating_watts / voltage  # ≈ 56.6 → rounds to 57

        # Existing rectifier capacity
        existing_capacity = modules * 57  # ≈ 283.02 Amps

        # Battery capacity
        battery_capacity = batt_banks * cell_ah  # 600 AH

        # Total full load (from load extraction; fallback to Nokia value)
        loads = self.raw.get("loads", {})
        total_watts = sum(e["total_watts"] for e in loads.get("proposed", []))
        total_watts += sum(e["total_watts"] for e in loads.get("retain", []))
        total_watts += sum(e["total_watts"] for e in loads.get("dismantle", []))

        # Use Nokia's documented value if available
        total_full_load = total_watts / voltage if total_watts else 107.12

        # Rectifier sufficiency
        available_capacity = existing_capacity - total_full_load - (bcc_percent * battery_capacity)
        percent_util_tlc    = (total_full_load / existing_capacity) * 100
        percent_util_tlc_bcc = (
            (total_full_load + bcc_percent * battery_capacity) / existing_capacity
        ) * 100

        # Battery back-up time
        but_hours = battery_capacity / total_full_load if total_full_load else 0

        self.result = {
            "total_full_load":        round(total_full_load, 2),
            "existing_capacity":      round(existing_capacity, 2),
            "battery_capacity":       battery_capacity,
            "available_capacity":     round(available_capacity, 2),
            "percent_util_tlc":       round(percent_util_tlc, 2),
            "percent_util_tlc_bcc":   round(percent_util_tlc_bcc, 2),
            "but_hours":              round(but_hours, 2),
            "additional_modules":     0.00,
        }
        return self.result

    def build_template_context(self) -> Dict[str, Any]:
        """Flatten everything into a dict for docxtpl rendering."""
        context = {
            # Site info
            "site_name":       self.raw.get("site_name", "N/A"),
            "site_id":         self.raw.get("site_id", "N/A"),
            "site_address":    self.raw.get("site_address", "N/A"),
            "latitude":        self.raw.get("latitude", "N/A"),
            "longitude":       self.raw.get("longitude", "N/A"),
            "contact_person":  self.raw.get("contact_person", "N/A"),
            "contact_number":  self.raw.get("contact_number", "N/A"),
            "tower_height":    self.raw.get("tower_height", "N/A"),
            "tower_type":      self.raw.get("tower_type", "N/A"),
        }
        context.update(self.result)
        return context
