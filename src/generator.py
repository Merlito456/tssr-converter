from src.template_filler import TemplateFiller
from pathlib import Path


class TSSRGenerator:
    def __init__(self, template_path: str):
        self.template_path = template_path

    def generate(self, context: dict, image_map: dict, output_path: str) -> str:
        filler = TemplateFiller(self.template_path)

        # 1. Text replacements
        filler.fill_text(context)

        # 2. Rectifier tables
        filler.fill_rectifier_1(context)
        filler.fill_rectifier_2(context)

        # 3. Image replacement
        # Flatten image_map into the keys the filler expects
        flat = self._flatten_images(image_map)
        filler.replace_images(flat)

        # 4. Save
        filler.save(output_path)
        return output_path

    def _flatten_images(self, image_map: dict) -> dict:
        """
        Ericsson image sections → Nokia slot keys.

        The Ericsson PDF has these sections (from ImageExtractor):
            vicinity_map_and_site_photos
            proposed_space
            room_layout
            cable_routing
        """
        vmp    = image_map.get("vicinity_map_and_site_photos", [])
        prop   = image_map.get("proposed_space", [])
        room   = image_map.get("room_layout", [])
        cable  = image_map.get("cable_routing", [])

        def pick(lst, i):
            if not lst:
                return None
            return lst[i] if i < len(lst) else lst[-1]

        return {
            "vicinity_map":     pick(vmp, 0),
            "site_photo_1":     pick(vmp, 1),
            "site_photo_2":     pick(vmp, 2),
            "proposed_space_1": pick(prop, 0),
            "proposed_space_2": pick(prop, 1),
            "trs_diagram":      pick(prop, 2),
            "room_layout":      pick(room, 0),
            "cable_routing":    pick(cable, 0),
        }
