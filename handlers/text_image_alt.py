"""
Text with Image Alt Handler — Image on RIGHT with alternate styling.

Template layouts (alternate styling variants):
    idx 22 — Text with Image Two Thirds Alt (image 2/3, text 1/3)
    idx 23 — Text with Image Half Alt (50/50 split)
    idx 24 — Text with Image One Third Alt (image 1/3, text 2/3)

These have the same spatial arrangement as the original Text with Image
layouts (19-21) but use a different visual treatment (alternate background/
styling). The placeholder indices differ from the originals:

    Layout 22: title=0, subtitle=31, content=33, picture=34, footer=10, slidenum=11
    Layout 23: title=0, subtitle=31, content=33, picture=32, footer=10, slidenum=11
    Layout 24: title=0, subtitle=31, content=10, picture=34, footer=20, slidenum=21

Detection: primarily by source layout name containing "alt". Falls back
to the same image+text heuristics as TextImageHandler but at slightly
lower confidence, so the original handler gets priority for non-Alt
source slides.

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.text_image import TextImageHandler
from utils.extractor import extract_images, extract_shapes_with_text

# Alt layout configs: (layout_index, content_ph_idx, picture_ph_idx)
ALT_LAYOUT_TWO_THIRDS = (22, 33, 34)   # Image takes 2/3
ALT_LAYOUT_HALF = (23, 33, 32)          # 50/50
ALT_LAYOUT_ONE_THIRD = (24, 10, 34)     # Image takes 1/3

# Same thresholds as original
TEXT_HEAVY_THRESHOLD = 200
TEXT_LIGHT_THRESHOLD = 60


class TextImageAltHandler(TextImageHandler):

    name = "Text with Image Alt"
    description = "Content slide with text and image (alternate styling)"
    layout_name = "Text with Image Half Alt"
    layout_index = 23  # Default, overridden by get_layout_index()

    # Title and subtitle indices are the same across all alt variants
    PH_TITLE = 0
    PH_SUBTITLE = 31
    # Content and footer indices vary per layout — handled in fill_slide

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides that should use the Alt text-with-image layout.

        Primary trigger: source layout name contains "alt".
        Secondary: same image+text heuristics as parent but at lower
        confidence so the original handler wins for non-Alt sources.
        """
        if slide_index == 0:
            return 0.0

        # Primary: source layout name
        layout_name = slide.slide_layout.name.lower()
        if "alt" in layout_name and "image" in layout_name:
            return 0.70

        # Secondary: fall back to parent's heuristics at lower confidence
        # (the original TextImageHandler should claim non-Alt slides)
        parent_score = super().detect(slide, slide_index)
        # Only claim if parent wouldn't — reduced confidence
        if parent_score > 0.5:
            return parent_score * 0.5  # e.g. 0.65 → 0.325 (below threshold)
        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract content using parent's logic, then remap layout variants
        to the Alt set.
        """
        # Use parent extraction — it sets _layout_variant to original tuples
        result = super().extract_content(slide, slide_index)

        # Remap to Alt layout variants
        from handlers.text_image import LAYOUT_TWO_THIRDS, LAYOUT_HALF, LAYOUT_ONE_THIRD
        variant = result.get("_layout_variant")

        if variant == LAYOUT_TWO_THIRDS or (variant and variant[0] == 19):
            result["_layout_variant"] = ALT_LAYOUT_TWO_THIRDS
        elif variant == LAYOUT_ONE_THIRD or (variant and variant[0] == 21):
            result["_layout_variant"] = ALT_LAYOUT_ONE_THIRD
        else:
            result["_layout_variant"] = ALT_LAYOUT_HALF

        return result

    # --- Dynamic layout ---

    def get_layout_index(self, content: dict) -> int:
        """Return the appropriate Alt layout index."""
        variant = content.get("_layout_variant", ALT_LAYOUT_HALF)
        return variant[0]

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill Alt layout placeholders. Content and picture indices
        vary per layout variant.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title (same across all variants)
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Subtitle (same across all variants)
        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        # Content body — index varies by variant
        variant = content.get("_layout_variant", ALT_LAYOUT_HALF)
        content_idx = variant[1]
        if content.get("content") and content_idx in placeholders:
            placeholders[content_idx].text = content["content"]

        # Picture — index varies by variant
        if content.get("image_blob"):
            picture_idx = variant[2]
            if picture_idx in placeholders:
                image_stream = io.BytesIO(content["image_blob"])
                placeholders[picture_idx].insert_picture(image_stream)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "content": "varies (33 or 10)",
            "picture": "varies (32 or 34)",
        }
