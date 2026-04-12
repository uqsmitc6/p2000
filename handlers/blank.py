"""
Blank Branded Handler — Slides with no meaningful text content.

Template layout 44 "Blank Branded":
    idx 10 — Footer (programme name)
    idx 11 — Slide number

Detection: slide has no meaningful text (only footer/slide-number text, or
completely empty). May have images, diagrams, or group shapes that need
to be preserved via the visual-shape-preservation system in converter.py.

This handler acts as a container: it creates the branded slide shell and
relies on _preserve_visual_shapes() to carry across any visual content.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class BlankBrandedHandler(SlideHandler):

    name = "Blank Branded"
    description = "Blank or image-only slide with no meaningful text"
    layout_name = "Blank Branded"
    layout_index = 44

    PH_FOOTER = 10
    PH_SLIDE_NUM = 11

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code", "presentation title",
    ]

    FOOTER_PATTERNS = [
        r"(?i)cricos", r"(?i)hbis\s+innovation",
        r"(?i)executive\s+education",
        r"(?i)presentation\s+title",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with no meaningful text content.

        Returns moderate confidence for truly blank or image-only slides.
        Low priority — other handlers should claim the slide first if
        they see text they can use.
        """
        if slide_index == 0:
            return 0.0

        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        # Filter to meaningful text
        meaningful = self._get_meaningful_text(shapes)

        if not meaningful:
            # No text at all
            if images:
                return 0.55  # Image-only slide — good candidate
            # Completely empty slide
            return 0.40

        # Calculate total meaningful text length
        total_text = sum(len(s["text"]) for s in meaningful)

        if total_text < 15:
            # Very little text — might be a label on a diagram
            if images:
                return 0.45
            return 0.30

        # Too much text for a blank slide — another handler should take it
        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract minimal content. Visual shapes are handled separately
        by _preserve_visual_shapes() in converter.py.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "footer": "",
            "has_images": bool(extract_images(slide)),
        }

        # Try to find footer text
        for s in shapes:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (10, 17):
                    result["footer"] = s["text"]
                    break
            elif any(re.search(p, s["text"]) for p in self.FOOTER_PATTERNS):
                result["footer"] = s["text"]
                break

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill only the footer/slide-number placeholders.
        Visual content is preserved by _preserve_visual_shapes().
        NEVER set font properties.
        """
        # Nothing to fill — template provides branded footer/slide number
        # and visual shapes are copied separately by the converter
        pass

    # --- Helpers ---

    def _get_meaningful_text(self, shapes: list) -> list:
        """Filter out noise and footer text, return shapes with real content."""
        meaningful = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if len(text) <= 1:
                continue
            if re.match(r"^\d{1,3}$", text):
                continue
            if any(re.search(p, text) for p in self.FOOTER_PATTERNS):
                continue
            meaningful.append(s)
        return meaningful

    def get_placeholder_map(self) -> dict:
        return {
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
