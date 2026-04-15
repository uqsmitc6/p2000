"""
Text with Block Handler — Text content with coloured text panel.

Template layouts (three colour variants):
    idx 32 — Text with Dark Purple Block
    idx 33 — Text with Neutral Block
    idx 34 — Text with Grey Block

Structure is nearly identical to Graph with Block (29-31), but:
    - Left content placeholder idx varies per variant (1, 32, 32)
    - Block text panel (idx 14) is slightly smaller (5.35in × 4.73in vs 5.85in × 5.59in)
    - These are text-focused (no graph expectation)

Placeholder mapping:
    idx  0 — Title (left side)
    idx 31 — Subtitle (left side)
    idx  1 — Content (left, Dark Purple variant only)
    idx 32 — Content (left, Neutral and Grey variants)
    idx 14 — Block text (right side, coloured panel)
    idx 19 — Footer
    idx 20 — Slide number

Detection: slide has two distinct text areas side by side, or source
layout name contains "text" and "block". Distinguished from Two Content
by the presence of a coloured panel indicator and from Graph with Block
by the absence of graph/chart content.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


# Colour variant configs: (layout_index, content_ph_idx)
TEXT_BLOCK_DARK_PURPLE = (32, 1)
TEXT_BLOCK_NEUTRAL = (33, 32)
TEXT_BLOCK_GREY = (34, 32)

DEFAULT_VARIANT = TEXT_BLOCK_DARK_PURPLE


class TextBlockHandler(SlideHandler):

    name = "Text with Block"
    description = "Text content with coloured text block panel"
    layout_name = "Text with Dark Purple Block"
    layout_index = 32  # Default, overridden by get_layout_index()

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_BLOCK_TEXT = 14   # Coloured text panel (right)
    PH_FOOTER = 19
    PH_SLIDE_NUM = 20

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    FOOTER_PATTERNS = [
        r"(?i)cricos", r"(?i)hbis\s+innovation",
        r"(?i)executive\s+education",
        r"(?i)presentation\s+title",
    ]

    SLIDE_MID = 12192000 // 2

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides suitable for text-with-block layout.

        Triggers:
        - Source layout name contains "text" and "block" (but not "graph")
        - Two substantial text areas side by side
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "text" in layout_name and "block" in layout_name and "graph" not in layout_name:
            return 0.75

        # Heuristic: two substantial text blocks on opposite sides
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 3:  # title + 2 text areas
            return 0.0

        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 2:
            return 0.0

        # Look for two substantial text blocks on opposite sides
        left_text = sum(
            len(s["text"]) for s in body_shapes
            if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID
        )
        right_text = sum(
            len(s["text"]) for s in body_shapes
            if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID
        )

        if left_text >= 50 and right_text >= 100:
            return 0.48  # Lower than Graph with Block to avoid conflicts
        if left_text >= 100 and right_text >= 50:
            return 0.48

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, left content, and right block text.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "content": "",
            "block_text": "",
            "footer": "",
            "_variant": DEFAULT_VARIANT,
        }

        # Determine colour variant
        layout_name = slide.slide_layout.name.lower()
        if "neutral" in layout_name:
            result["_variant"] = TEXT_BLOCK_NEUTRAL
        elif "grey" in layout_name or "gray" in layout_name:
            result["_variant"] = TEXT_BLOCK_GREY
        else:
            result["_variant"] = TEXT_BLOCK_DARK_PURPLE

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)
        if not meaningful:
            return result

        # Separate by role
        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in meaningful:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 19):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 20):
                    pass
                else:
                    body_shapes.append(s)
            elif self._is_footer_text(s["text"]):
                if not footer_text:
                    footer_text = s["text"]
            else:
                body_shapes.append(s)

        result["footer"] = footer_text

        # Fallback title
        if not title_shape and body_shapes:
            with_font = [s for s in body_shapes if s["font_size"]]
            if with_font:
                candidate = max(with_font, key=lambda s: s["font_size"])
                if len(candidate["text"]) < 120:
                    title_shape = candidate
                    body_shapes = [s for s in body_shapes if s is not title_shape]

        if title_shape:
            result["title"] = title_shape["text"]

        # Subtitle
        body_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more = len(body_shapes) > 1
            if is_short and is_near_top and has_more:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Split into left (content) and right (block text)
        left_texts = []
        right_texts = []

        for s in body_shapes:
            shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
            if shape_centre < self.SLIDE_MID:
                left_texts.append(s["text"])
            else:
                right_texts.append(s["text"])

        result["content"] = "\n".join(left_texts)
        result["block_text"] = "\n".join(right_texts)

        # If all text ended up on one side, use longest for block
        if result["content"] and not result["block_text"]:
            all_shapes = sorted(body_shapes, key=lambda s: len(s["text"]), reverse=True)
            if len(all_shapes) >= 2:
                result["block_text"] = all_shapes[0]["text"]
                result["content"] = "\n".join(s["text"] for s in all_shapes[1:])
            elif all_shapes:
                result["block_text"] = all_shapes[0]["text"]
                result["content"] = ""

        return result

    # --- Dynamic layout ---

    def get_layout_index(self, content: dict) -> int:
        """Return the appropriate colour variant layout index."""
        variant = content.get("_variant", DEFAULT_VARIANT)
        return variant[0]

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill text-with-block placeholders.
        Content placeholder index varies per colour variant.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        # Content — index varies by variant
        variant = content.get("_variant", DEFAULT_VARIANT)
        content_idx = variant[1]
        if content.get("content"):
            if content_idx in placeholders:
                placeholders[content_idx].text = content["content"]
            else:
                # Fallback: try all known content indices
                placed = False
                for fallback_idx in (1, 32, 10, 13):
                    if fallback_idx in placeholders and fallback_idx != content_idx:
                        placeholders[fallback_idx].text = content["content"]
                        placed = True
                        break
                if not placed:
                    import logging
                    logging.getLogger("uqslide.text_block").error(
                        "CONTENT LOSS: %d chars could not be placed — "
                        "placeholder %d not found. Available: %s",
                        len(content["content"]), content_idx,
                        list(placeholders.keys()),
                    )

        if content.get("block_text") and self.PH_BLOCK_TEXT in placeholders:
            placeholders[self.PH_BLOCK_TEXT].text = content["block_text"]

    # --- Helpers ---

    def _get_meaningful_shapes(self, shapes: list) -> list:
        meaningful = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if len(text) <= 1:
                continue
            if re.match(r"^\d{1,3}$", text) and len(text) <= 3:
                continue
            meaningful.append(s)
        return meaningful

    def _find_title(self, shapes: list):
        for s in shapes:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                return s
        sorted_by_top = sorted(shapes, key=lambda s: s["top"] or 0)
        for s in sorted_by_top:
            if len(s["text"]) < 120:
                if s.get("font_size") and s["font_size"] >= 20:
                    return s
                elif not s.get("font_size") and len(s["text"]) < 80:
                    return s
        return None

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "content": "varies (1 or 32)",
            "block_text": self.PH_BLOCK_TEXT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
