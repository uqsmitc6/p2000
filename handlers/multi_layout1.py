"""
Multi-layout 1 Handler — Four-panel staggered layout (Z-pattern).

Template layout 36 "Multi-layout 1":
    idx  0 — Title (full width, top)
    idx 34 — Subtitle (full width)
    idx 31 — Top-left block (OBJECT, 5.92in × 1.63in)
    idx 28 — Bottom-left block (BODY, 5.94in × 2.65in)
    idx 36 — Top-right block (BODY, 5.94in × 2.65in)
    idx 35 — Bottom-right block (OBJECT, 5.92in × 1.63in)
    idx 25 — Footer
    idx 26 — Slide number

The layout creates a Z-pattern: small block top-left, large block
top-right, large block bottom-left, small block bottom-right.

Detection: slide with 4 content areas arranged in a grid. Distinguished
from Two Content by having 4 areas rather than 2, and from Three Content
by having a 2×2 arrangement.

Content to look for:
    - SWOT analysis or similar 2×2 matrix
    - Four-quadrant comparison slides
    - Dashboard-style multi-panel information
    - Pros/cons with supporting detail per item

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class MultiLayout1Handler(SlideHandler):

    name = "Multi-layout 1"
    description = "Four-panel staggered layout (Z-pattern)"
    layout_name = "Multi-layout 1"
    layout_index = 36

    PH_TITLE = 0
    PH_SUBTITLE = 34
    PH_TOP_LEFT = 31
    PH_BOTTOM_LEFT = 28
    PH_TOP_RIGHT = 36
    PH_BOTTOM_RIGHT = 35
    PH_FOOTER = 25
    PH_SLIDE_NUM = 26

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

    SLIDE_WIDTH = 12192000
    SLIDE_HEIGHT = 6858000
    SLIDE_MID_X = SLIDE_WIDTH // 2
    SLIDE_MID_Y = SLIDE_HEIGHT // 2

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with four content quadrants.

        Content patterns:
        - Source layout name contains "multi-layout 1" or "multi layout 1"
        - 4 substantial text blocks arranged in a 2×2 grid
        - Each quadrant has at least 15 chars
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "multi-layout 1" in layout_name or "multi layout 1" in layout_name:
            return 0.75

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 5:  # title + 4 areas
            return 0.0

        # Find title
        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 4:
            return 0.0

        # Check 2×2 grid distribution
        tl = [s for s in body_shapes
              if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID_X
              and (s["top"] or 0) + (s["height"] or 0) / 2 < self.SLIDE_MID_Y]
        tr = [s for s in body_shapes
              if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID_X
              and (s["top"] or 0) + (s["height"] or 0) / 2 < self.SLIDE_MID_Y]
        bl = [s for s in body_shapes
              if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID_X
              and (s["top"] or 0) + (s["height"] or 0) / 2 >= self.SLIDE_MID_Y]
        br = [s for s in body_shapes
              if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID_X
              and (s["top"] or 0) + (s["height"] or 0) / 2 >= self.SLIDE_MID_Y]

        filled = sum(1 for q in [tl, tr, bl, br] if q)
        if filled == 4:
            # All four quadrants have content
            min_text = min(
                sum(len(s["text"]) for s in q)
                for q in [tl, tr, bl, br]
            )
            if min_text >= 15:
                return 0.55
            elif min_text >= 5:
                return 0.45

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """Extract title, subtitle, and four quadrant text areas."""
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "top_left": "",
            "bottom_left": "",
            "top_right": "",
            "bottom_right": "",
            "footer": "",
        }

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)
        if not meaningful:
            return result

        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in meaningful:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 25):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 26):
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
            if len(candidate["text"]) < 100 and len(body_shapes) > 1:
                near_top = (candidate["top"] or 0) < (title_shape["top"] or 0) + (title_shape["height"] or 0) * 2
                if near_top:
                    result["subtitle"] = candidate["text"]
                    body_shapes = body_shapes[1:]

        # Split into quadrants
        tl, tr, bl, br = [], [], [], []
        for s in body_shapes:
            cx = (s["left"] or 0) + (s["width"] or 0) / 2
            cy = (s["top"] or 0) + (s["height"] or 0) / 2
            if cx < self.SLIDE_MID_X:
                if cy < self.SLIDE_MID_Y:
                    tl.append(s)
                else:
                    bl.append(s)
            else:
                if cy < self.SLIDE_MID_Y:
                    tr.append(s)
                else:
                    br.append(s)

        for group in [tl, tr, bl, br]:
            group.sort(key=lambda s: (s["top"] or 0))

        result["top_left"] = "\n".join(s["text"] for s in tl)
        result["bottom_left"] = "\n".join(s["text"] for s in bl)
        result["top_right"] = "\n".join(s["text"] for s in tr)
        result["bottom_right"] = "\n".join(s["text"] for s in br)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill four-panel placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("top_left") and self.PH_TOP_LEFT in placeholders:
            placeholders[self.PH_TOP_LEFT].text = content["top_left"]

        if content.get("bottom_left") and self.PH_BOTTOM_LEFT in placeholders:
            placeholders[self.PH_BOTTOM_LEFT].text = content["bottom_left"]

        if content.get("top_right") and self.PH_TOP_RIGHT in placeholders:
            placeholders[self.PH_TOP_RIGHT].text = content["top_right"]

        if content.get("bottom_right") and self.PH_BOTTOM_RIGHT in placeholders:
            placeholders[self.PH_BOTTOM_RIGHT].text = content["bottom_right"]

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
            "top_left": self.PH_TOP_LEFT,
            "bottom_left": self.PH_BOTTOM_LEFT,
            "top_right": self.PH_TOP_RIGHT,
            "bottom_right": self.PH_BOTTOM_RIGHT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
