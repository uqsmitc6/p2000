"""
Icons + Two Contents Handler — Two main content columns with icon sidebar.

Template layout 12 "Icons + two contents":
    idx  0 — Title (top left, w=8.08in)
    idx 31 — Subtitle (full width, h=0.55in)
    idx 40 — Left content (large, 3.92in × 4.51in)
    idx 41 — Right content (large, 3.81in × 4.51in)
    idx 12 — Icon 1 (tiny 0.33in × 0.34in, right sidebar top)
    idx 10 — Icon 1 text (2.76in × 1.10in)
    idx 37 — Icon 2 (tiny 0.33in × 0.34in)
    idx 34 — Icon 2 text (2.76in × 1.10in)
    idx 38 — Icon 3 (tiny 0.33in × 0.34in)
    idx 35 — Icon 3 text (2.76in × 1.10in)
    idx 39 — Icon 4 (tiny 0.33in × 0.34in)
    idx 36 — Icon 4 text (2.76in × 1.10in)
    idx 17 — Footer
    idx 18 — Slide number

The icon placeholders (12, 37, 38, 39) are tiny squares for small
icon images. We leave these empty (they'll show the template default
or be cleaned up). The text placeholders beside them hold short
descriptions.

Detection: slide with two main content areas plus a sidebar of
short key points or stats. Distinguished from plain Two Content
by the presence of a sidebar with 3+ short text items.

Content to look for:
    - Two-column content with sidebar annotations
    - Main content with key metrics/stats sidebar
    - Feature comparison with highlights
    - Course content with learning outcomes sidebar

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class IconsTwoContentsHandler(SlideHandler):

    name = "Icons + Two Contents"
    description = "Two main content columns with icon sidebar"
    layout_name = "Icons + two contents"
    layout_index = 12

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT = 40
    PH_RIGHT = 41
    # Icon placeholders (tiny, for images — left empty)
    PH_ICONS = [12, 37, 38, 39]
    # Icon text placeholders
    PH_ICON_TEXTS = [10, 34, 35, 36]
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

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

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with two content columns plus a sidebar.

        Content patterns:
        - Source layout name contains "icons" and "two content"
        - Two substantial content areas plus 3+ short sidebar items
        - Sidebar items cluster on the right edge of the slide
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "icon" in layout_name and "two content" in layout_name:
            return 0.75

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 5:
            return 0.0

        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 5:
            return 0.0

        # Look for sidebar pattern: 3+ short texts clustered on the right
        right_edge = self.SLIDE_WIDTH * 0.7
        sidebar_shapes = [
            s for s in body_shapes
            if (s["left"] or 0) >= right_edge and len(s["text"]) < 100
        ]
        main_shapes = [s for s in body_shapes if s not in sidebar_shapes]

        if len(sidebar_shapes) >= 3 and len(main_shapes) >= 2:
            # Check main shapes span left/centre
            main_text = sum(len(s["text"]) for s in main_shapes)
            if main_text >= 100:
                return 0.50

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, two main content columns, and up to
        4 sidebar text items.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left": "",
            "right": "",
            "sidebar": [],  # Up to 4 short text items
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
                elif idx in (17,):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18,):
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

        # Separate sidebar (short texts on far right) from main content
        right_edge = self.SLIDE_WIDTH * 0.7
        sidebar_shapes = sorted(
            [s for s in body_shapes
             if (s["left"] or 0) >= right_edge and len(s["text"]) < 100],
            key=lambda s: (s["top"] or 0),
        )
        main_shapes = sorted(
            [s for s in body_shapes if s not in sidebar_shapes],
            key=lambda s: (s["top"] or 0, s["left"] or 0),
        )

        result["sidebar"] = [s["text"] for s in sidebar_shapes[:4]]

        # Split main content into left/right columns
        mid_x = self.SLIDE_WIDTH * 0.5
        left_texts = sorted(
            [s for s in main_shapes
             if (s["left"] or 0) + (s["width"] or 0) / 2 < mid_x],
            key=lambda s: (s["top"] or 0),
        )
        right_texts = sorted(
            [s for s in main_shapes
             if (s["left"] or 0) + (s["width"] or 0) / 2 >= mid_x],
            key=lambda s: (s["top"] or 0),
        )

        result["left"] = "\n".join(s["text"] for s in left_texts)
        result["right"] = "\n".join(s["text"] for s in right_texts)

        # If all text ended up on one side, redistribute
        if not result["left"] and result["right"]:
            all_main = [s["text"] for s in main_shapes]
            mid = max(1, len(all_main) // 2)
            result["left"] = "\n".join(all_main[:mid])
            result["right"] = "\n".join(all_main[mid:])
        elif result["left"] and not result["right"]:
            all_main = [s["text"] for s in main_shapes]
            mid = max(1, len(all_main) // 2)
            result["left"] = "\n".join(all_main[:mid])
            result["right"] = "\n".join(all_main[mid:])

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, subtitle, two content columns, and sidebar texts.
        Icon placeholders are left empty. NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left") and self.PH_LEFT in placeholders:
            placeholders[self.PH_LEFT].text = content["left"]

        if content.get("right") and self.PH_RIGHT in placeholders:
            placeholders[self.PH_RIGHT].text = content["right"]

        # Fill sidebar text placeholders
        sidebar = content.get("sidebar", [])
        for i, ph_idx in enumerate(self.PH_ICON_TEXTS):
            if i >= len(sidebar):
                break
            if sidebar[i] and ph_idx in placeholders:
                placeholders[ph_idx].text = sidebar[i]

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
            "left": self.PH_LEFT,
            "right": self.PH_RIGHT,
            "icons": self.PH_ICONS,
            "icon_texts": self.PH_ICON_TEXTS,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
