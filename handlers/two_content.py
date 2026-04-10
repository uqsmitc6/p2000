"""
Two Content Handler — Slides with two side-by-side content areas.

Template layout 7 "Two Content":
    idx  0 — Title
    idx 31 — Subtitle (optional)
    idx 10 — Left content column
    idx 32 — Right content column
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: slide has two distinct text groups positioned side-by-side,
or has a clear two-column layout. Also catches slides where academics
have placed two text boxes next to each other manually.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class TwoContentHandler(SlideHandler):

    name = "Two Content"
    description = "Content slide with two side-by-side content columns"
    layout_name = "Two Content"
    layout_index = 7

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT = 10
    PH_RIGHT = 32
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with two-column content.

        Heuristics:
        - Source layout name contains "two" or "column"
        - Multiple body-sized shapes arranged side by side (left halves differ)
        - Two placeholder bodies at similar vertical position but different horizontal
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if "two" in layout_name and ("content" in layout_name or "column" in layout_name):
            # Exclude "two third" / "third" patterns — those are Split Content
            if "third" not in layout_name:
                return 0.75

        shapes = extract_shapes_with_text(slide)
        if not shapes:
            return 0.0

        # Filter noise
        meaningful = [
            s for s in shapes
            if s["text"].strip()
            and not any(p in s["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(s["text"].strip()) > 5
        ]

        if len(meaningful) < 3:  # Need at least title + 2 body areas
            return 0.0

        # Find body shapes (exclude title and title-like shapes)
        title_shape = None
        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                title_shape = s
                break
        if not title_shape:
            sorted_by_top = sorted(meaningful, key=lambda s: s["top"] or 0)
            for s in sorted_by_top:
                if len(s["text"]) < 120:
                    if s.get("font_size") and s["font_size"] >= 20:
                        title_shape = s
                        break
                    elif not s.get("font_size") and len(s["text"]) < 80:
                        title_shape = s
                        break

        body_candidates = [s for s in meaningful if s is not title_shape]

        if len(body_candidates) < 2:
            return 0.0

        # Check if any two body shapes are side by side
        # (similar top position, different left positions, each taking ~half width)
        slide_width = 12192000  # Standard 16:9 width in EMU
        half_width = slide_width / 2

        for i, a in enumerate(body_candidates):
            for b in body_candidates[i + 1:]:
                if a["left"] is None or b["left"] is None:
                    continue
                if a["top"] is None or b["top"] is None:
                    continue

                # Similar vertical position (within 15% of slide height)
                top_diff = abs(a["top"] - b["top"])
                if top_diff > 1000000:  # ~1 inch tolerance
                    continue

                # Different horizontal positions
                left_diff = abs(a["left"] - b["left"])
                if left_diff < half_width * 0.3:
                    continue  # Not separated enough

                # Both shapes must have substantial text (not just a label)
                if len(a["text"]) < 30 or len(b["text"]) < 30:
                    continue

                # Shapes should have roughly equal widths (within 50%)
                if a["width"] and b["width"]:
                    width_ratio = min(a["width"], b["width"]) / max(a["width"], b["width"])
                    if width_ratio < 0.5:
                        continue  # Too unequal → Split Content territory

                return 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and two columns of content.

        Strategy:
        1. Find the title (placeholder idx 0/3/15 or largest font).
        2. Split remaining body shapes into left and right columns
           based on horizontal position relative to slide midpoint.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left_content": "",
            "right_content": "",
            "footer": "",
        }

        if not shapes:
            return result

        # Filter noise
        filtered = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if re.match(r"^\d{1,3}$", text) and len(text) <= 3:
                continue
            filtered.append(s)

        if not filtered:
            return result

        # Separate by role
        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in filtered:
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

        # Fallback title detection
        if not title_shape and body_shapes:
            with_font = [s for s in body_shapes if s["font_size"]]
            if with_font:
                candidate = max(with_font, key=lambda s: s["font_size"])
                if len(candidate["text"]) < 120:
                    title_shape = candidate
                    body_shapes = [s for s in body_shapes if s is not title_shape]

        if title_shape:
            result["title"] = title_shape["text"]

        # Subtitle detection
        body_shapes.sort(key=lambda s: (s["top"], s["left"]))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more = len(body_shapes) > 1
            if is_short and is_near_top and has_more:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Split remaining body shapes into left/right columns
        if body_shapes:
            slide_midpoint = 6096000  # Half of standard 16:9 width in EMU

            left_shapes = []
            right_shapes = []

            for s in body_shapes:
                # Use the centre of the shape to determine column
                shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
                if shape_centre < slide_midpoint:
                    left_shapes.append(s)
                else:
                    right_shapes.append(s)

            # Sort each column by vertical position
            left_shapes.sort(key=lambda s: (s["top"], s["left"]))
            right_shapes.sort(key=lambda s: (s["top"], s["left"]))

            result["left_content"] = "\n".join(s["text"] for s in left_shapes)
            result["right_content"] = "\n".join(s["text"] for s in right_shapes)

            # If everything ended up on one side, split evenly
            if result["left_content"] and not result["right_content"]:
                all_text = [s["text"] for s in left_shapes]
                mid = len(all_text) // 2
                result["left_content"] = "\n".join(all_text[:mid]) if mid > 0 else all_text[0]
                result["right_content"] = "\n".join(all_text[mid:]) if mid > 0 else ""

        return result

    def _is_footer_text(self, text: str) -> bool:
        footer_patterns = [
            r"(?i)cricos", r"(?i)hbis\s+innovation",
            r"(?i)executive\s+education",
            r"(?i)presentation\s+title",
        ]
        return any(re.search(p, text) for p in footer_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill Two Content placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left_content") and self.PH_LEFT in placeholders:
            placeholders[self.PH_LEFT].text = content["left_content"]

        if content.get("right_content") and self.PH_RIGHT in placeholders:
            placeholders[self.PH_RIGHT].text = content["right_content"]

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "left": self.PH_LEFT,
            "right": self.PH_RIGHT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
