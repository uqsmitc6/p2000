"""
Title Only Handler — Slides with just a title and possibly images/diagrams.

Template layout 43 "Title Only":
    idx  0 — Title
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: slide has a title but very little or no body text. The slide
may have images, diagrams, or charts that fill the content area. These
slides are distinguished from Section Dividers by not having the sparse,
structured section-number + title pattern.

Note: Images/charts cannot be transferred to the template automatically
(they'd need to be re-embedded). This handler preserves the title and
flags the slide as needing manual image re-insertion.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class TitleOnlyHandler(SlideHandler):

    name = "Title Only"
    description = "Slide with just a title (and possibly images that need manual placement)"
    layout_name = "Title Only"
    layout_index = 43

    PH_TITLE = 0
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
        Detect slides that have a title but minimal body text.
        These are typically diagram/image slides where the visual
        content can't be auto-extracted.
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if "title only" in layout_name:
            return 0.70

        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        if not shapes:
            # No text at all — might be a blank or image-only slide
            if images:
                return 0.30  # Has images but no text
            return 0.0

        # Filter noise
        meaningful = [
            s for s in shapes
            if s["text"].strip()
            and not any(p in s["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(s["text"].strip()) > 1
            and not re.match(r"^\d{1,3}$", s["text"].strip())
        ]

        if not meaningful:
            return 0.0

        # Separate title from body: title is likely the topmost short text
        sorted_by_top = sorted(meaningful, key=lambda s: s["top"] or 0)
        title_candidate = None
        body_text_total = 0

        for s in sorted_by_top:
            if title_candidate is None and len(s["text"]) < 120:
                title_candidate = s
            else:
                body_text_total += len(s["text"])

        has_images = bool(images)

        # Title only: has a title but very little body text
        if title_candidate and body_text_total < 30:
            if has_images:
                return 0.68  # Title + images, minimal text → strong Title Only
            elif len(meaningful) == 1:
                return 0.45  # Just a title, no images
            return 0.0

        # Also catch slides with 2 short shapes (title + subtitle-like)
        if len(meaningful) == 2 and has_images:
            lengths = sorted(len(s["text"]) for s in meaningful)
            if lengths[0] < 50 and lengths[1] < 120:
                return 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract just the title. Body content is intentionally empty
        since these slides typically have visual content that can't
        be auto-transferred.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "footer": "",
            "has_images": False,
        }

        # Check for images
        images = extract_images(slide)
        result["has_images"] = bool(images)

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

        # Find title
        title_shape = None
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
            elif not title_shape:
                # First non-footer text becomes title
                if not self._is_footer_text(s["text"]) and len(s["text"]) < 120:
                    title_shape = s

        result["footer"] = footer_text

        if title_shape:
            result["title"] = title_shape["text"]
        elif filtered:
            # Fallback: largest font
            with_font = [s for s in filtered if s["font_size"]]
            if with_font:
                result["title"] = max(with_font, key=lambda s: s["font_size"])["text"]

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
        """Fill Title Only placeholder. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
