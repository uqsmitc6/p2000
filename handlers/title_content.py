"""
Title and Content Handler — The standard content slide.

Template placeholders (layout index 6, "Title and Content"):
    idx  0 — Title
    idx 31 — Subtitle (optional — if absent, content shifts up)
    idx 10 — Content body (full width, 12.28" × 4.51")
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: This is the default/fallback handler. Any slide that isn't
a cover, section divider, thank you, or contents page is likely a
content slide. Most slides in any deck will be this type.

Design concerns (future):
    - Wall-of-text risk: content body is huge (12.28" × 4.51").
      Future versions should enforce visual communication rules
      (max bullet points, text density limits, suggest splitting).
    - Subtitle is optional: when absent, content shifts up 0.55"
      to reclaim the dead space.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images


class TitleContentHandler(SlideHandler):

    name = "Title and Content"
    description = "Standard content slide with title, optional subtitle, and body"
    layout_name = "Title and Content"
    layout_index = 6

    # Placeholder indices
    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_CONTENT = 10
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

    # Subtitle placeholder height — used to shift content up when no subtitle
    SUBTITLE_HEIGHT_EMU = 502920  # ~0.55 inches

    # Noise to filter
    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Title and Content is the default handler — it matches any slide
        that has a title and body content. It should have LOW priority
        so more specific handlers (cover, divider, thank you) win first.
        """
        # Never match first slide (that's the cover)
        if slide_index == 0:
            return 0.0

        texts = extract_text_elements(slide)
        images = extract_images(slide)

        if not texts:
            return 0.0

        # Filter noise
        meaningful = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 1
        ]

        if not meaningful:
            return 0.0

        # Basic heuristic: has at least one text element that looks like
        # a title (shorter text) and at least some body content
        has_short_text = any(len(t["text"]) < 80 for t in meaningful)
        has_body_text = any(len(t["text"]) > 30 for t in meaningful)

        # Content slides typically have more text than dividers
        total_text_len = sum(len(t["text"]) for t in meaningful)

        if has_short_text and has_body_text and total_text_len > 50:
            return 0.4  # Low confidence — acts as fallback

        if len(meaningful) >= 2:
            return 0.35

        return 0.1

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, optional subtitle, and body content from a slide.
        """
        texts = extract_text_elements(slide)

        # Expand and filter
        expanded = []
        for t in texts:
            lines = t["text"].split("\n") if "\n" in t["text"] else [t["text"]]
            for line in lines:
                line = line.strip()
                if line:
                    expanded.append({**t, "text": line})

        filtered = [
            t for t in expanded
            if not any(p in t["text"].lower().strip() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 0
        ]

        # Skip bare slide numbers
        filtered = [
            t for t in filtered
            if not (re.match(r"^\d{1,3}$", t["text"].strip()) and len(t["text"].strip()) <= 3)
        ]

        result = {
            "title": "",
            "subtitle": "",
            "content": "",
            "footer": "",
        }

        # Separate footer-like text
        content_texts = []
        for t in filtered:
            if self._is_footer_text(t["text"]):
                if not result["footer"]:
                    result["footer"] = t["text"]
            else:
                content_texts.append(t)

        if not content_texts:
            return result

        # Classification strategy:
        # 1. Group texts by their shape position (top coordinate)
        # 2. The topmost short text → title
        # 3. If there's a second short text near the top → subtitle
        # 4. Everything else → body content

        # Sort by vertical position
        content_texts.sort(key=lambda t: (t["top"], t["left"]))

        # Title: largest font, or topmost text if fonts are similar
        by_font = sorted(content_texts, key=lambda t: (t["font_size"] or 0), reverse=True)
        result["title"] = by_font[0]["text"]

        # Remaining texts
        remaining = [t for t in content_texts if t["text"] != result["title"]]

        # Check if the second text looks like a subtitle
        # (shorter, near the top, different from body content)
        if remaining:
            candidate = remaining[0]
            # Subtitle heuristic: short, near top, and there's more content below
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (slide.slide_height / 3) if hasattr(slide, 'slide_height') else True
            has_more_content = len(remaining) > 1

            if is_short and has_more_content:
                result["subtitle"] = candidate["text"]
                remaining = remaining[1:]

        # Body content: join remaining texts
        if remaining:
            result["content"] = "\n".join(t["text"] for t in remaining)

        return result

    def _is_footer_text(self, text: str) -> bool:
        """Check if text looks like a footer/programme name."""
        footer_patterns = [
            r"(?i)cricos", r"(?i)hbis\s+innovation",
            r"(?i)executive\s+education",
            r"(?i)presentation\s+title",
        ]
        return any(re.search(p, text) for p in footer_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill Title and Content placeholders.
        If no subtitle, shifts content body up to reclaim the space.
        """
        from pptx.util import Emu

        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Subtitle
        has_subtitle = bool(content.get("subtitle"))
        if has_subtitle and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]
        elif not has_subtitle and self.PH_SUBTITLE in placeholders:
            # Clear subtitle and shift content up
            placeholders[self.PH_SUBTITLE].text = ""

        # Content body
        if content.get("content") and self.PH_CONTENT in placeholders:
            ph = placeholders[self.PH_CONTENT]

            # Shift content up if no subtitle.
            # IMPORTANT: preserve inherited left/width when modifying position.
            if not has_subtitle:
                orig_left = ph.left
                orig_width = ph.width
                new_top = ph.top - self.SUBTITLE_HEIGHT_EMU
                new_height = ph.height + self.SUBTITLE_HEIGHT_EMU
                ph.top = new_top
                ph.left = orig_left
                ph.width = orig_width
                ph.height = new_height

            ph.text = content["content"]

        # Footer
        if content.get("footer") and self.PH_FOOTER in placeholders:
            placeholders[self.PH_FOOTER].text = content["footer"]

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "content": self.PH_CONTENT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
