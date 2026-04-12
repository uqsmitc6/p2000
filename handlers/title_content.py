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
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


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

        Strategy:
        1. Use source placeholder indices if available — idx 0/3/15 = title,
           idx 1/10 = content body. This is the most reliable signal.
        2. Fall back to font size / position heuristics for non-placeholder shapes.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "content": "",
            "footer": "",
        }

        if not shapes:
            return result

        # Filter out noise, bare slide numbers, and image license text
        filtered = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if re.match(r"^\d{1,3}$", text) and len(text) <= 3:
                continue
            if self._is_image_caption(text):
                continue
            filtered.append(s)

        if not filtered:
            return result

        # --- Separate by role using placeholder indices ---
        title_shape = None
        subtitle_shape = None
        body_shapes = []
        footer_text = ""

        for s in filtered:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    # TITLE / CENTER_TITLE
                    title_shape = s
                elif idx in (1, 10):
                    # BODY / CONTENT
                    body_shapes.append(s)
                elif idx in (17,):
                    # FOOTER
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18,):
                    # SLIDE NUMBER — skip
                    pass
                else:
                    # Unknown placeholder — treat as body
                    body_shapes.append(s)
            elif self._is_footer_text(s["text"]):
                if not footer_text:
                    footer_text = s["text"]
            else:
                # Non-placeholder shapes go to body
                body_shapes.append(s)

        result["footer"] = footer_text

        # --- Fallback title detection if no placeholder found ---
        if not title_shape and body_shapes:
            # Priority 1: Shape with largest font size
            with_font = [s for s in body_shapes if s["font_size"]]
            if with_font:
                candidate = max(with_font, key=lambda s: s["font_size"])
                # Only promote to title if it's reasonably short
                if len(candidate["text"]) < 120:
                    title_shape = candidate
                    body_shapes = [s for s in body_shapes if s is not title_shape]

            # Priority 2: Topmost short shape
            if not title_shape:
                sorted_by_top = sorted(body_shapes, key=lambda s: s["top"])
                for s in sorted_by_top:
                    if len(s["text"]) < 120:
                        title_shape = s
                        body_shapes = [bs for bs in body_shapes if bs is not s]
                        break

        if title_shape:
            result["title"] = title_shape["text"]

        # --- Identify subtitle ---
        # Sort body shapes by vertical position
        body_shapes.sort(key=lambda s: (s["top"], s["left"]))

        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more_content = len(body_shapes) > 1

            if is_short and is_near_top and has_more_content:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # --- Body content ---
        if body_shapes:
            result["content"] = "\n".join(s["text"] for s in body_shapes)

        return result

    def _is_image_caption(self, text: str) -> bool:
        """Check if text is an image license/caption that should be filtered."""
        caption_patterns = [
            r"(?i)image\s+licensed\s+through",
            r"(?i)adobe\s+stock[:\s]+\d",
            r"(?i)shutterstock[:\s]+\d",
            r"(?i)getty\s+images",
            r"(?i)©\s*\d{4}",
            r"(?i)source:\s*(http|www)",
        ]
        return any(re.search(p, text) for p in caption_patterns)

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

            # Override the master's level-1 default which is bold + accent1.
            # Body text should be regular weight.
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.font.bold = False

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
