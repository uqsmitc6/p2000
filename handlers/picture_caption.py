"""
Picture with Caption Handler — Image-focused slides with title and caption.

Template layout 18 "Picture with Caption":
    idx  0 — Title (w=8928943)
    idx 31 — Subtitle body (w=8928943, h=506132)
    idx 32 — Picture placeholder (w=11233149, h=4125914)
    idx  2 — Caption text (w=11200605, h=360000)
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: slide has a primary image AND either a title or short descriptive
text, but NOT substantial body content (that would go to Text with Image).
Typical examples: photo slides, diagram showcases, figure plates.

Distinguished from Text with Image by:
    - Caption is SHORT (<150 chars) vs body text (>150 chars)
    - Visual focus is the image, not the text
    - No bullet lists or structured content

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class PictureCaptionHandler(SlideHandler):

    name = "Picture with Caption"
    description = "Image-focused slide with title and short caption text"
    layout_name = "Picture with Caption"
    layout_index = 18

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_PICTURE = 32
    PH_CAPTION = 2
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

    # Max caption length — beyond this it's body text, not a caption
    MAX_CAPTION_CHARS = 150

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
        Detect slides with a prominent image and short descriptive text.

        Higher priority than Title Only (which also handles images) but
        lower priority than Text with Image (which handles text-heavy
        slides with images).
        """
        if slide_index == 0:
            return 0.0

        images = extract_images(slide)
        if not images:
            return 0.0

        # Need at least one substantial image (not just a logo)
        slide_width = 12192000   # Standard 13.33in slide
        slide_height = 6858000   # Standard 7.5in slide
        real_images = [
            img for img in images
            if img.get("width") and img.get("height")
            and img["width"] > slide_width * 0.15
            and img["height"] > slide_height * 0.15
        ]

        if not real_images:
            return 0.0

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_text(shapes)

        if not meaningful:
            # Image but no text — Blank Branded is better
            return 0.0

        # Separate title from other text
        title, other_texts = self._classify_text(meaningful)

        # Calculate total non-title text
        total_other = sum(len(t) for t in other_texts)

        if total_other > 300:
            # Too much text — Text with Image or Title & Content is better
            return 0.0

        if title and total_other <= self.MAX_CAPTION_CHARS:
            # Title + short caption + image — ideal for this handler
            return 0.62

        if title and total_other == 0:
            # Title + image, no caption — still works (caption optional)
            return 0.58

        if not title and total_other <= self.MAX_CAPTION_CHARS:
            # Just a caption with image, no clear title
            return 0.48

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, caption text, and primary image.
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "subtitle": "",
            "caption": "",
            "footer": "",
            "image_blob": None,
            "image_content_type": "image/png",
        }

        # Find the primary (largest) image
        images_with_blobs = [img for img in images if img.get("blob")]
        if images_with_blobs:
            primary = max(
                images_with_blobs,
                key=lambda i: (i.get("width") or 0) * (i.get("height") or 0),
            )
            result["image_blob"] = primary["blob"]
            result["image_content_type"] = primary.get("content_type") or "image/png"

        # Extract text
        meaningful = self._get_meaningful_text(shapes)

        # Find footer
        for s in shapes:
            if s["is_placeholder"] and s["shape_idx"] in (17, 10):
                result["footer"] = s["text"]
                break
            elif any(re.search(p, s["text"]) for p in self.FOOTER_PATTERNS):
                result["footer"] = s["text"]
                break

        if not meaningful:
            return result

        # Classify text into title + other
        title, other_texts = self._classify_text(meaningful)

        if title:
            result["title"] = title

        # If we have 2+ other texts, first short one is subtitle, rest is caption
        if len(other_texts) >= 2:
            first = other_texts[0]
            rest = other_texts[1:]
            if len(first) < 80:
                result["subtitle"] = first
                result["caption"] = "\n".join(rest)
            else:
                result["caption"] = "\n".join(other_texts)
        elif len(other_texts) == 1:
            text = other_texts[0]
            if len(text) < 80 and not result["title"]:
                # Short text with no title — use as title instead
                result["title"] = text
            elif len(text) < 80:
                # Short text — could be subtitle or caption
                result["subtitle"] = text
            else:
                result["caption"] = text

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, subtitle, picture, and caption placeholders.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("caption") and self.PH_CAPTION in placeholders:
            placeholders[self.PH_CAPTION].text = content["caption"]

        # Insert image into picture placeholder
        if content.get("image_blob") and self.PH_PICTURE in placeholders:
            image_stream = io.BytesIO(content["image_blob"])
            placeholders[self.PH_PICTURE].insert_picture(image_stream)

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

    def _classify_text(self, shapes: list) -> tuple[str, list[str]]:
        """
        Separate title from other text.

        Returns:
            (title_text, [other_texts])
        """
        title = None
        others = []

        # Check for placeholder-based title first
        for s in shapes:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                title = s["text"]
                break

        # Fallback: topmost short text
        if title is None:
            sorted_shapes = sorted(shapes, key=lambda s: s["top"] or 0)
            for s in sorted_shapes:
                if len(s["text"]) < 120:
                    title = s["text"]
                    break

        # Collect non-title texts
        for s in shapes:
            if s["text"] == title:
                continue
            others.append(s["text"])

        return title or "", others

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "picture": self.PH_PICTURE,
            "caption": self.PH_CAPTION,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
