"""
Picture with Pullout Handler — Full-bleed image with overlaid pullout text.

Template layout 17 "Picture with Pullout":
    idx 20 — Background picture (full slide, 13.33in × 6.98in)
    idx  0 — Title (left overlay, w=3.62in, h=1.36in)
    idx 31 — Pullout body text (left overlay, w=3.62in, h=2.77in)
    idx 10 — Secondary picture placeholder (off-screen at top=8.95in — not used)
    idx 17 — Footer
    idx 18 — Slide number

Detection: slide has a full-bleed or near-full-bleed image with overlaid
text. The key distinguishing feature is a large background image with
a small text area positioned on top of it — not beside it.

Distinguished from Text with Image by the overlay arrangement (text ON
the image) vs side-by-side (text BESIDE the image).

Content to look for:
    - Slides where academics have placed a full-width photo with a text
      box overlay (common for opening/impact slides)
    - Slides with a background image + a callout or key statistic
    - Source layout named "Picture with Pullout" or "Full bleed" or similar

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class PicturePulloutHandler(SlideHandler):

    name = "Picture with Pullout"
    description = "Full-bleed image with overlaid title and pullout text"
    layout_name = "Picture with Pullout"
    layout_index = 17

    PH_BG_PICTURE = 20  # Full-slide background image
    PH_TITLE = 0
    PH_BODY = 31        # Pullout/overlay text
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
        Detect slides with a full-bleed image and overlaid text.

        Content patterns to look for:
        - A very large image (>70% of slide in both dimensions) — the background
        - Small amount of text overlaid on top of it
        - NOT just an image-only slide (needs at least a title)
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "pullout" in layout_name:
            return 0.75
        if "full bleed" in layout_name or "fullbleed" in layout_name:
            return 0.70

        images = extract_images(slide)
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if not images:
            return 0.0

        # Look for a full-bleed or near-full-bleed image
        bg_images = [
            img for img in images
            if img.get("width") and img.get("height")
            and img["width"] > self.SLIDE_WIDTH * 0.7
            and img["height"] > self.SLIDE_HEIGHT * 0.5
        ]

        if not bg_images:
            return 0.0

        if not meaningful:
            return 0.0  # Image but no text — Blank Branded is better

        # For pullout: should have some text but NOT a lot
        total_text = sum(len(s["text"]) for s in meaningful)

        if total_text > 500:
            # Too much text — this is a content slide with a background, not a pullout
            return 0.0

        if total_text < 10:
            return 0.0  # Trivial text

        # Check if text shapes are positioned in a small area (overlay)
        # rather than spread across the full slide
        text_lefts = [(s["left"] or 0) for s in meaningful]
        text_widths = [(s["width"] or 0) for s in meaningful]

        # Most text should be in a narrow column (<50% of slide width)
        narrow_text = all(w < self.SLIDE_WIDTH * 0.5 for w in text_widths if w > 0)

        if narrow_text and len(meaningful) >= 1 and total_text >= 10:
            return 0.55

        # Wider text but still with a big background image
        if len(meaningful) <= 3 and total_text < 300:
            return 0.42

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, pullout body text, and background image.
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "body": "",
            "footer": "",
            "image_blob": None,
        }

        # Background image — largest by area
        images_with_blobs = [img for img in images if img.get("blob")]
        if images_with_blobs:
            primary = max(
                images_with_blobs,
                key=lambda i: (i.get("width") or 0) * (i.get("height") or 0),
            )
            result["image_blob"] = primary["blob"]

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
            sorted_by_top = sorted(body_shapes, key=lambda s: (s["top"] or 0))
            for s in sorted_by_top:
                if len(s["text"]) < 100:
                    title_shape = s
                    body_shapes = [bs for bs in body_shapes if bs is not s]
                    break

        if title_shape:
            result["title"] = title_shape["text"]

        # Body/pullout text
        body_shapes.sort(key=lambda s: (s["top"] or 0))
        result["body"] = "\n".join(s["text"] for s in body_shapes)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill pullout placeholders and insert background image.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("body") and self.PH_BODY in placeholders:
            placeholders[self.PH_BODY].text = content["body"]

        # Background image
        if content.get("image_blob") and self.PH_BG_PICTURE in placeholders:
            image_stream = io.BytesIO(content["image_blob"])
            placeholders[self.PH_BG_PICTURE].insert_picture(image_stream)

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

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "body": self.PH_BODY,
            "bg_picture": self.PH_BG_PICTURE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
