"""
Quote Handler — Quote slides with a quote and attribution.

Template layouts 38 "Quote 1" and 39 "Quote 2":
    idx 10 — Background picture (full slide)
    idx 16 — Quote text
    idx 17 — Attribution (author/source)
    idx 18 — Footer (programme name)
    idx 19 — Slide number

Both layouts are structurally identical (same placeholders), just
different background styling. We default to Quote 1 (idx 38).

Detection: slide has a short, prominent quote (often in quotation marks)
with an attribution line. May have "said", a dash before the author name,
or explicit quotation marks.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class QuoteHandler(SlideHandler):

    name = "Quote"
    description = "A quote slide with quoted text and attribution"
    layout_name = "Quote 1"
    layout_index = 38

    PH_PICTURE = 10
    PH_QUOTE = 16
    PH_ATTRIBUTION = 17
    PH_FOOTER = 18
    PH_SLIDE_NUM = 19

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # Patterns that suggest a WRAPPED quote (text starts AND ends with quote marks)
    WRAPPED_QUOTE_PATTERN = re.compile(
        r'^["\u201c\u2018\u00ab].*["\u201d\u2019\u00bb]$', re.DOTALL
    )

    # Patterns for INLINE quotes: lead-in text with embedded quote marks
    # e.g., 'Customer experience is a "customer\'s journey..." (Author, Year)'
    INLINE_QUOTE_PATTERN = re.compile(
        r'["\u201c\u2018\u00ab].{20,}["\u201d\u2019\u00bb]', re.DOTALL
    )

    # Citation at end: (Author Year) or (Author Year, page)
    CITATION_SUFFIX_PATTERN = re.compile(
        r'\([A-Z][a-z]+.*?\d{4}.*?\)\s*$'
    )

    # Patterns that suggest an attribution line
    ATTRIBUTION_PATTERNS = [
        r"^[-\u2013\u2014]\s*[A-Z]",        # Starts with dash + capitalised name
        r"(?i)^(professor|prof\.|dr\.?|sir|lord|dame)\s+[A-Z]",  # Academic title + name
        r"^[A-Z][a-z]+\s+[A-Z][a-z]+,\s*\d{4}",  # "Firstname Lastname, 2019"
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect quote slides based on text patterns and layout.
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if "quote" in layout_name:
            return 0.80

        texts = extract_text_elements(slide)
        if not texts:
            return 0.0

        meaningful = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 1
        ]

        if not meaningful:
            return 0.0

        # Look for a text element that is a wrapped quote (starts AND ends with quote marks)
        has_wrapped_quote = False
        has_inline_quote = False
        has_attribution = False
        has_citation_suffix = False
        quote_text_len = 0

        for t in meaningful:
            text = t["text"].strip()
            if self.WRAPPED_QUOTE_PATTERN.match(text):
                has_wrapped_quote = True
                quote_text_len = len(text)
            elif self.INLINE_QUOTE_PATTERN.search(text):
                has_inline_quote = True
                quote_text_len = len(text)
            if self.CITATION_SUFFIX_PATTERN.search(text):
                has_citation_suffix = True
            for pattern in self.ATTRIBUTION_PATTERNS:
                if re.search(pattern, text):
                    has_attribution = True
                    break

        # Total text on slide — true quotes are sparse
        total_text = sum(len(t["text"]) for t in meaningful)

        # Strong signal: wrapped quote + attribution + very sparse slide
        if has_wrapped_quote and has_attribution and len(meaningful) <= 3:
            if total_text < 400:
                return 0.70

        # Moderate: wrapped quote + very sparse (2 elements max, short total)
        if has_wrapped_quote and len(meaningful) <= 2:
            if total_text < 250:
                return 0.50

        # Inline quote with citation: single text element containing
        # quoted text + parenthetical citation on a sparse slide
        if has_inline_quote and has_citation_suffix and len(meaningful) <= 2:
            if total_text < 400:
                return 0.65

        # Inline quote with attribution as separate element
        if has_inline_quote and has_attribution and len(meaningful) <= 3:
            if total_text < 400:
                return 0.60

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract quote text, attribution, and optional background image.

        Strategy:
        1. Find the longest text element — likely the quote.
        2. Find short text that looks like an attribution (has dash prefix,
           or is a name, or follows the quote).
        3. Extract background image if present.
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "quote": "",
            "attribution": "",
            "footer": "",
            "image_blob": None,
        }

        # Background image (largest by area)
        real_images = [img for img in images if img.get("blob")]
        if real_images:
            primary = max(real_images, key=lambda i: (i["width"] or 0) * (i["height"] or 0))
            result["image_blob"] = primary["blob"]

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

        # Extract footer
        footer_text = ""
        content_shapes = []
        for s in filtered:
            if s["is_placeholder"] and s["shape_idx"] in (17, 18):
                if s["shape_idx"] == 17 and not footer_text:
                    footer_text = s["text"]
                continue
            content_shapes.append(s)

        result["footer"] = footer_text

        if not content_shapes:
            return result

        # Identify quote vs attribution
        # The quote is typically the longest text
        content_shapes.sort(key=lambda s: -len(s["text"]))

        quote_text = content_shapes[0]["text"]

        # Clean up quote marks for consistent output
        quote_text = quote_text.strip()

        result["quote"] = quote_text

        # Attribution: look for remaining short text
        if len(content_shapes) > 1:
            for s in content_shapes[1:]:
                text = s["text"].strip()
                # Check if it looks like an attribution
                is_attrib = False
                for pattern in self.ATTRIBUTION_PATTERNS:
                    if re.search(pattern, text):
                        is_attrib = True
                        break
                # Also treat any short remaining text as attribution
                if is_attrib or len(text) < 80:
                    # Clean leading dash
                    text = re.sub(r"^[-\u2013\u2014]\s*", "", text)
                    result["attribution"] = text
                    break

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill quote placeholders. Optionally insert background image.
        NEVER set font properties.
        """
        import io
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Quote text
        if content.get("quote") and self.PH_QUOTE in placeholders:
            placeholders[self.PH_QUOTE].text = content["quote"]

        # Attribution
        if content.get("attribution") and self.PH_ATTRIBUTION in placeholders:
            placeholders[self.PH_ATTRIBUTION].text = content["attribution"]

        # Background picture
        if content.get("image_blob") and self.PH_PICTURE in placeholders:
            image_stream = io.BytesIO(content["image_blob"])
            placeholders[self.PH_PICTURE].insert_picture(image_stream)

    def get_placeholder_map(self) -> dict:
        return {
            "quote": self.PH_QUOTE,
            "attribution": self.PH_ATTRIBUTION,
            "picture": self.PH_PICTURE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
