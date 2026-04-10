"""
References Handler — Academic citation and image credit slides.

Uses layout 6 "Title and Content":
    idx  0 — Title ("References" / "Selected References")
    idx 10 — Content (citation text)
    idx 31 — Subtitle (not used)
    idx 17 — Footer
    idx 18 — Slide number

Detection: slides containing academic citation patterns (DOIs, journal names,
year references, "et al.", "Source:", "References", image credit lines).
Typically found in the last few slides of a deck.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements


class ReferencesHandler(SlideHandler):

    name = "References"
    description = "Academic references and image credits slide"
    layout_name = "Title and Content"
    layout_index = 6

    PH_TITLE = 0
    PH_CONTENT = 10
    PH_SUBTITLE = 31
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # Title patterns that indicate a references slide
    REFERENCE_TITLE_PATTERNS = [
        r"(?i)^references?\s*$",
        r"(?i)^selected\s+references?\s*$",
        r"(?i)^bibliography\s*$",
        r"(?i)^sources?\s*$",
        r"(?i)^image\s+(credits?|sources?|references?)",
        r"(?i)^photo\s+credits?",
        r"(?i)^citations?\s*$",
        r"(?i)^further\s+reading",
        r"(?i)^recommended\s+reading",
        r"(?i)^reading\s+list",
    ]

    # Citation body patterns — each match adds to the score
    CITATION_PATTERNS = [
        r"\(\d{4}\)",                    # (2019)
        r",\s*\d{4}[,\.\)]",            # , 2019. or , 2019)
        r"(?i)\bet\s+al\.?",            # et al.
        r"doi[:\s]",                     # DOI references
        r"https?://doi\.org",            # DOI URLs
        r"(?i)journal\s+of\b",          # Journal of ...
        r"(?i)\bvol\.\s*\d",            # Vol. 12
        r"(?i)\bpp?\.\s*\d",            # p. 23 or pp. 23-45
        r"(?i)harvard\s+business\s+review",
        r"(?i)adobe\s+stock",           # Image credits
        r"(?i)licensed\s+(from|under)",  # Image licensing
        r"(?i)(unsplash|pexels|getty|shutterstock|istock)",
        r"(?i)image\s+(source|credit|licensed)",
        r"(?i)photo\s+(source|credit|by)",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect reference/citation slides.
        """
        if slide_index == 0:
            return 0.0

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

        # Check for reference title
        has_ref_title = False
        for t in meaningful:
            text = t["text"].strip()
            for pattern in self.REFERENCE_TITLE_PATTERNS:
                if re.search(pattern, text):
                    has_ref_title = True
                    break
            if has_ref_title:
                break

        # Count citation pattern matches in the body text
        all_text = " ".join(t["text"] for t in meaningful)
        citation_hits = 0
        for pattern in self.CITATION_PATTERNS:
            matches = re.findall(pattern, all_text)
            citation_hits += len(matches)

        # Strong: reference title + citation patterns
        if has_ref_title and citation_hits >= 2:
            return 0.85

        # Reference title alone (might just be a header)
        if has_ref_title:
            return 0.75

        # Heavy citations without explicit title (3+ patterns)
        if citation_hits >= 5:
            return 0.70

        # Moderate citations (could be inline references in a content slide)
        if citation_hits >= 3:
            # Only if the slide is mostly citations (short elements with years)
            year_elements = sum(
                1 for t in meaningful
                if re.search(r"\(\d{4}\)", t["text"]) or re.search(r",\s*\d{4}", t["text"])
            )
            if year_elements >= 2:
                return 0.60

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract reference title and citation text.
        """
        texts = extract_text_elements(slide)

        result = {
            "title": "References",
            "content": "",
            "footer": "",
        }

        if not texts:
            return result

        # Filter noise
        filtered = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 1
            and not re.match(r"^\d{1,3}$", t["text"].strip())
        ]

        if not filtered:
            return result

        # Separate title from body
        # Title: short text matching reference patterns, or the first short text
        title_text = None
        body_texts = []

        for t in filtered:
            text = t["text"].strip()
            if not title_text:
                for pattern in self.REFERENCE_TITLE_PATTERNS:
                    if re.search(pattern, text):
                        title_text = text
                        break
            if text != title_text:
                body_texts.append(text)

        if title_text:
            result["title"] = title_text
        elif body_texts:
            # Check if first element looks like a title (short, no year refs)
            first = body_texts[0]
            if len(first) < 60 and not re.search(r"\(\d{4}\)", first):
                result["title"] = first
                body_texts = body_texts[1:]

        # Join body text with newlines
        result["content"] = "\n".join(body_texts)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill reference slide placeholders.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("content") and self.PH_CONTENT in placeholders:
            placeholders[self.PH_CONTENT].text = content["content"]

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "content": self.PH_CONTENT,
            "subtitle": self.PH_SUBTITLE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
