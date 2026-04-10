"""
Thank You Handler — Contact/closing slide.

Template placeholders (layout index 42, "Thank You"):
    idx  0 — Title (defaults to "Contact")
    idx 10 — Name
    idx 16 — Job title
    idx 17 — Email
    idx 18 — Phone

Detection: Last slide in the deck, or slides with "thank you",
"questions", "contact", "Q&A" patterns. Also catches slides with
email/phone patterns.

Extraction: Pulls presenter/contact info from the slide itself,
or falls back to info extracted from the cover slide if available.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images


class ThankYouHandler(SlideHandler):

    name = "Thank You"
    description = "Closing/contact slide with presenter details"
    layout_name = "Thank You"
    layout_index = 42

    # Placeholder indices
    PH_TITLE = 0
    PH_NAME = 10
    PH_JOB_TITLE = 16
    PH_EMAIL = 17
    PH_PHONE = 18

    # --- Detection patterns ---

    # Closing patterns — these should match text where the closing phrase
    # is the PRIMARY content, not buried inside a longer sentence.
    CLOSING_PATTERNS = [
        r"(?i)^thank\s*you[!.\s]*$",     # "Thank you" as the whole text
        r"(?i)^thanks[!.\s]*$",           # "Thanks" as the whole text
        r"(?i)^questions[\s,?]*$",
        r"(?i)^q\s*(&|and)\s*a\b",
        r"(?i)^contact\s*$",              # "Contact" as the whole text
        r"(?i)^get\s+in\s+touch",
        r"(?i)^any\s+questions",
    ]

    EMAIL_PATTERN = r"[\w.+-]+@[\w-]+\.[\w.]+"
    PHONE_PATTERN = r"(\+?\d[\d\s\-()]{7,})"

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert",
        "click to add", "click icon",
        "\u00a9 the university", "cricos",
    ]

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect thank you / closing slides.
        """
        # Never match first slide
        if slide_index == 0:
            return 0.0

        texts = extract_text_elements(slide)

        meaningful = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 0
        ]

        all_text = " ".join(t["text"] for t in meaningful)

        # Strong signal: closing phrase
        has_closing = any(
            re.search(p, t["text"].strip())
            for t in meaningful
            for p in self.CLOSING_PATTERNS
        )

        # Supporting signal: email or phone present
        has_email = bool(re.search(self.EMAIL_PATTERN, all_text))
        has_phone = bool(re.search(self.PHONE_PATTERN, all_text))

        # Disqualifier: if any single text element is long (>100 chars),
        # this is a content slide, not a closing slide
        has_long_text = any(len(t["text"].strip()) > 100 for t in meaningful)

        # Layout name signal
        layout_name = slide.slide_layout.name.lower()
        is_thank_layout = "thank" in layout_name or "contact" in layout_name

        if is_thank_layout:
            return 0.95

        if has_long_text:
            return 0.05  # Content slides with long text are not Thank You

        if has_closing and (has_email or has_phone):
            return 0.9

        if has_closing and len(meaningful) <= 5:
            return 0.7

        # Email + phone without closing phrase — could be a contact slide
        if has_email and has_phone and len(meaningful) <= 6:
            return 0.6

        return 0.05

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract contact/presenter information from a closing slide.
        """
        texts = extract_text_elements(slide)

        # Flatten and filter
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
            and not re.match(r"^\d{1,3}$", t["text"].strip())  # skip slide numbers
        ]

        result = {
            "title": "Contact",  # Default
            "name": "",
            "job_title": "",
            "email": "",
            "phone": "",
        }

        for t in filtered:
            text = t["text"].strip()

            # Closing phrase → title
            if any(re.search(p, text) for p in self.CLOSING_PATTERNS):
                result["title"] = text
                continue

            # Email
            email_match = re.search(self.EMAIL_PATTERN, text)
            if email_match and not result["email"]:
                result["email"] = email_match.group(0)
                # If text is just the email, skip further classification
                if text == result["email"]:
                    continue

            # Phone
            phone_match = re.search(self.PHONE_PATTERN, text)
            if phone_match and not result["phone"]:
                result["phone"] = phone_match.group(0).strip()
                if text == result["phone"]:
                    continue

            # Name detection: "Professor X", "Dr X", or short proper-cased text
            if self._looks_like_name(text) and not result["name"]:
                result["name"] = text
                continue

            # Job title detection
            if self._looks_like_job_title(text) and not result["job_title"]:
                result["job_title"] = text
                continue

        return result

    def _looks_like_name(self, text: str) -> bool:
        """Check if text looks like a person's name."""
        name_patterns = [
            r"(?i)^(a/?)?professor\s",
            r"(?i)^(associate|assistant)\s+professor\s",
            r"(?i)^dr\.?\s",
            r"(?i)^(mr|mrs|ms|miss)\.?\s",
        ]
        if any(re.search(p, text) for p in name_patterns):
            return True

        # Short proper-cased text (2-4 words, each capitalised)
        words = text.split()
        if 2 <= len(words) <= 5 and all(w[0].isupper() for w in words if w):
            return True

        return False

    def _looks_like_job_title(self, text: str) -> bool:
        """Check if text looks like a job title."""
        title_patterns = [
            r"(?i)^(senior|junior|associate|assistant|lead|head|chief|director)",
            r"(?i)(manager|coordinator|officer|analyst|consultant|advisor|lecturer|fellow)",
            r"(?i)(school\s+of|faculty\s+of|department\s+of|centre\s+for)",
            r"(?i)(business\s+school|executive\s+education)",
        ]
        return any(re.search(p, text) for p in title_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill Thank You placeholders with extracted content.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("name") and self.PH_NAME in placeholders:
            placeholders[self.PH_NAME].text = content["name"]

        if content.get("job_title") and self.PH_JOB_TITLE in placeholders:
            placeholders[self.PH_JOB_TITLE].text = content["job_title"]

        if content.get("email") and self.PH_EMAIL in placeholders:
            placeholders[self.PH_EMAIL].text = content["email"]

        if content.get("phone") and self.PH_PHONE in placeholders:
            placeholders[self.PH_PHONE].text = content["phone"]

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "name": self.PH_NAME,
            "job_title": self.PH_JOB_TITLE,
            "email": self.PH_EMAIL,
            "phone": self.PH_PHONE,
        }
