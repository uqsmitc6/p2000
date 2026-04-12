"""
Quote 2 Handler — Alternate quote styling.

Template layout 39 "Quote 2":
    Same placeholder structure as Quote 1 (layout 38):
    idx 10 — Background picture (full slide)
    idx 16 — Quote text
    idx 17 — Attribution
    idx 18 — Footer
    idx 19 — Slide number

The visual difference is in background styling. Detection triggers
on source layout name containing "quote 2". Otherwise falls back to
Quote 1 (which is the default).

CRITICAL RULE: Never set font properties on placeholders.
"""

from handlers.quote import QuoteHandler


class Quote2Handler(QuoteHandler):

    name = "Quote 2"
    description = "A quote slide with alternate background styling"
    layout_name = "Quote 2"
    layout_index = 39

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides that should use Quote 2 specifically.
        Only triggers when source layout explicitly names "quote 2".
        For all other quote slides, the original QuoteHandler wins.
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "quote 2" in layout_name or "quote2" in layout_name:
            return 0.82  # Slightly above Quote 1's 0.80

        # Don't compete with Quote 1 for generic quote detection
        return 0.0
