"""
Base class for slide type handlers.

Each handler knows how to:
1. Detect whether an input slide matches its type
2. Extract content from a matching slide
3. Create a brand-compliant output slide using the template
"""

from abc import ABC, abstractmethod


class SlideHandler(ABC):
    """Base class for all slide type handlers."""

    # Subclasses must set these
    name: str = ""
    description: str = ""
    layout_name: str = ""  # Name in the template's slide master
    layout_index: int = 0  # Index in slide_layouts

    @abstractmethod
    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract content from an input slide.
        Returns a dict of classified content fields.
        """
        pass

    @abstractmethod
    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill a template slide's placeholders with extracted content.
        MUST only set .text — never set font properties.
        """
        pass

    @abstractmethod
    def detect(self, slide, slide_index: int) -> float:
        """
        Return a confidence score (0.0–1.0) that this input slide
        matches this handler's type.

        For Phase 1, only the cover handler needs detection.
        Later handlers will use heuristics + optional AI fallback.
        """
        pass

    def get_placeholder_map(self) -> dict:
        """Return the placeholder index mapping for this layout."""
        return {}
