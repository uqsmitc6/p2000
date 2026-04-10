# Slide type handlers
from handlers.base import SlideHandler
from handlers.cover1 import Cover1Handler
from handlers.section_divider import SectionDividerHandler
from handlers.title_content import TitleContentHandler
from handlers.thank_you import ThankYouHandler
from handlers.text_image import TextImageHandler
from handlers.title_table import TitleTableHandler
from handlers.two_content import TwoContentHandler
from handlers.quote import QuoteHandler
from handlers.title_only import TitleOnlyHandler
from handlers.split_content import SplitContentHandler
from handlers.acknowledgement import AcknowledgementHandler
from handlers.references import ReferencesHandler

# Registry of all available handlers.
# Order matters for tie-breaking: more specific handlers are checked first.
# Title and Content is intentionally last — it's the fallback.
HANDLER_REGISTRY = {
    "Acknowledgement of Country": AcknowledgementHandler,
    "Cover 1": Cover1Handler,
    "Section Divider": SectionDividerHandler,
    "Thank You": ThankYouHandler,
    "Text with Image": TextImageHandler,
    "Title and Table": TitleTableHandler,
    "Two Content": TwoContentHandler,
    "Split Content": SplitContentHandler,
    "Quote": QuoteHandler,
    "References": ReferencesHandler,
    "Title Only": TitleOnlyHandler,
    "Title and Content": TitleContentHandler,
}


def get_all_handlers():
    """Return instantiated handlers for all registered slide types."""
    return {name: cls() for name, cls in HANDLER_REGISTRY.items()}
