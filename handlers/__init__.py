# Slide type handlers
from handlers.base import SlideHandler
from handlers.cover1 import Cover1Handler
from handlers.section_divider import SectionDividerHandler
from handlers.title_content import TitleContentHandler
from handlers.thank_you import ThankYouHandler

# Registry of all available handlers.
# Order matters for tie-breaking: more specific handlers are checked first.
# Title and Content is intentionally last — it's the fallback.
HANDLER_REGISTRY = {
    "Cover 1": Cover1Handler,
    "Section Divider": SectionDividerHandler,
    "Thank You": ThankYouHandler,
    "Title and Content": TitleContentHandler,
}


def get_all_handlers():
    """Return instantiated handlers for all registered slide types."""
    return {name: cls() for name, cls in HANDLER_REGISTRY.items()}
