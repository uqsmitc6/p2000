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
from handlers.blank import BlankBrandedHandler
from handlers.picture_caption import PictureCaptionHandler
from handlers.three_content import ThreeContentHandler
from handlers.two_content_horizontal import TwoContentHorizontalHandler
from handlers.text_image_alt import TextImageAltHandler
from handlers.graph_block import GraphBlockHandler
from handlers.text_block import TextBlockHandler
from handlers.quote2 import Quote2Handler
from handlers.top_image_content import TopImageContentHandler
from handlers.intro_two_content import IntroTwoContentHandler
from handlers.picture_pullout import PicturePulloutHandler
from handlers.process_diagram import ProcessDiagramHandler
from handlers.three_pullouts import ThreePulloutsHandler
from handlers.two_graphs import TwoGraphsHandler
from handlers.image_collage import ImageCollageHandler
from handlers.text_four_images import TextFourImagesHandler
from handlers.multi_layout1 import MultiLayout1Handler
from handlers.multi_layout2 import MultiLayout2Handler
from handlers.icons_two_contents import IconsTwoContentsHandler
from handlers.three_col_text_images import ThreeColTextImagesHandler
from handlers.icons_text import IconsTextHandler
from handlers.order import OrderHandler

# Registry of all available handlers.
# Order matters for tie-breaking: more specific handlers are checked first.
# Title and Content is intentionally last — it's the fallback.
# Blank Branded is just above T&C — it's a low-priority catch-all.
HANDLER_REGISTRY = {
    "Acknowledgement of Country": AcknowledgementHandler,
    "Cover 1": Cover1Handler,
    "Section Divider": SectionDividerHandler,
    "Thank You": ThankYouHandler,
    "Top Image + Content": TopImageContentHandler,
    "Intro + Two Content": IntroTwoContentHandler,
    "Picture with Pullout": PicturePulloutHandler,
    "Picture with Caption": PictureCaptionHandler,
    "Text with Image": TextImageHandler,
    "Text with Image Alt": TextImageAltHandler,
    "Graph with Block": GraphBlockHandler,
    "Text with Block": TextBlockHandler,
    "Title and Table": TitleTableHandler,
    "Two Content": TwoContentHandler,
    "Three Content": ThreeContentHandler,
    "Two Content Horizontal": TwoContentHorizontalHandler,
    "Split Content": SplitContentHandler,
    "Quote": QuoteHandler,
    "Quote 2": Quote2Handler,
    "References": ReferencesHandler,
    "Title, Subtitle, 2 Graphs": TwoGraphsHandler,
    "Three Pullouts": ThreePulloutsHandler,
    "Process Diagram": ProcessDiagramHandler,
    "Order": OrderHandler,
    "Icons & Text": IconsTextHandler,
    "Three Column Text & Images": ThreeColTextImagesHandler,
    "Icons + Two Contents": IconsTwoContentsHandler,
    "Multi-layout 1": MultiLayout1Handler,
    "Multi-layout 2": MultiLayout2Handler,
    "Text with 4 Images": TextFourImagesHandler,
    "Image Collage": ImageCollageHandler,
    "Title Only": TitleOnlyHandler,
    "Blank Branded": BlankBrandedHandler,
    "Title and Content": TitleContentHandler,
}


def get_all_handlers():
    """Return instantiated handlers for all registered slide types."""
    return {name: cls() for name, cls in HANDLER_REGISTRY.items()}
