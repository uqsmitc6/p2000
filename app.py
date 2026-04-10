"""
UQ Slide Converter — Streamlit App

Upload a PowerPoint file → get a brand-compliant version back.
Optionally uses Claude Vision API for smarter slide classification.
"""

import os
import logging
import streamlit as st

APP_VERSION = "0.5.1"

# --- Logging setup ---
# Logs go to stdout → visible in Render's log viewer
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("uqslide.app")

# --- Page config ---
st.set_page_config(
    page_title="UQ Slide Converter",
    page_icon="🟣",
    layout="centered",
)

# --- Styling ---
st.markdown("""
<style>
    h1 { color: #51247A; }
    .stDownloadButton > button {
        background-color: #51247A;
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        font-size: 1rem;
    }
    .stDownloadButton > button:hover {
        background-color: #3b1a5a;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.title("UQ Slide Converter")
st.caption(f"v{APP_VERSION}")
st.markdown(
    "Upload a PowerPoint file and get a brand-compliant version using the "
    "official UQ Business School template."
)

# --- API key ---
# Priority: env var → Streamlit secrets → sidebar input
api_key = os.environ.get("ANTHROPIC_API_KEY")
if not api_key:
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY", None)
    except Exception:
        pass

if not api_key:
    with st.sidebar:
        st.markdown("### Settings")
        api_key = st.text_input(
            "Anthropic API key (optional)",
            type="password",
            help="Enables AI-powered slide classification using Claude Vision. "
                 "Without a key, the tool uses heuristics only.",
        )

if api_key:
    st.success("AI classification enabled (Claude Vision)")
else:
    st.info(
        "**Heuristic mode** — slide classification uses pattern matching only. "
        "Add an Anthropic API key in the sidebar for AI-powered classification "
        "using slide screenshots."
    )

# --- Supported types ---
with st.expander("Supported slide types"):
    st.markdown("""
- **Cover 1** — Programme title slide (first slide)
- **Acknowledgement of Country** — Auto-inserted as slide 2 with standard UQ wording
- **Section Divider** — Section break with number and title
- **Title and Content** — Standard content slide (default fallback)
- **Text with Image** — Text alongside an image (auto-selects 1/3, 1/2, or 2/3 variant)
- **Two Content** — Two equal columns of text
- **Split Content** — Asymmetric columns (1/3 + 2/3 or 2/3 + 1/3)
- **Title and Table** — Slide with a data table
- **Title Only** — Title with image/diagram, minimal body text
- **Quote** — Quotation with attribution
- **References** — Academic citations, image credits, and bibliography slides
- **Thank You** — Closing/contact slide (last slide)

Slides that don't match a known type will be skipped and listed.
    """)

def _get_title(detail: dict) -> str:
    """Extract a display title from a report detail entry."""
    content = detail.get("content", {})
    if not content:
        return detail.get("preview", "")
    return (
        content.get("title", "") or
        content.get("section_num", "") or
        content.get("name", "") or
        detail.get("preview", "")
    )[:60]


# --- File upload ---
uploaded_file = st.file_uploader(
    "Upload your PowerPoint file",
    type=["pptx"],
    help="Drag and drop or click to browse. Only .pptx files are supported.",
)

if uploaded_file is not None:
    input_bytes = uploaded_file.getvalue()
    input_filename = uploaded_file.name

    st.markdown(f"**Uploaded:** `{input_filename}` ({len(input_bytes) / 1024:.0f} KB)")
    logger.info("File uploaded: %s (%d KB)", input_filename, len(input_bytes) // 1024)

    if st.button("Convert presentation", type="primary"):
        logger.info("Conversion started: %s (API=%s)", input_filename, bool(api_key))
        status_area = st.empty()

        def update_status(msg):
            status_area.text(msg)

        with st.spinner("Analysing and converting slides..."):
            try:
                from converter import convert_presentation
                output_bytes, report = convert_presentation(
                    input_bytes,
                    api_key=api_key if api_key else None,
                    progress_callback=update_status,
                )

                status_area.empty()

                total = len(report["details"])
                converted = report["slides_converted"]
                flagged = report["slides_flagged"]
                skipped = report["slides_skipped"]
                api_calls = report.get("api_calls", 0)

                # --- Summary ---
                cols = st.columns(4 if api_calls else 3)
                cols[0].metric("Converted", converted)
                cols[1].metric("Flagged", flagged)
                cols[2].metric("Skipped", skipped)
                if api_calls:
                    cols[3].metric("AI calls", api_calls)

                # --- Download ---
                output_filename = input_filename.replace(".pptx", "_BRANDED.pptx")
                st.download_button(
                    label=f"Download {output_filename}",
                    data=output_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

                # --- Confident conversions ---
                confident = [d for d in report["details"] if d["status"] == "converted"]
                if confident:
                    with st.expander(f"Converted slides ({len(confident)})", expanded=False):
                        for d in confident:
                            title = _get_title(d)
                            method = d.get("classification_method", "heuristic")
                            tag = " 🤖" if method == "api" else ""
                            st.markdown(
                                f"- **Slide {d['slide']}** → {d['handler']}{tag} — *{title}*"
                            )

                # --- Flagged slides ---
                flagged_items = [d for d in report["details"] if d["status"] == "flagged"]
                if flagged_items:
                    with st.expander(f"Flagged for review ({len(flagged_items)})", expanded=True):
                        st.warning(
                            "These slides were converted but classification confidence was low. "
                            "Check they've been assigned to the right layout."
                        )
                        for d in flagged_items:
                            title = _get_title(d)
                            conf = d.get("confidence", 0)
                            method = d.get("classification_method", "heuristic")
                            tag = " 🤖" if method == "api" else ""
                            st.markdown(
                                f"- **Slide {d['slide']}** → {d['handler']} "
                                f"({conf:.0%}{tag}) — *{title}*"
                            )

                # --- Skipped slides ---
                skipped_items = [d for d in report["details"] if d["status"] == "skipped"]
                if skipped_items:
                    with st.expander(f"Skipped — not recognised ({len(skipped_items)})", expanded=True):
                        st.error(
                            "These slides are **not included** in the output. "
                            "They may need to be added manually."
                        )
                        for d in skipped_items:
                            preview = d.get("preview", "(no text)")
                            reason = d.get("api_reason", d.get("reason", ""))
                            reason_str = f" — *{reason}*" if reason else ""
                            st.markdown(
                                f"- **Slide {d['slide']}**: {preview}{reason_str}"
                            )

                # --- Errors ---
                errors = report.get("errors", [])
                if errors:
                    with st.expander(f"Errors ({len(errors)})", expanded=True):
                        st.warning(
                            "Some issues occurred during conversion. "
                            "The output may be incomplete."
                        )
                        for err in errors:
                            st.markdown(f"- {err}")
                    logger.warning("Conversion completed with %d errors", len(errors))

            except Exception as e:
                logger.error("Conversion failed: %s", e, exc_info=True)
                st.error(f"Something went wrong: {str(e)}")
                st.exception(e)


# --- Footer ---
st.markdown("---")
st.caption("UQ Business School — Learning Design Team")
