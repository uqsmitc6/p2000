"""
UQ Slide Converter — Streamlit App

Upload a PowerPoint file → get a brand-compliant version back.
Optionally uses Claude Vision API for smarter slide classification.
"""

import base64
import io
import os
import logging
import streamlit as st

APP_VERSION = "0.9.4"

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
    layout="wide",
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
    /* Make the file uploader drop zone more prominent */
    [data-testid="stFileUploader"] {
        border: 2px dashed #51247A;
        border-radius: 10px;
        padding: 1rem;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #3b1a5a;
        background-color: #f9f5ff;
    }
    /* Viewer slide image styling */
    .slide-viewer img {
        border: 1px solid #ddd;
        border-radius: 4px;
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

with st.expander("System diagnostics"):
    import subprocess
    try:
        lo_result = subprocess.run(
            ["libreoffice", "--version"],
            capture_output=True, text=True, timeout=5,
        )
        if lo_result.returncode == 0:
            st.success(f"LibreOffice: {lo_result.stdout.strip()}")
        else:
            st.error(f"LibreOffice: error code {lo_result.returncode}")
    except FileNotFoundError:
        st.error("LibreOffice: NOT INSTALLED")
    except Exception as e:
        st.error(f"LibreOffice: {e}")

    try:
        ppm_result = subprocess.run(
            ["pdftoppm", "-v"],
            capture_output=True, text=True, timeout=5,
        )
        ppm_version = ppm_result.stderr.strip() or ppm_result.stdout.strip() or "available"
        st.success(f"pdftoppm: {ppm_version}")
    except FileNotFoundError:
        st.error("pdftoppm (poppler-utils): NOT INSTALLED")
    except Exception as e:
        st.error(f"pdftoppm: {e}")

    if api_key:
        st.success("Anthropic API key: configured")
    else:
        st.warning("Anthropic API key: not set (heuristics only)")

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
    "Drag and drop your PowerPoint file here, or click to browse",
    type=["pptx"],
    help="Only .pptx files are supported. The file will be converted to UQ brand format.",
)

if uploaded_file is not None:
    input_bytes = uploaded_file.getvalue()
    input_filename = uploaded_file.name

    st.markdown(f"**Uploaded:** `{input_filename}` ({len(input_bytes) / 1024:.0f} KB)")
    logger.info("File uploaded: %s (%d KB)", input_filename, len(input_bytes) // 1024)

    # Estimate slide count for memory warning
    try:
        from pptx import Presentation as _Prs
        _tmp_prs = _Prs(io.BytesIO(input_bytes))
        _est_slides = len(_tmp_prs.slides)
        del _tmp_prs
    except Exception:
        _est_slides = 0

    # Option to skip AI features for large decks (saves memory)
    use_ai = api_key
    if api_key and _est_slides > 25:
        st.warning(
            f"This deck has **{_est_slides} slides**. AI classification and "
            "verification require significant memory. If conversion fails, "
            "try with AI features disabled."
        )
        use_ai = st.checkbox(
            "Enable AI classification & verification",
            value=False,
            help="Uses Claude Vision for slide classification and post-conversion QA. "
                 "Disable for large decks to avoid memory issues on the free tier.",
        )
        if use_ai:
            use_ai = api_key  # Pass the actual key

    if st.button("Convert presentation", type="primary"):
        logger.info("Conversion started: %s (API=%s, slides=%d)",
                     input_filename, bool(use_ai), _est_slides)
        status_area = st.empty()

        def update_status(msg):
            status_area.text(msg)

        with st.spinner("Analysing and converting slides..."):
            try:
                from converter import convert_presentation
                output_bytes, report = convert_presentation(
                    input_bytes,
                    api_key=use_ai if use_ai else None,
                    progress_callback=update_status,
                )

                status_area.empty()

                # Store results in session state so they persist across reruns
                st.session_state["output_bytes"] = output_bytes
                st.session_state["report"] = report
                st.session_state["output_filename"] = input_filename.replace(
                    ".pptx", "_BRANDED.pptx"
                )

                # Free converter memory — images are now in session state
                import gc
                del output_bytes, report
                gc.collect()

            except Exception as e:
                logger.error("Conversion failed: %s", e, exc_info=True)
                st.error(f"Something went wrong: {str(e)}")
                st.exception(e)

# --- Display results from session state ---
if "output_bytes" in st.session_state and "report" in st.session_state:
    report = st.session_state["report"]
    output_bytes = st.session_state["output_bytes"]
    output_filename = st.session_state["output_filename"]

    converted = report["slides_converted"]
    flagged = report["slides_flagged"]
    skipped = report["slides_skipped"]
    api_calls = report.get("api_calls", 0)

    # --- Summary ---
    v_summary = report.get("verification_summary", {})
    has_verification = bool(v_summary)
    num_cols = 3 + (1 if api_calls else 0) + (1 if has_verification else 0)
    cols = st.columns(num_cols)
    col_idx = 0
    cols[col_idx].metric("Converted", converted); col_idx += 1
    cols[col_idx].metric("Flagged", flagged); col_idx += 1
    cols[col_idx].metric("Skipped", skipped); col_idx += 1
    if api_calls:
        cols[col_idx].metric("AI calls", api_calls); col_idx += 1
    if has_verification:
        v_issues_count = v_summary.get("issues_found", 0)
        cols[col_idx].metric("QA issues", v_issues_count)

    # --- Download buttons ---
    dl_col1, dl_col2 = st.columns([1, 1])
    with dl_col1:
        st.download_button(
            label=f"Download {output_filename}",
            data=output_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    with dl_col2:
        # Generate feedback report as markdown
        feedback_lines = []
        feedback_lines.append(f"# UQ Slide Converter — Feedback Report")
        feedback_lines.append(f"")
        source_name = output_filename.replace("_BRANDED.pptx", ".pptx")
        feedback_lines.append(f"**Source file:** {source_name}")
        feedback_lines.append(f"**Converted:** {converted} slides | **Flagged:** {flagged} | **Skipped:** {skipped}")
        if api_calls:
            feedback_lines.append(f"**AI calls:** {api_calls}")
        if has_verification:
            v_issues_count = v_summary.get("issues_found", 0)
            v_passed_count = v_summary.get("passed", 0)
            v_total = v_summary.get("total", 0)
            feedback_lines.append(f"**Verification:** {v_passed_count} passed, {v_issues_count} issues ({v_total} slides checked)")
        feedback_lines.append("")

        # Slide-by-slide details
        feedback_lines.append("---")
        feedback_lines.append("")
        feedback_lines.append("## Slide Details")
        feedback_lines.append("")
        for detail in report["details"]:
            slide_num = detail["slide"]
            handler = detail.get("handler", "?")
            status = detail["status"]
            title = _get_title(detail)
            conf = detail.get("confidence", 0)
            method = detail.get("classification_method", "heuristic")
            method_tag = " (AI)" if method == "api" else " (heuristic)"

            v = detail.get("verification", {})
            if v.get("pass") is True:
                v_status = "PASS"
            elif v.get("pass") is False:
                sev = v.get("severity", "unknown")
                v_status = f"ISSUE ({sev})"
            else:
                v_status = "not verified"

            feedback_lines.append(f"### Slide {slide_num} — {handler}: {title}")
            feedback_lines.append(f"")
            feedback_lines.append(f"- **Status:** {status} ({conf:.0%}{method_tag})")
            feedback_lines.append(f"- **Verification:** {v_status}")

            if v.get("pass") is False:
                issues = v.get("issues", [])
                for issue in issues:
                    feedback_lines.append(f"  - {issue}")

            feedback_lines.append("")

        # Flagged slides section
        flagged_items_fb = [d for d in report["details"] if d["status"] == "flagged"]
        if flagged_items_fb:
            feedback_lines.append("---")
            feedback_lines.append("")
            feedback_lines.append("## Flagged for Review")
            feedback_lines.append("")
            for d in flagged_items_fb:
                title = _get_title(d)
                conf = d.get("confidence", 0)
                feedback_lines.append(f"- **Slide {d['slide']}** → {d.get('handler', '?')} ({conf:.0%}) — {title}")
            feedback_lines.append("")

        # Skipped slides section
        skipped_items_fb = [d for d in report["details"] if d["status"] == "skipped"]
        if skipped_items_fb:
            feedback_lines.append("---")
            feedback_lines.append("")
            feedback_lines.append("## Skipped Slides (not in output)")
            feedback_lines.append("")
            for d in skipped_items_fb:
                preview = d.get("preview", "(no text)")
                reason = d.get("api_reason", d.get("reason", ""))
                reason_str = f" — {reason}" if reason else ""
                feedback_lines.append(f"- **Slide {d['slide']}**: {preview}{reason_str}")
            feedback_lines.append("")

        # Errors section
        errors_fb = report.get("errors", [])
        if errors_fb:
            feedback_lines.append("---")
            feedback_lines.append("")
            feedback_lines.append("## Errors")
            feedback_lines.append("")
            for err in errors_fb:
                feedback_lines.append(f"- {err}")
            feedback_lines.append("")

        feedback_md = "\n".join(feedback_lines)
        feedback_filename = output_filename.replace("_BRANDED.pptx", "_FEEDBACK.md")
        st.download_button(
            label=f"Download feedback report",
            data=feedback_md,
            file_name=feedback_filename,
            mime="text/markdown",
        )

    # --- In-browser slide comparison viewer ---
    source_images = report.get("source_images", {})
    output_images = report.get("output_images", {})

    if source_images and output_images:
        st.markdown("---")
        st.markdown("### Slide Comparison Viewer")
        st.caption("Side-by-side comparison of original → branded slides. Click any slide to expand.")

        # Build list of viewable slides (converted/flagged with both source and output images)
        viewable_slides = []
        for detail in report["details"]:
            if detail["status"] not in ("converted", "flagged"):
                continue
            slide_num = detail["slide"]
            output_idx = detail.get("output_index")
            if not isinstance(slide_num, int) or output_idx is None:
                continue  # Skip auto-inserted slides (AoC)
            source_idx = slide_num - 1
            if source_idx in source_images and output_idx in output_images:
                viewable_slides.append(detail)

        if viewable_slides:
            # Severity icons
            severity_icon = {
                "ok": "✅", "minor": "🟡", "major": "🟠", "critical": "🔴",
            }

            # --- Filter controls ---
            filter_col1, filter_col2 = st.columns([1, 3])
            with filter_col1:
                show_filter = st.selectbox(
                    "Show",
                    ["All slides", "Issues only", "Passed only"],
                    index=0,
                    key="viewer_filter",
                )

            # Apply filter
            if show_filter == "Issues only":
                filtered_slides = [
                    d for d in viewable_slides
                    if d.get("verification", {}).get("pass") is False
                ]
            elif show_filter == "Passed only":
                filtered_slides = [
                    d for d in viewable_slides
                    if d.get("verification", {}).get("pass") is True
                ]
            else:
                filtered_slides = viewable_slides

            st.caption(f"Showing {len(filtered_slides)} of {len(viewable_slides)} slides")

            # --- All slides as expandable sections ---
            for i, detail in enumerate(filtered_slides):
                source_idx = detail["slide"] - 1
                output_idx = detail.get("output_index")
                handler = detail.get("handler", "?")
                title = _get_title(detail)[:45]
                v = detail.get("verification", {})

                # Build expander label with status icon
                if v.get("pass") is True:
                    icon = "✅"
                elif v.get("pass") is False:
                    sev = v.get("severity", "unknown")
                    icon = severity_icon.get(sev, "❓")
                else:
                    icon = "—"

                label = f"{icon} Slide {detail['slide']} → {handler}: {title}"

                # Auto-expand slides with issues
                has_issues = v.get("pass") is False
                with st.expander(label, expanded=has_issues):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**Original**")
                        st.image(source_images[source_idx], use_container_width=True)
                    with col2:
                        st.markdown("**Branded**")
                        st.image(output_images[output_idx], use_container_width=True)

                    # Verification feedback
                    if v:
                        if v.get("pass") is True:
                            st.success("Verification passed — no issues found")
                        elif v.get("pass") is False:
                            sev = v.get("severity", "unknown")
                            sev_icon = severity_icon.get(sev, "❓")
                            issues = v.get("issues", [])
                            st.error(f"{sev_icon} **{sev.title()}** — {len(issues)} issue(s)")
                            for issue in issues:
                                st.markdown(f"- {issue}")
                    else:
                        st.info("No verification data (AI verification not enabled)")

    # --- Summary expanders (below viewer) ---
    st.markdown("---")

    # --- Confident conversions ---
    confident = [d for d in report["details"] if d["status"] in ("converted", "replaced")]
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
        with st.expander(f"Flagged for review ({len(flagged_items)})", expanded=False):
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
                if "rendering failed" in err.lower():
                    st.error(f"🖼️ {err}")
                else:
                    st.markdown(f"- {err}")
            logger.warning("Conversion completed with %d errors", len(errors))

    # --- Verification Summary (text-based, below viewer) ---
    verification = report.get("verification", [])
    v_summary = report.get("verification_summary", {})
    if verification:
        v_issues = [v for v in verification if v.get("pass") is False]
        v_passed = [v for v in verification if v.get("pass") is True]

        severity_icon = {"ok": "✅", "minor": "🟡", "major": "🟠", "critical": "🔴"}

        if v_issues:
            with st.expander(
                f"🔍 Verification issues ({len(v_issues)} of {len(verification)} slides)",
                expanded=True,
            ):
                st.info(
                    "AI compared each original slide against its branded version "
                    "and found the following issues."
                )
                for v in v_issues:
                    sev = v.get("severity", "unknown")
                    icon = severity_icon.get(sev, "❓")
                    handler = v.get("handler", "")
                    issues_text = "; ".join(v.get("issues", []))
                    st.markdown(
                        f"- {icon} **Slide {v['source_slide']}** → {handler} "
                        f"({sev}) — {issues_text}"
                    )

        if v_passed:
            with st.expander(
                f"✅ Verification passed ({len(v_passed)} slides)", expanded=False
            ):
                for v in v_passed:
                    handler = v.get("handler", "")
                    st.markdown(
                        f"- **Slide {v['source_slide']}** → {handler} — OK"
                    )

        if v_summary:
            st.caption(
                f"Verification: {v_summary.get('passed', 0)} passed, "
                f"{v_summary.get('issues_found', 0)} issues, "
                f"{v_summary.get('errors', 0)} errors "
                f"({v_summary.get('total', 0)} slides checked)"
            )


# --- Admin: API Cost Tracker (sidebar) ---
with st.sidebar:
    st.markdown("---")
    st.markdown("### Admin")

    try:
        from utils.cost_logger import (
            get_cost_summary, get_cost_log, clear_cost_log,
            export_cost_log_csv, COST_LOG_FILE,
        )

        summary = get_cost_summary()

        if summary["total_calls"] > 0:
            st.metric("Total API cost", f"${summary['total_cost_usd']:.4f}")
            st.caption(
                f"{summary['total_calls']} calls · "
                f"{summary['total_input_tokens']:,} in / "
                f"{summary['total_output_tokens']:,} out tokens"
            )
            st.caption(f"Log: `{COST_LOG_FILE}`")

            # Breakdown by purpose
            by_purpose = summary.get("by_purpose", {})
            if by_purpose:
                purpose_lines = []
                for purpose, stats in sorted(by_purpose.items()):
                    purpose_lines.append(
                        f"- **{purpose.title()}**: {stats['calls']} calls — "
                        f"${stats['cost_usd']:.4f}"
                    )
                with st.expander("Cost by purpose"):
                    st.markdown("\n".join(purpose_lines))

            # Breakdown by date
            by_date = summary.get("by_date", {})
            if by_date:
                with st.expander("Cost by date"):
                    for date_str in sorted(by_date.keys(), reverse=True):
                        stats = by_date[date_str]
                        st.markdown(
                            f"- **{date_str}**: {stats['calls']} calls — "
                            f"${stats['cost_usd']:.4f}"
                        )

            # Breakdown by filename
            by_filename = summary.get("by_filename", {})
            if by_filename:
                with st.expander("Cost by file"):
                    for fname in sorted(by_filename.keys()):
                        stats = by_filename[fname]
                        st.markdown(
                            f"- **{fname}**: {stats['calls']} calls — "
                            f"${stats['cost_usd']:.4f}"
                        )

            # Recent calls
            log = get_cost_log()
            if log:
                with st.expander(f"Recent calls ({min(len(log), 20)} of {len(log)})"):
                    for entry in log[:20]:
                        ts = entry.get("timestamp", "")[:19].replace("T", " ")
                        purpose = entry.get("purpose", "?")
                        slide = entry.get("slide_info", "")
                        cost = entry.get("total_cost_usd", 0)
                        tokens = entry.get("total_tokens", 0)
                        st.caption(
                            f"{ts} · {purpose} · {slide} · "
                            f"{tokens:,} tokens · ${cost:.4f}"
                        )

            # Export + Clear buttons
            col_export, col_clear = st.columns(2)
            with col_export:
                csv_data = export_cost_log_csv()
                if csv_data:
                    st.download_button(
                        "Export CSV",
                        data=csv_data,
                        file_name="api_costs.csv",
                        mime="text/csv",
                        type="secondary",
                    )
            with col_clear:
                if st.button("Clear log", type="secondary"):
                    clear_cost_log()
                    st.rerun()
        else:
            st.caption("No API calls logged yet.")

    except Exception as e:
        st.caption(f"Cost tracker unavailable: {e}")


# --- Footer ---
st.markdown("---")
st.caption("UQ Business School — Learning Design Team")
