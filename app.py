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

APP_VERSION = "0.9.6"

# --- Logging setup ---
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
    [data-testid="stFileUploader"] {
        border: 2px dashed #51247A;
        border-radius: 10px;
        padding: 1rem;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #3b1a5a;
        background-color: #f9f5ff;
    }
    .slide-viewer img {
        border: 1px solid #ddd;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.title("UQ Slide Converter")
st.caption(f"v{APP_VERSION}")

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


# --- Top-level tabs ---
tab_converter, tab_costs = st.tabs(["Converter", "API Costs"])

# ============================================================
# TAB 1: Converter
# ============================================================
with tab_converter:
    if api_key:
        st.success("AI classification enabled (Claude Vision)")
    else:
        st.info(
            "**Heuristic mode** — slide classification uses pattern matching only. "
            "Add an Anthropic API key in the sidebar for AI-powered classification "
            "using slide screenshots."
        )

    st.markdown(
        "Upload a PowerPoint file and get a brand-compliant version using the "
        "official UQ Business School template."
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

        # Option to skip AI features for large decks
        use_ai = api_key
        file_mb = len(input_bytes) / (1024 * 1024)
        if api_key and (_est_slides > 50 or file_mb > 100):
            st.warning(
                f"This deck has **{_est_slides} slides** ({file_mb:.0f} MB). "
                "AI classification and verification will render slides to images. "
                "For very large files, this may take several minutes."
            )
            use_ai = st.checkbox(
                "Enable AI classification & verification",
                value=True,
                help="Uses Claude Vision for slide classification and post-conversion QA. "
                     "Images are rendered to disk to support large files.",
            )
            if use_ai:
                use_ai = api_key  # Pass the actual key
        elif api_key and _est_slides > 25:
            use_ai = st.checkbox(
                "Enable AI classification & verification",
                value=True,
                help="Uses Claude Vision for slide classification and post-conversion QA.",
            )
            if use_ai:
                use_ai = api_key

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
                        filename=input_filename,
                    )

                    status_area.empty()

                    # Store results in session state so they persist across reruns
                    st.session_state["output_bytes"] = output_bytes
                    st.session_state["report"] = report
                    st.session_state["output_filename"] = input_filename.replace(
                        ".pptx", "_BRANDED.pptx"
                    )

                    # Free converter memory
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
        source_render_dir = report.get("source_render_dir")
        output_render_dir = report.get("output_render_dir")
        num_source_rendered = report.get("num_source_rendered", 0)
        num_output_rendered = report.get("num_output_rendered", 0)

        # Also support legacy in-memory dicts for backward compat
        source_images = report.get("source_images", {})
        output_images = report.get("output_images", {})
        has_viewer_data = (
            (source_render_dir and num_source_rendered > 0 and
             output_render_dir and num_output_rendered > 0) or
            (source_images and output_images)
        )

        if has_viewer_data:
            st.markdown("---")
            st.markdown("### Slide Comparison Viewer")
            st.caption("Side-by-side comparison of original → branded slides. Click any slide to expand.")

            def _get_slide_img(render_dir, idx, legacy_dict):
                if render_dir:
                    from utils.renderer import load_slide_image
                    img = load_slide_image(render_dir, idx)
                    if img:
                        return img
                return legacy_dict.get(idx)

            viewable_slides = []
            for detail in report["details"]:
                if detail["status"] not in ("converted", "flagged"):
                    continue
                slide_num = detail["slide"]
                output_idx = detail.get("output_index")
                if not isinstance(slide_num, int) or output_idx is None:
                    continue
                viewable_slides.append(detail)

            if viewable_slides:
                severity_icon = {
                    "ok": "✅", "minor": "🟡", "major": "🟠", "critical": "🔴",
                }

                filter_col1, filter_col2 = st.columns([1, 3])
                with filter_col1:
                    show_filter = st.selectbox(
                        "Show",
                        ["All slides", "Issues only", "Passed only"],
                        index=0,
                        key="viewer_filter",
                    )

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

                for i, detail in enumerate(filtered_slides):
                    source_idx = detail["slide"] - 1
                    output_idx = detail.get("output_index")
                    handler = detail.get("handler", "?")
                    title = _get_title(detail)[:45]
                    v = detail.get("verification", {})

                    if v.get("pass") is True:
                        icon = "✅"
                    elif v.get("pass") is False:
                        sev = v.get("severity", "unknown")
                        icon = severity_icon.get(sev, "❓")
                    else:
                        icon = "—"

                    label = f"{icon} Slide {detail['slide']} → {handler}: {title}"

                    has_issues = v.get("pass") is False
                    with st.expander(label, expanded=has_issues):
                        col1, col2 = st.columns(2)
                        src_img = _get_slide_img(source_render_dir, source_idx, source_images)
                        out_img = _get_slide_img(output_render_dir, output_idx, output_images)
                        with col1:
                            st.markdown("**Original**")
                            if src_img:
                                st.image(src_img, use_container_width=True)
                            else:
                                st.info("Source image not available")
                        with col2:
                            st.markdown("**Branded**")
                            if out_img:
                                st.image(out_img, use_container_width=True)
                            else:
                                st.info("Output image not available")

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

        # --- Verification Summary ---
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


# ============================================================
# TAB 2: API Costs (persistent log)
# ============================================================
with tab_costs:
    try:
        import pandas as pd
        from utils.cost_logger import (
            get_cost_summary, get_cost_log, clear_cost_log,
            export_cost_log_csv, COST_LOG_FILE,
        )

        entries = get_cost_log()
        summary = get_cost_summary(entries)

        if summary["total_calls"] == 0:
            st.info("No API calls logged yet. Convert a file with AI enabled to start tracking costs.")
        else:
            # --- Top-level metrics ---
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total cost", f"${summary['total_cost_usd']:.4f}")
            m2.metric("Total calls", f"{summary['total_calls']:,}")
            m3.metric("Input tokens", f"{summary['total_input_tokens']:,}")
            m4.metric("Output tokens", f"{summary['total_output_tokens']:,}")

            st.caption(f"Log file: `{COST_LOG_FILE}` (persistent across deployments)")

            # --- Breakdowns side-by-side ---
            bcol1, bcol2, bcol3 = st.columns(3)

            with bcol1:
                st.markdown("**By purpose**")
                by_purpose = summary.get("by_purpose", {})
                if by_purpose:
                    purpose_rows = []
                    for purpose, stats in sorted(by_purpose.items()):
                        purpose_rows.append({
                            "Purpose": purpose.title(),
                            "Calls": stats["calls"],
                            "Cost (USD)": f"${stats['cost_usd']:.4f}",
                        })
                    st.dataframe(pd.DataFrame(purpose_rows), hide_index=True, use_container_width=True)

            with bcol2:
                st.markdown("**By date**")
                by_date = summary.get("by_date", {})
                if by_date:
                    date_rows = []
                    for date_str in sorted(by_date.keys(), reverse=True):
                        stats = by_date[date_str]
                        date_rows.append({
                            "Date": date_str,
                            "Calls": stats["calls"],
                            "Cost (USD)": f"${stats['cost_usd']:.4f}",
                        })
                    st.dataframe(pd.DataFrame(date_rows), hide_index=True, use_container_width=True)

            with bcol3:
                st.markdown("**By file**")
                by_filename = summary.get("by_filename", {})
                if by_filename:
                    file_rows = []
                    for fname in sorted(by_filename.keys()):
                        stats = by_filename[fname]
                        file_rows.append({
                            "File": fname[:40],
                            "Calls": stats["calls"],
                            "Cost (USD)": f"${stats['cost_usd']:.4f}",
                        })
                    st.dataframe(pd.DataFrame(file_rows), hide_index=True, use_container_width=True)

            # --- Full call log as searchable table ---
            st.markdown("---")
            st.markdown("### All API calls")

            # Build dataframe from entries (already newest-first)
            df_rows = []
            for entry in entries:
                ts = entry.get("timestamp", "")[:19].replace("T", " ")
                df_rows.append({
                    "Time": ts,
                    "Purpose": entry.get("purpose", ""),
                    "Slide": entry.get("slide_info", ""),
                    "File": entry.get("filename", "")[:35],
                    "Model": entry.get("model", ""),
                    "In tokens": entry.get("input_tokens", 0),
                    "Out tokens": entry.get("output_tokens", 0),
                    "Cost (USD)": entry.get("total_cost_usd", 0),
                })

            df = pd.DataFrame(df_rows)

            # Filter controls
            filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 2])
            with filter_col1:
                purpose_filter = st.selectbox(
                    "Filter by purpose",
                    ["All"] + sorted(set(df["Purpose"].unique())),
                    key="cost_purpose_filter",
                )
            with filter_col2:
                file_filter = st.selectbox(
                    "Filter by file",
                    ["All"] + sorted(set(df["File"].unique())),
                    key="cost_file_filter",
                )

            filtered_df = df.copy()
            if purpose_filter != "All":
                filtered_df = filtered_df[filtered_df["Purpose"] == purpose_filter]
            if file_filter != "All":
                filtered_df = filtered_df[filtered_df["File"] == file_filter]

            st.dataframe(
                filtered_df,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Cost (USD)": st.column_config.NumberColumn(format="$%.6f"),
                    "In tokens": st.column_config.NumberColumn(format="%d"),
                    "Out tokens": st.column_config.NumberColumn(format="%d"),
                },
            )
            st.caption(f"Showing {len(filtered_df)} of {len(df)} entries")

            # --- Export + Clear ---
            st.markdown("---")
            action_col1, action_col2, action_col3 = st.columns([1, 1, 4])
            with action_col1:
                csv_data = export_cost_log_csv()
                if csv_data:
                    st.download_button(
                        "Export CSV",
                        data=csv_data,
                        file_name="api_costs.csv",
                        mime="text/csv",
                    )
            with action_col2:
                if st.button("Clear log", type="secondary", key="clear_cost_log"):
                    clear_cost_log()
                    st.rerun()

        # ============================================================
        # Conversion History (within API Costs tab)
        # ============================================================
        st.markdown("---")
        st.markdown("### Conversion History")
        st.caption("Track quality improvement across conversion runs")

        try:
            from utils.conversion_logger import (
                get_conversion_history, get_file_progression, clear_conversion_history,
            )

            history = get_conversion_history()

            if not history:
                st.info("No conversions logged yet. Run a conversion to start tracking.")
            else:
                # Summary table — all conversions, newest first
                hist_rows = []
                filenames_seen = set()
                for entry in history:
                    ts = entry.get("timestamp", "")[:19].replace("T", " ")
                    fname = entry.get("filename", "unknown")
                    filenames_seen.add(fname)
                    sev = entry.get("issues_by_severity", {})
                    hist_rows.append({
                        "Time": ts,
                        "File": fname[:45],
                        "Converted": entry.get("slides_converted", 0),
                        "Flagged": entry.get("slides_flagged", 0),
                        "Skipped": entry.get("slides_skipped", 0),
                        "Critical": sev.get("critical", 0),
                        "Major": sev.get("major", 0),
                        "Minor": sev.get("minor", 0),
                        "Total Issues": entry.get("verification_issues", 0),
                        "Errors": len(entry.get("errors", [])),
                    })

                st.dataframe(
                    pd.DataFrame(hist_rows),
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "Critical": st.column_config.NumberColumn(help="Critical severity issues"),
                        "Major": st.column_config.NumberColumn(help="Major severity issues"),
                    },
                )
                st.caption(f"{len(history)} conversion(s) logged")

                # Per-file progression
                if len(filenames_seen) > 0:
                    st.markdown("#### Per-file progression")
                    st.caption("Select a file to see how issues changed across runs")

                    selected_file = st.selectbox(
                        "Select file",
                        sorted(filenames_seen),
                        key="history_file_select",
                    )

                    if selected_file:
                        progression = get_file_progression(selected_file)
                        if len(progression) > 1:
                            prog_df = pd.DataFrame(progression)
                            st.dataframe(prog_df, hide_index=True, use_container_width=True)

                            # Show trend
                            first = progression[0]
                            last = progression[-1]
                            delta_critical = last["critical"] - first["critical"]
                            delta_major = last["major"] - first["major"]
                            delta_total = last["total_issues"] - first["total_issues"]

                            tcol1, tcol2, tcol3 = st.columns(3)
                            tcol1.metric("Critical", last["critical"], delta=delta_critical, delta_color="inverse")
                            tcol2.metric("Major", last["major"], delta=delta_major, delta_color="inverse")
                            tcol3.metric("Total issues", last["total_issues"], delta=delta_total, delta_color="inverse")
                        elif len(progression) == 1:
                            st.dataframe(pd.DataFrame(progression), hide_index=True, use_container_width=True)
                            st.caption("Only one conversion logged — run again after fixes to see progression.")
                        else:
                            st.info("No history for this file.")

                # Clear button
                st.markdown("---")
                if st.button("Clear conversion history", type="secondary", key="clear_conv_history"):
                    clear_conversion_history()
                    st.rerun()

        except Exception as e:
            st.warning(f"Conversion history unavailable: {e}")
            logger.error("Conversion history error: %s", e, exc_info=True)

    except ImportError:
        st.warning(
            "pandas is required for the API Costs tab. "
            "Install it with: `pip install pandas`"
        )
    except Exception as e:
        st.error(f"Cost tracker error: {e}")
        logger.error("Cost tracker tab error: %s", e, exc_info=True)


# --- Sidebar: quick-glance cost metric ---
with st.sidebar:
    st.markdown("---")
    st.markdown("### Admin")
    try:
        from utils.cost_logger import get_cost_summary, COST_LOG_FILE
        _sidebar_summary = get_cost_summary()
        if _sidebar_summary["total_calls"] > 0:
            st.metric("Total API cost", f"${_sidebar_summary['total_cost_usd']:.4f}")
            st.caption(
                f"{_sidebar_summary['total_calls']} calls · "
                f"See **API Costs** tab for details"
            )
        else:
            st.caption("No API calls logged yet.")
    except Exception:
        st.caption("Cost tracker unavailable.")


# --- Footer ---
st.markdown("---")
st.caption("UQ Business School — Learning Design Team")
