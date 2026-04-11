"""
Slide-to-image renderer using LibreOffice headless.

Converts a PPTX file to PNG images (one per slide) using LibreOffice's
built-in export. This gives pixel-perfect rendering of themes, gradients,
fonts, and layout — exactly what Claude Vision needs for classification.

Requires LibreOffice to be installed (available via system package or Docker).
"""

import io
import os
import glob
import logging
import shutil
import tempfile
import subprocess
from pathlib import Path

logger = logging.getLogger("uqslide.renderer")


def render_slides_to_images(
    input_bytes: bytes,
    dpi: int = 150,
    timeout: int = 120,
) -> tuple[list[bytes], str]:
    """
    Render all slides in a PPTX to PNG images.

    Args:
        input_bytes: Raw bytes of the .pptx file
        dpi: Resolution for export (150 = good balance of quality/size)
        timeout: Max seconds for LibreOffice conversion (increased to 120 for large decks)

    Returns:
        Tuple of (list of PNG image bytes, diagnostic_message).
        Returns (empty list, error_message) if rendering failed.
    """
    tmpdir = tempfile.mkdtemp(prefix="uqslide_render_")
    diag = ""

    try:
        # Check available disk space and memory
        import shutil as _shutil
        disk = _shutil.disk_usage(tmpdir)
        free_mb = disk.free / (1024 * 1024)
        logger.info("Render tmpdir: %s — free disk: %.0f MB", tmpdir, free_mb)
        if free_mb < 100:
            diag = f"Low disk space: {free_mb:.0f} MB free"
            logger.warning(diag)

        # Write input PPTX to temp file
        input_path = os.path.join(tmpdir, "input.pptx")
        with open(input_path, "wb") as f:
            f.write(input_bytes)
        logger.info("Wrote input PPTX: %d bytes to %s", len(input_bytes), input_path)

        # Convert to PNG using LibreOffice headless
        cmd = [
            "libreoffice",
            "--headless",
            "--norestore",
            "--convert-to", "png",
            "--outdir", tmpdir,
            input_path,
        ]

        # Set HOME to tmpdir to avoid LibreOffice profile lock issues
        env = os.environ.copy()
        env["HOME"] = tmpdir
        # Disable Java (common source of crashes in headless mode)
        env["SAL_USE_VCLPLUGIN"] = "svp"

        logger.info("Running LibreOffice PNG export: %s", " ".join(cmd))
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            env=env,
        )

        logger.info("LibreOffice PNG exit code: %d", result.returncode)
        if result.stdout:
            logger.info("LibreOffice stdout: %s", result.stdout[:500])
        if result.stderr:
            logger.info("LibreOffice stderr: %s", result.stderr[:500])

        if result.returncode != 0:
            diag = f"LibreOffice PNG export failed (code {result.returncode}): {(result.stderr or result.stdout or 'no output')[:300]}"
            logger.warning(diag)

        # List ALL files in tmpdir for debugging
        all_files = os.listdir(tmpdir)
        logger.info("Files in tmpdir after PNG export: %s", all_files)

        # Collect output images
        png_files = sorted(glob.glob(os.path.join(tmpdir, "*.png")))

        if not png_files:
            logger.warning("LibreOffice produced no PNG files — trying PDF route")
            images, pdf_diag = _render_via_pdf(input_path, tmpdir, dpi, timeout, env)
            if not images:
                diag = (diag + " | " if diag else "") + f"PDF fallback also failed: {pdf_diag}"
            return images, diag or pdf_diag

        # If LibreOffice output a single image, try PDF route for per-slide
        if len(png_files) == 1:
            logger.info("LibreOffice produced 1 PNG — trying PDF route for per-slide images")
            images, pdf_diag = _render_via_pdf(input_path, tmpdir, dpi, timeout, env)
            if images:
                return images, pdf_diag or "PDF route: success"
            # Fall back to single image
            logger.info("PDF route failed, using single PNG")

        # Read all PNGs in order
        images = []
        for png_path in png_files:
            with open(png_path, "rb") as f:
                images.append(f.read())

        diag = diag or f"Rendered {len(images)} slide images via LibreOffice"
        return images, diag

    except subprocess.TimeoutExpired:
        diag = f"LibreOffice timed out after {timeout} seconds"
        logger.error(diag)
        return [], diag
    except FileNotFoundError:
        diag = "LibreOffice not found — is it installed?"
        logger.error(diag)
        return [], diag
    except Exception as e:
        diag = f"Unexpected rendering error: {e}"
        logger.error(diag, exc_info=True)
        return [], diag
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _render_via_pdf(
    input_path: str,
    tmpdir: str,
    dpi: int,
    timeout: int,
    env: dict,
) -> tuple[list[bytes], str]:
    """
    Fallback: Convert PPTX → PDF → per-page PNGs.

    LibreOffice reliably converts PPTX to multi-page PDF.
    We then use pdftoppm (poppler) or Pillow to split pages.

    Returns:
        Tuple of (list of PNG bytes, diagnostic_message)
    """
    diag = ""
    pdf_dir = os.path.join(tmpdir, "pdf_out")
    os.makedirs(pdf_dir, exist_ok=True)

    # Step 1: PPTX → PDF
    cmd_pdf = [
        "libreoffice",
        "--headless",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", pdf_dir,
        input_path,
    ]

    logger.info("Running LibreOffice PDF export: %s", " ".join(cmd_pdf))
    result = subprocess.run(
        cmd_pdf,
        capture_output=True,
        text=True,
        timeout=timeout,
        env=env,
    )

    logger.info("LibreOffice PDF exit code: %d", result.returncode)
    if result.stdout:
        logger.info("LO PDF stdout: %s", result.stdout[:500])
    if result.stderr:
        logger.info("LO PDF stderr: %s", result.stderr[:500])

    pdf_dir_files = os.listdir(pdf_dir)
    logger.info("Files in pdf_dir: %s", pdf_dir_files)

    pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    if not pdf_files:
        diag = f"PDF conversion produced no output (code {result.returncode}): {(result.stderr or result.stdout or 'no output')[:300]}"
        logger.error(diag)
        return [], diag

    pdf_path = pdf_files[0]
    pdf_size = os.path.getsize(pdf_path)
    logger.info("PDF created: %s (%d bytes)", pdf_path, pdf_size)

    # Step 2: PDF → PNGs using pdftoppm (poppler-utils)
    png_dir = os.path.join(tmpdir, "png_out")
    os.makedirs(png_dir, exist_ok=True)

    # Try pdftoppm first (best quality)
    try:
        cmd_png = [
            "pdftoppm",
            "-png",
            "-r", str(dpi),
            pdf_path,
            os.path.join(png_dir, "slide"),
        ]

        logger.info("Running pdftoppm: %s", " ".join(cmd_png))
        ppm_result = subprocess.run(
            cmd_png,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        logger.info("pdftoppm exit code: %d", ppm_result.returncode)
        if ppm_result.stderr:
            logger.info("pdftoppm stderr: %s", ppm_result.stderr[:300])

        png_files = sorted(glob.glob(os.path.join(png_dir, "slide-*.png")))
        if png_files:
            logger.info("pdftoppm produced %d slide images", len(png_files))
            images = []
            for png_path in png_files:
                with open(png_path, "rb") as f:
                    images.append(f.read())
            return images, f"PDF route: {len(images)} images via pdftoppm"
        else:
            diag = f"pdftoppm produced no PNGs (code {ppm_result.returncode}): {(ppm_result.stderr or 'no output')[:200]}"
            logger.warning(diag)
    except FileNotFoundError:
        diag = "pdftoppm not found"
        logger.warning(diag)

    # Fallback: use pdf2image (Python library, wraps poppler)
    try:
        from pdf2image import convert_from_path
        pil_images = convert_from_path(pdf_path, dpi=dpi)
        images = []
        for img in pil_images:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            images.append(buf.getvalue())
        return images, f"PDF route: {len(images)} images via pdf2image"
    except ImportError:
        logger.warning("pdf2image not available")
    except Exception as e:
        logger.warning("pdf2image failed: %s", e)

    # Last resort: use Pillow to try to read the PDF (limited support)
    try:
        from PIL import Image
        img = Image.open(pdf_path)
        images = []
        try:
            page_num = 0
            while True:
                img.seek(page_num)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                images.append(buf.getvalue())
                page_num += 1
        except EOFError:
            pass
        if images:
            return images, f"PDF route: {len(images)} images via Pillow"
        return [], diag or "Pillow could not extract pages from PDF"
    except Exception as e:
        return [], diag or f"All PDF render methods failed: {e}"


def render_single_slide(
    input_bytes: bytes,
    slide_index: int,
    dpi: int = 150,
    timeout: int = 120,
) -> bytes | None:
    """
    Render a single slide to PNG.

    This renders ALL slides then returns only the requested one.
    For batch processing, use render_slides_to_images() instead.

    Args:
        input_bytes: Raw bytes of the .pptx file
        slide_index: 0-based slide index
        dpi: Resolution for export
        timeout: Max seconds for conversion

    Returns:
        PNG bytes for the requested slide, or None if rendering failed.
    """
    all_images, _ = render_slides_to_images(input_bytes, dpi=dpi, timeout=timeout)

    if not all_images or slide_index >= len(all_images):
        return None

    return all_images[slide_index]


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is installed and accessible."""
    try:
        result = subprocess.run(
            ["libreoffice", "--version"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False
