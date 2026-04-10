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
import shutil
import tempfile
import subprocess
from pathlib import Path


def render_slides_to_images(
    input_bytes: bytes,
    dpi: int = 150,
    timeout: int = 60,
) -> list[bytes]:
    """
    Render all slides in a PPTX to PNG images.

    Args:
        input_bytes: Raw bytes of the .pptx file
        dpi: Resolution for export (150 = good balance of quality/size)
        timeout: Max seconds for LibreOffice conversion

    Returns:
        List of PNG image bytes, one per slide, in order.
        Returns empty list if LibreOffice is not available or conversion fails.
    """
    tmpdir = tempfile.mkdtemp(prefix="uqslide_render_")

    try:
        # Write input PPTX to temp file
        input_path = os.path.join(tmpdir, "input.pptx")
        with open(input_path, "wb") as f:
            f.write(input_bytes)

        # Convert to PNG using LibreOffice headless
        # This produces input_Slide1.png, input_Slide2.png, etc.
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

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            env=env,
        )

        if result.returncode != 0:
            # LibreOffice sometimes outputs one PNG for the whole deck
            # Check if we at least got something
            pass

        # Collect output images
        # LibreOffice may output a single PNG or multiple depending on version
        png_files = sorted(glob.glob(os.path.join(tmpdir, "*.png")))

        if not png_files:
            return []

        # If LibreOffice output a single image, that's the whole deck rendered
        # as one page — we need individual slides. Try PDF intermediate approach.
        if len(png_files) == 1:
            # Try via PDF → per-page PNGs
            return _render_via_pdf(input_path, tmpdir, dpi, timeout, env)

        # Read all PNGs in order
        images = []
        for png_path in png_files:
            with open(png_path, "rb") as f:
                images.append(f.read())

        return images

    except subprocess.TimeoutExpired:
        return []
    except FileNotFoundError:
        # LibreOffice not installed
        return []
    except Exception:
        return []
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _render_via_pdf(
    input_path: str,
    tmpdir: str,
    dpi: int,
    timeout: int,
    env: dict,
) -> list[bytes]:
    """
    Fallback: Convert PPTX → PDF → per-page PNGs.

    LibreOffice reliably converts PPTX to multi-page PDF.
    We then use pdftoppm (poppler) or Pillow to split pages.
    """
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

    result = subprocess.run(
        cmd_pdf,
        capture_output=True,
        text=True,
        timeout=timeout,
        env=env,
    )

    pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    if not pdf_files:
        return []

    pdf_path = pdf_files[0]

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

        subprocess.run(
            cmd_png,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        png_files = sorted(glob.glob(os.path.join(png_dir, "slide-*.png")))
        if png_files:
            images = []
            for png_path in png_files:
                with open(png_path, "rb") as f:
                    images.append(f.read())
            return images
    except FileNotFoundError:
        pass

    # Fallback: use pdf2image (Python library, wraps poppler)
    try:
        from pdf2image import convert_from_path
        pil_images = convert_from_path(pdf_path, dpi=dpi)
        images = []
        for img in pil_images:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            images.append(buf.getvalue())
        return images
    except ImportError:
        pass

    # Last resort: use Pillow to try to read the PDF (limited support)
    try:
        from PIL import Image
        img = Image.open(pdf_path)
        # PDF in Pillow only gives page 0 by default
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
        return images if images else []
    except Exception:
        return []


def render_single_slide(
    input_bytes: bytes,
    slide_index: int,
    dpi: int = 150,
    timeout: int = 60,
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
    all_images = render_slides_to_images(input_bytes, dpi=dpi, timeout=timeout)

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
