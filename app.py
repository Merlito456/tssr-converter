"""
app.py — Streamlit UI for the Ericsson → Nokia TSSR Converter.

Upload an Ericsson Hermes TSSR PDF.
The app:
  1. Extracts site info, rectifier data, and load tables
  2. Extracts embedded images (vicinity map, photos, layouts)
  3. Runs DC load calculations
  4. Fills the Nokia T7 TSSR template (.docx) as-is — no manual tagging
  5. Returns the ready-to-review Nokia TSSR
"""

import os
import tempfile
import traceback
from pathlib import Path

import streamlit as st

from src.extractor import EricssonExtractor
from src.image_extractor import ImageExtractor
from src.processor import TSSRProcessor
from src.generator import TSSRGenerator


# ============================================================
# CONFIG
# ============================================================
TEMPLATE_PATH = "templates/nokia_t7_template.docx"
TEMP_DIR = "assets/temp_images"

st.set_page_config(
    page_title="Ericsson → Nokia TSSR Converter",
    page_icon="📡",
    layout="wide",
)


# ============================================================
# HELPERS
# ============================================================
def safe_image(container, *args, **kwargs):
    """
    Streamlit-compatible image renderer.
    Tries the modern `use_container_width` arg first, falls back
    to the deprecated `use_column_width` for older versions.
    """
    try:
        container.image(*args, use_container_width=True, **kwargs)
    except TypeError:
        try:
            container.image(*args, use_column_width=True, **kwargs)
        except TypeError:
            # Last resort: no sizing kwarg at all
            container.image(*args, **kwargs)


def ensure_dir(path: str):
    Path(path).mkdir(parents=True, exist_ok=True)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown(f"**Template:** `{TEMPLATE_PATH}`")

    if not Path(TEMPLATE_PATH).exists():
        st.error("Template not found in repo. Upload it below.")
        uploaded_tpl = st.file_uploader(
            "Upload Nokia T7 template (.docx)",
            type=["docx"],
            key="tpl_uploader",
        )
        if uploaded_tpl is not None:
            ensure_dir("templates")
            with open(TEMPLATE_PATH, "wb") as f:
                f.write(uploaded_tpl.getbuffer())
            st.success("Template saved. Click 'Reboot' in the top-right menu.")
            st.stop()
    else:
        st.success("Template ready ✅")

    st.divider()
    st.markdown("### 📋 What this app does")
    st.markdown(
        "1. Reads an Ericsson Hermes TSSR PDF\n"
        "2. Extracts site info, rectifiers, loads, and images\n"
        "3. Runs Nokia DC load calculations\n"
        "4. Fills the Nokia T7 template **as-is** (no tagging required)\n"
        "5. Returns a ready-to-review `.docx`"
    )

    st.divider()
    with st.expander("🔧 Debug options"):
        show_raw_data = st.checkbox("Show raw extracted data", value=False)
        show_images = st.checkbox("Show extracted images", value=True)


# ============================================================
# MAIN
# ============================================================
st.title("📡 Ericsson → Nokia TSSR Converter")
st.caption(
    "Upload an Ericsson Hermes TSSR PDF. The app extracts site data, "
    "performs DC calculations, extracts photos, and generates a "
    "Nokia-formatted TSSR (.docx)."
)

uploaded_pdf = st.file_uploader("Upload Ericsson TSSR PDF", type=["pdf"])

if uploaded_pdf is not None:
    # Save to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_pdf.getbuffer())
        pdf_path = tmp.name

    st.info(f"Loaded: **{uploaded_pdf.name}**")

    # ---------------------------------------------------------
    # CONVERT BUTTON
    # ---------------------------------------------------------
    if st.button("🚀 Convert to Nokia TSSR", type="primary"):

        progress = st.progress(0, text="Starting...")

        # ---------- 1. TEXT EXTRACTION ----------
        progress.progress(15, text="📄 Extracting site & rectifier data...")
        try:
            extractor = EricssonExtractor(pdf_path)
            data = extractor.extract_all()
            st.success(
                f"Extracted site: **{data.get('site_name', 'N/A')}** "
                f"({data.get('site_id', 'N/A')})"
            )
        except Exception as e:
            st.error(f"Extraction failed: {e}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            st.stop()

        # ---------- 2. IMAGE EXTRACTION ----------
        progress.progress(40, text="🖼️ Extracting images from PDF...")
        try:
            img_extractor = ImageExtractor(pdf_path, output_dir=TEMP_DIR)
            image_map = img_extractor.extract_all()
            total_imgs = sum(len(v) for v in image_map.values())
            st.info(
                f"Found **{total_imgs}** images across "
                f"**{len(image_map)}** sections"
            )
        except Exception as e:
            st.error(f"Image extraction failed: {e}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            image_map = {}
            img_extractor = None

        # ---------- 3. IMAGE PREVIEW ----------
        if show_images and image_map:
            with st.expander("Preview extracted images", expanded=False):
                for section, files in image_map.items():
                    st.markdown(f"**{section}** — {len(files)} image(s)")
                    if not files:
                        continue
                    cols = st.columns(min(len(files), 3))
                    for i, f in enumerate(files[:3]):
                        with cols[i]:
                            try:
                                safe_image(cols[i], f)
                            except Exception as img_err:
                                st.warning(f"Cannot show {Path(f).name}: {img_err}")
                            st.caption(Path(f).name)

        # ---------- 4. CALCULATIONS ----------
        progress.progress(65, text="🧮 Running DC calculations...")
        try:
            processor = TSSRProcessor(data)
            processor.calculate_dc_load()
            context = processor.build_template_context()
        except Exception as e:
            st.error(f"Calculation failed: {e}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            st.stop()

        # Show metrics
        st.subheader("📊 Calculation Results")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Full Load", f"{context.get('total_full_load', 0)} A")
        c2.metric("Available Cap.", f"{context.get('available_capacity', 0)} A")
        c3.metric("Battery Back-Up", f"{context.get('but_hours', 0)} hr")
        c4.metric("Utilization", f"{context.get('percent_util_tlc', 0)}%")

        if show_raw_data:
            with st.expander("🔎 Raw extracted data"):
                st.json(data)
            with st.expander("🔎 Full template context"):
                safe_ctx = {
                    k: (str(v) if not isinstance(v, (str, int, float, list, dict))
                        else v)
                    for k, v in context.items()
                }
                st.json(safe_ctx)

        # ---------- 5. GENERATE DOCX ----------
        progress.progress(85, text="📝 Generating Nokia TSSR...")
        try:
            generator = TSSRGenerator(TEMPLATE_PATH)
            output_path = (
                f"MIN{data.get('site_id', 'XXXX')}_NOKIA_TSSR.docx"
            )
            generator.generate(context, image_map, output_path)
        except Exception as e:
            st.error(f"Document generation failed: {e}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            st.stop()

        progress.progress(100, text="✅ Done!")

        # ---------- 6. DOWNLOAD ----------
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Nokia TSSR (.docx)",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document",
            )
        st.success("Your Nokia TSSR is ready. Download it above.")

        # ---------- 7. CLEANUP ----------
        try:
            if img_extractor:
                img_extractor.cleanup()
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(output_path):
                os.remove(output_path)
        except Exception:
            pass  # don't fail the UX over cleanup
