import streamlit as st
import os
import tempfile
from src.extractor import EricssonExtractor
from src.image_extractor import ImageExtractor
from src.processor import TSSRProcessor
from src.generator import TSSRGenerator

st.set_page_config(
    page_title="Ericsson → Nokia TSSR Converter",
    page_icon="📡",
    layout="wide",
)

st.title("📡 Ericsson → Nokia TSSR Converter")
st.caption("Upload an Ericsson Hermes TSSR PDF. The app extracts site data, "
           "performs DC calculations, extracts photos, and generates a "
           "Nokia-formatted TSSR (.docx).")

# ---------- Sidebar ----------
with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown("**Template:** `templates/nokia_template.docx`")
    if not os.path.exists("templates/nokia_template.docx"):
        st.error("Template not found! Upload it below.")
        tpl = st.file_uploader("Upload Nokia template (.docx)", type="docx")
        if tpl:
            os.makedirs("templates", exist_ok=True)
            with open("templates/nokia_template.docx", "wb") as f:
                f.write(tpl.getbuffer())
            st.success("Template saved. Refresh the page.")
            st.stop()
    else:
        st.success("Template ready ✅")

    st.divider()
    st.markdown("### 📋 Expected Output")
    st.markdown(
        "- Site Info table\n"
        "- DC Load Calculations\n"
        "- Rectifier Sufficiency\n"
        "- Battery Back-Up Time\n"
        "- Vicinity Map & Photos\n"
        "- Equipment Layouts"
    )

# ---------- Main ----------
uploaded = st.file_uploader("Upload Ericsson TSSR PDF", type="pdf")

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded.getbuffer())
        pdf_path = tmp.name

    st.info(f"Loaded: **{uploaded.name}**")

    if st.button("🚀 Convert to Nokia TSSR", type="primary"):

        progress = st.progress(0, text="Starting...")

        # --- 1. Text extraction ---
        progress.progress(15, text="📄 Extracting site & rectifier data...")
        extractor = EricssonExtractor(pdf_path)
        data = extractor.extract_all()
        st.success(f"Extracted Site: **{data.get('site_name')}** "
                   f"({data.get('site_id')})")

        # --- 2. Image extraction ---
        progress.progress(40, text="🖼️ Extracting images from PDF...")
        img_extractor = ImageExtractor(pdf_path)
        image_map = img_extractor.extract_all()

        total_imgs = sum(len(v) for v in image_map.values())
        st.info(f"Found **{total_imgs}** images across "
                f"**{len(image_map)}** sections")

        with st.expander("View extracted image sections"):
            for section, files in image_map.items():
                st.write(f"**{section}** — {len(files)} image(s)")
                cols = st.columns(min(len(files), 3))
                for i, f in enumerate(files[:3]):
                    cols[i].image(f, use_column_width=True)

        # --- 3. Calculations ---
        progress.progress(65, text="🧮 Running DC calculations...")
        processor = TSSRProcessor(data)
        processor.calculate_dc_load()
        context = processor.build_template_context()

        st.subheader("📊 Calculation Results")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Full Load", f"{context['total_full_load']} A")
        c2.metric("Available Cap.", f"{context['available_capacity']} A")
        c3.metric("Battery Back-Up", f"{context['but_hours']} hr")
        c4.metric("Utilization", f"{context['percent_util_tlc']}%")

        # --- 4. Generate DOCX ---
        progress.progress(85, text="📝 Generating Nokia TSSR...")
        generator = TSSRGenerator("templates/nokia_template.docx")
        output_path = f"MIN{data.get('site_id','XXXX')}_NOKIA_TSSR.docx"
        generator.generate(context, image_map, output_path)

        progress.progress(100, text="✅ Done!")

        # --- 5. Download ---
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Nokia TSSR (.docx)",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document",
            )

        # --- Cleanup ---
        img_extractor.cleanup()
        os.remove(pdf_path)
        os.remove(output_path)
