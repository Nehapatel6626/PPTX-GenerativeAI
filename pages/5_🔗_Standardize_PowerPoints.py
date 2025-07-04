
import streamlit as st
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt
import uuid
from pathlib import Path
import re

# Setup temp folder
TEMP_DIR = Path("temp_pptx")
TEMP_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="🧠 AI PPTX Formatter", page_icon="🤖")
st.title("🤖 Smart PPTX Styler with AI Layout & Clean Formatting")

file1 = st.file_uploader("📤 Upload Content File (File 1)", type=["pptx"])
file2 = st.file_uploader("🎨 Upload Theme File (File 2)", type=["pptx"])

if file1 and file2:
    file1_path = TEMP_DIR / f"file1_{uuid.uuid4().hex}.pptx"
    file2_path = TEMP_DIR / f"file2_{uuid.uuid4().hex}.pptx"
    output_path = TEMP_DIR / f"output_{uuid.uuid4().hex}.pptx"

    with open(file1_path, "wb") as f:
        f.write(file1.read())
    with open(file2_path, "wb") as f:
        f.write(file2.read())

    # Load presentations
    prs_content = Presentation(file1_path)
    prs_theme = Presentation(file2_path)

    # Remove default slides from theme/template
    while len(prs_theme.slides) > 0:
        r_id = prs_theme.slides._sldIdLst[0].rId
        prs_theme.part.drop_rel(r_id)
        del prs_theme.slides._sldIdLst[0]

    # AI layout logic
    def choose_best_layout_ai(slide):
        text_boxes = [s for s in slide.shapes if s.has_text_frame]
        all_text = "\n".join([s.text.strip() for s in text_boxes if s.text.strip()])
        text_len = len(all_text)

        lower = all_text.lower()
        if "introduction" in lower or "overview" in lower:
            return 0  # Title Slide
        if "case study" in lower or "conclusion" in lower:
            return 3  # Section Header
        if "question" in lower or "answer key" in lower:
            return 5  # Comparison
        if re.search(r"[\u2022\-\*\•]", all_text) or all_text.count("\n") > 5:
            return 4  # Two Content
        if text_len < 120:
            return 6  # Title Only
        if text_len > 400:
            return 2  # Title and Content
        return 1  # Fallback: Title and Content

    st.markdown("### 🔍 AI is analyzing slide content and applying formatting...")

    for idx, slide in enumerate(prs_content.slides):
        layout_index = choose_best_layout_ai(slide)
        layout = prs_theme.slide_layouts[layout_index]
        new_slide = prs_theme.slides.add_slide(layout)

        # Extract title and body text
        title_text = ""
        body_text = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                if not title_text:
                    title_text = shape.text.strip()
                else:
                    body_text += shape.text.strip() + "\n"

        # Optional: truncate long body
        if len(body_text) > 1000:
            body_text = body_text[:1000] + "\n[...]"

        # Apply title
        try:
            if new_slide.shapes.title:
                title_shape = new_slide.shapes.title
                title_shape.text = title_text.strip()
                title_shape.text_frame.word_wrap = True
                title_shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except:
            pass

        # Remove all extra placeholders (e.g., "Click to add text")
        for shape in list(new_slide.placeholders):
            if shape.placeholder_format.type not in [0, 1, 2]:  # Only keep title, subtitle, body
                sp = shape._element
                sp.getparent().remove(sp)

        # Apply body text (to BODY placeholder only)
        for shape in new_slide.placeholders:
            if shape.placeholder_format.type == 2:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = body_text.strip()
                run.font.size = Pt(14)  # Clamp font size to readable level
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                break

        layout_name = prs_theme.slide_layouts[layout_index].name
        st.success(f"✅ Slide {idx+1}: Applied layout — `{layout_name}`")

    # Save output
    prs_theme.save(output_path)
    st.success("🎉 Final PPTX generated successfully!")

    with open(output_path, "rb") as f:
        st.download_button("⬇️ Download Cleaned Themed PPTX", f, file_name="themed_output_cleaned.pptx")
