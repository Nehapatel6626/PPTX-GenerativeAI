from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import streamlit as st
import os

def get_theme_font(presentation):
    """Safely get theme font from presentation"""
    try:
        # Try to get from theme if available
        return presentation.slide_master.theme.fonts.latin_typeface
    except AttributeError:
        try:
            # Fallback to first slide's title font
            for slide in presentation.slides:
                if slide.shapes.title:
                    return slide.shapes.title.text_frame.paragraphs[0].font.name
        except:
            # Final fallback to Calibri
            return "Calibri"

def adjust_text_frame(text_frame, target_font):
    """Ensure text fits within slide boundaries"""
    for paragraph in text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.font.name = target_font
        paragraph.font.size = Pt(12)
        paragraph.space_after = Pt(0)
        paragraph.font.color.rgb = RGBColor(0, 0, 0)  # Black text
    
    # Shrink text if needed
    if len(text_frame.text) > 500:
        for paragraph in text_frame.paragraphs:
            paragraph.font.size = Pt(10)

def standardize_ppt(source_path, target_path, output_path):
    """Improved standardization with proper layout handling"""
    source = Presentation(source_path)
    target = Presentation(target_path)
    
    # Get target properties safely
    target_font = get_theme_font(target)
    margin = Inches(0.5)
    
    for source_slide in source.slides:
        # Create blank slide
        blank_slide_layout = target.slide_layouts[6]
        new_slide = target.slides.add_slide(blank_slide_layout)
        
        # Calculate available space
        slide_width = target.slide_width - (2 * margin)
        left_pos = margin
        
        # Add title if exists
        if source_slide.shapes.title:
            title = new_slide.shapes.add_textbox(
                left_pos, margin, slide_width, Inches(0.8)
            )
            title.text = source_slide.shapes.title.text
            adjust_text_frame(title.text_frame, target_font)
            top_pos = margin + Inches(1)
        else:
            top_pos = margin
        
        # Add content
        content_height = target.slide_height - top_pos - margin
        content = new_slide.shapes.add_textbox(
            left_pos, top_pos, slide_width, content_height
        )
        
        # Aggregate all text content
        all_text = []
        for shape in source_slide.shapes:
            if shape.has_text_frame and shape != source_slide.shapes.title:
                all_text.append(shape.text)
        
        content.text = "\n\n".join(all_text)
        adjust_text_frame(content.text_frame, target_font)
    
    target.save(output_path)

# Streamlit UI
st.set_page_config(page_title="PPT Standardizer Pro", layout="wide")
st.title("🎯 PowerPoint Standardizer Pro")

col1, col2 = st.columns(2)
with col1:
    source_ppt = st.file_uploader("Upload Content PPT", type=["pptx"])
with col2:
    target_ppt = st.file_uploader("Upload Design Template PPT", type=["pptx"])

if st.button("✨ Standardize Presentation", type="primary"):
    if source_ppt and target_ppt:
        with st.spinner("Creating professionally formatted slides..."):
            # Save temp files
            source_path = "temp_source.pptx"
            target_path = "temp_target.pptx"
            with open(source_path, "wb") as f:
                f.write(source_ppt.getvalue())
            with open(target_path, "wb") as f:
                f.write(target_ppt.getvalue())
            
            # Process files
            output_path = "professional_output.pptx"
            standardize_ppt(source_path, target_path, output_path)
            
            # Show download button
            with open(output_path, "rb") as f:
                ppt_bytes = f.read()
            
            st.success("Formatting complete!")
            st.download_button(
                label="⬇️ Download Professional PPT",
                data=ppt_bytes,
                file_name="formatted_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
            # Clean up temp files
            os.remove(source_path)
            os.remove(target_path)
    else:
        st.warning("Please upload both files to proceed")

st.sidebar.markdown("""
### Key Features:
- **Perfect Layout Adaptation**
- **Automatic Text Formatting**
- **Theme Consistency**
- **Professional Output**
- **Robust Error Handling**
""")