import streamlit as st
from pptx import Presentation
import io
import base64

def standardize_ppt(source_ppt, target_ppt):
    """
    Transfers slides from source_ppt (no theme) to target_ppt (with theme),
    applying the target's theme and formatting.
    """
    target_presentation = Presentation(target_ppt)
    source_presentation = Presentation(source_ppt)
    
    # Use the target presentation's theme and layouts
    for slide in source_presentation.slides:
        # Add a new slide to the target presentation with a default layout
        new_slide = target_presentation.slides.add_slide(target_presentation.slide_layouts[1])  # Title and Content
        
        # Copy shapes (text, images, tables) from source to target
        for shape in slide.shapes:
            if shape.has_text_frame:
                new_shape = new_slide.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                new_shape.text_frame.text = shape.text_frame.text
                # Apply target font (simplified)
                for paragraph in new_shape.text_frame.paragraphs:
                    paragraph.font.name = "Calibri"  # Default font (adjust as needed)
            
            elif shape.has_table:
                # Copy tables (simplified)
                table = shape.table
                new_table = new_slide.shapes.add_table(
                    table.rows.count, table.columns.count,
                    shape.left, shape.top, shape.width, shape.height
                ).table
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.cell(i, j).text = cell.text
    
    return target_presentation

def get_binary_file_downloader_html(bin_file, file_label='File'):
    """Generates a download link for a binary file."""
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{bin_file}">Download {file_label}</a>'
    return href

# Streamlit UI
st.title("📊 PowerPoint Standardizer")
st.markdown("""
Upload two PowerPoint files:
1. **Source PPT** (no theme, content to transfer).
2. **Target PPT** (with theme, defines formatting).
""")

# File Upload
source_ppt = st.file_uploader("Upload Source PPT (No Theme)", type=["pptx"])
target_ppt = st.file_uploader("Upload Target PPT (With Theme)", type=["pptx"])

if st.button("Standardize PowerPoints") and source_ppt and target_ppt:
    with st.spinner("Processing..."):
        # Save uploaded files temporarily
        with open("temp_source.pptx", "wb") as f:
            f.write(source_ppt.getbuffer())
        with open("temp_target.pptx", "wb") as f:
            f.write(target_ppt.getbuffer())
        
        # Standardize
        standardized_ppt = standardize_ppt("temp_source.pptx", "temp_target.pptx")
        standardized_ppt.save("standardized_output.pptx")
        
        # Provide download link
        st.success("Done!")
        st.markdown(get_binary_file_downloader_html("standardized_output.pptx", "Standardized PPT"), unsafe_allow_html=True)

# Sidebar note
st.sidebar.markdown("""
### Instructions:
1. Upload the **unformatted PPT** (content to transfer).
2. Upload the **formatted PPT** (theme to apply).
3. Click **Standardize PowerPoints**.
4. Download the result.
""")