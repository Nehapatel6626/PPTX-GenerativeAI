import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import tempfile
import os
from uuid import uuid4

# Load the API key securely from environment variable
api_key = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=api_key)

# Default theme configurations
DEFAULT_THEMES = {
    "Professional Blue": {
        "background": RGBColor(12, 35, 64),
        "title_color": RGBColor(255, 255, 255),
        "text_color": RGBColor(200, 200, 200),
        "accent": RGBColor(0, 112, 192),
        "title_font": "Calibri",
        "text_font": "Calibri"
    },
    "Modern Green": {
        "background": RGBColor(22, 82, 66),
        "title_color": RGBColor(255, 255, 255),
        "text_color": RGBColor(220, 220, 220),
        "accent": RGBColor(76, 175, 80),
        "title_font": "Arial",
        "text_font": "Arial"
    },
    "Light Corporate": {
        "background": RGBColor(255, 255, 255),
        "title_color": RGBColor(13, 71, 161),
        "text_color": RGBColor(33, 33, 33),
        "accent": RGBColor(25, 118, 210),
        "title_font": "Segoe UI",
        "text_font": "Segoe UI"
    },
    "Dark Tech": {
        "background": RGBColor(33, 33, 33),
        "title_color": RGBColor(0, 200, 255),
        "text_color": RGBColor(200, 200, 200),
        "accent": RGBColor(0, 150, 255),
        "title_font": "Consolas",
        "text_font": "Consolas"
    }
}

def hex_to_rgb(hex_color):
    """Convert hex color to RGBColor"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

def extract_theme_from_docx(uploaded_file):
    """Extract theme colors and fonts from an uploaded Word document"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name
    
    doc = Document(tmp_file_path)
    theme = {
        "background": RGBColor(255, 255, 255),
        "title_color": RGBColor(0, 0, 0),
        "text_color": RGBColor(0, 0, 0),
        "accent": RGBColor(79, 129, 189),
        "title_font": "Calibri",
        "text_font": "Calibri",
        "template_path": tmp_file_path
    }
    
    try:
        for para in doc.paragraphs:
            if para.text.strip():
                try:
                    run = para.runs[0]
                    if run.font.size and run.font.size.pt > 14:
                        theme["title_color"] = run.font.color.rgb
                        theme["title_font"] = run.font.name or "Calibri"
                    else:
                        theme["text_color"] = run.font.color.rgb
                        theme["text_font"] = run.font.name or "Calibri"
                except:
                    continue
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    try:
                        run = cell.paragraphs[0].runs[0]
                        if run.font.color.rgb:
                            theme["accent"] = run.font.color.rgb
                            break
                    except:
                        continue
    except Exception as e:
        st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
    return theme

def generate_document_content(topic, page_count):
    model = genai.GenerativeModel('gemini-3.5-flash-lite')
    prompt = f"""Create a comprehensive document on '{topic}' with exactly {page_count} pages. 
    For each page, provide:
    1. A clear title in [Title:] format
    2. 3-5 detailed bullet points in [Content:] format (each point should be 2-3 lines/40-60 words)
    3. Optional notes in [Notes:] format
    4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
    Structure your response like this:
    [Title:] Page Title
    [Layout:] title-content
    [Content:]
    - Main Point 1: Detailed explanation spanning 1-2 lines with supporting information that provides context and value to the audience. This makes each point substantial.
    - Main Point 2: Another complete thought with sufficient detail to stand alone as a mini-paragraph, giving the audience concrete information they can use.
    - Main Point 3: Final point with enough depth to be meaningful, typically consisting of 2-3 sentences that develop a complete idea.
    [Notes:] Additional notes here
    
    Important guidelines:
    - Do not give only placeholders or labels. Write full, rich content for each bullet point.
    - Apply the same for the Questionnaire page too, don't just keep placeholders, generate 10 questions.
    - Avoid general outlines — generate fully written content as if it's going directly into a document.
    - Each bullet point should be 2-3 lines (30-50 words)
    - Provide complete thoughts with supporting details
    - Maintain parallel structure across points
    - Avoid single-sentence bullet points
    - Focus on substance over brevity
    
    Include these sections (adjust based on requested page count):
    - Title page
    - Introduction/Overview
    - Key Concepts
    - Detailed Analysis
    - Case Studies/Examples
    - Applications
    - Challenges
    - Future Trends
    - Conclusion
    - Questionnaire 

    After generating the content for the above document sections, titled [Title:], generate a Questionnaire Page containing:
    - 10 multiple-choice questions testing the user’s understanding of the document
    - Each question should have 4 options (A, B, C, D)
    - Clearly indicate the correct answer after each question using the format: [Correct Answer: C]
    - Make the questions relevant and cover all key areas: introduction, concepts, analysis, case studies, applications, challenges, and future trends.
    - Make sure the questions test the user's comprehension of the previous pages' content and are well-distributed across different topics in the document.

    Please ensure the document content is presented first, followed by a clear separator, and then the questionnaire.

    Begin the content generation now.
    """
    
    response = model.generate_content(prompt)
    return response.text

def parse_document_content(document_text):
    pages = []
    current_page = {}
    
    for line in document_text.split('\n'):
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('[Title:]'):
            if current_page:
                pages.append(current_page)
            current_page = {
                'title': line.replace('[Title:]', '').strip(),
                'content': [],
                'notes': '',
                'layout': 'title-content'
            }
        elif line.startswith('[Content:]'):
            content = line.replace('[Content:]', '').strip()
            if content:
                current_page['content'].append(content)
        elif line.startswith('[Notes:]'):
            current_page['notes'] = line.replace('[Notes:]', '').strip()
        elif line.startswith('[Layout:]'):
            layout = line.replace('[Layout:]', '').strip().lower()
            valid_layouts = ['title-only', 'title-content', 'two-column', 'section-header']
            current_page['layout'] = layout if layout in valid_layouts else 'title-content'
        elif current_page.get('content') is not None and line.startswith('-'):
            point = line[1:].strip()
            if ':' in point:
                point = point.split(':')[0].strip()
            current_page['content'].append(point)
    
    if current_page:
        pages.append(current_page)
    
    return pages

def create_word_document(pages_data, theme):
    """Create a Word document"""
    doc = Document()
    
    # Set document margins
    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
    
    for page in pages_data:
        # Add title
        title = doc.add_heading(page['title'], level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.runs[0]
        title_run.font.color.rgb = theme['title_color']
        title_run.font.size = Pt(18)
        title_run.font.name = theme['title_font']
        
        # Add content
        if page['layout'] != 'title-only':
            if page['layout'] == 'two-column':
                mid_point = len(page['content']) // 2
                left_content = page['content'][:mid_point]
                right_content = page['content'][mid_point:]
                
                # Create a two-column table
                table = doc.add_table(rows=max(len(left_content), len(right_content)), cols=2)
                table.autofit = True
                
                for i in range(max(len(left_content), len(right_content))):
                    left_cell = table.rows[i].cells[0]
                    right_cell = table.rows[i].cells[1]
                    
                    if i < len(left_content):
                        p = left_cell.add_paragraph(f"• {left_content[i]}")
                        p.runs[0].font.color.rgb = theme['text_color']
                        p.runs[0].font.size = Pt(12)
                        p.runs[0].font.name = theme['text_font']
                    
                    if i < len(right_content):
                        p = right_cell.add_paragraph(f"• {right_content[i]}")
                        p.runs[0].font.color.rgb = theme['text_color']
                        p.runs[0].font.size = Pt(12)
                        p.runs[0].font.name = theme['text_font']
            else:
                for point in page['content']:
                    p = doc.add_paragraph(f"• {point}")
                    p.runs[0].font.color.rgb = theme['text_color']
                    p.runs[0].font.size = Pt(12)
                    p.runs[0].font.name = theme['text_font']
        
        # Add notes as footer if present
        if page['notes']:
            section = doc.sections[-1]
            footer = section.footer
            footer_para = footer.paragraphs[0]
            footer_para.text = page['notes']
            footer_para.runs[0].font.size = Pt(10)
            footer_para.runs[0].font.name = theme['text_font']
        
        doc.add_page_break()
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    # Clean up temporary template file
    if "template_path" in theme and os.path.exists(theme["template_path"]):
        os.unlink(theme["template_path"])
    
    return doc_io

def main():
    st.set_page_config(page_title="Word Document Generator", layout="wide")
    
    st.title("Advanced Word Document Generator")
    st.markdown("Create professional Word documents with AI")
    
    # Initialize session state for custom themes if not exists
    if 'custom_themes' not in st.session_state:
        st.session_state.custom_themes = {}
    
    # Combine default and custom themes
    ALL_THEMES = {**DEFAULT_THEMES, **st.session_state.custom_themes}
    col1, col2 = st.columns([3, 1])
    
    with col1:
        topic = st.text_input("Document Topic:", 
                            placeholder="Enter your topic (e.g., 'AI in Healthcare')")
        
        with st.expander("Advanced Options"):
            page_count = st.slider("Number of Pages:", 5, 20, 10)
            
            theme_option = st.radio("Theme Selection Method:", 
                                  ["Predefined Theme", "Custom Theme", "Example-Based Theme"])
            
            theme = DEFAULT_THEMES["Professional Blue"]
            
            if theme_option == "Predefined Theme":
                theme_name = st.selectbox("Select Theme:", list(DEFAULT_THEMES.keys()))
                theme = DEFAULT_THEMES[theme_name]
            elif theme_option == "Custom Theme":
                theme_name = st.selectbox("Select Custom Theme:", 
                                        ["Create New..."] + list(st.session_state.custom_themes.keys()))
                
                if theme_name == "Create New...":
                    with st.form("custom_theme_form"):
                        new_theme_name = st.text_input("Theme Name")
                        bg_color = st.color_picker("Background Color", "#0C2340")
                        title_color = st.color_picker("Title Color", "#FFFFFF")
                        text_color = st.color_picker("Text Color", "#C8C8C8")
                        accent_color = st.color_picker("Accent Color", "#0070C0")
                        title_font = st.text_input("Title Font", "Calibri")
                        text_font = st.text_input("Text Font", "Calibri")
                        
                        if st.form_submit_button("Save Custom Theme"):
                            if new_theme_name:
                                st.session_state.custom_themes[new_theme_name] = {
                                    "background": hex_to_rgb(bg_color),
                                    "title_color": hex_to_rgb(title_color),
                                    "text_color": hex_to_rgb(text_color),
                                    "accent": hex_to_rgb(accent_color),
                                    "title_font": title_font,
                                    "text_font": text_font
                                }
                                st.success(f"Theme '{new_theme_name}' saved successfully!")
                            else:
                                st.warning("Please enter a theme name")
                elif theme_name in st.session_state.custom_themes:
                    theme = st.session_state.custom_themes[theme_name]
            else:
                uploaded_file = st.file_uploader("Upload Word Document Template", type=["docx"])
                if uploaded_file:
                    with st.spinner("Extracting theme from template..."):
                        theme = extract_theme_from_docx(uploaded_file)
                        st.success("Theme extracted from template!")
                else:
                    st.info("Please upload a Word document to extract its theme")
    
    with col2:
        pass  # Placeholder for logo or branding if needed
    
    if st.button("Generate Word Document", type="primary"):
        if not topic:
            st.warning("Please enter a topic first!")
        elif theme_option == "Example-Based Theme" and not uploaded_file:
            st.warning("Please upload a Word document template first")
        else:
            with st.spinner(f"Creating {page_count}-page Word document about '{topic}'..."):
                try:
                    document_text = generate_document_content(topic, page_count)
                    pages_data = parse_document_content(document_text)
                    
                    with st.expander("Page Overview (Detailed)"):
                        for i, page in enumerate(pages_data, 1):
                            st.subheader(f"Page {i}: {page['title']}")
                            st.markdown("**Content:**")
                            for point in page.get('content', []):
                                st.markdown(f"- {point}")
                            if page.get('notes'):
                                st.markdown(f"**Notes:** {page['notes']}")
                            st.markdown("---")
                    
                    doc_file = create_word_document(pages_data, theme)
                    
                    st.success("Word document generated successfully!")
                    
                    st.download_button(
                        label="Download Word Document",
                        data=doc_file,
                        file_name=f"{topic.replace(' ', '_')}_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()