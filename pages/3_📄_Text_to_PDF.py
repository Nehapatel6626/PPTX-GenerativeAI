import streamlit as st
import google.generativeai as genai
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch  # Added import for inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from docx import Document
from docx.shared import RGBColor
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
        "background": colors.HexColor("#0C2340"),
        "title_color": colors.HexColor("#FFFFFF"),
        "text_color": colors.HexColor("#C8C8C8"),
        "accent": colors.HexColor("#0070C0"),
        "title_font": "Helvetica-Bold",
        "text_font": "Helvetica"
    },
    "Modern Green": {
        "background": colors.HexColor("#165242"),
        "title_color": colors.HexColor("#FFFFFF"),
        "text_color": colors.HexColor("#DCDCDC"),
        "accent": colors.HexColor("#4CAF50"),
        "title_font": "Helvetica-Bold",
        "text_font": "Helvetica"
    },
    "Light Corporate": {
        "background": colors.HexColor("#FFFFFF"),
        "title_color": colors.HexColor("#0D47A1"),
        "text_color": colors.HexColor("#212121"),
        "accent": colors.HexColor("#1976D2"),
        "title_font": "Times-Bold",
        "text_font": "Times-Roman"
    },
    "Dark Tech": {
        "background": colors.HexColor("#212121"),
        "title_color": colors.HexColor("#00C8FF"),
        "text_color": colors.HexColor("#C8C8C8"),
        "accent": colors.HexColor("#0096FF"),
        "title_font": "Courier-Bold",
        "text_font": "Courier"
    }
}

def hex_to_rgb(hex_color):
    """Convert hex color to RGBColor for python-docx"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

def hex_to_reportlab_color(hex_color):
    """Convert hex color to ReportLab color"""
    return colors.HexColor(hex_color)

def extract_theme_from_docx(uploaded_file):
    """Extract theme colors and fonts from an uploaded Word document"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name
    
    doc = Document(tmp_file_path)
    theme = {
        "background": colors.HexColor("#FFFFFF"),
        "title_color": colors.HexColor("#000000"),
        "text_color": colors.HexColor("#000000"),
        "accent": colors.HexColor("#4F81BD"),
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
                        theme["title_color"] = colors.HexColor(f"#{run.font.color.rgb:06x}")
                        theme["title_font"] = run.font.name or "Calibri"
                    else:
                        theme["text_color"] = colors.HexColor(f"#{run.font.color.rgb:06x}")
                        theme["text_font"] = run.font.name or "Calibri"
                except:
                    continue
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    try:
                        run = cell.paragraphs[0].runs[0]
                        if run.font.color.rgb:
                            theme["accent"] = colors.HexColor(f"#{run.font.color.rgb:06x}")
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

def create_pdf_document(pages_data, theme):
    """Create a PDF document"""
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    # Define custom styles
    title_style = ParagraphStyle(
        name='Title',
        fontName=theme['title_font'],
        fontSize=18,
        textColor=theme['title_color'],
        alignment=1,  # Center
        spaceAfter=12
    )
    content_style = ParagraphStyle(
        name='Content',
        fontName=theme['text_font'],
        fontSize=12,
        textColor=theme['text_color'],
        leading=14,
        spaceAfter=8,
        bulletFontName=theme['text_font'],
        bulletFontSize=12,
        leftIndent=10
    )
    
    story = []
    
    for page in pages_data:
        # Add title
        story.append(Paragraph(page['title'], title_style))
        story.append(Spacer(1, 0.2 * inch))
        
        # Add content
        if page['layout'] != 'title-only':
            if page['layout'] == 'two-column':
                mid_point = len(page['content']) // 2
                left_content = page['content'][:mid_point]
                right_content = page['content'][mid_point:]
                
                # Simulate two-column layout with separate lists
                for i in range(max(len(left_content), len(right_content))):
                    if i < len(left_content):
                        story.append(Paragraph(f"• {left_content[i]}", content_style))
                    if i < len(right_content):
                        story.append(Paragraph(f"• {right_content[i]}", content_style))
            else:
                content_items = [ListItem(Paragraph(point, content_style)) for point in page['content']]
                story.append(ListFlowable(content_items, bulletType='bullet', start='bulletchar'))
        
        story.append(Spacer(1, 0.3 * inch))
    
    doc.build(story)
    pdf_io.seek(0)
    
    # Clean up temporary template file
    if "template_path" in theme and os.path.exists(theme["template_path"]):
        os.unlink(theme["template_path"])
    
    return pdf_io

def main():
    st.set_page_config(page_title="PDF Document Generator", layout="wide")
    
    st.title("Advanced PDF Document Generator")
    st.markdown("Create professional PDF documents with AI")
    
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
                        title_font = st.text_input("Title Font", "Helvetica-Bold")
                        text_font = st.text_input("Text Font", "Helvetica")
                        
                        if st.form_submit_button("Save Custom Theme"):
                            if new_theme_name:
                                st.session_state.custom_themes[new_theme_name] = {
                                    "background": hex_to_reportlab_color(bg_color),
                                    "title_color": hex_to_reportlab_color(title_color),
                                    "text_color": hex_to_reportlab_color(text_color),
                                    "accent": hex_to_reportlab_color(accent_color),
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
    
    if st.button("Generate PDF Document", type="primary"):
        if not topic:
            st.warning("Please enter a topic first!")
        elif theme_option == "Example-Based Theme" and not uploaded_file:
            st.warning("Please upload a Word document template first")
        else:
            with st.spinner(f"Creating {page_count}-page PDF document about '{topic}'..."):
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
                    
                    pdf_file = create_pdf_document(pages_data, theme)
                    
                    st.success("PDF document generated successfully!")
                    
                    st.download_button(
                        label="Download PDF",
                        data=pdf_file,
                        file_name=f"{topic.replace(' ', '_')}_document.pdf",
                        mime="application/pdf"
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()