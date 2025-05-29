import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import os
from docx import Document
from fpdf import FPDF
from io import BytesIO

# Configure Google Gemini API key (for text generation)
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

def generate_content(topic, output_format, slide_count=None, word_page_count=None):
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    if output_format == "pptx":
        prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
        For each slide, provide:
        1. A clear title in [Title:] format
        2. 3-5 detailed bullet points in [Content:] format
        3. Optional speaker notes in [Notes:] format
        4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
        
        Structure:
        [Title:] Slide Title
        [Layout:] title-content
        [Content:]
        - Point 1
        - Point 2
        - Point 3
        [Notes:] Notes here
        """
    elif output_format == "docx":
        prompt = f"""Create a comprehensive document on '{topic}' with approximately {word_page_count} pages of content.
        Provide the content in the following format:
        
        [Heading:] Main Heading
        [Subheading:] Subheading if any
        [Content:] 
        Paragraph 1 text here.
        - Bullet point 1
        - Bullet point 2
        Paragraph 2 text here.
        """
    elif output_format == "pdf":
        prompt = f"""Create a comprehensive report on '{topic}' suitable for a PDF document.
        Provide the content in the following format:
        
        [Heading:] Main Heading
        [Subheading:] Subheading if any
        [Content:] 
        Paragraph 1 text here.
        - Bullet point 1
        - Bullet point 2
        Paragraph 2 text here.
        """
    
    response = model.generate_content(prompt)
    return response.text

def parse_pptx_content(slide_text):
    slides = []
    current_slide = {}
    for line in slide_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('[Title:]'):
            if current_slide:
                slides.append(current_slide)
            current_slide = {'title': line.replace('[Title:]', '').strip(), 'content': [], 'notes': '', 'layout': 'title-content'}
        elif line.startswith('[Content:]'):
            content = line.replace('[Content:]', '').strip()
            if content:
                current_slide['content'].append(content)
        elif line.startswith('[Notes:]'):
            current_slide['notes'] = line.replace('[Notes:]', '').strip()
        elif line.startswith('[Layout:]'):
            layout = line.replace('[Layout:]', '').strip().lower()
            current_slide['layout'] = layout
        elif line.startswith('-'):
            current_slide['content'].append(line[1:].strip())
    if current_slide:
        slides.append(current_slide)
    return slides

def parse_doc_content(content_text):
    sections = []
    current_section = {}
    
    for line in content_text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('[Heading:]'):
            if current_section:
                sections.append(current_section)
            current_section = {'heading': line.replace('[Heading:]', '').strip(), 'subheading': '', 'content': []}
        elif line.startswith('[Subheading:]'):
            current_section['subheading'] = line.replace('[Subheading:]', '').strip()
        elif line.startswith('[Content:]'):
            content = line.replace('[Content:]', '').strip()
            if content:
                current_section['content'].append(content)
        elif line.startswith('-'):
            current_section['content'].append(line[1:].strip())
        else:
            current_section['content'].append(line)
    
    if current_section:
        sections.append(current_section)
    return sections

def create_pptx(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content layout
        
        # Set title
        slide.shapes.title.text = slide_info['title']
        slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Add bullet points to content placeholder
        content_placeholder = slide.placeholders[1].text_frame
        content_placeholder.clear()
        for point in slide_info['content']:
            p = content_placeholder.add_paragraph()
            p.text = point
            p.level = 0
        
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def create_docx(sections):
    doc = Document()
    
    for section in sections:
        # Add heading
        doc.add_heading(section['heading'], level=1)
        
        # Add subheading if exists
        if section['subheading']:
            doc.add_heading(section['subheading'], level=2)
        
        # Add content
        for content in section['content']:
            if content.startswith('-'):
                # This is a bullet point
                doc.add_paragraph(content[1:].strip(), style='List Bullet')
            else:
                # Regular paragraph
                doc.add_paragraph(content)
        
        # Add space between sections
        doc.add_paragraph()
    
    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)
    return docx_io

def create_pdf(sections):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    for section in sections:
        # Add heading
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(200, 10, txt=section['heading'], ln=1, align='L')
        pdf.set_font("Arial", size=12)
        
        # Add subheading if exists
        if section['subheading']:
            pdf.set_font("Arial", 'I', 14)
            pdf.cell(200, 10, txt=section['subheading'], ln=1, align='L')
            pdf.set_font("Arial", size=12)
        
        # Add content
        for content in section['content']:
            if content.startswith('-'):
                # This is a bullet point
                pdf.cell(10)  # indent
                pdf.multi_cell(0, 10, txt=f"• {content[1:].strip()}")
            else:
                # Regular paragraph
                pdf.multi_cell(0, 10, txt=content)
        
        # Add space between sections
        pdf.ln(5)
    
    pdf_io = io.BytesIO()
    pdf.output(pdf_io)
    pdf_io.seek(0)
    return pdf_io

def main():
    st.title("Multi-Format Document Generator with Gemini")
    
    # Document type selection
    doc_type = st.radio("Select Output Format", ["PPTX Presentation", "Word Document", "PDF Report"])
    
    topic = st.text_input("Enter Document Topic", "")
    
    # Conditional parameters based on document type
    if doc_type == "PPTX Presentation":
        slide_count = st.slider("Number of Slides", 5, 15, 7)
    elif doc_type == "Word Document":
        word_page_count = st.slider("Approximate Page Count", 1, 10, 3)
    
    if st.button("Generate Document"):
        if not topic.strip():
            st.warning("Please enter a topic")
            return
        
        with st.spinner("Generating content..."):
            if doc_type == "PPTX Presentation":
                output_format = "pptx"
                content = generate_content(topic, output_format, slide_count=slide_count)
                slides = parse_pptx_content(content)
                
                st.success("Presentation content generated!")
                st.header("Slides Preview")
                for slide in slides:
                    st.subheader(slide['title'])
                    st.markdown("\n".join(f"- {c}" for c in slide['content']))
                
                with st.spinner("Creating PowerPoint file..."):
                    pptx_file = create_pptx(slides)
                
                st.success("PowerPoint generated!")
                st.download_button(
                    "Download PPTX", 
                    data=pptx_file, 
                    file_name=f"{topic.replace(' ','_')}.pptx", 
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            elif doc_type == "Word Document":
                output_format = "docx"
                content = generate_content(topic, output_format, word_page_count=word_page_count)
                sections = parse_doc_content(content)
                
                st.success("Document content generated!")
                st.header("Document Preview")
                for section in sections:
                    st.subheader(section['heading'])
                    if section['subheading']:
                        st.write(section['subheading'])
                    for content in section['content']:
                        if content.startswith('-'):
                            st.markdown(f"- {content[1:]}")
                        else:
                            st.write(content)
                
                with st.spinner("Creating Word document..."):
                    docx_file = create_docx(sections)
                
                st.success("Word document generated!")
                st.download_button(
                    "Download DOCX", 
                    data=docx_file, 
                    file_name=f"{topic.replace(' ','_')}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            elif doc_type == "PDF Report":
                output_format = "pdf"
                content = generate_content(topic, output_format)
                sections = parse_doc_content(content)
                
                st.success("Report content generated!")
                st.header("Report Preview")
                for section in sections:
                    st.subheader(section['heading'])
                    if section['subheading']:
                        st.write(section['subheading'])
                    for content in section['content']:
                        if content.startswith('-'):
                            st.markdown(f"- {content[1:]}")
                        else:
                            st.write(content)
                
                with st.spinner("Creating PDF document..."):
                    pdf_file = create_pdf(sections)
                
                st.success("PDF generated!")
                st.download_button(
                    "Download PDF", 
                    data=pdf_file, 
                    file_name=f"{topic.replace(' ','_')}.pdf", 
                    mime="application/pdf"
                )

if __name__ == "__main__":
    main()