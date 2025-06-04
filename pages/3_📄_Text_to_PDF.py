import streamlit as st
import google.generativeai as genai
from fpdf import FPDF
from io import BytesIO
import os

# Configure Google Gemini API key
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

def generate_content(topic, page_count):
    model = genai.GenerativeModel('gemini-2.0-flash')
    prompt = f"""Create a comprehensive document on '{topic}' with approximately {page_count} pages of content.
    Provide the content in the following format:
    [Heading:] Main Heading
    [Subheading:] Subheading if any
    [Content:] 
    Paragraph 1 text here.
    - Bullet point 1
    - Bullet point 2
    Paragraph 2 text here.
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Error generating content: {str(e)}")
        return None

def create_pdf(topic, sections):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        for section in sections:
            if section['heading']:
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=section['heading'], ln=1, align='L')
                pdf.set_font("Arial", size=12)
            if section['subheading']:
                pdf.set_font("Arial", 'I', 14)
                pdf.cell(200, 10, txt=section['subheading'], ln=1, align='L')
                pdf.set_font("Arial", size=12)
            for content in section['content']:
                if content.startswith('•'):
                    pdf.cell(10)
                    pdf.multi_cell(0, 10, txt=content)
                else:
                    pdf.multi_cell(0, 10, txt=content)
            pdf.ln(5)
        
        pdf_io = BytesIO()
        pdf.output(pdf_io)
        pdf_io.seek(0)
        return pdf_io
    except Exception as e:
        st.error(f"Error creating PDF: {str(e)}")
        return None

def main():
    st.header("Text to PDF Report")
    topic = st.text_input("Enter Report Topic", key="pdf_topic")
    page_count = st.slider("Approximate Page Count", 1, 10, 3, key="pdf_pages")
    
    if st.button("Generate PDF Report", key="pdf_generate"):
        if not topic.strip():
            st.warning("Please enter a topic")
            return
        
        with st.spinner("Generating report content..."):
            content = generate_content(topic, page_count)
            if not content:
                return
            
            sections = []
            current_section = {'heading': '', 'subheading': '', 'content': []}
            
            for line in content.split('\n'):
                line = line.strip()
                if not line:
                    continue
                if line.startswith('[Heading:]'):
                    if current_section['heading'] or current_section['content']:
                        sections.append(current_section)
                    current_section = {
                        'heading': line.replace('[Heading:]', '').strip(),
                        'subheading': '',
                        'content': []
                    }
                elif line.startswith('[Subheading:]'):
                    current_section['subheading'] = line.replace('[Subheading:]', '').strip()
                elif line.startswith('[Content:]'):
                    content_text = line.replace('[Content:]', '').strip()
                    if content_text:
                        current_section['content'].append(content_text)
                elif line.startswith('-'):
                    current_section['content'].append(f"• {line[1:].strip()}")
                else:
                    current_section['content'].append(line)
            
            if current_section['heading'] or current_section['content']:
                sections.append(current_section)
            
            st.success("Report content generated!")
            
            pdf_io = create_pdf(topic, sections)
            if pdf_io:
                st.download_button(
                    "Download PDF", 
                    data=pdf_io, 
                    file_name=f"{topic.replace(' ','_')}.pdf", 
                    mime="application/pdf"
                )

if __name__ == "__main__":
    main()