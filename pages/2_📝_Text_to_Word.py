import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO
import os

# Configure Google Gemini API key
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

st.header("Text to Word Document")
topic = st.text_input("Enter Document Topic", key="word_topic")
page_count = st.slider("Approximate Page Count", 1, 10, 3, key="word_pages")

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

if st.button("Generate Word Document", key="word_generate"):
    if not topic.strip():
        st.warning("Please enter a topic")
        return
    
    with st.spinner("Generating document content..."):
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
        
        st.success("Document content generated!")
        
        # Create DOCX
        try:
            doc = Document()
            for section in sections:
                if section['heading']:
                    doc.add_heading(section['heading'], level=1)
                if section['subheading']:
                    doc.add_heading(section['subheading'], level=2)
                for content in section['content']:
                    if content.startswith('•'):
                        doc.add_paragraph(content, style='List Bullet')
                    else:
                        doc.add_paragraph(content)
                doc.add_paragraph()
            
            docx_io = BytesIO()
            doc.save(docx_io)
            docx_io.seek(0)
            
            st.download_button(
                "Download DOCX", 
                data=docx_io, 
                file_name=f"{topic.replace(' ','_')}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error creating Word document: {str(e)}")