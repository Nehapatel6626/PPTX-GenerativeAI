# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN
# import io
# import os
# from docx import Document
# from fpdf import FPDF
# from io import BytesIO

# # Configure Google Gemini API key
# genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# # Set page config
# st.set_page_config(page_title="Document Generator", layout="wide")

# # Initialize session state for page navigation
# if 'current_page' not in st.session_state:
#     st.session_state.current_page = "PPTX"

# # Sidebar navigation
# def sidebar_nav():
#     st.sidebar.title("Navigation")
#     if st.sidebar.button("PPTX Generator"):
#         st.session_state.current_page = "PPTX"
#     if st.sidebar.button("Word Generator"):
#         st.session_state.current_page = "Word"
#     if st.sidebar.button("PDF Generator"):
#         st.session_state.current_page = "PDF"

# # Common content generation function
# def generate_content(topic, output_format, slide_count=None, page_count=None):
#     model = genai.GenerativeModel('gemini-2.0-flash')
    
#     if output_format == "pptx":
#         prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
#         For each slide, provide:
#         1. A clear title in [Title:] format
#         2. 3-5 detailed bullet points in [Content:] format
#         3. Optional speaker notes in [Notes:] format
#         4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
#         """
#     elif output_format in ["docx", "pdf"]:
#         prompt = f"""Create a comprehensive document on '{topic}' with approximately {page_count} pages of content.
#         Provide the content in the following format:
#         [Heading:] Main Heading
#         [Subheading:] Subheading if any
#         [Content:] 
#         Paragraph 1 text here.
#         - Bullet point 1
#         - Bullet point 2
#         Paragraph 2 text here.
#         """
    
#     try:
#         response = model.generate_content(prompt)
#         return response.text
#     except Exception as e:
#         st.error(f"Error generating content: {str(e)}")
#         return None

# # PPTX Page
# def pptx_page():
#     st.header("Text to PowerPoint Presentation")
#     topic = st.text_input("Enter Presentation Topic", key="pptx_topic")
#     slide_count = st.slider("Number of Slides", 5, 15, 7, key="pptx_slides")
    
#     if st.button("Generate Presentation", key="pptx_generate"):
#         if not topic.strip():
#             st.warning("Please enter a topic")
#             return
        
#         with st.spinner("Generating presentation content..."):
#             content = generate_content(topic, "pptx", slide_count=slide_count)
#             if not content:
#                 return
            
#             slides = []
#             current_slide = {}
#             for line in content.split('\n'):
#                 line = line.strip()
#                 if not line:
#                     continue
#                 if line.startswith('[Title:]'):
#                     if current_slide:
#                         slides.append(current_slide)
#                     current_slide = {
#                         'title': line.replace('[Title:]', '').strip(),
#                         'content': [],
#                         'notes': '',
#                         'layout': 'title-content'
#                     }
#                 elif line.startswith('[Content:]'):
#                     content_text = line.replace('[Content:]', '').strip()
#                     if content_text:
#                         current_slide['content'].append(content_text)
#                 elif line.startswith('-'):
#                     current_slide['content'].append(line[1:].strip())
            
#             if current_slide:
#                 slides.append(current_slide)
            
#             st.success("Presentation content generated!")
            
#             # Create PPTX
#             try:
#                 prs = Presentation()
#                 prs.slide_width = Inches(13.33)
#                 prs.slide_height = Inches(7.5)
                
#                 for slide_info in slides:
#                     slide = prs.slides.add_slide(prs.slide_layouts[1])
#                     slide.shapes.title.text = slide_info['title']
#                     content_placeholder = slide.placeholders[1].text_frame
#                     content_placeholder.clear()
#                     for point in slide_info['content']:
#                         p = content_placeholder.add_paragraph()
#                         p.text = point
#                         p.level = 0
                
#                 pptx_io = BytesIO()
#                 prs.save(pptx_io)
#                 pptx_io.seek(0)
                
#                 st.download_button(
#                     "Download PPTX", 
#                     data=pptx_io, 
#                     file_name=f"{topic.replace(' ','_')}.pptx", 
#                     mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
#                 )
#             except Exception as e:
#                 st.error(f"Error creating PowerPoint: {str(e)}")

# # Word Page
# def word_page():
#     st.header("Text to Word Document")
#     topic = st.text_input("Enter Document Topic", key="word_topic")
#     page_count = st.slider("Approximate Page Count", 1, 10, 3, key="word_pages")
    
#     if st.button("Generate Word Document", key="word_generate"):
#         if not topic.strip():
#             st.warning("Please enter a topic")
#             return
        
#         with st.spinner("Generating document content..."):
#             content = generate_content(topic, "docx", page_count=page_count)
#             if not content:
#                 return
            
#             sections = []
#             current_section = {'heading': '', 'subheading': '', 'content': []}
            
#             for line in content.split('\n'):
#                 line = line.strip()
#                 if not line:
#                     continue
#                 if line.startswith('[Heading:]'):
#                     if current_section['heading'] or current_section['content']:
#                         sections.append(current_section)
#                     current_section = {
#                         'heading': line.replace('[Heading:]', '').strip(),
#                         'subheading': '',
#                         'content': []
#                     }
#                 elif line.startswith('[Subheading:]'):
#                     current_section['subheading'] = line.replace('[Subheading:]', '').strip()
#                 elif line.startswith('[Content:]'):
#                     content_text = line.replace('[Content:]', '').strip()
#                     if content_text:
#                         current_section['content'].append(content_text)
#                 elif line.startswith('-'):
#                     current_section['content'].append(f"• {line[1:].strip()}")
#                 else:
#                     current_section['content'].append(line)
            
#             if current_section['heading'] or current_section['content']:
#                 sections.append(current_section)
            
#             st.success("Document content generated!")
            
#             # Create DOCX
#             try:
#                 doc = Document()
#                 for section in sections:
#                     if section['heading']:
#                         doc.add_heading(section['heading'], level=1)
#                     if section['subheading']:
#                         doc.add_heading(section['subheading'], level=2)
#                     for content in section['content']:
#                         if content.startswith('•'):
#                             doc.add_paragraph(content, style='List Bullet')
#                         else:
#                             doc.add_paragraph(content)
#                     doc.add_paragraph()
                
#                 docx_io = BytesIO()
#                 doc.save(docx_io)
#                 docx_io.seek(0)
                
#                 st.download_button(
#                     "Download DOCX", 
#                     data=docx_io, 
#                     file_name=f"{topic.replace(' ','_')}.docx", 
#                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#                 )
#             except Exception as e:
#                 st.error(f"Error creating Word document: {str(e)}")

# # PDF Page
# def pdf_page():
#     st.header("Text to PDF Report")
#     topic = st.text_input("Enter Report Topic", key="pdf_topic")
#     page_count = st.slider("Approximate Page Count", 1, 10, 3, key="pdf_pages")
    
#     if st.button("Generate PDF Report", key="pdf_generate"):
#         if not topic.strip():
#             st.warning("Please enter a topic")
#             return
        
#         with st.spinner("Generating report content..."):
#             content = generate_content(topic, "pdf", page_count=page_count)
#             if not content:
#                 return
            
#             sections = []
#             current_section = {'heading': '', 'subheading': '', 'content': []}
            
#             for line in content.split('\n'):
#                 line = line.strip()
#                 if not line:
#                     continue
#                 if line.startswith('[Heading:]'):
#                     if current_section['heading'] or current_section['content']:
#                         sections.append(current_section)
#                     current_section = {
#                         'heading': line.replace('[Heading:]', '').strip(),
#                         'subheading': '',
#                         'content': []
#                     }
#                 elif line.startswith('[Subheading:]'):
#                     current_section['subheading'] = line.replace('[Subheading:]', '').strip()
#                 elif line.startswith('[Content:]'):
#                     content_text = line.replace('[Content:]', '').strip()
#                     if content_text:
#                         current_section['content'].append(content_text)
#                 elif line.startswith('-'):
#                     current_section['content'].append(f"• {line[1:].strip()}")
#                 else:
#                     current_section['content'].append(line)
            
#             if current_section['heading'] or current_section['content']:
#                 sections.append(current_section)
            
#             st.success("Report content generated!")
            
#             # Create PDF
#             try:
#                 pdf = FPDF()
#                 pdf.add_page()
#                 pdf.set_font("Arial", size=12)
                
#                 for section in sections:
#                     if section['heading']:
#                         pdf.set_font("Arial", 'B', 16)
#                         pdf.cell(200, 10, txt=section['heading'], ln=1, align='L')
#                         pdf.set_font("Arial", size=12)
#                     if section['subheading']:
#                         pdf.set_font("Arial", 'I', 14)
#                         pdf.cell(200, 10, txt=section['subheading'], ln=1, align='L')
#                         pdf.set_font("Arial", size=12)
#                     for content in section['content']:
#                         if content.startswith('•'):
#                             pdf.cell(10)
#                             pdf.multi_cell(0, 10, txt=content)
#                         else:
#                             pdf.multi_cell(0, 10, txt=content)
#                     pdf.ln(5)
                
#                 pdf_io = BytesIO()
#                 pdf.output(pdf_io)
#                 pdf_io.seek(0)
                
#                 st.download_button(
#                     "Download PDF", 
#                     data=pdf_io, 
#                     file_name=f"{topic.replace(' ','_')}.pdf", 
#                     mime="application/pdf"
#                 )
#             except Exception as e:
#                 st.error(f"Error creating PDF: {str(e)}")

# # Main App
# def main():
#     sidebar_nav()
    
#     if st.session_state.current_page == "PPTX":
#         pptx_page()
#     elif st.session_state.current_page == "Word":
#         word_page()
#     elif st.session_state.current_page == "PDF":
#         pdf_page()

# if __name__ == "__main__":
#     main()





import streamlit as st

# Set page config
st.set_page_config(page_title="Document Generator", layout="wide")
st.image("https://huggingface.co/spaces/YashMK89/PPTX_generativeai/resolve/main/SVNIT_IS_DS_Projects.png", width=150)

# Sidebar navigation
page = st.sidebar.selectbox("Choose a feature", ["Introduction", "Text to PPT", "Text to PDF", "Text to Word", "Pexels Image Search"])

# Introduction Page
if page == "Introduction":
    st.title("Welcome to Document Generator")
    st.write("Your one-stop tool for creating documents and finding images with ease!")
    
    st.markdown("""
    ### About Document Generator
    The **Document Generator** is a powerful Streamlit app designed to simplify document creation and image sourcing. Whether you're preparing presentations, reports, or visual content, this app has you covered with the following features:
    
    - **Text to PPT**: Transform your text into professional PowerPoint slides with a single click.
    - **Text to PDF**: Convert your text into beautifully formatted PDF documents.
    - **Text to Word**: Generate Word documents from your text for easy editing and sharing.
    - **Pexels Image Search**: Search for high-quality images from Pexels and download them directly for your projects.
    
    Use the sidebar to navigate to the desired feature and start creating!
    """)
    
    st.markdown("---")
    st.subheader("Get Started")
    st.write("Select a feature from the sidebar to begin generating documents or searching for images.")
    
# Placeholder for other pages
elif page == "pages/1_📊_Text_to_PPT.py":
    st.title("Text to PPT")
    st.write("Enter your text below to generate a PowerPoint presentation.")
    # Add PPT generation logic here (e.g., using python-pptx)
    
elif page == "pages/3_📄_Text_to_PDF.py":
    st.title("Text to PDF")
    st.write("Enter your text below to generate a PDF document.")
    # Add PDF generation logic here (e.g., using reportlab)
    
elif page == "pages/2_📝_Text_to_Word.py":
    st.title("Text to Word")
    st.write("Enter your text below to generate a Word document.")
    # Add Word generation logic here (e.g., using python-docx)
    
elif page == "pages/4_🖼_text_to_Image_Search.py":
    st.title("Pexels Image Search")
    st.write("Search for images on Pexels and download them for your projects.")
    # Add Pexels API integration here

# You can keep any shared configuration or session state management here