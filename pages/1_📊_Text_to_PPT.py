import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import os

# Configure Google Gemini API key
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

def generate_content(topic, slide_count):
    model = genai.GenerativeModel('gemini-2.0-flash')
    prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
    For each slide, provide:
    1. A clear title in [Title:] format
    2. 3-5 detailed bullet points in [Content:] format
    3. Optional speaker notes in [Notes:] format
    4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Error generating content: {str(e)}")
        return None

def create_presentation(topic, slides):
    try:
        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        for slide_info in slides:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = slide_info['title']
            content_placeholder = slide.placeholders[1].text_frame
            content_placeholder.clear()
            for point in slide_info['content']:
                p = content_placeholder.add_paragraph()
                p.text = point
                p.level = 0
        
        pptx_io = BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        return pptx_io
    except Exception as e:
        st.error(f"Error creating PowerPoint: {str(e)}")
        return None

def main():
    st.header("Text to PowerPoint Presentation")
    topic = st.text_input("Enter Presentation Topic", key="pptx_topic")
    slide_count = st.slider("Number of Slides", 5, 15, 7, key="pptx_slides")
    
    if st.button("Generate Presentation", key="pptx_generate"):
        if not topic.strip():
            st.warning("Please enter a topic")
            return
        
        with st.spinner("Generating presentation content..."):
            content = generate_content(topic, slide_count)
            if not content:
                return
            
            slides = []
            current_slide = {}
            for line in content.split('\n'):
                line = line.strip()
                if not line:
                    continue
                if line.startswith('[Title:]'):
                    if current_slide:
                        slides.append(current_slide)
                    current_slide = {
                        'title': line.replace('[Title:]', '').strip(),
                        'content': [],
                        'notes': '',
                        'layout': 'title-content'
                    }
                elif line.startswith('[Content:]'):
                    content_text = line.replace('[Content:]', '').strip()
                    if content_text:
                        current_slide['content'].append(content_text)
                elif line.startswith('-'):
                    current_slide['content'].append(line[1:].strip())
            
            if current_slide:
                slides.append(current_slide)
            
            st.success("Presentation content generated!")
            
            pptx_io = create_presentation(topic, slides)
            if pptx_io:
                st.download_button(
                    "Download PPTX", 
                    data=pptx_io, 
                    file_name=f"{topic.replace(' ','_')}.pptx", 
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

if __name__ == "__main__":
    main()