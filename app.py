# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN
# import io
# import os
# import requests
# from PIL import Image
# from io import BytesIO

# # Configure Google Gemini API key (for text generation)
# genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# # Hugging Face API info (for image generation)
# HF_API_URL = "https://api-inference.huggingface.co/models/stabilityai/stable-diffusion-2"
# HF_API_TOKEN = os.getenv("HF_API_TOKEN")  # Hugging Face token in env variable

# # Function to generate slide text content using Gemini
# def generate_slide_content(topic, slide_count):
#     model = genai.GenerativeModel('gemini-2.0-flash')
#     prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
#     For each slide, provide:
#     1. A clear title in [Title:] format
#     2. 3-5 detailed bullet points in [Content:] format
#     3. Optional speaker notes in [Notes:] format
#     4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
#     Structure:
#     [Title:] Slide Title
#     [Layout:] title-content
#     [Content:]
#     - Point 1
#     - Point 2
#     - Point 3
#     [Notes:] Notes here
#     """

#     response = model.generate_content(prompt)
#     return response.text

# # Parse the Gemini-generated text into slide dicts
# def parse_slide_content(slide_text):
#     slides = []
#     current_slide = {}
#     for line in slide_text.split('\n'):
#         line = line.strip()
#         if not line:
#             continue
#         if line.startswith('[Title:]'):
#             if current_slide:
#                 slides.append(current_slide)
#             current_slide = {'title': line.replace('[Title:]', '').strip(), 'content': [], 'notes': '', 'layout': 'title-content'}
#         elif line.startswith('[Content:]'):
#             content = line.replace('[Content:]', '').strip()
#             if content:
#                 current_slide['content'].append(content)
#         elif line.startswith('[Notes:]'):
#             current_slide['notes'] = line.replace('[Notes:]', '').strip()
#         elif line.startswith('[Layout:]'):
#             layout = line.replace('[Layout:]', '').strip().lower()
#             current_slide['layout'] = layout
#         elif line.startswith('-'):
#             current_slide['content'].append(line[1:].strip())
#     if current_slide:
#         slides.append(current_slide)
#     return slides

# # Generate image from Hugging Face API given a prompt (e.g., slide title)
# def generate_image(prompt):
#     headers = {"Authorization": f"Bearer {HF_API_TOKEN}"}
#     payload = {"inputs": prompt}
#     response = requests.post(HF_API_URL, headers=headers, json=payload)
#     if response.status_code == 200:
#         # The HF API returns raw image bytes or base64 depending on model
#         # For stabilityai/stable-diffusion-2, it's an image in bytes directly
#         return Image.open(BytesIO(response.content))
#     else:
#         st.error(f"Image generation failed: {response.status_code} {response.text}")
#         return None

# # Create PPTX, add slides, and add images for each slide
# def create_pptx_with_images(slides_data):
#     prs = Presentation()
#     prs.slide_width = Inches(13.33)
#     prs.slide_height = Inches(7.5)
    
#     for slide_info in slides_data:
#         slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content layout
        
#         # Set title
#         slide.shapes.title.text = slide_info['title']
#         slide.shapes.title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
#         # Add bullet points to content placeholder
#         content_placeholder = slide.placeholders[1].text_frame
#         content_placeholder.clear()
#         for point in slide_info['content']:
#             p = content_placeholder.add_paragraph()
#             p.text = point
#             p.level = 0
        
#         # Generate image for slide title
#         image = generate_image(slide_info['title'])
        
#         if image:
#             # Save image temporarily in memory
#             img_byte_arr = io.BytesIO()
#             image.save(img_byte_arr, format='PNG')
#             img_byte_arr.seek(0)
            
#             # Insert image on slide - position bottom right (adjust as needed)
#             left = Inches(7)
#             top = Inches(3)
#             width = Inches(6)
#             slide.shapes.add_picture(img_byte_arr, left, top, width=width)
        
#     pptx_io = io.BytesIO()
#     prs.save(pptx_io)
#     pptx_io.seek(0)
#     return pptx_io

# def main():
#     st.title("PPT Generator with Gemini + Hugging Face Images")
    
#     topic = st.text_input("Enter Presentation Topic", "")
#     slide_count = st.slider("Number of Slides", 5, 15, 7)
    
#     if st.button("Generate Presentation"):
#         if not topic.strip():
#             st.warning("Please enter a topic")
#             return
        
#         with st.spinner("Generating slide content..."):
#             slide_text = generate_slide_content(topic, slide_count)
#             slides = parse_slide_content(slide_text)
        
#         st.success("Slide content generated!")
        
#         # Show slide preview with images
#         st.header("Slides Preview with Images")
#         for slide in slides:
#             st.subheader(slide['title'])
#             st.markdown("\n".join(f"- {c}" for c in slide['content']))
            
#             img = generate_image(slide['title'])
#             if img:
#                 st.image(img, width=400)
#             else:
#                 st.info("No image generated")
        
#         # Generate PPTX with images
#         with st.spinner("Generating PowerPoint file..."):
#             pptx_file = create_pptx_with_images(slides)
        
#         st.success("PowerPoint generated!")
#         st.download_button("Download PPTX", data=pptx_file, file_name=f"{topic.replace(' ','_')}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# if __name__ == "__main__":
#     main()



import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import os
import requests
from PIL import Image
from io import BytesIO

# Configure Google Gemini API key (for text generation)
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# Hugging Face API info (for image generation)
# Updated to use Stable Diffusion XL which is more reliable
HF_API_URL = "https://api-inference.huggingface.co/models/stabilityai/stable-diffusion-xl-base-1.0"
HF_API_TOKEN = os.getenv("HF_API_TOKEN")  # Hugging Face token in env variable

# Function to generate slide text content using Gemini
def generate_slide_content(topic, slide_count):
    model = genai.GenerativeModel('gemini-2.0-flash')
    prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
    For each slide, provide:
    1. A clear title in [Title:] format
    2. 3-5 detailed bullet points in [Content:] format
    3. Optional speaker notes in [Notes:] format
    4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    5. An image description for the slide in [ImagePrompt:] format (be specific about what visual would best represent this slide)
    
    Structure:
    [Title:] Slide Title
    [Layout:] title-content
    [ImagePrompt:] A detailed description of an appropriate image for this slide
    [Content:]
    - Point 1
    - Point 2
    - Point 3
    [Notes:] Notes here
    """

    response = model.generate_content(prompt)
    return response.text

# Parse the Gemini-generated text into slide dicts
def parse_slide_content(slide_text):
    slides = []
    current_slide = {}
    for line in slide_text.split('\n'):
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
                'layout': 'title-content',
                'image_prompt': ''
            }
        elif line.startswith('[Content:]'):
            content = line.replace('[Content:]', '').strip()
            if content:
                current_slide['content'].append(content)
        elif line.startswith('[Notes:]'):
            current_slide['notes'] = line.replace('[Notes:]', '').strip()
        elif line.startswith('[Layout:]'):
            layout = line.replace('[Layout:]', '').strip().lower()
            current_slide['layout'] = layout
        elif line.startswith('[ImagePrompt:]'):
            current_slide['image_prompt'] = line.replace('[ImagePrompt:]', '').strip()
        elif line.startswith('-'):
            current_slide['content'].append(line[1:].strip())
    if current_slide:
        slides.append(current_slide)
    return slides

# Generate image from Hugging Face API given a prompt
def generate_image(prompt, retries=3):
    headers = {"Authorization": f"Bearer {HF_API_TOKEN}"}
    payload = {
        "inputs": prompt,
        "options": {"wait_for_model": True},
        "parameters": {
            "width": 512,
            "height": 512,
            "num_inference_steps": 25
        }
    }
    
    for attempt in range(retries):
        try:
            response = requests.post(HF_API_URL, headers=headers, json=payload)
            
            if response.status_code == 200:
                return Image.open(BytesIO(response.content))
            elif response.status_code == 503:  # Model loading
                wait_time = 10  # seconds
                st.warning(f"Model is loading, waiting {wait_time} seconds... (attempt {attempt + 1}/{retries})")
                time.sleep(wait_time)
                continue
            else:
                st.error(f"Image generation failed: {response.status_code} {response.text}")
                return None
        except Exception as e:
            st.error(f"Error during image generation: {str(e)}")
            return None
    
    return None

# Create PPTX, add slides, and add images for each slide
def create_pptx_with_images(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    for slide_info in slides_data:
        # Choose layout based on the slide_info['layout']
        # For simplicity, we'll use title+content layout for all slides
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content layout
        
        # Set title
        title = slide.shapes.title
        title.text = slide_info['title']
        title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title.text_frame.paragraphs[0].font.size = Pt(36)
        
        # Add bullet points to content placeholder
        content_placeholder = slide.placeholders[1].text_frame
        content_placeholder.clear()
        for point in slide_info['content']:
            p = content_placeholder.add_paragraph()
            p.text = point
            p.level = 0
            p.font.size = Pt(24)
        
        # Generate image for slide using the specific image prompt
        if slide_info.get('image_prompt'):
            image_prompt = slide_info['image_prompt']
        else:
            # Fallback to title if no specific image prompt
            image_prompt = slide_info['title']
        
        with st.spinner(f"Generating image for: {image_prompt}"):
            image = generate_image(image_prompt)
        
        if image:
            # Save image temporarily in memory
            img_byte_arr = io.BytesIO()
            image.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            
            # Insert image on slide - position bottom right (adjust as needed)
            left = Inches(7)
            top = Inches(2)  # Higher up to allow more space
            width = Inches(6)
            height = Inches(4.5)
            slide.shapes.add_picture(img_byte_arr, left, top, width=width, height=height)
        
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def main():
    st.title("PPT Generator with Gemini + Hugging Face Images")
    
    topic = st.text_input("Enter Presentation Topic", "")
    slide_count = st.slider("Number of Slides", 5, 15, 7)
    
    if st.button("Generate Presentation"):
        if not topic.strip():
            st.warning("Please enter a topic")
            return
        
        with st.spinner("Generating slide content..."):
            slide_text = generate_slide_content(topic, slide_count)
            slides = parse_slide_content(slide_text)
        
        st.success(f"Generated {len(slides)} slides!")
        
        # Show slide preview with images
        st.header("Slides Preview with Images")
        for i, slide in enumerate(slides, 1):
            st.subheader(f"Slide {i}: {slide['title']}")
            st.markdown("\n".join(f"- {c}" for c in slide['content']))
            
            # Use specific image prompt if available, otherwise use title
            image_prompt = slide.get('image_prompt', slide['title'])
            st.caption(f"Image prompt: {image_prompt}")
            
            with st.spinner(f"Generating image for slide {i}..."):
                img = generate_image(image_prompt)
            
            if img:
                st.image(img, width=400, caption=f"Generated image for: {slide['title']}")
            else:
                st.warning("Could not generate image for this slide")
            
            st.divider()
        
        # Generate PPTX with images
        with st.spinner("Generating PowerPoint file..."):
            pptx_file = create_pptx_with_images(slides)
        
        st.success("PowerPoint generated!")
        st.download_button(
            "Download PPTX", 
            data=pptx_file, 
            file_name=f"{topic.replace(' ','_')}.pptx", 
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

if __name__ == "__main__":
    import time  # Add this at the top with other imports
    main()