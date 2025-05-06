import gradio as gr
import os
from pptx import Presentation
from pptx.util import Inches
import google.generativeai as genai

# Initialize Google Generative AI
GOOGLE_API_KEY = os.getenv("AIzaSyDMRDzpIx2OuiO8xoV8UntyRYhl7LbRup8")
genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel("gemini-pro")

# Function to generate slides using Gemini
def generate_slide_content(topic):
    prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
    For each slide, provide:
    1. A clear title in [Title:] format
    2. 3-5 detailed bullet points in [Content:] format (each point should have a brief explanation)
    3. Optional speaker notes in [Notes:] format
    4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
    Structure your response like this:
    [Title:] Slide Title
    [Layout:] title-content
    [Content:]
    - Main Point 1: Brief explanation of this point (keep to one line)
    - Main Point 2: Brief explanation of this point (keep to one line)
    - Main Point 3: Brief explanation of this point (keep to one line)
    [Notes:] Additional notes here
    
    Include these sections (adjust based on requested slide count):
    - Title slide
    - Introduction/Overview
    - Key Concepts (with detailed explanations)
    - Detailed Analysis (with supporting points)
    - Case Studies/Examples (with key takeaways)
    - Applications (with practical implementations)
    - Challenges (with potential solutions)
    - Future Trends (with supporting evidence)
    - Conclusion (with summary points)
    - Q&A
    
    Ensure each bullet point follows the format:
    - Main Point: Explanation (keep explanation concise but informative)"""
    response = model.generate_content(prompt)
    return response.text

# Function to create a PowerPoint from the generated content
def create_ppt_from_text(content, filename="output.pptx"):
    prs = Presentation()
    slides = content.strip().split("Slide")[1:]  # Skip any leading text

    for slide in slides:
        lines = slide.strip().split("\n")
        title = ""
        bullets = []

        for line in lines:
            if "Title:" in line:
                title = line.split("Title:")[1].strip()
            elif "Bullet" in line and ":" in line:
                bullets.append(line.split(":", 1)[1].strip())

        slide_layout = prs.slide_layouts[1]  # Title and Content
        slide_obj = prs.slides.add_slide(slide_layout)
        slide_obj.shapes.title.text = title
        content_box = slide_obj.placeholders[1]
        content_box.text = bullets[0] if bullets else ""

        for bullet in bullets[1:]:
            p = content_box.text_frame.add_paragraph()
            p.text = bullet

    prs.save(filename)
    return filename

# Wrapper for Gradio interface
def generate_ppt(topic):
    content = generate_slide_content(topic)
    ppt_path = create_ppt_from_text(content)
    return ppt_path

# Gradio UI
interface = gr.Interface(
    fn=generate_ppt,
    inputs=gr.Textbox(label="Enter your presentation topic"),
    outputs=gr.File(label="Download your PowerPoint"),
    title="AI PowerPoint Generator",
    description="Generate a 5-slide PowerPoint presentation using Google Gemini and download the PPTX file."
)

if __name__ == "__main__":
    interface.launch()
