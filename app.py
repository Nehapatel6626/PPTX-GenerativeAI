import gradio as gr
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
import io
import re
import tempfile
import os

# Setup Google API key (replace with your actual key)
genai.configure(api_key="AIzaSyDMRDzpIx2OuiO8xoV8UntyRYhl7LbRup8")

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

custom_themes = {}

def hex_to_rgb(hex_color):
    """Convert hex color to RGBColor"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

def extract_theme_from_pptx(uploaded_file):
    """Extract theme colors and fonts from an uploaded PowerPoint file"""
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name
    prs = Presentation(tmp_file_path)
    theme = {
        "background": RGBColor(0, 0, 0),  # Default black if not found
        "title_color": RGBColor(0, 0, 0),
        "text_color": RGBColor(0, 0, 0),
        "accent": RGBColor(0, 0, 0),
        "title_font": "Calibri",
        "text_font": "Calibri"
    }
    try:
        slide_master = prs.slide_master
        theme["background"] = slide_master.background.fill.fore_color.rgb
        for shape in slide_master.shapes:
            if shape.has_text_frame and shape.text.strip():
                if shape.name.lower() == 'title':
                    theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                    theme["title_font"] = shape.text_frame.paragraphs[0].font.name
                else:
                    theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                    theme["text_font"] = shape.text_frame.paragraphs[0].font.name
        for shape in slide_master.shapes:
            if hasattr(shape, 'fill'):
                if shape.fill.type == 1:  # Solid fill
                    theme["accent"] = shape.fill.fore_color.rgb
                    break
    except Exception as e:
        print(f"Warning: Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    os.unlink(tmp_file_path)
    return theme

def generate_slide_content(topic, slide_count):
    model = genai.GenerativeModel('gemini-2.0-flash')
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
                'layout': 'title-content'
            }
        elif line.startswith('[Content:]'):
            content = line.replace('[Content:]', '').strip()
            if content:
                current_slide['content'].append(content)
        elif line.startswith('[Notes:]'):
            current_slide['notes'] = line.replace('[Notes:]', '').strip()
        elif line.startswith('[Layout:]'):
            layout = line.replace('[Layout:]', '').strip().lower()
            valid_layouts = ['title-only', 'title-content', 'two-column', 'section-header']
            current_slide['layout'] = layout if layout in valid_layouts else 'title-content'
        elif current_slide.get('content') is not None and line.startswith('-'):
            current_slide['content'].append(line.strip())
    if current_slide:
        slides.append(current_slide)
    return slides

def create_detailed_pptx(slides_data, theme):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    layout_mapping = {
        'title-only': 0,
        'title-content': 1,
        'section-header': 2,
        'two-column': 3
    }

    for slide_info in slides_data:
        layout = slide_info.get('layout', 'title-content')
        layout_idx = layout_mapping.get(layout, 1)
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = theme["background"]

        title = slide.shapes.title
        title.text = slide_info['title']
        title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
        title.text_frame.paragraphs[0].font.size = Pt(36)
        title.text_frame.paragraphs[0].font.bold = True
        if "title_font" in theme:
            title.text_frame.paragraphs[0].font.name = theme["title_font"]

        title.text_frame.word_wrap = True
        title.text_frame.auto_size = True
        title.top = Inches(0.5)
        title.height = Inches(1.5)
        title.width = Inches(11)

        if layout_idx == 3:
            content = slide_info.get('content', [])
            mid_point = len(content) // 2
            left_content = content[:mid_point]
            right_content = content[mid_point:]

            left_body = slide.placeholders[1]
            left_tf = left_body.text_frame
            left_tf.clear()
            left_body.width = Inches(5.5)
            left_body.left = Inches(0.5)

            right_body = slide.placeholders[2]
            right_tf = right_body.text_frame
            right_tf.clear()
            right_body.width = Inches(5.5)
            right_body.left = Inches(7.0)

            for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
                for point in content_part:
                    p = tf.add_paragraph()
                    point_text = point.replace('- ', '').strip()
                    if len(point_text) > 80:
                        parts = [point_text[i:i+80] for i in range(0, len(point_text), 80)]
                        p.text = parts[0]
                        for part in parts[1:]:
                            new_p = tf.add_paragraph()
                            new_p.text = part
                            new_p.level = 1
                    else:
                        p.text = point_text
                    p.level = 0
                    p.font.color.rgb = theme["text_color"]
                    p.font.size = Pt(18)
                    if "text_font" in theme:
                        p.font.name = theme["text_font"]
                    p.space_after = Pt(6)
                    p.space_before = Pt(6)
        elif layout_idx != 0:
            body = slide.placeholders[1]
            tf = body.text_frame
            tf.clear()
            body.top = Inches(2.0)
            body.width = Inches(11)
            body.height = Inches(4.5)
            for point in slide_info.get('content', []):
                p = tf.add_paragraph()
                point_text = point.replace('- ', '').strip()
                if len(point_text) > 100:
                    parts = [point_text[i:i+100] for i in range(0, len(point_text), 100)]
                    p.text = parts[0]
                    for part in parts[1:]:
                        new_p = tf.add_paragraph()
                        new_p.text = part
                        new_p.level = 1
                        new_p.font.color.rgb = theme["text_color"]
                        new_p.font.size = Pt(16)
                        if "text_font" in theme:
                            new_p.font.name = theme["text_font"]
                else:
                    p.text = point_text
                p.level = 0
                p.font.color.rgb = theme["text_color"]
                p.font.size = Pt(18)
                if "text_font" in theme:
                    p.font.name = theme["text_font"]
                p.space_after = Pt(6)
                p.space_before = Pt(6)

        if slide_info.get('notes'):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_info['notes']

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def process_input(topic, slide_count, theme_option, theme_name, bg_color, title_color, text_color, accent_color,
                  title_font, text_font, uploaded_file):
    if not topic:
        return "Please enter a topic.", None, None

    ALL_THEMES = {**DEFAULT_THEMES, **custom_themes}
    theme = DEFAULT_THEMES["Professional Blue"]

    if theme_option == "Predefined Theme":
        theme = DEFAULT_THEMES[theme_name]
    elif theme_option == "Custom Theme":
        if theme_name == "Create New...":
            custom_theme_name = "Custom_" + str(len(custom_themes))
            custom_themes[custom_theme_name] = {
                "background": hex_to_rgb(bg_color),
                "title_color": hex_to_rgb(title_color),
                "text_color": hex_to_rgb(text_color),
                "accent": hex_to_rgb(accent_color),
                "title_font": title_font,
                "text_font": text_font
            }
            theme = custom_themes[custom_theme_name]
        else:
            theme = custom_themes.get(theme_name, theme)
    elif theme_option == "Example-Based Theme":
        if uploaded_file is None:
            return "Please upload a PowerPoint template.", None, None
        theme = extract_theme_from_pptx(uploaded_file)

    slide_text = generate_slide_content(topic, slide_count)
    slides_data = parse_slide_content(slide_text)
    pptx_file = create_detailed_pptx(slides_data, theme)

    overview = ""
    for i, slide in enumerate(slides_data, 1):
        overview += f"Slide {i}: {slide['title']}\nContent:\n"
        for point in slide.get('content', []):
            overview += f"- {point}\n"
        if slide.get('notes'):
            overview += f"Notes: {slide['notes']}\n"
        overview += "---\n"

    download_name = f"{topic.replace(' ', '_')}_presentation.pptx"
    return overview, pptx_file, download_name

# Define Gradio interface
theme_options = ["Predefined Theme", "Custom Theme", "Example-Based Theme"]
predefined_themes = list(DEFAULT_THEMES.keys())

demo = gr.Interface(
    fn=process_input,
    inputs=[
        gr.Textbox(label="Presentation Topic"),
        gr.Slider(minimum=5, maximum=20, value=10, label="Number of Slides"),
        gr.Radio(choices=theme_options, label="Theme Selection Method"),
        gr.Dropdown(choices=predefined_themes, label="Select Predefined Theme", visible=True),
        gr.ColorPicker(value="#0C2340", label="Background Color"),
        gr.ColorPicker(value="#FFFFFF", label="Title Color"),
        gr.ColorPicker(value="#C8C8C8", label="Text Color"),
        gr.ColorPicker(value="#0070C0", label="Accent Color"),
        gr.Textbox(value="Calibri", label="Title Font"),
        gr.Textbox(value="Calibri", label="Text Font"),
        gr.File(label="Upload PowerPoint Template", file_types=[".pptx"], visible=False)
    ],
    outputs=[
        gr.Textbox(label="Slide Overview (Detailed)"),
        gr.File(label="Download PowerPoint", file_types=[".pptx"]),
        gr.Textbox(label="Filename")
    ],
    allow_flagging="never"
)

if __name__ == "__main__":
    demo.launch()