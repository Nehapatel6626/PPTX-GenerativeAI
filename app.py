# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
# from pptx.enum.shapes import MSO_SHAPE_TYPE
# from pptx.oxml.xmlchemy import OxmlElement
# import io
# import re
# import tempfile
# import os

# # Load the API key securely from environment variable
# api_key = os.getenv("GOOGLE_API_KEY")
# genai.configure(api_key=api_key)

# # Default theme configurations
# DEFAULT_THEMES = {
#     "Professional Blue": {
#         "background": RGBColor(12, 35, 64),
#         "title_color": RGBColor(255, 255, 255),
#         "text_color": RGBColor(200, 200, 200),
#         "accent": RGBColor(0, 112, 192),
#         "title_font": "Calibri",
#         "text_font": "Calibri"
#     },
#     "Modern Green": {
#         "background": RGBColor(22, 82, 66),
#         "title_color": RGBColor(255, 255, 255),
#         "text_color": RGBColor(220, 220, 220),
#         "accent": RGBColor(76, 175, 80),
#         "title_font": "Arial",
#         "text_font": "Arial"
#     },
#     "Light Corporate": {
#         "background": RGBColor(255, 255, 255),
#         "title_color": RGBColor(13, 71, 161),
#         "text_color": RGBColor(33, 33, 33),
#         "accent": RGBColor(25, 118, 210),
#         "title_font": "Segoe UI",
#         "text_font": "Segoe UI"
#     },
#     "Dark Tech": {
#         "background": RGBColor(33, 33, 33),
#         "title_color": RGBColor(0, 200, 255),
#         "text_color": RGBColor(200, 200, 200),
#         "accent": RGBColor(0, 150, 255),
#         "title_font": "Consolas",
#         "text_font": "Consolas"
#     }
# }

# # def hex_to_rgb(hex_color):
# #     """Convert hex color to RGBColor"""
# #     hex_color = hex_color.lstrip('#')
# #     return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

# # def extract_theme_from_pptx(uploaded_file):
# #     """Extract theme colors and fonts from an uploaded PowerPoint file"""
# #     with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
# #         tmp_file.write(uploaded_file.read())
# #         tmp_file_path = tmp_file.name
    
# #     prs = Presentation(tmp_file_path)
# #     theme = {
# #         "background": RGBColor(0, 0, 0),  # Default black if not found
# #         "title_color": RGBColor(0, 0, 0),
# #         "text_color": RGBColor(0, 0, 0),
# #         "accent": RGBColor(0, 0, 0),
# #         "title_font": "Calibri",
# #         "text_font": "Calibri"
# #     }
    
# #     try:
# #         # Get colors from master slide
# #         slide_master = prs.slide_master
# #         theme["background"] = slide_master.background.fill.fore_color.rgb
        
# #         # Try to get title and text colors from placeholders
# #         for shape in slide_master.shapes:
# #             if shape.has_text_frame and shape.text.strip():
# #                 if shape.name.lower() == 'title':
# #                     theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
# #                     theme["title_font"] = shape.text_frame.paragraphs[0].font.name
# #                 else:
# #                     theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
# #                     theme["text_font"] = shape.text_frame.paragraphs[0].font.name
        
# #         # Try to get accent color from first shape with fill
# #         for shape in slide_master.shapes:
# #             if hasattr(shape, 'fill'):
# #                 if shape.fill.type == 1:  # Solid fill
# #                     theme["accent"] = shape.fill.fore_color.rgb
# #                     break
    
# #     except Exception as e:
# #         st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
# #     os.unlink(tmp_file_path)
# #     return theme

# def generate_slide_content(topic, slide_count):
#     model = genai.GenerativeModel('gemini-2.0-flash')
#     prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
#     For each slide, provide:
#     1. A clear title in [Title:] format
#     2. 3-5 detailed bullet points in [Content:] format (each point should have a brief explanation)
#     3. Optional speaker notes in [Notes:] format
#     4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
#     Structure your response like this:
#     [Title:] Slide Title
#     [Layout:] title-content
#     [Content:]
#     - Main Point 1: Brief explanation of this point (keep to one line)
#     - Main Point 2: Brief explanation of this point (keep to one line)
#     - Main Point 3: Brief explanation of this point (keep to one line)
#     [Notes:] Additional notes here
    
#     Include these sections (adjust based on requested slide count):
#     - Title slide
#     - Introduction/Overview
#     - Key Concepts (with detailed explanations)
#     - Detailed Analysis (with supporting points)
#     - Case Studies/Examples (with key takeaways)
#     - Applications (with practical implementations)
#     - Challenges (with potential solutions)
#     - Future Trends (with supporting evidence)
#     - Conclusion (with summary points)
#     - Q&A
    
#     Ensure each bullet point follows the format:
#     - Main Point: Explanation (keep explanation concise but informative)"""
    
#     response = model.generate_content(prompt)
#     return response.text

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
#             current_slide = {
#                 'title': line.replace('[Title:]', '').strip(),
#                 'content': [],
#                 'notes': '',
#                 'layout': 'title-content'
#             }
#         elif line.startswith('[Content:]'):
#             content = line.replace('[Content:]', '').strip()
#             if content:
#                 current_slide['content'].append(content)
#         elif line.startswith('[Notes:]'):
#             current_slide['notes'] = line.replace('[Notes:]', '').strip()
#         elif line.startswith('[Layout:]'):
#             layout = line.replace('[Layout:]', '').strip().lower()
#             valid_layouts = ['title-only', 'title-content', 'two-column', 'section-header']
#             current_slide['layout'] = layout if layout in valid_layouts else 'title-content'
#         elif current_slide.get('content') is not None and line.startswith('-'):
#             current_slide['content'].append(line.strip())
    
#     if current_slide:
#         slides.append(current_slide)
    
#     return slides

# def create_detailed_pptx(slides_data, theme):
#     prs = Presentation()
    
#     # Set widescreen layout (16:9 aspect ratio)
#     prs.slide_width = Inches(13.33)
#     prs.slide_height = Inches(7.5)
    
#     # Layout mapping - now using layout indices from the presentation
#     # First, let's get the available layout indices from the presentation
#     layout_indices = {
#         'title-only': 0,  # Typically index 0 is title slide
#         'title-content': 1,  # Typically index 1 is title and content
#         'section-header': 2,  # Might not exist in all templates
#         'two-column': 3  # Might not exist in all templates
#     }
    
#     # Verify which layouts actually exist in the presentation
#     available_layouts = {}
#     for name, idx in layout_indices.items():
#         if idx < len(prs.slide_layouts):
#             available_layouts[name] = idx
    
#     # Default to title-content if requested layout isn't available
#     default_layout_idx = available_layouts.get('title-content', 0)
    
#     for slide_info in slides_data:
#         layout = slide_info.get('layout', 'title-content').lower()
#         layout_idx = available_layouts.get(layout, default_layout_idx)
        
#         try:
#             slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
#         except IndexError:
#             # Fallback to default layout if requested layout doesn't exist
#             slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
#         # Apply background
#         background = slide.background
#         fill = background.fill
#         fill.solid()
#         fill.fore_color.rgb = theme["background"]
        
#         # Set title with theme and proper formatting
#         title = slide.shapes.title
#         title.text = slide_info['title']
#         title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
#         title.text_frame.paragraphs[0].font.size = Pt(36)
#         title.text_frame.paragraphs[0].font.bold = True
#         if "title_font" in theme:
#             title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
#         # Adjust title position to prevent overflow
#         title.text_frame.word_wrap = True
#         title.text_frame.auto_size = True
#         title.top = Inches(0.5)
#         title.height = Inches(1.5)
#         title.width = Inches(11)  # Leave some margin on sides
        
#         # Set content based on layout with proper text formatting
#         if layout_idx == 3:  # Two column layout
#             content = slide_info.get('content', [])
#             mid_point = len(content) // 2
#             left_content = content[:mid_point]
#             right_content = content[mid_point:]
            
#             # Left column
#             left_body = slide.placeholders[1]
#             left_tf = left_body.text_frame
#             left_tf.clear()
#             left_body.width = Inches(5.5)  # Narrower columns
#             left_body.left = Inches(0.5)
            
#             # Right column
#             right_body = slide.placeholders[2]
#             right_tf = right_body.text_frame
#             right_tf.clear()
#             right_body.width = Inches(5.5)  # Narrower columns
#             right_body.left = Inches(7.0)
            
#             for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
#                 for point in content_part:
#                     p = tf.add_paragraph()
#                     point_text = point.replace('- ', '').strip()
#                     # Split point if it's too long
#                     if len(point_text) > 80:
#                         parts = [point_text[i:i+80] for i in range(0, len(point_text), 80)]
#                         p.text = parts[0]
#                         for part in parts[1:]:
#                             new_p = tf.add_paragraph()
#                             new_p.text = part
#                             new_p.level = 1  # Indent wrapped lines
#                     else:
#                         p.text = point_text
                    
#                     p.level = 0
#                     p.font.color.rgb = theme["text_color"]
#                     p.font.size = Pt(18)
#                     if "text_font" in theme:
#                         p.font.name = theme["text_font"]
#                     p.space_after = Pt(6)  # Reduced spacing
#                     p.space_before = Pt(6)
        
#         elif layout_idx != 0:  # Not title-only
#             body = slide.placeholders[1]
#             tf = body.text_frame
#             tf.clear()
#             body.top = Inches(2.0)  # Position below title
#             body.width = Inches(11)  # Leave margins
#             body.height = Inches(4.5)  # Limit height
            
#             for point in slide_info.get('content', []):
#                 p = tf.add_paragraph()
#                 point_text = point.replace('- ', '').strip()
                
#                 # Handle long points by splitting them
#                 if len(point_text) > 100:
#                     parts = [point_text[i:i+100] for i in range(0, len(point_text), 100)]
#                     p.text = parts[0]
#                     for part in parts[1:]:
#                         new_p = tf.add_paragraph()
#                         new_p.text = part
#                         new_p.level = 1  # Indent wrapped lines
#                         new_p.font.color.rgb = theme["text_color"]
#                         new_p.font.size = Pt(16)  # Slightly smaller for wrapped text
#                         if "text_font" in theme:
#                             new_p.font.name = theme["text_font"]
#                 else:
#                     p.text = point_text
                
#                 p.level = 0
#                 p.font.color.rgb = theme["text_color"]
#                 p.font.size = Pt(18)
#                 if "text_font" in theme:
#                     p.font.name = theme["text_font"]
#                 p.space_after = Pt(6)  # Reduced spacing
#                 p.space_before = Pt(6)
        
#         # Add notes if available
#         if slide_info.get('notes'):
#             notes_slide = slide.notes_slide
#             notes_slide.notes_text_frame.text = slide_info['notes']
    
#     pptx_io = io.BytesIO()
#     prs.save(pptx_io)
#     pptx_io.seek(0)
#     return pptx_io

# def main():
#     st.set_page_config(page_title="Advanced PPTX Generator", layout="wide")
    
#     st.title("Advanced PowerPoint Generator")
#     st.markdown("Create professional presentations with AI")
    
#     # Initialize session state for custom themes if not exists
#     if 'custom_themes' not in st.session_state:
#         st.session_state.custom_themes = {}
    
#     # Combine default and custom themes
#     ALL_THEMES = {**DEFAULT_THEMES, **st.session_state.custom_themes}
    
#     col1, col2 = st.columns([3, 1])
    
#     with col1:
#         topic = st.text_input("Presentation Topic:", 
#                             placeholder="Enter your topic (e.g., 'AI in Healthcare')")
        
#         with st.expander("Advanced Options"):
#             slide_count = st.slider("Number of Slides:", 5, 20, 10)
            
#             # Theme selection with example-based option
#             theme_option = st.radio("Theme Selection Method:", 
#                                   ["Predefined Theme", "Custom Theme", "Example-Based Theme"])
            
#             theme = DEFAULT_THEMES["Professional Blue"]  # Default theme
            
#             if theme_option == "Predefined Theme":
#                 theme_name = st.selectbox("Select Theme:", list(DEFAULT_THEMES.keys()))
#                 theme = DEFAULT_THEMES[theme_name]
#             elif theme_option == "Custom Theme":
#                 theme_name = st.selectbox("Select Custom Theme:", 
#                                         ["Create New..."] + list(st.session_state.custom_themes.keys()))
                
#                 if theme_name == "Create New...":
#                     with st.form("custom_theme_form"):
#                         new_theme_name = st.text_input("Theme Name")
#                         bg_color = st.color_picker("Background Color", "#0C2340")
#                         title_color = st.color_picker("Title Color", "#FFFFFF")
#                         text_color = st.color_picker("Text Color", "#C8C8C8")
#                         accent_color = st.color_picker("Accent Color", "#0070C0")
#                         title_font = st.text_input("Title Font", "Calibri")
#                         text_font = st.text_input("Text Font", "Calibri")
                        
#                         if st.form_submit_button("Save Custom Theme"):
#                             if new_theme_name:
#                                 st.session_state.custom_themes[new_theme_name] = {
#                                     "background": hex_to_rgb(bg_color),
#                                     "title_color": hex_to_rgb(title_color),
#                                     "text_color": hex_to_rgb(text_color),
#                                     "accent": hex_to_rgb(accent_color),
#                                     "title_font": title_font,
#                                     "text_font": text_font
#                                 }
#                                 st.success(f"Theme '{new_theme_name}' saved successfully!")
#                             else:
#                                 st.warning("Please enter a theme name")
#                 elif theme_name in st.session_state.custom_themes:
#                     theme = st.session_state.custom_themes[theme_name]
#             else:  # Example-Based Theme
#                 uploaded_file = st.file_uploader("Upload PowerPoint Template", type=["pptx"])
#                 if uploaded_file:
#                     with st.spinner("Extracting theme from template..."):
#                         theme = extract_theme_from_pptx(uploaded_file)
#                         st.success("Theme extracted from template!")
#                 else:
#                     st.info("Please upload a PowerPoint file to extract its theme")
    
#     with col2:
#         st.image("https://via.placeholder.com/150?text=PPT", width=150)
    
#     if st.button("Generate Presentation", type="primary"):
#         if not topic:
#             st.warning("Please enter a topic first!")
#         elif theme_option == "Example-Based Theme" and not uploaded_file:
#             st.warning("Please upload a PowerPoint template file first")
#         else:
#             with st.spinner(f"Creating {slide_count}-slide presentation about '{topic}'..."):
#                 try:
#                     slide_text = generate_slide_content(topic, slide_count)
#                     slides_data = parse_slide_content(slide_text)
                    
#                     # Show slide overview with detailed content
#                     with st.expander("Slide Overview (Detailed)"):
#                         for i, slide in enumerate(slides_data, 1):
#                             st.subheader(f"Slide {i}: {slide['title']}")
#                             st.markdown("**Content:**")
#                             for point in slide.get('content', []):
#                                 if ":" in point:
#                                     main_point, explanation = point.split(":", 1)
#                                     st.markdown(f"- **{main_point.strip()}**: {explanation.strip()}")
#                                 else:
#                                     st.markdown(f"- {point}")
#                             if slide.get('notes'):
#                                 st.markdown(f"**Notes:** {slide['notes']}")
#                             st.markdown("---")
                    
#                     pptx_file = create_detailed_pptx(slides_data, theme)
                    
#                     st.success("Presentation generated successfully!")
                    
#                     st.download_button(
#                         label="Download PowerPoint",
#                         data=pptx_file,
#                         file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
#                         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
#                     )
                    
#                 except Exception as e:
#                     st.error(f"An error occurred: {str(e)}")

# if __name__ == "__main__":
#     main()
import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.xmlchemy import OxmlElement
import io
import re
import tempfile
import os

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
        # Get colors from master slide
        slide_master = prs.slide_master
        theme["background"] = slide_master.background.fill.fore_color.rgb
        
        # Try to get title and text colors from placeholders
        for shape in slide_master.shapes:
            if shape.has_text_frame and shape.text.strip():
                if shape.name.lower() == 'title':
                    theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                    theme["title_font"] = shape.text_frame.paragraphs[0].font.name
                else:
                    theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                    theme["text_font"] = shape.text_frame.paragraphs[0].font.name
        
        # Try to get accent color from first shape with fill
        for shape in slide_master.shapes:
            if hasattr(shape, 'fill'):
                if shape.fill.type == 1:  # Solid fill
                    theme["accent"] = shape.fill.fore_color.rgb
                    break
    
    except Exception as e:
        st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
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
    
    # Set widescreen layout (16:9 aspect ratio)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Layout mapping - now using layout indices from the presentation
    layout_indices = {
        'title-only': 0,  # Typically index 0 is title slide
        'title-content': 1,  # Typically index 1 is title and content
        'section-header': 2,  # Might not exist in all templates
        'two-column': 3  # Might not exist in all templates
    }
    
    # Verify which layouts actually exist in the presentation
    available_layouts = {}
    for name, idx in layout_indices.items():
        if idx < len(prs.slide_layouts):
            available_layouts[name] = idx
    
    # Default to title-content if requested layout isn't available
    default_layout_idx = available_layouts.get('title-content', 0)
    
    for slide_info in slides_data:
        layout = slide_info.get('layout', 'title-content').lower()
        layout_idx = available_layouts.get(layout, default_layout_idx)
        
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
        except IndexError:
            slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
        # Apply background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = theme["background"]
        
        # Set title with theme and proper formatting
        title = slide.shapes.title
        title.text = slide_info['title']
        title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
        title.text_frame.paragraphs[0].font.size = Pt(36)
        title.text_frame.paragraphs[0].font.bold = True
        if "title_font" in theme:
            title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
        # Adjust title position to prevent overflow
        title.text_frame.word_wrap = True
        title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        title.top = Inches(0.5)
        title.height = Inches(1.5)
        title.width = Inches(11)
        
        # Set content based on layout with proper text formatting
        if layout_idx == 3:  # Two column layout
            content = slide_info.get('content', [])
            mid_point = len(content) // 2
            left_content = content[:mid_point]
            right_content = content[mid_point:]
            
            # Left column
            left_body = slide.placeholders[1]
            left_tf = left_body.text_frame
            left_tf.clear()
            left_body.width = Inches(5.5)  # Narrower columns
            left_body.left = Inches(0.5)
            
            # Right column
            right_body = slide.placeholders[2]
            right_tf = right_body.text_frame
            right_tf.clear()
            right_body.width = Inches(5.5)  # Narrower columns
            right_body.left = Inches(7.0)
            
            for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
                for point in content_part:
                    p = tf.add_paragraph()
                    point_text = point.replace('- ', '').strip()
                    # Split point if it's too long
                    if len(point_text) > 90:
                        parts = re.findall('.{1,90}(?:\s+|$)', point_text)
                        for part in parts:
                            p = tf.add_paragraph()
                            p.text = part.strip()
                    else:
                        p.text = point_text
                    p.font.size = Pt(16)
                    p.font.color.rgb = theme["text_color"]
                    p.font.name = theme["text_font"] if "text_font" in theme else 'Arial'
                    p.alignment = PP_ALIGN.LEFT
        
        else:  # Title-content or other layouts
            body = slide.shapes.placeholders[1]
            tf = body.text_frame
            tf.clear()
            
            for content_point in slide_info.get('content', []):
                p = tf.add_paragraph()
                p.text = content_point.strip()
                p.font.size = Pt(18)
                p.font.color.rgb = theme["text_color"]
                p.font.name = theme["text_font"]
                p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

    return prs

# Main Streamlit interface
st.title('AI-Powered PPTX Generator')

# Theme selection
selected_theme_name = st.selectbox('Select PowerPoint Theme', list(DEFAULT_THEMES.keys()))
theme = DEFAULT_THEMES[selected_theme_name]

# File upload for pptx-based theme extraction (optional)
uploaded_pptx = st.file_uploader("Upload a PowerPoint to Extract Theme (Optional)", type=["pptx"])
if uploaded_pptx:
    theme = extract_theme_from_pptx(uploaded_pptx)

# User inputs for the presentation
topic = st.text_input('Enter the topic of the presentation:')
slide_count = st.slider('Number of slides', min_value=1, max_value=50, value=10)

if st.button("Generate PowerPoint"):
    if not topic:
        st.warning("Please enter a topic for the presentation.")
    else:
        slide_text = generate_slide_content(topic, slide_count)
        slides_data = parse_slide_content(slide_text)
        presentation = create_detailed_pptx(slides_data, theme)
        
        # Save presentation to buffer
        pptx_buffer = io.BytesIO()
        presentation.save(pptx_buffer)
        pptx_buffer.seek(0)
        
        # Download link
        st.download_button(
            label="Download PPTX",
            data=pptx_buffer,
            file_name=f"{topic}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
