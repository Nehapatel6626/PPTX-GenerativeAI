# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
# from pptx.enum.text import MSO_AUTO_SIZE
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

# def hex_to_rgb(hex_color):
#     """Convert hex color to RGBColor"""
#     hex_color = hex_color.lstrip('#')
#     return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

# def extract_theme_from_pptx(uploaded_file):
#     """Extract theme colors and fonts from an uploaded PowerPoint file"""
#     with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
#         tmp_file.write(uploaded_file.read())
#         tmp_file_path = tmp_file.name
    
#     prs = Presentation(tmp_file_path)
#     theme = {
#         "background": RGBColor(255, 255, 255),  # Default white if not found
#         "title_color": RGBColor(0, 0, 0),       # Default black
#         "text_color": RGBColor(0, 0, 0),        # Default black
#         "accent": RGBColor(79, 129, 189),       # Default blue
#         "title_font": "Calibri",
#         "text_font": "Calibri",
#         "template_path": tmp_file_path  # Store the template path for later use
#     }
    
#     try:
#         # Get colors from master slide
#         slide_master = prs.slide_master
        
#         # Handle background color
#         if hasattr(slide_master.background, 'fill'):
#             if slide_master.background.fill.type == 1:  # Solid fill
#                 theme["background"] = slide_master.background.fill.fore_color.rgb
        
#         # Try to get title and text colors from placeholders
#         for shape in slide_master.shapes:
#             if shape.has_text_frame and shape.text.strip():
#                 try:
#                     if shape.name.lower() == 'title':
#                         theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                         theme["title_font"] = shape.text_frame.paragraphs[0].font.name
#                     else:
#                         theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                         theme["text_font"] = shape.text_frame.paragraphs[0].font.name
#                 except:
#                     continue
        
#         # Try to get accent color from first shape with fill
#         for shape in slide_master.shapes:
#             if hasattr(shape, 'fill'):
#                 try:
#                     if shape.fill.type == 1:  # Solid fill
#                         theme["accent"] = shape.fill.fore_color.rgb
#                         break
#                 except:
#                     continue
    
#     except Exception as e:
#         st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
#     return theme

# def generate_slide_content(topic, slide_count):
#     model = genai.GenerativeModel('gemini-2.0-flash')
#     prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
#     For each slide, provide:
#     1. A clear title in [Title:] format
#     2. 3-5 detailed bullet points in [Content:] format (each point should be 2-3 lines/40-60 words)
#     3. Optional speaker notes in [Notes:] format
#     4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
#     Structure your response like this:
#     [Title:] Slide Title
#     [Layout:] title-content
#     [Content:]
#     - Main Point 1: Detailed explanation spanning 1-2 lines with supporting information that provides context and value to the audience. This makes each point substantial.
#     - Main Point 2: Another complete thought with sufficient detail to stand alone as a mini-paragraph, giving the audience concrete information they can use.
#     - Main Point 3: Final point with enough depth to be meaningful, typically consisting of 2-3 sentences that develop a complete idea.
#     [Notes:] Additional notes here
    
#     Important guidelines:
#     - Do not give only placeholders or labels. Write full, rich content for each bullet point.
#     - Apply the same for the Questionnaire slide too, don't just keep placeholders, generate 10 questions.
#     - Avoid general outlines — generate fully written content as if it's going directly into a slide.
#     - Each bullet point should be 2-3 lines (30-50 words)
#     - Provide complete thoughts with supporting details
#     - Maintain parallel structure across points
#     - Avoid single-sentence bullet points
#     - Focus on substance over brevity
    
#     Include these sections (adjust based on requested slide count):
#     - Title slide
#     - Introduction/Overview
#     - Key Concepts
#     - Detailed Analysis
#     - Case Studies/Examples
#     - Applications
#     - Challenges
#     - Future Trends
#     - Conclusion
#     - Questionnaire 

#     After generating the content for the above presentation sections,titled [Title:], generate a Questionnaire Slide containing:
#     - 10 multiple-choice questions testing the user’s understanding of the presentation
#     - Each question should have 4 options (A, B, C, D)
#     - Clearly indicate the correct answer after each question using the format: [Correct Answer: C]
#     - Make the questions relevant and cover all key areas: introduction, concepts, analysis, case studies, applications, challenges, and future trends.
#     - Make sure the questions test the user's comprehension of the previous slides' content and are well-distributed across different topics in the presentation.

#     Please ensure the presentation content is presented first, followed by a clear separator, and then the questionnaire.

#     Begin the content generation now.
#     """
    
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
#             # Simplify bullet points - remove any explanations after colons
#             point = line[1:].strip()
#             if ':' in point:
#                 point = point.split(':')[0].strip()
#             current_slide['content'].append(point)
    
#     if current_slide:
#         slides.append(current_slide)
    
#     return slides
    
# def create_detailed_pptx(slides_data, theme, branding_options=None):
#     """Create PowerPoint using the uploaded template"""
#     # Use the template if one was uploaded
#     if "template_path" in theme and os.path.exists(theme["template_path"]):
#         prs = Presentation(theme["template_path"])
#     else:
#         prs = Presentation()
    
#     # Set widescreen layout (16:9 aspect ratio)
#     prs.slide_width = Inches(13.33)
#     prs.slide_height = Inches(7.5)
    
#     # Layout mapping
#     layout_indices = {
#         'title-only': 0,
#         'title-content': 1,
#         'section-header': 2,
#         'two-column': 3
#     }
    
#     available_layouts = {}
#     for name, idx in layout_indices.items():
#         if idx < len(prs.slide_layouts):
#             available_layouts[name] = idx
    
#     default_layout_idx = available_layouts.get('title-content', 0)
    
#     for slide_info in slides_data:
#         layout = slide_info.get('layout', 'title-content').lower()
#         layout_idx = available_layouts.get(layout, default_layout_idx)
        
#         try:
#             slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
#         except IndexError:
#             slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
#         # Set title
#         title = slide.shapes.title
#         title.text = slide_info['title']
#         title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
#         # Only apply custom formatting if not using a template
#         if "template_path" not in theme:
#             # Apply background
#             background = slide.background
#             fill = background.fill
#             fill.solid()
#             fill.fore_color.rgb = theme["background"]
            
#             # Format title
#             title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
#             title.text_frame.paragraphs[0].font.size = Pt(36)
#             title.text_frame.paragraphs[0].font.bold = True
#             if "title_font" in theme:
#                 title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
#         # Add logo if provided
#         if branding_options and branding_options.get('logo_path'):
#             add_logo_to_slide(slide, branding_options['logo_path'], 
#                             branding_options.get('logo_position', 'top-right'))
        
#         # Set content based on layout
#         if layout_idx == 3:  # Two column layout
#             content = slide_info.get('content', [])
#             mid_point = len(content) // 2
#             left_content = content[:mid_point]
#             right_content = content[mid_point:]
            
#             left_body = slide.placeholders[1]
#             left_tf = left_body.text_frame
#             left_tf.clear()
            
#             right_body = slide.placeholders[2]
#             right_tf = right_body.text_frame
#             right_tf.clear()
            
#             for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
#                 for point in content_part:
#                     p = tf.add_paragraph()
#                     point_text = point.replace('- ', '').strip()
#                     p.text = point_text
#                     p.level = 0
#                     p.alignment = PP_ALIGN.JUSTIFY 
#                     if "template_path" not in theme:  # Only apply custom formatting if no template
#                         p.font.color.rgb = theme["text_color"]
#                         p.font.size = Pt(18)
#                         if "text_font" in theme:
#                             p.font.name = theme["text_font"]
        
#         elif layout_idx != 0:  # Not title-only
#             body = slide.placeholders[1]
#             tf = body.text_frame
#             tf.clear()
            
#             for point in slide_info.get('content', []):
#                 p = tf.add_paragraph()
#                 point_text = point.replace('- ', '').strip()
#                 p.text = point_text
#                 p.level = 0
#                 p.alignment = PP_ALIGN.JUSTIFY 
#                 if "template_path" not in theme:  # Only apply custom formatting if no template
#                     p.font.color.rgb = theme["text_color"]
#                     p.font.size = Pt(18)
#                     if "text_font" in theme:
#                         p.font.name = theme["text_font"]
        
#         # Add notes if available
#         if slide_info.get('notes'):
#             notes_slide = slide.notes_slide
#             notes_slide.notes_text_frame.text = slide_info['notes']
    
#     pptx_io = io.BytesIO()
#     prs.save(pptx_io)
#     pptx_io.seek(0)
    
#     # Clean up temporary template file
#     if "template_path" in theme and os.path.exists(theme["template_path"]):
#         os.unlink(theme["template_path"])
    
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
#         # st.image("https://huggingface.co/spaces/YashMK89/PPTX_generativeai/resolve/main/SVNIT_IS_DS_Projects.png", width=150)
    
#      if st.button("Generate Presentation", type="primary"):
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
#                                 st.markdown(f"- {point}")
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

######################################################################################################################################################################
# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
# from pptx.enum.text import MSO_AUTO_SIZE
# from pptx.oxml.xmlchemy import OxmlElement
# import io
# import re
# import tempfile
# import os

# # Load the API key securely from environment variable
# api_key = os.getenv("GOOGLE_API_KEY")
# if not api_key:
#     st.error("GOOGLE_API_KEY environment variable not set!")
#     st.stop()
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

# def hex_to_rgb(hex_color):
#     """Convert hex color to RGBColor"""
#     hex_color = hex_color.lstrip('#')
#     return RGBColor(*tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

# def extract_theme_from_pptx(uploaded_file):
#     """Extract theme colors and fonts from an uploaded PowerPoint file"""
#     with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
#         tmp_file.write(uploaded_file.read())
#         tmp_file_path = tmp_file.name
    
#     prs = Presentation(tmp_file_path)
#     theme = {
#         "background": RGBColor(255, 255, 255),  # Default white if not found
#         "title_color": RGBColor(0, 0, 0),       # Default black
#         "text_color": RGBColor(0, 0, 0),        # Default black
#         "accent": RGBColor(79, 129, 189),       # Default blue
#         "title_font": "Calibri",
#         "text_font": "Calibri",
#         "template_path": tmp_file_path  # Store the template path for later use
#     }
    
#     try:
#         # Get colors from master slide
#         slide_master = prs.slide_master
        
#         # Handle background color
#         if hasattr(slide_master.background, 'fill'):
#             if slide_master.background.fill.type == 1:  # Solid fill
#                 theme["background"] = slide_master.background.fill.fore_color.rgb
        
#         # Try to get title and text colors from placeholders
#         for shape in slide_master.shapes:
#             if shape.has_text_frame and shape.text.strip():
#                 try:
#                     if 'title' in shape.name.lower():
#                         theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                         theme["title_font"] = shape.text_frame.paragraphs[0].font.name
#                     else:
#                         theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                         theme["text_font"] = shape.text_frame.paragraphs[0].font.name
#                 except:
#                     continue
        
#         # Try to get accent color from first shape with fill
#         for shape in slide_master.shapes:
#             if hasattr(shape, 'fill'):
#                 try:
#                     if shape.fill.type == 1:  # Solid fill
#                         theme["accent"] = shape.fill.fore_color.rgb
#                         break
#                 except:
#                     continue
    
#     except Exception as e:
#         st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
#     return theme

# def generate_slide_content(topic, slide_count):
#     model = genai.GenerativeModel('gemini-2.0-flash')
#     prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
#     For each slide, provide:
#     1. A clear title in [Title:] format
#     2. 3-5 detailed bullet points in [Content:] format (each point should be 2-3 lines/40-60 words)
#     3. Optional speaker notes in [Notes:] format
#     4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
#     Structure your response like this:
#     [Title:] Slide Title
#     [Layout:] title-content
#     [Content:]
#     - Main Point 1: Detailed explanation spanning 1-2 lines with supporting information that provides context and value to the audience. This makes each point substantial.
#     - Main Point 2: Another complete thought with sufficient detail to stand alone as a mini-paragraph, giving the audience concrete information they can use.
#     - Main Point 3: Final point with enough depth to be meaningful, typically consisting of 2-3 sentences that develop a complete idea.
#     [Notes:] Additional notes here
    
#     Important guidelines:
#     - Do not give only placeholders or labels. Write full, rich content for each bullet point.
#     - Apply the same for the Questionnaire slide too, don't just keep placeholders, generate 10 questions.
#     - Avoid general outlines — generate fully written content as if it's going directly into a slide.
#     - Each bullet point should be 2-3 lines (30-50 words)
#     - Provide complete thoughts with supporting details
#     - Maintain parallel structure across points
#     - Avoid single-sentence bullet points
#     - Focus on substance over brevity
    
#     Include these sections (adjust based on requested slide count):
#     - Title slide
#     - Introduction/Overview
#     - Key Concepts
#     - Detailed Analysis
#     - Case Studies/Examples
#     - Applications
#     - Challenges
#     - Future Trends
#     - Conclusion
    
#     After generating the content for the above presentation sections, add a clear separator:
#     --- QUESTIONNAIRE ---
    
#     Then generate 10 multiple-choice questions testing the user's understanding of the presentation:
#     Format each question like:
#     [Question: Your question text here?]
#     - [A] Option A
#     - [B] Option B
#     - [C] Option C
#     - [D] Option D
#     [Correct: C]
    
#     Ensure:
#     - Questions cover all key areas: introduction, concepts, analysis, case studies, applications, challenges, and future trends
#     - Each question has exactly 4 options
#     - Clearly indicate the correct answer
#     - Questions test comprehension of the presentation content
    
#     Please ensure the presentation content is presented first, followed by the separator, and then the questionnaire.

#     Begin the content generation now.
#     """
    
#     response = model.generate_content(prompt)
#     return response.text

# def parse_slide_content(slide_text):
#     slides = []
#     current_slide = {}
#     questionnaire = []
    
#     # Split into main content and questionnaire
#     parts = slide_text.split('--- QUESTIONNAIRE ---')
#     main_content = parts[0].strip()
#     questionnaire_content = parts[1].strip() if len(parts) > 1 else ""
    
#     # Parse main slides
#     for line in main_content.split('\n'):
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
#             # Simplify bullet points - remove any explanations after colons
#             point = line[1:].strip()
#             if ':' in point:
#                 point = point.split(':')[0].strip()
#             current_slide['content'].append(point)
    
#     if current_slide:
#         slides.append(current_slide)
    
#     # Parse questionnaire if exists
#     if questionnaire_content:
#         current_question = {}
#         for line in questionnaire_content.split('\n'):
#             line = line.strip()
#             if not line:
#                 continue
                
#             if line.startswith('[Question:'):
#                 if current_question:
#                     questionnaire.append(current_question)
#                 question_text = line.replace('[Question:', '').replace(']', '').strip()
#                 current_question = {
#                     'text': question_text,
#                     'options': [],
#                     'correct': ''
#                 }
#             elif line.startswith('- [A]'):
#                 current_question['options'].append(line.replace('- [A]', '').strip())
#             elif line.startswith('- [B]'):
#                 current_question['options'].append(line.replace('- [B]', '').strip())
#             elif line.startswith('- [C]'):
#                 current_question['options'].append(line.replace('- [C]', '').strip())
#             elif line.startswith('- [D]'):
#                 current_question['options'].append(line.replace('- [D]', '').strip())
#             elif line.startswith('[Correct:'):
#                 current_question['correct'] = line.replace('[Correct:', '').replace(']', '').strip()
        
#         if current_question:
#             questionnaire.append(current_question)
    
#     return slides, questionnaire

# def create_questionnaire_slide(prs, questions, theme):
#     """Create dedicated questionnaire slide"""
#     slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title + Content layout
    
#     # Set title
#     title = slide.shapes.title
#     title.text = "Knowledge Check"
    
#     # Only apply custom formatting if not using a template
#     if "template_path" not in theme:
#         title.text_frame.paragraphs[0].font.color.rgb = theme["accent"]
#         title.text_frame.paragraphs[0].font.size = Pt(36)
#         title.text_frame.paragraphs[0].font.bold = True
#         if "title_font" in theme:
#             title.text_frame.paragraphs[0].font.name = theme["title_font"]
    
#     # Create content
#     body = slide.placeholders[1]
#     tf = body.text_frame
#     tf.clear()
    
#     for i, q in enumerate(questions):
#         # Add question
#         p = tf.add_paragraph()
#         p.text = f"{i+1}. {q['text']}"
#         p.level = 0
#         p.font.bold = True
#         p.space_after = Pt(8)
        
#         # Only apply custom formatting if not using a template
#         if "template_path" not in theme:
#             p.font.color.rgb = theme["text_color"]
#             p.font.size = Pt(20)
#             if "text_font" in theme:
#                 p.font.name = theme["text_font"]
        
#         # Add options
#         for j, opt in enumerate(q['options']):
#             p = tf.add_paragraph()
#             p.text = f"   {chr(65+j)}. {opt}"
#             p.level = 1
#             p.space_after = Pt(4)
            
#             # Only apply custom formatting if not using a template
#             if "template_path" not in theme:
#                 p.font.color.rgb = theme["text_color"]
#                 p.font.size = Pt(18)
#                 if "text_font" in theme:
#                     p.font.name = theme["text_font"]
        
#         # Add space between questions
#         p = tf.add_paragraph()
#         p.text = ""
#         p.space_after = Pt(16)
    
#     # Add correct answers to speaker notes
#     notes_slide = slide.notes_slide
#     notes_text = "Correct Answers:\n\n"
#     for i, q in enumerate(questions):
#         notes_text += f"{i+1}. {q['correct']}\n"
#     notes_slide.notes_text_frame.text = notes_text
    
#     return slide

# def create_detailed_pptx(slides_data, questions, theme, branding_options=None):
#     """Create PowerPoint using the uploaded template"""
#     # Use the template if one was uploaded
#     if "template_path" in theme and os.path.exists(theme["template_path"]):
#         prs = Presentation(theme["template_path"])
#     else:
#         prs = Presentation()
    
#     # Set widescreen layout (16:9 aspect ratio)
#     prs.slide_width = Inches(13.33)
#     prs.slide_height = Inches(7.5)
    
#     # Layout mapping
#     layout_indices = {
#         'title-only': 0,
#         'title-content': 1,
#         'section-header': 2,
#         'two-column': 3
#     }
    
#     available_layouts = {}
#     for name, idx in layout_indices.items():
#         if idx < len(prs.slide_layouts):
#             available_layouts[name] = idx
    
#     default_layout_idx = available_layouts.get('title-content', 0)
    
#     # Create main slides
#     for slide_info in slides_data:
#         layout = slide_info.get('layout', 'title-content').lower()
#         layout_idx = available_layouts.get(layout, default_layout_idx)
        
#         try:
#             slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
#         except IndexError:
#             slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
#         # Set title
#         title = slide.shapes.title
#         title.text = slide_info['title']
#         title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
#         # Only apply custom formatting if not using a template
#         if "template_path" not in theme:
#             # Apply background
#             background = slide.background
#             fill = background.fill
#             fill.solid()
#             fill.fore_color.rgb = theme["background"]
            
#             # Format title
#             title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
#             title.text_frame.paragraphs[0].font.size = Pt(36)
#             title.text_frame.paragraphs[0].font.bold = True
#             if "title_font" in theme:
#                 title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
#         # Add logo if provided
#         if branding_options and branding_options.get('logo_path'):
#             add_logo_to_slide(slide, branding_options['logo_path'], 
#                             branding_options.get('logo_position', 'top-right'))
        
#         # Set content based on layout
#         if layout_idx == 3:  # Two column layout
#             content = slide_info.get('content', [])
#             mid_point = len(content) // 2
#             left_content = content[:mid_point]
#             right_content = content[mid_point:]
            
#             left_body = slide.placeholders[1]
#             left_tf = left_body.text_frame
#             left_tf.clear()
            
#             right_body = slide.placeholders[2]
#             right_tf = right_body.text_frame
#             right_tf.clear()
            
#             for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
#                 for point in content_part:
#                     p = tf.add_paragraph()
#                     point_text = point.replace('- ', '').strip()
#                     p.text = point_text
#                     p.level = 0
#                     p.alignment = PP_ALIGN.JUSTIFY 
#                     if "template_path" not in theme:  # Only apply custom formatting if no template
#                         p.font.color.rgb = theme["text_color"]
#                         p.font.size = Pt(18)
#                         if "text_font" in theme:
#                             p.font.name = theme["text_font"]
        
#         elif layout_idx != 0:  # Not title-only
#             body = slide.placeholders[1]
#             tf = body.text_frame
#             tf.clear()
            
#             for point in slide_info.get('content', []):
#                 p = tf.add_paragraph()
#                 point_text = point.replace('- ', '').strip()
#                 p.text = point_text
#                 p.level = 0
#                 p.alignment = PP_ALIGN.JUSTIFY 
#                 if "template_path" not in theme:  # Only apply custom formatting if no template
#                     p.font.color.rgb = theme["text_color"]
#                     p.font.size = Pt(18)
#                     if "text_font" in theme:
#                         p.font.name = theme["text_font"]
        
#         # Add notes if available
#         if slide_info.get('notes'):
#             notes_slide = slide.notes_slide
#             notes_slide.notes_text_frame.text = slide_info['notes']
    
#     # Add questionnaire slide if questions exist
#     if questions:
#         create_questionnaire_slide(prs, questions, theme)
    
#     pptx_io = io.BytesIO()
#     prs.save(pptx_io)
#     pptx_io.seek(0)
    
#     # Clean up temporary template file
#     if "template_path" in theme and os.path.exists(theme["template_path"]):
#         os.unlink(theme["template_path"])
    
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
#             uploaded_file = None

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
#         st.markdown("### Instructions")
#         st.markdown("1. Enter your presentation topic")
#         st.markdown("2. Select slide count and theme")
#         st.markdown("3. Click 'Generate Presentation'")
#         st.markdown("4. Download your PowerPoint file")
    
#     if st.button("Generate Presentation", type="primary", key="generate_btn"):
#         if not topic:
#             st.warning("Please enter a topic first!")
#         elif theme_option == "Example-Based Theme" and not uploaded_file:
#             st.warning("Please upload a PowerPoint template file first")
#         else:
#             with st.spinner(f"Creating {slide_count}-slide presentation about '{topic}'..."):
#                 try:
#                     slide_text = generate_slide_content(topic, slide_count)
#                     slides_data, questionnaire = parse_slide_content(slide_text)
                    
#                     # Show slide overview with detailed content
#                     with st.expander("Slide Overview (Detailed)"):
#                         for i, slide in enumerate(slides_data, 1):
#                             st.subheader(f"Slide {i}: {slide['title']}")
#                             st.markdown("**Content:**")
#                             for point in slide.get('content', []):
#                                 st.markdown(f"- {point}")
#                             if slide.get('notes'):
#                                 st.markdown(f"**Notes:** {slide['notes']}")
#                             st.markdown("---")
                        
#                         if questionnaire:
#                             st.subheader("Questionnaire")
#                             for i, q in enumerate(questionnaire, 1):
#                                 st.markdown(f"**{i}. {q['text']}**")
#                                 for j, opt in enumerate(q['options']):
#                                     st.markdown(f"- {chr(65+j)}. {opt}")
#                                 st.markdown(f"*Correct: {q['correct']}*")
#                                 st.markdown("---")
                    
#                     pptx_file = create_detailed_pptx(slides_data, questionnaire, theme)
                    
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


###################################################################################################################################################################
import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
import io
import re
import tempfile
import os
from PIL import Image
import time
import json
from hashlib import sha256

# Load the API key securely from environment variable
api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    st.error("GOOGLE_API_KEY environment variable not set!")
    st.stop()
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

# PowerPoint compatibility options
PPT_VERSIONS = {
    "Modern (2013+)": {"aspect_ratio": (13.33, 7.5)},
    "Standard (2007-2013)": {"aspect_ratio": (10, 7.5)},
    "Legacy (2003)": {"aspect_ratio": (10, 7.5)}
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
        "background": RGBColor(255, 255, 255),  # Default white if not found
        "title_color": RGBColor(0, 0, 0),       # Default black
        "text_color": RGBColor(0, 0, 0),        # Default black
        "accent": RGBColor(79, 129, 189),       # Default blue
        "title_font": "Calibri",
        "text_font": "Calibri",
        "template_path": tmp_file_path  # Store the template path for later use
    }
    
    try:
        # Get colors from master slide
        slide_master = prs.slide_master
        
        # Handle background color
        if hasattr(slide_master.background, 'fill'):
            if slide_master.background.fill.type == 1:  # Solid fill
                theme["background"] = slide_master.background.fill.fore_color.rgb
        
        # Try to get title and text colors from placeholders
        for shape in slide_master.shapes:
            if shape.has_text_frame and shape.text.strip():
                try:
                    if 'title' in shape.name.lower():
                        theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                        theme["title_font"] = shape.text_frame.paragraphs[0].font.name
                    else:
                        theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
                        theme["text_font"] = shape.text_frame.paragraphs[0].font.name
                except:
                    continue
        
        # Try to get accent color from first shape with fill
        for shape in slide_master.shapes:
            if hasattr(shape, 'fill'):
                try:
                    if shape.fill.type == 1:  # Solid fill
                        theme["accent"] = shape.fill.fore_color.rgb
                        break
                except:
                    continue
    
    except Exception as e:
        st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
    return theme

@st.cache_data(ttl=3600, show_spinner="Generating slide content...")
def generate_slide_content(topic, slide_count, _model=None):
    """Generate slide content with caching"""
    if _model is None:
        _model = genai.GenerativeModel('gemini-2.0-flash')
    
    prompt = f"""Create a comprehensive presentation on '{topic}' with exactly {slide_count} slides. 
    For each slide, provide:
    1. A clear title in [Title:] format
    2. 3-5 detailed bullet points in [Content:] format (each point should be 2-3 lines/40-60 words)
    3. Optional speaker notes in [Notes:] format
    4. Layout suggestion in [Layout:] format (title-only, title-content, two-column, section-header)
    
    Structure your response like this:
    [Title:] Slide Title
    [Layout:] title-content
    [Content:]
    - Main Point 1: Detailed explanation spanning 1-2 lines with supporting information that provides context and value to the audience. This makes each point substantial.
    - Main Point 2: Another complete thought with sufficient detail to stand alone as a mini-paragraph, giving the audience concrete information they can use.
    - Main Point 3: Final point with enough depth to be meaningful, typically consisting of 2-3 sentences that develop a complete idea.
    [Notes:] Additional notes here
    
    Important guidelines:
    - Do not give only placeholders or labels. Write full, rich content for each bullet point.
    - Apply the same for the Questionnaire slide too, don't just keep placeholders, generate 10 questions.
    - Avoid general outlines — generate fully written content as if it's going directly into a slide.
    - Each bullet point should be 2-3 lines (30-50 words)
    - Provide complete thoughts with supporting details
    - Maintain parallel structure across points
    - Avoid single-sentence bullet points
    - Focus on substance over brevity
    
    Include these sections (adjust based on requested slide count):
    - Title slide
    - Introduction/Overview
    - Key Concepts
    - Detailed Analysis
    - Case Studies/Examples
    - Applications
    - Challenges
    - Future Trends
    - Conclusion
    
    After generating the content for the above presentation sections, add a clear separator:
    --- QUESTIONNAIRE ---
    
    Then generate 10 multiple-choice questions testing the user's understanding of the presentation:
    Format each question like:
    [Question: Your question text here?]
    - [A] Option A
    - [B] Option B
    - [C] Option C
    - [D] Option D
    [Correct: C]
    
    Ensure:
    - Questions cover all key areas: introduction, concepts, analysis, case studies, applications, challenges, and future trends
    - Each question has exactly 4 options
    - Clearly indicate the correct answer
    - Questions test comprehension of the presentation content
    
    Finally, add a separator:
    --- ANSWER KEY ---
    
    Then provide an answer key listing all correct answers in the format:
    [Answer Key]
    1. C
    2. A
    3. B
    ... [continue for all 10 questions]
    
    Please ensure the presentation content is presented first, followed by the questionnaire, then the answer key.

    Begin the content generation now.
    """
    
    response = _model.generate_content(prompt)
    return response.text

def parse_slide_content(slide_text):
    slides = []
    current_slide = {}
    questionnaire = []
    answer_key = []
    
    # Split into sections
    parts = slide_text.split('--- QUESTIONNAIRE ---')
    main_content = parts[0].strip()
    remaining = parts[1].strip() if len(parts) > 1 else ""
    
    # Split remaining into questionnaire and answer key
    qa_parts = remaining.split('--- ANSWER KEY ---')
    questionnaire_content = qa_parts[0].strip() if qa_parts else ""
    answer_key_content = qa_parts[1].strip() if len(qa_parts) > 1 else ""
    
    # Parse main slides
    for line in main_content.split('\n'):
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
            # Simplify bullet points - remove any explanations after colons
            point = line[1:].strip()
            if ':' in point:
                point = point.split(':')[0].strip()
            current_slide['content'].append(point)
    
    if current_slide:
        slides.append(current_slide)
    
    # Parse questionnaire if exists
    if questionnaire_content:
        current_question = {}
        for line in questionnaire_content.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            if line.startswith('[Question:'):
                if current_question:
                    questionnaire.append(current_question)
                question_text = line.replace('[Question:', '').replace(']', '').strip()
                current_question = {
                    'text': question_text,
                    'options': [],
                    'correct': ''
                }
            elif line.startswith('- [A]'):
                current_question['options'].append(line.replace('- [A]', '').strip())
            elif line.startswith('- [B]'):
                current_question['options'].append(line.replace('- [B]', '').strip())
            elif line.startswith('- [C]'):
                current_question['options'].append(line.replace('- [C]', '').strip())
            elif line.startswith('- [D]'):
                current_question['options'].append(line.replace('- [D]', '').strip())
            elif line.startswith('[Correct:'):
                current_question['correct'] = line.replace('[Correct:', '').replace(']', '').strip()
        
        if current_question:
            questionnaire.append(current_question)
    
    # Parse answer key if exists
    if answer_key_content:
        # Extract answer key lines
        for line in answer_key_content.split('\n'):
            line = line.strip()
            if line.startswith('[') or not line:
                continue
            if re.match(r'^\d+\.\s*[A-D]$', line):
                answer_key.append(line.split('.')[1].strip())
    
    return slides, questionnaire, answer_key

def add_title_separator(slide, title_shape, accent_color):
    """Add a decorative line under the title"""
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        title_shape.left, 
        title_shape.top + title_shape.height - Pt(5),
        title_shape.width,
        Pt(3)
    )
    fill = line.fill
    fill.solid()
    fill.fore_color.rgb = accent_color
    line.line.fill.background()

def create_question_slide(prs, question, question_num, theme):
    """Create a slide for a single question"""
    # Try to use Title + Content layout, fallback to blank if not available
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title + Content layout
    except IndexError:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Remove unused placeholders
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.text.strip():
            sp = shape._element
            sp.getparent().remove(sp)
    
    # Set title
    title = slide.shapes.title if hasattr(slide.shapes, 'title') else None
    if title:
        title.text = f"Question {question_num}"
    
    # Only apply custom formatting if not using a template
    if "template_path" not in theme and title:
        title.text_frame.paragraphs[0].font.color.rgb = theme["accent"]
        title.text_frame.paragraphs[0].font.size = Pt(36)
        title.text_frame.paragraphs[0].font.bold = True
        if "title_font" in theme:
            title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
        # Add decorative line under title
        add_title_separator(slide, title, theme["accent"])
    
    # Create content
    left = Inches(1)
    top = Inches(2) if title else Inches(0.5)
    width = prs.slide_width - Inches(2)
    height = prs.slide_height - Inches(3)
    
    content_box = slide.shapes.add_textbox(left, top, width, height)
    tf = content_box.text_frame
    tf.word_wrap = True
    
    # Add question text
    p = tf.add_paragraph()
    p.text = question['text']
    p.font.bold = True
    p.space_after = Pt(24)
    
    # Only apply custom formatting if not using a template
    if "template_path" not in theme:
        p.font.color.rgb = theme["text_color"]
        p.font.size = Pt(24)
        if "text_font" in theme:
            p.font.name = theme["text_font"]
    
    # Add options
    for j, opt in enumerate(question['options']):
        p = tf.add_paragraph()
        p.text = f"{chr(65+j)}. {opt}"
        p.level = 0
        p.space_before = Pt(12)
        p.space_after = Pt(12)
        
        # Only apply custom formatting if not using a template
        if "template_path" not in theme:
            p.font.color.rgb = theme["text_color"]
            p.font.size = Pt(20)
            if "text_font" in theme:
                p.font.name = theme["text_font"]
    
    # Add correct answer to speaker notes
    if hasattr(slide, 'notes_slide'):
        notes_slide = slide.notes_slide
        notes_text = f"Correct Answer: {question['correct']}\n"
        if hasattr(notes_slide, 'notes_text_frame'):
            notes_slide.notes_text_frame.text = notes_text
    
    return slide

def create_answer_key_slide(prs, questions, theme):
    """Create answer key slide with all correct answers"""
    # Try to use Title + Content layout, fallback to blank if not available
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title + Content layout
    except IndexError:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Remove unused placeholders
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.text.strip():
            sp = shape._element
            sp.getparent().remove(sp)
    
    # Set title
    title = slide.shapes.title if hasattr(slide.shapes, 'title') else None
    if title:
        title.text = "Answer Key"
    
    # Only apply custom formatting if not using a template
    if "template_path" not in theme and title:
        title.text_frame.paragraphs[0].font.color.rgb = theme["accent"]
        title.text_frame.paragraphs[0].font.size = Pt(36)
        title.text_frame.paragraphs[0].font.bold = True
        if "title_font" in theme:
            title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
        # Add decorative line under title
        add_title_separator(slide, title, theme["accent"])
    
    # Create content
    left = Inches(1.5)
    top = Inches(2) if title else Inches(0.5)
    width = prs.slide_width - Inches(3)
    height = prs.slide_height - Inches(3)
    
    content_box = slide.shapes.add_textbox(left, top, width, height)
    tf = content_box.text_frame
    tf.word_wrap = True
    
    # Add answer key
    for i, q in enumerate(questions):
        p = tf.add_paragraph()
        p.text = f"Question {i+1}: {q['correct']}"
        p.level = 0
        p.space_after = Pt(18)
        
        # Only apply custom formatting if not using a template
        if "template_path" not in theme:
            p.font.color.rgb = theme["text_color"]
            p.font.size = Pt(24)
            p.font.bold = True
            if "text_font" in theme:
                p.font.name = theme["text_font"]
    
    return slide

def create_detailed_pptx(slides_data, questions, theme, branding_options=None, ppt_version="Modern (2013+)"):
    """Create PowerPoint using the uploaded template"""
    # Use the template if one was uploaded
    if "template_path" in theme and os.path.exists(theme["template_path"]):
        prs = Presentation(theme["template_path"])
    else:
        prs = Presentation()
    
    # Set aspect ratio based on selected version
    aspect_ratio = PPT_VERSIONS[ppt_version]["aspect_ratio"]
    prs.slide_width = Inches(aspect_ratio[0])
    prs.slide_height = Inches(aspect_ratio[1])
    
    # Layout mapping
    layout_indices = {
        'title-only': 0,
        'title-content': 1,
        'section-header': 2,
        'two-column': 3
    }
    
    available_layouts = {}
    for name, idx in layout_indices.items():
        if idx < len(prs.slide_layouts):
            available_layouts[name] = idx
    
    default_layout_idx = available_layouts.get('title-content', 0)
    
    # Create main slides
    for i, slide_info in enumerate(slides_data, 1):
        layout = slide_info.get('layout', 'title-content').lower()
        layout_idx = available_layouts.get(layout, default_layout_idx)
        
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
        except IndexError:
            slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
        # Remove unused placeholders
        for shape in slide.shapes:
            if shape.has_text_frame and not shape.text.strip():
                sp = shape._element
                sp.getparent().remove(sp)
        
        # Try to get the title placeholder
        title_placeholder = slide.shapes.title if hasattr(slide.shapes, 'title') else None
        
        # Set title if placeholder exists
        if title_placeholder is not None:
            title_placeholder.text = slide_info['title']
            title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Only apply custom formatting if not using a template
            if "template_path" not in theme:
                # Format title
                title_placeholder.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
                title_placeholder.text_frame.paragraphs[0].font.size = Pt(36)
                title_placeholder.text_frame.paragraphs[0].font.bold = True
                if "title_font" in theme:
                    title_placeholder.text_frame.paragraphs[0].font.name = theme["title_font"]
                
                # Add decorative line under title
                add_title_separator(slide, title_placeholder, theme["accent"])
        else:
            # Add title manually if no placeholder
            left = Inches(1)
            top = Inches(0.5)
            width = prs.slide_width - Inches(2)
            height = Inches(1)
            title_box = slide.shapes.add_textbox(left, top, width, height)
            title_tf = title_box.text_frame
            p = title_tf.add_paragraph()
            p.text = slide_info['title']
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(36)
            if "title_font" in theme:
                p.font.name = theme["title_font"]
            if "title_color" in theme:
                p.font.color.rgb = theme["title_color"]
        
        # Only apply custom formatting if not using a template
        if "template_path" not in theme:
            # Apply background
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = theme["background"]
        
        # Set content based on layout
        if layout == 'two-column':
            content = slide_info.get('content', [])
            mid_point = len(content) // 2
            left_content = content[:mid_point]
            right_content = content[mid_point:]
            
            # Create left column
            left = Inches(0.5)
            top = Inches(2)
            width = prs.slide_width / 2 - Inches(1)
            height = prs.slide_height - Inches(3)
            left_box = slide.shapes.add_textbox(left, top, width, height)
            left_tf = left_box.text_frame
            
            # Create right column
            left = prs.slide_width / 2 + Inches(0.5)
            right_box = slide.shapes.add_textbox(left, top, width, height)
            right_tf = right_box.text_frame
            
            for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
                for point in content_part:
                    p = tf.add_paragraph()
                    point_text = point.replace('- ', '').strip()
                    p.text = point_text
                    p.level = 0
                    p.alignment = PP_ALIGN.JUSTIFY
                    p.space_after = Pt(12)
                    if "template_path" not in theme:  # Only apply custom formatting if no template
                        p.font.color.rgb = theme["text_color"]
                        p.font.size = Pt(18)
                        if "text_font" in theme:
                            p.font.name = theme["text_font"]
        
        elif layout != 'title-only':  # Content slides
            # Create content area
            left = Inches(1)
            top = Inches(2)
            width = prs.slide_width - Inches(2)
            height = prs.slide_height - Inches(3)
            content_box = slide.shapes.add_textbox(left, top, width, height)
            tf = content_box.text_frame
            
            for point in slide_info.get('content', []):
                p = tf.add_paragraph()
                point_text = point.replace('- ', '').strip()
                p.text = point_text
                p.level = 0
                p.alignment = PP_ALIGN.JUSTIFY
                p.space_after = Pt(12)
                if "template_path" not in theme:  # Only apply custom formatting if no template
                    p.font.color.rgb = theme["text_color"]
                    p.font.size = Pt(18)
                    if "text_font" in theme:
                        p.font.name = theme["text_font"]
        
        # Add notes if available
        if slide_info.get('notes') and hasattr(slide, 'notes_slide'):
            notes_slide = slide.notes_slide
            if hasattr(notes_slide, 'notes_text_frame'):
                notes_slide.notes_text_frame.text = slide_info['notes']
    
    # Add question slides (one per question)
    if questions:
        # Add section header
        try:
            section_slide = prs.slides.add_slide(prs.slide_layouts[2])  # Section header layout
        except IndexError:
            section_slide = prs.slides.add_slide(prs.slide_layouts[0])  # Fallback to title slide
        
        title_placeholder = section_slide.shapes.title if hasattr(section_slide.shapes, 'title') else None
        if title_placeholder is not None:
            title_placeholder.text = "Knowledge Check"
            
            # Format section header
            if "template_path" not in theme:
                title_placeholder.text_frame.paragraphs[0].font.color.rgb = theme["accent"]
                title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
        
        # Add each question on a separate slide
        for i, question in enumerate(questions, 1):
            create_question_slide(prs, question, i, theme)
        
        # Add answer key slide
        create_answer_key_slide(prs, questions, theme)
    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    # Clean up temporary template file
    if "template_path" in theme and os.path.exists(theme["template_path"]):
        os.unlink(theme["template_path"])
    
    return pptx_io

def regenerate_slide(slide_index, topic, slide_count, theme, ppt_version):
    """Regenerate a specific slide using AI"""
    model = genai.GenerativeModel('gemini-2.0-flash')
    prompt = f"""We are creating a presentation on '{topic}' with {slide_count} slides. 
    Please regenerate only slide number {slide_index+1} with the same structure as the others.
    The slide should have:
    - A clear title in [Title:] format
    - 3-5 detailed bullet points in [Content:] format
    - Optional speaker notes in [Notes:] format
    - Layout suggestion in [Layout:] format
    
    Format:
    [Title:] Slide Title
    [Layout:] layout-type
    [Content:]
    - Point 1
    - Point 2
    [Notes:] Optional notes
    
    Make sure the content is relevant to the overall topic and flows with the presentation.
    """
    
    response = model.generate_content(prompt)
    new_slide_text = response.text
    
    # Parse the new slide
    new_slide = {'title': '', 'content': [], 'notes': '', 'layout': 'title-content'}
    for line in new_slide_text.split('\n'):
        line = line.strip()
        if line.startswith('[Title:]'):
            new_slide['title'] = line.replace('[Title:]', '').strip()
        elif line.startswith('[Content:]'):
            pass  # Content follows on subsequent lines
        elif line.startswith('[Notes:]'):
            new_slide['notes'] = line.replace('[Notes:]', '').strip()
        elif line.startswith('[Layout:]'):
            layout = line.replace('[Layout:]', '').strip().lower()
            valid_layouts = ['title-only', 'title-content', 'two-column', 'section-header']
            new_slide['layout'] = layout if layout in valid_layouts else 'title-content'
        elif line.startswith('-'):
            point = line[1:].strip()
            if ':' in point:
                point = point.split(':')[0].strip()
            new_slide['content'].append(point)
    
    return new_slide

def main():
    st.set_page_config(page_title="Advanced PPTX Generator", layout="wide")
    
    st.title("Advanced PowerPoint Generator")
    st.markdown("Create professional presentations with AI")
    
    # Initialize session state for custom themes if not exists
    if 'custom_themes' not in st.session_state:
        st.session_state.custom_themes = {}
    
    # Initialize session state for presentation data
    if 'slides_data' not in st.session_state:
        st.session_state.slides_data = []
    if 'questionnaire' not in st.session_state:
        st.session_state.questionnaire = []
    if 'answer_key' not in st.session_state:
        st.session_state.answer_key = []
    if 'pptx_file' not in st.session_state:
        st.session_state.pptx_file = None
    if 'regenerate_map' not in st.session_state:
        st.session_state.regenerate_map = {}
    
    # Combine default and custom themes
    ALL_THEMES = {**DEFAULT_THEMES, **st.session_state.custom_themes}
    col1, col2 = st.columns([3, 1])
    
    with col1:
        topic = st.text_input("Presentation Topic:", 
                            placeholder="Enter your topic (e.g., 'AI in Healthcare')")
        
        with st.expander("Advanced Options"):
            slide_count = st.slider("Number of Slides:", 5, 20, 10)
            
            # Theme selection with example-based option
            theme_option = st.radio("Theme Selection Method:", 
                                  ["Predefined Theme", "Custom Theme", "Example-Based Theme"])
            
            theme = DEFAULT_THEMES["Professional Blue"]  # Default theme
            uploaded_file = None

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
                        title_font = st.text_input("Title Font", "Calibri")
                        text_font = st.text_input("Text Font", "Calibri")
                        
                        if st.form_submit_button("Save Custom Theme"):
                            if new_theme_name:
                                st.session_state.custom_themes[new_theme_name] = {
                                    "background": hex_to_rgb(bg_color),
                                    "title_color": hex_to_rgb(title_color),
                                    "text_color": hex_to_rgb(text_color),
                                    "accent": hex_to_rgb(accent_color),
                                    "title_font": title_font,
                                    "text_font": text_font
                                }
                                st.success(f"Theme '{new_theme_name}' saved successfully!")
                            else:
                                st.warning("Please enter a theme name")
                elif theme_name in st.session_state.custom_themes:
                    theme = st.session_state.custom_themes[theme_name]
            else:  # Example-Based Theme
                uploaded_file = st.file_uploader("Upload PowerPoint Template", type=["pptx"])
                if uploaded_file:
                    with st.spinner("Extracting theme from template..."):
                        theme = extract_theme_from_pptx(uploaded_file)
                        st.success("Theme extracted from template!")
                else:
                    st.info("Please upload a PowerPoint file to extract its theme")
            
            # PowerPoint version compatibility
            ppt_version = st.selectbox(
                "PowerPoint Version Compatibility:",
                list(PPT_VERSIONS.keys()),
                index=0
            )
            st.caption("Modern (2013+): Widescreen (16:9)\nStandard (2007-2013): Standard (4:3)\nLegacy (2003): Standard (4:3)")
    
    with col2:
        st.markdown("### Instructions")
        st.markdown("1. Enter your presentation topic")
        st.markdown("2. Select slide count and theme")
        st.markdown("3. Click 'Generate Presentation'")
        st.markdown("4. Download your PowerPoint file")
        st.markdown("5. Use 'Regenerate' buttons to improve specific slides")
    
    generate_btn = st.button("Generate Presentation", type="primary", key="generate_btn")
    if generate_btn:
        if not topic:
            st.warning("Please enter a topic first!")
        elif theme_option == "Example-Based Theme" and not uploaded_file:
            st.warning("Please upload a PowerPoint template file first")
        else:
            with st.spinner("Generating your presentation..."):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Generate slide content
                status_text.text("Generating content with AI...")
                progress_bar.progress(20)
                model = genai.GenerativeModel('gemini-2.0-flash')
                slide_text = generate_slide_content(topic, slide_count, _model=model)
                
                # Parse slide content
                status_text.text("Parsing slide content...")
                progress_bar.progress(40)
                slides_data, questionnaire, answer_key = parse_slide_content(slide_text)
                
                # Store in session state
                st.session_state.slides_data = slides_data
                st.session_state.questionnaire = questionnaire
                st.session_state.answer_key = answer_key
                st.session_state.regenerate_map = {i: False for i in range(len(slides_data))}
                
                # Create PowerPoint
                status_text.text("Creating PowerPoint file...")
                progress_bar.progress(70)
                pptx_file = create_detailed_pptx(slides_data, questionnaire, theme, ppt_version=ppt_version)
                st.session_state.pptx_file = pptx_file
                
                progress_bar.progress(100)
                status_text.text("Done!")
                time.sleep(1)
                progress_bar.empty()
                status_text.empty()
                
                st.success("Presentation generated successfully!")
    
    # Show slide overview with regeneration options
    if st.session_state.slides_data:
        with st.expander("Slide Overview (Click to Expand)", expanded=True):
            st.markdown(f"**Total Slides:** {len(st.session_state.slides_data)} content slides + "
                        f"{len(st.session_state.questionnaire)} question slides + 1 answer key slide")
            
            for i, slide in enumerate(st.session_state.slides_data):
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.subheader(f"Slide {i+1}: {slide['title']}")
                    st.markdown("**Content:**")
                    for point in slide.get('content', []):
                        st.markdown(f"- {point}")
                    if slide.get('notes'):
                        st.markdown(f"**Notes:** {slide['notes']}")
                
                with col2:
                    if st.button(f"Regenerate", key=f"reg_{i}"):
                        with st.spinner(f"Regenerating slide {i+1}..."):
                            new_slide = regenerate_slide(i, topic, slide_count, theme, ppt_version)
                            st.session_state.slides_data[i] = new_slide
                            st.session_state.regenerate_map[i] = True
                            st.rerun()
                    
                    # Show regeneration status
                    if st.session_state.regenerate_map.get(i, False):
                        st.success("Regenerated!")
                
                st.markdown("---")
            
            if st.session_state.questionnaire:
                st.subheader("Questionnaire")
                st.markdown(f"**Total Questions:** {len(st.session_state.questionnaire)}")
                for i, q in enumerate(st.session_state.questionnaire, 1):
                    st.subheader(f"Question {i}")
                    st.markdown(f"**{q['text']}**")
                    for j, opt in enumerate(q['options']):
                        st.markdown(f"{chr(65+j)}. {opt}")
                    st.markdown(f"*Correct: {q['correct']}*")
                    st.markdown("---")
                
                if st.session_state.answer_key:
                    st.subheader("Answer Key")
                    for i, ans in enumerate(st.session_state.answer_key, 1):
                        st.markdown(f"{i}. {ans}")
    
    # Rebuild PPTX if any slides were regenerated
    if any(st.session_state.regenerate_map.values()):
        with st.spinner("Updating presentation with regenerated slides..."):
            pptx_file = create_detailed_pptx(
                st.session_state.slides_data,
                st.session_state.questionnaire,
                theme,
                ppt_version=ppt_version
            )
            st.session_state.pptx_file = pptx_file
            st.session_state.regenerate_map = {i: False for i in range(len(st.session_state.slides_data))}
            st.success("Presentation updated with regenerated slides!")
    
    # Download button
    if st.session_state.pptx_file:
        st.download_button(
            label="Download PowerPoint",
            data=st.session_state.pptx_file,
            file_name=f"{topic.replace(' ', '_') if topic else 'presentation'}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

if __name__ == "__main__":
    main()