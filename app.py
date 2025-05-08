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
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import datetime
import json
from PIL import Image

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

# Template configurations
TEMPLATES = {
    "Default": None,
    "Business": "templates/business.pptx",
    "Creative": "templates/creative.pptx",
    "Academic": "templates/academic.pptx",
    "Minimalist": "templates/minimalist.pptx"
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
    model = genai.GenerativeModel('gemini-pro')
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

def add_logo_to_slide(slide, logo_path, position='top-right', size=(1.5, 1.5)):
    """Add logo to slide at specified position"""
    left, top = 0, 0
    width, height = Inches(size[0]), Inches(size[1])
    
    if position == 'top-right':
        left = slide.slide_width - width - Inches(0.5)
        top = Inches(0.5)
    elif position == 'top-left':
        left = Inches(0.5)
        top = Inches(0.5)
    elif position == 'bottom-right':
        left = slide.slide_width - width - Inches(0.5)
        top = slide.slide_height - height - Inches(0.5)
    elif position == 'bottom-left':
        left = Inches(0.5)
        top = slide.slide_height - height - Inches(0.5)
    elif position == 'center':
        left = (slide.slide_width - width) / 2
        top = (slide.slide_height - height) / 2
    
    slide.shapes.add_picture(logo_path, left, top, width=width, height=height)

def create_detailed_pptx(slides_data, theme, template=None, logo_path=None, branding_options=None):
    """Create PowerPoint with enhanced features"""
    if template and os.path.exists(template):
        prs = Presentation(template)
    else:
        prs = Presentation()
    
    # Set widescreen layout (16:9 aspect ratio)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Apply branding if specified
    if branding_options:
        for slide_master in prs.slide_masters:
            if branding_options.get('background_color'):
                slide_master.background.fill.solid()
                slide_master.background.fill.fore_color.rgb = branding_options['background_color']
    
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
    
    for slide_info in slides_data:
        layout = slide_info.get('layout', 'title-content').lower()
        layout_idx = available_layouts.get(layout, default_layout_idx)
        
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
        except IndexError:
            slide = prs.slides.add_slide(prs.slide_layouts[default_layout_idx])
        
        # Apply theme
        if not template:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = theme["background"]
        
        # Set title
        title = slide.shapes.title
        title.text = slide_info['title']
        if not template:
            title.text_frame.paragraphs[0].font.color.rgb = theme["title_color"]
            title.text_frame.paragraphs[0].font.size = Pt(36)
            title.text_frame.paragraphs[0].font.bold = True
            if "title_font" in theme:
                title.text_frame.paragraphs[0].font.name = theme["title_font"]
        
        # Add logo if provided
        if logo_path and os.path.exists(logo_path):
            add_logo_to_slide(slide, logo_path, branding_options.get('logo_position', 'top-right'))
        
        # Set content based on layout
        if layout_idx == 3:  # Two column layout
            content = slide_info.get('content', [])
            mid_point = len(content) // 2
            left_content = content[:mid_point]
            right_content = content[mid_point:]
            
            left_body = slide.placeholders[1]
            left_tf = left_body.text_frame
            left_tf.clear()
            
            right_body = slide.placeholders[2]
            right_tf = right_body.text_frame
            right_tf.clear()
            
            for content_part, tf in [(left_content, left_tf), (right_content, right_tf)]:
                for point in content_part:
                    p = tf.add_paragraph()
                    point_text = point.replace('- ', '').strip()
                    p.text = point_text
                    p.level = 0
                    if not template:
                        p.font.color.rgb = theme["text_color"]
                        p.font.size = Pt(18)
                        if "text_font" in theme:
                            p.font.name = theme["text_font"]
        
        elif layout_idx != 0:  # Not title-only
            body = slide.placeholders[1]
            tf = body.text_frame
            tf.clear()
            
            for point in slide_info.get('content', []):
                p = tf.add_paragraph()
                point_text = point.replace('- ', '').strip()
                p.text = point_text
                p.level = 0
                if not template:
                    p.font.color.rgb = theme["text_color"]
                    p.font.size = Pt(18)
                    if "text_font" in theme:
                        p.font.name = theme["text_font"]
        
        # Add notes if available
        if slide_info.get('notes'):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_info['notes']
    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def send_email_with_attachment(receiver_email, subject, message, attachment_data, filename):
    """Send email with attachment"""
    sender_email = st.secrets.get("EMAIL_ADDRESS")
    password = st.secrets.get("EMAIL_PASSWORD")
    
    if not sender_email or not password:
        st.error("Email configuration not set up properly")
        return False
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(message, 'plain'))
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    msg.attach(part)
    
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Failed to send email: {str(e)}")
        return False

def save_to_history(user_id, presentation_data):
    """Save presentation to user history"""
    if 'history' not in st.session_state:
        st.session_state.history = {}
    
    if user_id not in st.session_state.history:
        st.session_state.history[user_id] = []
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.history[user_id].append({
        'timestamp': timestamp,
        'topic': presentation_data['topic'],
        'slides': presentation_data['slides'],
        'theme': presentation_data['theme']
    })

def get_download_link(file_data, filename, text):
    """Generate download link"""
    b64 = base64.b64encode(file_data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{text}</a>'

def main():
    st.set_page_config(page_title="Advanced PPTX Generator", layout="wide")
    
    # Initialize session state
    if 'history' not in st.session_state:
        st.session_state.history = {}
    if 'custom_themes' not in st.session_state:
        st.session_state.custom_themes = {}
    if 'logo_path' not in st.session_state:
        st.session_state.logo_path = None
    
    st.title("🚀 Advanced PowerPoint Generator")
    st.markdown("Create professional presentations with AI in minutes")
    
    # User authentication (simplified for demo)
    user_id = st.text_input("Enter your user ID (email or username):")
    if not user_id:
        st.warning("Please enter a user ID to enable history and sharing features")
    
    # Main layout
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # Presentation topic
        topic = st.text_input("Presentation Topic:", 
                            placeholder="Enter your topic (e.g., 'AI in Healthcare')")
        
        with st.expander("⚙️ Advanced Options"):
            # Slide count
            slide_count = st.slider("Number of Slides:", 5, 20, 10)
            
            # Template selection
            template_option = st.selectbox("Template:", list(TEMPLATES.keys()))
            
            # Theme selection
            theme_option = st.radio("Theme Selection:", 
                                  ["Predefined Theme", "Custom Theme", "Extract from PPTX"])
            
            theme = DEFAULT_THEMES["Professional Blue"]  # Default
            
            if theme_option == "Predefined Theme":
                theme_name = st.selectbox("Select Theme:", list(DEFAULT_THEMES.keys()))
                theme = DEFAULT_THEMES[theme_name]
            elif theme_option == "Custom Theme":
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
                            st.success(f"Theme '{new_theme_name}' saved!")
            else:
                uploaded_pptx = st.file_uploader("Upload PowerPoint Template", type=["pptx"])
                if uploaded_pptx:
                    theme = extract_theme_from_pptx(uploaded_pptx)
            
            # Branding options
            with st.expander("🏢 Branding Options"):
                logo_file = st.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])
                if logo_file:
                    # Save logo temporarily
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
                        tmp_logo.write(logo_file.read())
                        st.session_state.logo_path = tmp_logo.name
                
                logo_position = st.selectbox("Logo Position", 
                                          ["top-right", "top-left", "bottom-right", "bottom-left"])
                
                branding_options = {
                    'logo_position': logo_position,
                    'background_color': st.color_picker("Brand Background Color", "#FFFFFF")
                }
    
    with col2:
        st.image("https://via.placeholder.com/150?text=PPT", width=150)
        st.markdown("### Features:")
        st.markdown("- AI-powered content")
        st.markdown("- Custom templates")
        st.markdown("- Branding options")
        st.markdown("- Share via email")
        st.markdown("- Presentation history")
    
    # Generate button
    if st.button("✨ Generate Presentation", type="primary"):
        if not topic:
            st.warning("Please enter a topic first!")
        else:
            with st.spinner("🧠 Generating your presentation..."):
                try:
                    # Generate content
                    slide_text = generate_slide_content(topic, slide_count)
                    slides_data = parse_slide_content(slide_text)
                    
                    # Create presentation
                    pptx_file = create_detailed_pptx(
                        slides_data, 
                        theme,
                        template=TEMPLATES[template_option],
                        logo_path=st.session_state.get('logo_path'),
                        branding_options=branding_options
                    )
                    
                    # Save to history
                    if user_id:
                        save_to_history(user_id, {
                            'topic': topic,
                            'slides': slides_data,
                            'theme': theme
                        })
                    
                    # Display preview
                    st.success("🎉 Presentation generated successfully!")
                    
                    with st.expander("📝 Slide Preview"):
                        for i, slide in enumerate(slides_data, 1):
                            st.subheader(f"Slide {i}: {slide['title']}")
                            for point in slide.get('content', []):
                                st.markdown(f"- {point}")
                            if slide.get('notes'):
                                st.caption(f"Notes: {slide['notes']}")
                            st.divider()
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=pptx_file,
                        file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    
                    # Sharing options
                    with st.expander("📤 Share Presentation"):
                        share_option = st.radio("Share via:", ["Email", "Link"])
                        
                        if share_option == "Email":
                            receiver_email = st.text_input("Recipient Email:")
                            subject = st.text_input("Subject:", f"Presentation: {topic}")
                            message = st.text_area("Message:", f"Here's the presentation about {topic} that you requested.")
                            
                            if st.button("Send Email"):
                                if send_email_with_attachment(
                                    receiver_email,
                                    subject,
                                    message,
                                    pptx_file.getvalue(),
                                    f"{topic}_presentation.pptx"
                                ):
                                    st.success("Email sent successfully!")
                        else:
                            st.markdown(get_download_link(
                                pptx_file.getvalue(),
                                f"{topic}_presentation.pptx",
                                "Copy this link to share your presentation"
                            ), unsafe_allow_html=True)
                    
                    # History section
                    if user_id and user_id in st.session_state.history:
                        with st.expander("🕒 Presentation History"):
                            for i, item in enumerate(st.session_state.history[user_id][-3:][::-1], 1):
                                st.markdown(f"**{i}. {item['topic']}** ({item['timestamp']})")
                                if st.button(f"Regenerate {i}", key=f"regenerate_{i}"):
                                    st.session_state.regenerate_topic = item['topic']
                                    st.experimental_rerun()
                
                except Exception as e:
                    st.error(f"❌ An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
# import streamlit as st
# import google.generativeai as genai
# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
# from pptx.enum.shapes import MSO_SHAPE_TYPE
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
#         "background": RGBColor(0, 0, 0),  # Default black if not found
#         "title_color": RGBColor(0, 0, 0),
#         "text_color": RGBColor(0, 0, 0),
#         "accent": RGBColor(0, 0, 0),
#         "title_font": "Calibri",
#         "text_font": "Calibri"
#     }
    
#     try:
#         # Get colors from master slide
#         slide_master = prs.slide_master
#         theme["background"] = slide_master.background.fill.fore_color.rgb
        
#         # Try to get title and text colors from placeholders
#         for shape in slide_master.shapes:
#             if shape.has_text_frame and shape.text.strip():
#                 if shape.name.lower() == 'title':
#                     theme["title_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                     theme["title_font"] = shape.text_frame.paragraphs[0].font.name
#                 else:
#                     theme["text_color"] = shape.text_frame.paragraphs[0].font.color.rgb
#                     theme["text_font"] = shape.text_frame.paragraphs[0].font.name
        
#         # Try to get accent color from first shape with fill
#         for shape in slide_master.shapes:
#             if hasattr(shape, 'fill'):
#                 if shape.fill.type == 1:  # Solid fill
#                     theme["accent"] = shape.fill.fore_color.rgb
#                     break
    
#     except Exception as e:
#         st.warning(f"Couldn't fully extract theme: {str(e)}. Using default colors where needed.")
    
#     os.unlink(tmp_file_path)
#     return theme

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
#         title.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
#         title.top = Inches(0.5)
#         title.height = Inches(1.5)
#         title.width = Inches(11)
        
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
#                    #Set the bullet text (even if it's long, let PowerPoint wrap it)
                     # p.text = point_text
                     # p.level = 0  # ensure it becomes a bullet point

#                     p.font.size = Pt(16)
#                     p.font.color.rgb = theme["text_color"]
#                     p.font.name = theme["text_font"] if "text_font" in theme else 'Arial'
#                     p.alignment = PP_ALIGN.LEFT
        
#         else:  # Title-content or other layouts
#             body = slide.shapes.placeholders[1]
#             tf = body.text_frame
#             tf.clear()
            
#             for content_point in slide_info.get('content', []):
#                 p = tf.add_paragraph()
#                 p.text = content_point.strip()
#                 p.font.size = Pt(18)
#                 p.font.color.rgb = theme["text_color"]
#                 p.font.name = theme["text_font"]
#                 p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

#     return prs

# # Main Streamlit interface
# st.title('AI-Powered PPTX Generator')

# # Theme selection
# selected_theme_name = st.selectbox('Select PowerPoint Theme', list(DEFAULT_THEMES.keys()))
# theme = DEFAULT_THEMES[selected_theme_name]

# # File upload for pptx-based theme extraction (optional)
# uploaded_pptx = st.file_uploader("Upload a PowerPoint to Extract Theme (Optional)", type=["pptx"])
# if uploaded_pptx:
#     theme = extract_theme_from_pptx(uploaded_pptx)

# # User inputs for the presentation
# topic = st.text_input('Enter the topic of the presentation:')
# slide_count = st.slider('Number of slides', min_value=1, max_value=50, value=10)

# if st.button("Generate PowerPoint"):
#     if not topic:
#         st.warning("Please enter a topic for the presentation.")
#     else:
#         slide_text = generate_slide_content(topic, slide_count)
#         slides_data = parse_slide_content(slide_text)
#         presentation = create_detailed_pptx(slides_data, theme)
        
#         # Save presentation to buffer
#         pptx_buffer = io.BytesIO()
#         presentation.save(pptx_buffer)
#         pptx_buffer.seek(0)
        
#         # Download link
#         st.download_button(
#             label="Download PPTX",
#             data=pptx_buffer,
#             file_name=f"{topic}_presentation.pptx",
#             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
#         )
