import streamlit as st
import requests
from PIL import Image
from io import BytesIO
import os

# Set up the page
st.set_page_config(page_title="Pexels Image Search", layout="wide")
st.title("📷 Pexels Image Search with Attribution")
st.write("Search for images and properly credit the photographers")

# Get API key from Hugging Face secrets
api_key = os.environ.get("PEXELS_API_KEY")

if not api_key:
    st.error("Pexels API key not found. Please set it in Hugging Face secrets.")
    st.stop()

def search_pexels(query, per_page=9, page=1):
    headers = {"Authorization": api_key}
    url = f"https://api.pexels.com/v1/search?query={query}&per_page={per_page}&page={page}"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Error searching Pexels: {str(e)}")
        return None

# Initialize session state
if 'page' not in st.session_state:
    st.session_state.page = 1
if 'selected_image' not in st.session_state:
    st.session_state.selected_image = None
if 'show_modal' not in st.session_state:
    st.session_state.show_modal = False

# Search interface
search_query = st.text_input("🔍 Search for images", placeholder="Enter search term...", key="search_input")

if search_query:
    with st.spinner("Searching Pexels..."):
        results = search_pexels(search_query, page=st.session_state.page)
    
    if results and 'photos' in results:
        st.success(f"Found {results['total_results']} results")
        
        # Display in a grid
        cols = st.columns(3)
        for idx, photo in enumerate(results['photos']):
            with cols[idx % 3]:
                try:
                    img_url = photo['src']['medium']
                    response = requests.get(img_url, timeout=10)
                    img = Image.open(BytesIO(response.content))
                    
                    # Make image clickable
                    clicked = st.image(
                        img,
                        caption=f"{photo['photographer']} - {photo['alt'][:50]}...",
                        use_container_width=True
                    )
                    
                    if clicked:
                        st.session_state.selected_image = photo
                        st.session_state.show_modal = True
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"Error loading image: {str(e)}")
        
        # Pagination controls
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            if st.session_state.page > 1:
                if st.button("⬅️ Previous Page"):
                    st.session_state.page -= 1
                    st.rerun()
        with col3:
            if results['total_results'] > st.session_state.page * 9:
                if st.button("Next Page ➡️"):
                    st.session_state.page += 1
                    st.rerun()
        with col2:
            st.write(f"Page {st.session_state.page} of {(results['total_results']//9)+1}")
    
    elif results:
        st.warning("No results found")

# Modal for image details
if st.session_state.show_modal and st.session_state.selected_image:
    photo = st.session_state.selected_image
    
    # Custom CSS for the modal
    st.markdown("""
    <style>
        .modal {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 80%;
            max-width: 800px;
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            z-index: 1000;
        }
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0,0,0,0.5);
            z-index: 999;
        }
        .close-btn {
            position: absolute;
            top: 10px;
            right: 10px;
            background: #ff4b4b;
            color: white;
            border: none;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            font-size: 16px;
            cursor: pointer;
        }
        .download-btn {
            background: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            margin-top: 10px;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Modal content
    st.markdown(f"""
    <div class="modal-overlay" onclick="document.querySelector('div.modal-overlay').style.display = 'none'"></div>
    <div class="modal">
        <button class="close-btn" onclick="document.querySelector('div.modal-overlay').style.display = 'none'">×</button>
        <h2>{photo['alt']}</h2>
        <div style="display: flex; gap: 20px;">
            <div style="flex: 2;">
                <img src="{photo['src']['large']}" style="width: 100%; border-radius: 8px;"/>
            </div>
            <div style="flex: 1;">
                <h3>Photographer</h3>
                <p><a href="{photo['photographer_url']}" target="_blank">{photo['photographer']}</a></p>
                
                <h3>Image Details</h3>
                <p>Size: {photo['width']} × {photo['height']}</p>
                <p>ID: {photo['id']}</p>
                
                <a href="{photo['src']['original']}" download="pexels-{photo['id']}.jpg" class="download-btn">
                    Download Original
                </a>
                
                <p style="margin-top: 20px;"><small>
                    <a href="https://www.pexels.com/license/" target="_blank">License Information</a>
                </small></p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Handle modal close
    if st.button("Close Modal", key="close_modal"):
        st.session_state.show_modal = False
        st.rerun()

# Add some CSS styling for the main content
st.markdown("""
<style>
    .stImage>img {
        border-radius: 8px;
        box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1);
        transition: transform 0.3s;
        cursor: pointer;
    }
    .stImage>img:hover {
        transform: scale(1.02);
    }
    [data-testid="stExpander"] {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 10px;
        margin-top: 5px;
    }
    .st-emotion-cache-16txtl3 {
        padding: 2rem 1rem;
    }
    .stButton>button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)