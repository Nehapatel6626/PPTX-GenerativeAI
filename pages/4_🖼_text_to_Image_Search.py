import streamlit as st
import requests
from PIL import Image
from io import BytesIO
import os

# Set up the page with default Streamlit styling
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

# Search interface
search_query = st.text_input("🔍 Search for images", placeholder="Enter search term...")

# Initialize session state
if 'page' not in st.session_state:
    st.session_state.page = 1
if 'selected_image' not in st.session_state:
    st.session_state.selected_image = None

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
                    
                    # Display image with clickable option
                    clicked = st.image(
                        img,
                        caption=f"{photo['photographer']} - {photo['alt'][:50]}...",
                        use_container_width=True
                    )
                    
                    # Options expander
                    with st.expander("⚙️ Options"):
                        st.download_button(
                            label="Download Original",
                            data=requests.get(photo['src']['original']).content,
                            file_name=f"pexels-{photo['id']}.jpg",
                            mime="image/jpeg",
                            key=f"dl_{photo['id']}"
                        )
                        if st.button("View Details", key=f"view_{photo['id']}"):
                            st.session_state.selected_image = photo
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

# Image details sidebar (using default Streamlit styling)
if st.session_state.selected_image:
    photo = st.session_state.selected_image
    with st.sidebar:
        st.title("Image Details")
        
        try:
            img_url = photo['src']['large']
            response = requests.get(img_url, timeout=10)
            img = Image.open(BytesIO(response.content))
            st.image(img, use_container_width=True)
        except Exception as e:
            st.error(f"Error loading image: {str(e)}")
        
        # Attribution details
        st.subheader("Attribution")
        st.markdown(f"""
        **Photographer:** [{photo['photographer']}]({photo['photographer_url']})  
        **Profile:** [View on Pexels]({photo['photographer_url']})  
        **Photo:** [View Original]({photo['url']})  
        **Size:** {photo['width']} × {photo['height']}  
        **License:** [Free to use](https://www.pexels.com/license/)  
        """)
        
        # Download button
        st.download_button(
            label="Download Original Quality",
            data=requests.get(photo['src']['original']).content,
            file_name=f"pexels-{photo['id']}-original.jpg",
            mime="image/jpeg",
            use_container_width=True
        )
        
        if st.button("Close Details", use_container_width=True):
            st.session_state.selected_image = None
            st.rerun()

# Minimal styling that doesn't override defaults
st.markdown("""
<style>
    .stImage>img {
        border-radius: 8px;
        transition: transform 0.3s;
    }
    .stImage>img:hover {
        transform: scale(1.02);
    }
</style>
""", unsafe_allow_html=True)