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

# Search interface
search_query = st.text_input("🔍 Search for images", placeholder="Enter search term...")

if search_query:
    with st.spinner("Searching Pexels..."):
        results = search_pexels(search_query)
    
    if results and 'photos' in results:
        st.success(f"Found {results['total_results']} results")
        
        # Display in a grid
        cols = st.columns(3)
        for idx, photo in enumerate(results['photos']):
            with cols[idx % 3]:
                # Display the image
                try:
                    img_url = photo['src']['medium']
                    response = requests.get(img_url, timeout=10)
                    img = Image.open(BytesIO(response.content))
                    st.image(img, caption=photo['alt'], use_column_width=True)
                    
                    # Display attribution
                    with st.expander("ℹ️ Attribution"):
                        st.markdown(f"""
                        **Photographer:** [{photo['photographer']}]({photo['photographer_url']})  
                        **Pexels Profile:** [{photo['photographer_url']}]({photo['photographer_url']})  
                        **License:** [Free to use](https://www.pexels.com/license/)  
                        """)
                except Exception as e:
                    st.error(f"Error loading image: {str(e)}")
    elif results:
        st.warning("No results found")

# Add some CSS styling
st.markdown("""
<style>
    .stImage>img {
        border-radius: 8px;
        box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1);
        transition: transform 0.3s;
    }
    .stImage>img:hover {
        transform: scale(1.02);
    }
    [data-testid="stExpander"] {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)