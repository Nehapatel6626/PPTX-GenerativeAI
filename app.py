import streamlit as st

# Set page config
st.set_page_config(page_title="Document Generator", layout="wide")
st.image("SVNIT logo.png", width=150)
# Introduction Page
st.title("Welcome to Document Generator")
st.write("Your one-stop tool for creating documents and finding images with ease!")

st.markdown("""
### About Document Generator
The **Document Generator** is a powerful Streamlit app designed to simplify document creation and image sourcing. Whether you're preparing presentations, reports, or visual content, this app has you covered with the following features:

- **Text to PPT**: Transform your text into professional PowerPoint slides with a single click.
- **Text to PDF**: Convert your text into beautifully formatted PDF documents.
- **Text to Word**: Generate Word documents from your text for easy editing and sharing.
- **Pexels Image Search**: Search for high-quality images from Pexels and download them directly for your projects.
- **Standardize_PowerPoints** : Smart PPTX Styler with AI Layout & Clean Formatting.

Get started by exploring these features in the full application!
""")

st.markdown("---")
st.subheader("Get Started")
st.write("This is the introduction to the Document Generator. Contact the developer for access to the full app with document creation and image search capabilities.")