import streamlit as st
import streamlit.components.v1 as components
import asyncio
from script import html_to_pptx
import os
import base64
import html

# Page Config
st.set_page_config(page_title="HTML to PPTX Converter", page_icon="📊", layout="wide")

# Custom CSS to mimic the previous design
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        background-color: #c2410c;
        color: white;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 0.75rem;
        border: none;
        transition: all 0.2s;
    }
    .stButton>button:hover {
        background-color: #9a3412;
        color: white;
    }
    .main {
        background-color: #f8fafc;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #1e293b;'>HTML ➔ <span style='color: #c2410c;'>PPTX</span></h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #64748b; margin-bottom: 2rem;'>Convertissez votre code HTML en présentations PowerPoint éditables.</p>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("📝 Saisie du contenu")

    # Input options
    option = st.radio("Méthode d'entrée :", ("Coller du code HTML", "Importer un fichier .html"))

    html_content = ""
    if option == "Coller du code HTML":
        html_content = st.text_area("Code HTML :", placeholder="<section class='slide'>...</section>", height=400)
    else:
        uploaded_file = st.file_uploader("Choisissez un fichier HTML", type="html")
        if uploaded_file is not None:
            html_content = uploaded_file.read().decode("utf-8")

    if st.button("🚀 Générer le PPTX"):
        if not html_content.strip():
            st.error("Veuillez fournir du contenu HTML.")
        else:
            # Ensure Playwright browsers are installed (for Streamlit Cloud)
            import subprocess
            try:
                # Try to install browser AND dependencies at runtime
                # Note: install-deps requires sudo which might fail on some cloud hosts,
                # but 'install chromium' is usually enough if packages.txt is correct.
                subprocess.run(["playwright", "install", "chromium"], check=True)
            except Exception as e:
                st.warning(f"Note: Playwright installation may have issues: {e}")

            with st.spinner("⏳ Conversion en cours... Cela peut prendre quelques secondes."):
                output_file = "presentation_generee.pptx"
                try:
                    # Run the async conversion
                    asyncio.run(html_to_pptx(html_content, output_file))

                    if os.path.exists(output_file):
                        with open(output_file, "rb") as f:
                            btn = st.download_button(
                                label="⬇️ Télécharger le PowerPoint",
                                data=f,
                                file_name="presentation_generee.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        st.success("✅ Conversion terminée !")
                    else:
                        st.error("Erreur lors de la génération du fichier.")
                except Exception as e:
                    st.error(f"Une erreur est survenue : {str(e)}")

with col2:
    st.subheader("👁️ Prévisualisation")
    if html_content:
        # Wrap the user's content in a full HTML page for the component
        # We use JS to scale the 1280x720 slide to fit the component width/height
        preview_template = f"""
        <html>
            <head>
                <script src='https://cdn.tailwindcss.com'></script>
                <style>
                    body {{
                        margin: 0;
                        padding: 0;
                        background-color: #f1f5f9;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        height: 100vh;
                        overflow: hidden;
                    }}
                    #preview-container {{
                        width: 1280px;
                        height: 720px;
                        background: white;
                        position: relative;
                        overflow: hidden;
                        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
                        transform-origin: center center;
                    }}
                </style>
                <script>
                    function scalePreview() {{
                        const container = document.getElementById('preview-container');
                        const padding = 20;
                        const availableWidth = window.innerWidth - padding;
                        const availableHeight = window.innerHeight - padding;
                        const scale = Math.min(availableWidth / 1280, availableHeight / 720);
                        container.style.transform = `scale(${{scale}})`;
                    }}
                    window.addEventListener('resize', scalePreview);
                    window.addEventListener('load', scalePreview);
                </script>
            </head>
            <body>
                <div id="preview-container">
                    {html_content}
                </div>
            </body>
        </html>
        """
        components.html(preview_template, height=600)
    else:
        st.info("Collez du code pour voir l'aperçu ici.")

st.divider()
st.markdown("""
<div style='text-align: center; color: #94a3b8; font-size: 0.8rem;'>
    Propulsé par Playwright & python-pptx
</div>
""", unsafe_allow_html=True)
