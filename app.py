import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

# SIDEBAR
with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")
    st.info("Zalijepi Claude ključ ovdje")

geolocator = Nominatim(user_agent="lokacija_generator_hr")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Unos lokacije")
    address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")
    
    if st.button("🔍 Pronađi lokaciju", type="primary", use_container_width=True):
        with st.spinner("Tražim lokaciju..."):
            location = geolocator.geocode(address)
            if location:
                st.session_state.location = location
                st.success(f"✅ Pronađeno: {location.address}")
                st.caption(f"Koordinate: {location.latitude:.6f}, {location.longitude:.6f}")

with col2:
    st.subheader("Generirani opis (Claude)")
    
    if 'location' in st.session_state and api_key:
        if st.button("🚀 Generiraj opis i Word dokument", type="primary", use_container_width=True):
            with st.spinner("Claude radi..."):
                try:
                    headers = {
                        "x-api-key": api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json"
                    }
                    
                    prompt = f"""Napiši detaljan formalan opis lokacije na hrvatskom jeziku u stilu sigurnosnog elaborata.

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude}, {st.session_state.location.longitude}

Koristi strukturu i stil kao u primjeru za stadion u Kranjčevićevoj ulici."""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-sonnet-4-6",
                            "max_tokens": 4000,
                            "temperature": 0.7,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=80
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.success("✅ Opis generiran!")
                        st.text_area("Pregled teksta:", opis, height=400)

                        # ==================== WORD DOKUMENT ====================
                        doc = Document()
                        doc.add_heading('Snimka postojećeg stanja', 0)
                        doc.add_heading('Opis lokacije', level=1)
                        doc.add_paragraph(opis)

                        # Formatiranje
                        for paragraph in doc.paragraphs:
                            if paragraph.text.strip():
                                paragraph.style = 'Normal'
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                        filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        # Spremanje preko BytesIO (riješava encoding probleme)
                        buffer = io.BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)

                        st.download_button(
                            label="💾 Preuzmi Word dokument (.docx)",
                            data=buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"Claude greška: {response.status_code}")
                        
                except Exception as e:
                    st.error(f"Greška: {str(e)}")
    else:
        st.info("Unesi API ključ i pronađi lokaciju.")
