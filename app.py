import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
import io

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

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

with col2:
    st.subheader("Generirani opis")
    
    if 'location' in st.session_state and api_key:
        if st.button("🚀 Generiraj dokument sa tablicama", type="primary", use_container_width=True):
            with st.spinner("Claude radi... (25-45 sekundi)"):
                try:
                    headers = {
                        "x-api-key": api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json"
                    }
                    
                    prompt = f"""Napiši **cijeli sigurnosni elaborat** točno u stilu i formatu dokumenta koji ti je dan kao primjer (o stadionu u Kranjčevićevoj ulici).

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude}, {st.session_state.location.longitude}

Koristi istu strukturu:
- Naslove "Tablica 1. Opis lokacije", "Tablica 2. Osnovne karakteristike" itd.
- Ispod svakog naslova tablice piši sadržaj u obliku:
  **Opis lokacije:** tekst...
  **Opis okolnih građevina...:** tekst...
  **Načini pristupa:** tekst...
  itd.

Piši vrlo formalno, birokratski, sa sličnim rečenicama kao u primjeru. Koristi • za nabrajanja gdje je prikladno."""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-sonnet-4-6",
                            "max_tokens": 5000,
                            "temperature": 0.6,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=100
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.text_area("Generirani tekst:", opis, height=500)

                        # ==================== WORD DOKUMENT ====================
                        doc = Document()
                        doc.add_heading('Snimka postojećeg stanja', 0)

                        # Dodaj generirani tekst
                        doc.add_paragraph(opis)

                        filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        buffer = io.BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)

                        st.download_button(
                            label="💾 Preuzmi Word dokument sa tablicama",
                            data=buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"Claude greška: {response.status_code}")
                except Exception as e:
                    st.error(f"Greška: {e}")
    else:
        st.info("Unesi API ključ i pronađi lokaciju.")
