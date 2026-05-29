import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
import io

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije - Normalan Tekst")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

geolocator = Nominatim(user_agent="lokacija_generator_hr")

address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

if st.button("🚀 Generiraj normalan, povezan tekst", type="primary"):
    if not api_key:
        st.error("Unesi API ključ!")
    else:
        with st.spinner("Claude radi..."):
            try:
                location = geolocator.geocode(address)
                
                headers = {
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                }
                
                prompt = f"""Napiši **normalan, tekući, povezan tekst** na hrvatskom jeziku, baš kao u originalnom Word dokumentu koji sam ti poslao na početku (o stadionu).

Lokacija: {location.address}
Koordinate: {location.latitude:.6f}, {location.longitude:.6f}

Piši **isključivo normalne odlomke**, bez tablica, bez **boldanja**, bez | znakova i bez markdowna. 
Tekst mora biti smislen, formalan i birokratski, u stilu službenog elaborata. 
Svaka sekcija (Opis lokacije, Okolne građevine, Načini pristupa, Frekvencija prometa, Stanje kriminaliteta) neka bude u obliku 4-8 povezanih rečenica."""

                response = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers,
                    json={
                        "model": "claude-sonnet-4-6",
                        "max_tokens": 4000,
                        "temperature": 0.6,
                        "messages": [{"role": "user", "content": prompt}]
                    },
                    timeout=90
                )

                if response.status_code == 200:
                    tekst = response.json()["content"][0]["text"]
                    st.text_area("Generirani tekst:", tekst, height=700)

                    # Word
                    doc = Document()
                    doc.add_heading('Snimka postojećeg stanja', 0)
                    doc.add_paragraph(tekst)
                    
                    filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)

                    st.download_button("💾 Preuzmi Word dokument", buffer, filename)
                else:
                    st.error(f"Greška: {response.status_code}")
            except Exception as e:
                st.error(f"Greška: {e}")
