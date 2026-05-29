import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
import io

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")
    st.caption("Ako dobiješ timeout, pokušaj ponovno ili smanji zahtjev.")

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
        if st.button("🚀 Generiraj dokument (sa tablicama)", type="primary", use_container_width=True):
            with st.spinner("Claude radi... (može potrajati do 50 sekundi)"):
                try:
                    headers = {
                        "x-api-key": api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json"
                    }
                    
                    prompt = f"""Napiši sigurnosni elaborat za lokaciju u stilu primjera koji si dobio (stadion u Kranjčevićevoj).

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude:.6f}, {st.session_state.location.longitude:.6f}

Koristi strukturu sa Tablicama:
- Tablica 1. Opis lokacije
- Tablica 2. Osnovne karakteristike
- Tablica 3. Građevinske karakteristike (prilagodi)
- Tablica 4. Instalacije
- Tablica 5. Namjena
- Tablica 6. Radni procesi
- Tablica 7. Vrste i visina vrijednosti
- Tablica 11. Kritične točke itd.

Piši formalno na hrvatskom, koristi slične fraze kao u primjeru."""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-sonnet-4-6",
                            "max_tokens": 3500,      # smanjeno
                            "temperature": 0.65,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=120                 # povećano
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.text_area("Generirani tekst:", opis, height=500)

                        doc = Document()
                        doc.add_heading('Snimka postojećeg stanja', 0)
                        doc.add_paragraph(opis)

                        filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        buffer = io.BytesIO()
                        doc.save(buffer)
                        buffer.seek(0)

                        st.download_button(
                            label="💾 Preuzmi Word dokument",
                            data=buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"Claude greška: {response.status_code} - {response.text[:200]}")
                except requests.exceptions.Timeout:
                    st.error("⏱️ Timeout - Claude je predugo odgovarao. Pokušaj ponovno.")
                except Exception as e:
                    st.error(f"Greška: {e}")
    else:
        st.info("Unesi API ključ i pronađi lokaciju.")
