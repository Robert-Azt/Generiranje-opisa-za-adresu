import streamlit as st
import requests
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from docx import Document
from datetime import datetime

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

# SIDEBAR
with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")
    st.info("Zalijepi svoj Claude ključ ovdje")

# GEOLOCATION
geolocator = Nominatim(user_agent="lokacija_generator_hr")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Unos lokacije")
    address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")
    
    if st.button("🔍 Pronađi lokaciju i generiraj opis", type="primary", use_container_width=True):
        with st.spinner("Tražim lokaciju..."):
            try:
                location = geolocator.geocode(address)
                if location:
                    st.session_state.location = location
                    st.success(f"✅ Pronađeno: {location.address}")
                    st.caption(f"Koordinate: {location.latitude:.6f}, {location.longitude:.6f}")
                else:
                    st.error("Lokacija nije pronađena.")
            except Exception as e:
                st.error(f"Greška pri traženju lokacije: {e}")

with col2:
    st.subheader("Generirani opis")
    
    if 'location' in st.session_state and api_key:
        if st.button("🚀 Generiraj opis pomoću Claudea", type="primary", use_container_width=True):
            with st.spinner("Claude radi... (15-40 sekundi)"):
                try:
                    headers = {
                        "x-api-key": api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json"
                    }
                    
                    prompt = f"""Napiši formalan opis lokacije na hrvatskom jeziku za sigurnosnu analizu.

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude}, {st.session_state.location.longitude}

Sekcije:
1. Opis lokacije
2. Opis okolnih građevina, površina i okoliša
3. Načini pristupa
4. Frekvencija prometa
5. Stanje kriminaliteta"""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-3-5-sonnet-20240620",
                            "max_tokens": 3000,
                            "temperature": 0.7,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=60
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.text_area("Rezultat:", opis, height=600)
                        
                        # Word
                        doc = Document()
                        doc.add_heading('Opis lokacije', 0)
                        doc.add_paragraph(opis)
                        filename = f"opis_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        doc.save(filename)
                        
                        with open(filename, "rb") as f:
                            st.download_button("Preuzmi Word", f, filename)
                    else:
                        st.error(f"API greška: {response.status_code}")
                        st.write(response.text)
                        
                except Exception as e:
                    st.error(f"Greška: {str(e)}")
    else:
        st.info("Unesite API ključ i pronađite lokaciju.")

st.caption("Debug verzija • Ako i dalje ne radi, javi mi što točno vidiš")
