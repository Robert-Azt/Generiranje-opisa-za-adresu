import streamlit as st
import requests
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from docx import Document
from datetime import datetime
import io

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

# SIDEBAR
with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")
    st.info("Zalijepi Claude ključ ovdje")

# GEOLOCATION
geolocator = Nominatim(user_agent="lokacija_generator_hr")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Unos lokacije")
    address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")
    
    if st.button("🔍 Pronađi lokaciju", type="primary", use_container_width=True):
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
                st.error(f"Greška: {e}")

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
                    
                    prompt = f"""Napiši detaljan, formalan opis lokacije na hrvatskom jeziku za sigurnosnu analizu.

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude}, {st.session_state.location.longitude}

Napiši sljedeće sekcije:
1. Opis lokacije
2. Opis okolnih građevina, površina i okoliša
3. Načini pristupa
4. Frekvencija prometa (radni dani, vikend, noć)
5. Stanje kriminaliteta u okolnom prostoru

Koristi formalni stil kao u primjeru za stadion."""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-3-5-sonnet-20240620",
                            "max_tokens": 3500,
                            "temperature": 0.7,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=70
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.success("✅ Opis generiran!")
                        st.text_area("Generirani opis:", opis, height=550)
                        
                        # === POPRAVLJENO SPREMANJE WORD DOKUMENTA ===
                        doc = Document()
                        doc.add_heading('Snimka postojećeg stanja', 0)
                        doc.add_heading('Opis lokacije', level=1)
                        doc.add_paragraph(opis)
                        
                        filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        
                        # Spremanje u memory buffer da izbjegnemo encoding probleme
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
                        st.error(f"API greška: {response.status_code}")
                        st.write(response.text[:500])
                        
                except Exception as e:
                    st.error(f"Greška: {str(e)}")
    else:
        if not api_key:
            st.warning("⚠️ Unesi Anthropic API ključ u sidebar.")
        else:
            st.info("Prvo pronađi lokaciju.")

st.caption("Popravljena encoding verzija")
