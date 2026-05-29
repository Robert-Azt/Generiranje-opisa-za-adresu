import streamlit as st
import requests
from geopy.geocoders import Nominatim
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
    
    if st.button("🔍 Pronađi lokaciju", type="primary", use_container_width=True):
        with st.spinner("Tražim lokaciju..."):
            location = geolocator.geocode(address)
            if location:
                st.session_state.location = location
                st.success(f"✅ Pronađeno: {location.address}")
                st.caption(f"Koordinate: {location.latitude:.6f}, {location.longitude:.6f}")
            else:
                st.error("Lokacija nije pronađena.")

with col2:
    st.subheader("Generirani opis (Claude)")
    
    if 'location' in st.session_state and api_key:
        if st.button("🚀 Generiraj opis pomoću Claudea", type="primary", use_container_width=True):
            with st.spinner("Claude radi... (15-40 sekundi)"):
                try:
                    headers = {
                        "x-api-key": api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json"
                    }
                    
                    prompt = f"""Napiši detaljan formalan opis lokacije na hrvatskom jeziku u stilu sigurnosne analize.

Lokacija: {st.session_state.location.address}
Koordinate: {st.session_state.location.latitude}, {st.session_state.location.longitude}

Napiši ove sekcije:
1. Opis lokacije
2. Opis okolnih građevina, površina i okoliša
3. Načini pristupa
4. Frekvencija prometa radnim danom, vikendom i noću
5. Stanje kriminaliteta u okolnom prostoru

Koristi formalni, stručni stil."""

                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers=headers,
                        json={
                            "model": "claude-3-5-sonnet-20240620",
                            "max_tokens": 4000,
                            "temperature": 0.7,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=80
                    )

                    if response.status_code == 200:
                        opis = response.json()["content"][0]["text"]
                        st.success("✅ Opis uspješno generiran!")
                        st.text_area("Generirani tekst:", opis, height=600)
                        
                        # Download kao TXT (privremeno, da izbjegnemo encoding problem)
                        st.download_button(
                            label="💾 Preuzmi kao .txt datoteku",
                            data=opis,
                            file_name=f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                            mime="text/plain"
                        )
                    else:
                        st.error(f"Claude greška: {response.status_code}")
                        st.write(response.text[:300])
                        
                except Exception as e:
                    st.error(f"Greška: {str(e)}")
    else:
        st.info("Unesi API ključ i pronađi lokaciju.")

st.caption("Verzija bez Worda (zbog encoding problema)")
