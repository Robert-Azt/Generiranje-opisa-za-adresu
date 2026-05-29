import streamlit as st
from docx import Document
from docx.shared import Pt
from datetime import datetime
import requests
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# ====================== KONFIG ======================
st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije za Sigurnosnu Analizu")

# AI Provider
ai_provider = st.sidebar.selectbox("AI Model", ["Claude 3.5 Sonnet", "Claude 3 Opus", "OpenAI GPT-4o"])
api_key = st.sidebar.text_input("Anthropic / OpenAI API Key", type="password", value="")

geolocator = Nominatim(user_agent="lokacija_generator_hr")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

def get_ai_description(location_data, prompt):
    if not api_key:
        st.error("Molimo unesite API ključ.")
        return None
    
    try:
        if "Claude" in ai_provider:
            # Anthropic Claude
            headers = {
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            }
            
            model = "claude-3-5-sonnet-20240620" if "Sonnet" in ai_provider else "claude-3-opus-20240229"
            
            payload = {
                "model": model,
                "max_tokens": 4000,
                "temperature": 0.7,
                "messages": [
                    {"role": "user", "content": prompt}
                ]
            }
            
            response = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers=headers,
                json=payload
            )
            
            if response.status_code == 200:
                return response.json()["content"][0]["text"]
            else:
                st.error(f"Claude greška: {response.text}")
                return None
                
        else:
            # OpenAI fallback
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-4o",
                temperature=0.7,
                messages=[{"role": "user", "content": prompt}]
            )
            return response.choices[0].message.content
            
    except Exception as e:
        st.error(f"Greška pri pozivu AI: {e}")
        return None

# ====================== GLAVNI DIO ======================
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Unos lokacije")
    input_type = st.radio("Način unosa", ["Adresa", "Koordinate"])
    
    if input_type == "Adresa":
        address = st.text_input("Unesite adresu", "Kranjčevićeva ulica, Zagreb")
        if st.button("🔍 Pronađi i generiraj opis"):
            location = geocode(address)
    else:
        lat = st.number_input("Latitude", value=45.7925, format="%.6f")
        lon = st.number_input("Longitude", value=15.9622, format="%.6f")
        if st.button("🔍 Pronađi i generiraj opis"):
            location = geolocator.reverse((lat, lon), exactly_one=True)

    if 'location' in locals() and location:
        st.success(f"✅ Pronađeno: **{location.address}**")
        st.info(f"Koordinate: {location.latitude:.6f}, {location.longitude:.6f}")

with col2:
    if 'location' in locals() and location:
        st.subheader("Generirani opis (Claude)")
        
        location_data = {
            "adresa": location.address,
            "grad": location.raw.get('address', {}).get('city') or location.raw.get('address', {}).get('town', 'Zagreb'),
            "lat": location.latitude,
            "lon": location.longitude,
            "quart": location.raw.get('address', {}).get('suburb', '')
        }
        
        prompt = f"""
Ti si stručnjak za izradu sigurnosnih analiza i elaborata za sportske objekte u Hrvatskoj.
Napiši **vrlo formalan, stručan i detaljan opis lokacije** na hrvatskom jeziku, u stilu službenog dokumenta (kao primjer koji ti je dan ranije).

Lokacija: {location_data['adresa']}
Grad: {location_data['grad']}
Četvrt: {location_data['quart']}
Koordinate: {location_data['lat']}, {location_data['lon']}

Napiši sljedeće sekcije:
1. Opis lokacije
2. Opis okolnih građevina, površina i okoliša
3. Načini pristupa
4. Frekvencija prometa (radni dani, vikend, noć)
5. Stanje kriminaliteta u okolnom prostoru

Koristi formalni, birokratski jezik, slično primjeru iz dokumenta o stadionu u Kranjčevićevoj ulici.
"""

        if st.button("🚀 Generiraj opis pomoću Claudea"):
            with st.spinner("Claude radi... (može potrajati 10-20 sekundi)"):
                opis = get_ai_description(location_data, prompt)
                
                if opis:
                    st.text_area("Generirani opis", opis, height=500)
                    
                    # Word dokument
                    if st.button("💾 Spremi kao Word dokument"):
                        doc = Document()
                        doc.add_heading('Snimka postojećeg stanja', 0)
                        doc.add_heading('Opis lokacije', level=1)
                        doc.add_paragraph(opis)
                        
                        filename = f"Opis_lokacije_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                        doc.save(filename)
                        
                        with open(filename, "rb") as f:
                            st.download_button(
                                label="Preuzmi .docx datoteku",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )