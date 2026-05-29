import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
import io

st.set_page_config(page_title="Generator Sigurnosnog Elaborata", layout="wide")
st.title("🗺️ Generator Sigurnosnog Elaborata (Tablični format)")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

geolocator = Nominatim(user_agent="lokacija_generator_hr")

address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

if st.button("🚀 Generiraj elaborat u tabličnom formatu", type="primary"):
    if not api_key:
        st.error("Unesi Anthropic API ključ!")
    else:
        with st.spinner("Claude generira elaborat u željenom formatu... (30-60 sekundi)"):
            try:
                location = geolocator.geocode(address)
                if not location:
                    st.error("Lokacija nije pronađena.")
                    st.stop()

                headers = {
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                }
                
                prompt = f"""Napiši sigurnosni elaborat **točno u stilu i formatu** kao na slikama koje si poslao i u originalnom Word dokumentu o stadionu.

Lokacija: {location.address}
Koordinate: {location.latitude:.6f}, {location.longitude:.6f}

Koristi isključivo format sa tablicama od dva stupca:
- Lijevi stupac: naziv polja (boldano)
- Desni stupac: opis

Počni sa:
# SIGURNOSNI ELABORAT
## Snimka postojećeg stanja

Zatim:
**Tablica 2.1. Opis lokacije**  
| Opis lokacije: | [tekst] |
| Opis okolnih građevina, površina i okoliša: | [tekst sa • nabrajanjima] |
| Načini pristupa: | [tekst] |
| Frekvencija prometa radnim danom, vikendom, noću: | [tekst] |
| Stanje kriminaliteta u okolnom prostoru: | [tekst] |

Nastavi sa ostalim tablicama (2.2, 2.3, itd.) koliko je moguće. Piši formalno, sažeto i birokratski na hrvatskom jeziku."""

                response = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers,
                    json={
                        "model": "claude-sonnet-4-6",
                        "max_tokens": 6000,
                        "temperature": 0.5,
                        "messages": [{"role": "user", "content": prompt}]
                    },
                    timeout=120
                )

                if response.status_code == 200:
                    tekst = response.json()["content"][0]["text"]
                    st.text_area("Generirani elaborat:", tekst, height=800)

                    doc = Document()
                    doc.add_heading('SIGURNOSNI ELABORAT', 0)
                    doc.add_paragraph(tekst)

                    filename = f"Sigurnosni_Elaborat_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)

                    st.download_button("💾 Preuzmi Word dokument", buffer, filename)
                else:
                    st.error(f"Claude greška: {response.status_code}")
            except Exception as e:
                st.error(f"Greška: {e}")
