import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

geolocator = Nominatim(user_agent="lokacija_generator_hr")
address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

SECTIONS = [
    ("Opis lokacije", "opis_lokacije"),
    ("Opis okolnih građevina, površina i okoliša", "opis_okolnih_gradjevina"),
    ("Načini pristupa", "nacini_pristupa"),
    ("Frekvencija prometa radnim danom, vikendom, noću", "frekvencija_prometa"),
    ("Stanje kriminaliteta u okolnom prostoru", "stanje_kriminaliteta"),
]

def set_cell_bold_label(cell, label, text):
    """Dodaje bold labelu i normalan tekst u ćeliju tablice."""
    paragraph = cell.paragraphs[0]
    paragraph.clear()
    # Bold label
    run_label = paragraph.add_run(f"{label}: ")
    run_label.bold = True
    run_label.font.size = Pt(10)
    # Normalan tekst
    run_text = paragraph.add_run(text)
    run_text.bold = False
    run_text.font.size = Pt(10)

def set_table_header(table, title):
    """Postavlja bold naslov u prvu ćeliju tablice."""
    cell = table.cell(0, 0)
    paragraph = cell.paragraphs[0]
    paragraph.clear()
    run = paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(10)

def add_section_table(doc, table_number, table_title, section_title, rows_data):
    """
    Dodaje naslov sekcije i tablicu u Word dokument,
    identično formatu originalnog dokumenta.
    rows_data: lista tuplea (label, tekst)
    """
    # Naslov sekcije (Heading 3)
    heading = doc.add_heading(section_title, level=3)
    heading.runs[0].font.size = Pt(12)

    # Tablica s jednim stupcem, n+1 redaka (zaglavlje + n redaka podataka)
    table = doc.add_table(rows=1 + len(rows_data), cols=1)
    table.style = "Table Grid"

    # Zaglavlje tablice
    header_cell = table.cell(0, 0)
    header_para = header_cell.paragraphs[0]
    header_para.clear()
    run = header_para.add_run(f"Tablica {table_number}. {table_title}")
    run.bold = True
    run.font.size = Pt(10)

    # Redci podataka
    for i, (label, tekst) in enumerate(rows_data):
        cell = table.cell(i + 1, 0)
        set_cell_bold_label(cell, label, tekst)

    doc.add_paragraph()  # razmak između tablica

def generate_text(api_key, location_address, lat, lon):
    """Poziva Claude API i vraća strukturirani JSON s tekstovima po sekcijama."""
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    prompt = f"""Ti si stručnjak za izradu sigurnosnih elaborata i procjena ugroženosti u Hrvatskoj. 
Piši formalno, birokratski i detaljno, kao u službenom elaboratu.

Lokacija: {location_address}
Koordinate: {lat:.6f}, {lon:.6f}

Generiraj sadržaj za tablicu 2.1 "Opis lokacije" koja sadrži sljedeće retke.
Za svaki redak napiši 4-8 povezanih, smislenih rečenica u obliku tekućeg odlomka.
NE koristi markdown, boldanje, tablice, bullet points ili bilo kakvo formatiranje.
Piši ISKLJUČIVO čist tekst u odlomcima.

Odgovori SAMO s JSON objektom (bez ikakvih dodatnih komentara ili markdown backtick oznaka):
{{
  "opis_lokacije": "...",
  "opis_okolnih_gradjevina": "...",
  "nacini_pristupa": "...",
  "frekvencija_prometa": "...",
  "stanje_kriminaliteta": "..."
}}

Svaka vrijednost mora biti čist tekst bez posebnih znakova formatiranja."""

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

    if response.status_code != 200:
        raise Exception(f"API greška: {response.status_code} - {response.text}")

    raw = response.json()["content"][0]["text"].strip()

    # Ukloni eventualne markdown backtick ograde
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    data = json.loads(raw)
    return data


if st.button("🚀 Generiraj elaborat", type="primary"):
    if not api_key:
        st.error("Unesi API ključ!")
    else:
        with st.spinner("Geocodiranje adrese..."):
            location = geolocator.geocode(address)
            if not location:
                st.error("Adresa nije pronađena. Pokušaj s drugom adresom.")
                st.stop()

        st.success(f"📍 Pronađena lokacija: {location.address}")

        with st.spinner("Claude generira tekst..."):
            try:
                data = generate_text(api_key, location.address, location.latitude, location.longitude)
            except json.JSONDecodeError as e:
                st.error(f"Greška pri parsiranju JSON odgovora: {e}")
                st.stop()
            except Exception as e:
                st.error(f"Greška: {e}")
                st.stop()

        # Prikaz u Streamlitu
        st.subheader("📄 Generirani sadržaj")
        for label, key in SECTIONS:
            with st.expander(f"**{label}**", expanded=True):
                st.write(data.get(key, ""))

        # Izrada Word dokumenta
        doc = Document()

        # Naslov dokumenta
        title = doc.add_heading("Snimka postojećeg stanja", level=1)
        title.runs[0].font.size = Pt(14)

        # Tablica 2.1 - Opis lokacije
        add_section_table(
            doc,
            table_number="2.1",
            table_title="Opis lokacije",
            section_title="Opis lokacije",
            rows_data=[
                ("Opis lokacije", data.get("opis_lokacije", "")),
                ("Opis okolnih građevina, površina i okoliša", data.get("opis_okolnih_gradjevina", "")),
                ("Načini pristupa", data.get("nacini_pristupa", "")),
                ("Frekvencija prometa radnim danom, vikendom, noću", data.get("frekvencija_prometa", "")),
                ("Stanje kriminaliteta u okolnom prostoru", data.get("stanje_kriminaliteta", "")),
            ]
        )

        # Spremi Word dokument
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
        st.success("✅ Dokument je spreman za preuzimanje!")
