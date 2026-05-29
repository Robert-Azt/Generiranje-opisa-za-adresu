import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
from docx.shared import Pt
import io
import json

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

geolocator = Nominatim(user_agent="lokacija_generator_hr")
address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

# Sve tablice i njihovi redci (label, json_key)
TABLES = [
    {
        "number": "2.1",
        "title": "Opis lokacije",
        "section": "Opis lokacije",
        "rows": [
            ("Opis lokacije",                                              "opis_lokacije"),
            ("Opis okolnih građevina, površina i okoliša",                 "opis_okolnih_gradjevina"),
            ("Načini pristupa",                                            "nacini_pristupa"),
            ("Frekvencija prometa radnim danom, vikendom, noću",           "frekvencija_prometa"),
            ("Stanje kriminaliteta u okolnom prostoru",                    "stanje_kriminaliteta"),
        ]
    },
    {
        "number": "2.2",
        "title": "Osnovne karakteristike",
        "section": "Osnovne karakteristike",
        "rows": [
            ("Prostorna organiziranost", "prostorna_organiziranost"),
            ("Veličina i namjena",       "velicina_i_namjena"),
        ]
    },
    {
        "number": "2.3",
        "title": "Građevinske karakteristike",
        "section": "Građevinske karakteristike",
        "rows": [
            ("Vrsta materijala",                                                         "vrsta_materijala"),
            ("Nagib terena",                                                             "nagib_terena"),
            ("Postojeći elementi javne površine (stepenice, podvožnjaci, objekti) i dr.", "elementi_javne_povrsine"),
        ]
    },
    {
        "number": "2.4",
        "title": "Instalacije",
        "section": "Instalacije",
        "rows": [
            ("Električne instalacije",                       "elektricne_instalacije"),
            ("Ostale instalacije (plin, voda, kanalizacija)", "ostale_instalacije"),
        ]
    },
    {
        "number": "2.5",
        "title": "Namjena",
        "section": "Namjena",
        "rows": [
            ("Opća namjena",                  "opca_namjena"),
            ("Namjena pojedinih prostora",    "namjena_pojedinih_prostora"),
            ("Radno vrijeme",                 "radno_vrijeme"),
            ("Put kretanja osoba i vozila",   "put_kretanja"),
            ("Način zaključavanja prostora",  "nacin_zakljucavanja"),
        ]
    },
    {
        "number": "2.6",
        "title": "Radni procesi",
        "section": "Radni procesi",
        "rows": [
            ("Popis i opis procesa i postupaka bitnih za sigurnost", "radni_procesi"),
        ]
    },
    {
        "number": "2.7",
        "title": "Vrste i visina vrijednosti",
        "section": "Vrsta i visina vrijednosti",
        "rows": [
            ("Vrste vrijednosti",       "vrste_vrijednosti"),
            ("Visina vrijednosti",      "visina_vrijednosti"),
            ("Način čuvanja vrijednosti", "nacin_cuvanja_vrijednosti"),
        ]
    },
    {
        "number": "2.8",
        "title": "Organizacija sigurnosti",
        "section": "Organizacija sigurnosti",
        "rows": [
            ("Tjelesna zaštita",      "tjelesna_zastita"),
            ("Tehnička zaštita",      "tehnicka_zastita"),
            ("Organizacijske mjere",  "organizacijske_mjere"),
        ]
    },
    {
        "number": "2.9",
        "title": "Stanje dokumentiranosti",
        "section": "Stanje dokumentiranosti",
        "rows": [
            ("Postojeći sustavi zaštite",                              "postojeci_sustavi_zastite"),
            ("Dokumentiranost postojećih sustava tehničke zaštite",    "dokumentiranost_sustava"),
        ]
    },
    {
        "number": "2.10",
        "title": "Uočeni nedostaci",
        "section": "Uočeni nedostaci",
        "rows": [
            ("Uočeni nedostaci", "uoceni_nedostaci"),
        ]
    },
    {
        "number": "2.11",
        "title": "Kritične točke i ugroženi prostori",
        "section": "Kritične točke i ugroženi prostori",
        "rows": [
            ("Kritične točke",    "kriticne_tocke"),
            ("Ugroženi prostori", "ugrozeni_prostori"),
        ]
    },
]

# Svi JSON ključevi koje Claude treba generirati
ALL_KEYS = [key for table in TABLES for (_, key) in table["rows"]]


def add_section_table(doc, table):
    """Dodaje naslov sekcije i tablicu identičnu formatu originalnog dokumenta."""
    # Heading naslov sekcije
    heading = doc.add_heading(table["section"], level=3)
    for run in heading.runs:
        run.font.size = Pt(12)

    n_rows = 1 + len(table["rows"])  # zaglavlje + redci
    tbl = doc.add_table(rows=n_rows, cols=1)
    tbl.style = "Table Grid"

    # Zaglavlje tablice (bold)
    header_cell = tbl.cell(0, 0)
    header_para = header_cell.paragraphs[0]
    header_para.clear()
    run = header_para.add_run(f"Tablica {table['number']}. {table['title']}")
    run.bold = True
    run.font.size = Pt(10)

    # Redci s bold labelom i normalnim tekstom
    for i, (label, key) in enumerate(table["rows"]):
        cell = tbl.cell(i + 1, 0)
        para = cell.paragraphs[0]
        para.clear()
        r_label = para.add_run(f"{label}: ")
        r_label.bold = True
        r_label.font.size = Pt(10)
        r_text = para.add_run(table["data"].get(key, ""))
        r_text.bold = False
        r_text.font.size = Pt(10)

    doc.add_paragraph()  # razmak


def generate_text(api_key, location_address, lat, lon):
    """Poziva Claude API i vraća JSON sa svim sekcijama."""
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    keys_list = "\n".join(f'  "{k}": "..."' for k in ALL_KEYS)

    prompt = f"""Ti si stručnjak za izradu sigurnosnih elaborata i procjena ugroženosti u Hrvatskoj.
Piši formalno, birokratski i detaljno, kao u službenom elaboratu o procjeni sigurnosnih rizika nekretnine.

Lokacija: {location_address}
Koordinate: {lat:.6f}, {lon:.6f}

Za svaki od dolje navedenih ključeva napiši 4-8 povezanih, smislenih rečenica prilagođenih toj specifičnoj lokaciji.
Tekst mora biti u obliku tekućeg odlomka, bez markdowna, bez bullet pointova, bez boldanja, bez tablica.
Piši isključivo čist tekst.

Odgovori SAMO s JSON objektom, bez ikakvih komentara, bez markdown backtick oznaka:
{{
{keys_list}
}}"""

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers=headers,
        json={
            "model": "claude-sonnet-4-6",
            "max_tokens": 8000,
            "temperature": 0.6,
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=120
    )

    if response.status_code != 200:
        raise Exception(f"API greška: {response.status_code} - {response.text}")

    raw = response.json()["content"][0]["text"].strip()

    # Ukloni eventualne markdown backtick ograde
    if "```" in raw:
        parts = raw.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("{"):
                raw = part
                break

    return json.loads(raw)


if st.button("🚀 Generiraj elaborat", type="primary"):
    if not api_key:
        st.error("Unesi API ključ!")
        st.stop()

    with st.spinner("Geocodiranje adrese..."):
        location = geolocator.geocode(address)
        if not location:
            st.error("Adresa nije pronađena. Pokušaj s drugom adresom.")
            st.stop()

    st.success(f"📍 {location.address}")

    with st.spinner("Claude generira tekst za svih 11 tablica... (može potrajati ~30s)"):
        try:
            data = generate_text(api_key, location.address, location.latitude, location.longitude)
        except json.JSONDecodeError as e:
            st.error(f"Greška pri parsiranju JSON odgovora: {e}")
            st.stop()
        except Exception as e:
            st.error(f"Greška: {e}")
            st.stop()

    # Prikaži u Streamlitu po tablicama
    st.subheader("📄 Generirani sadržaj")
    for table in TABLES:
        with st.expander(f"Tablica {table['number']} — {table['title']}", expanded=False):
            for label, key in table["rows"]:
                st.markdown(f"**{label}:**")
                st.write(data.get(key, ""))

    # Izrada Word dokumenta
    doc = Document()
    title = doc.add_heading("Snimka postojećeg stanja", level=1)
    for run in title.runs:
        run.font.size = Pt(14)

    for table in TABLES:
        table["data"] = data
        add_section_table(doc, table)

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
    st.success("✅ Dokument spreman!")
