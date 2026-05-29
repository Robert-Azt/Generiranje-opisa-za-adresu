import streamlit as st
import requests
from geopy.geocoders import Nominatim
from datetime import datetime
from docx import Document
from docx.shared import Pt
import io
import json
import concurrent.futures

st.set_page_config(page_title="Generator Opisa Lokacije", layout="wide")
st.title("🗺️ Generator Opisa Lokacije")

with st.sidebar:
    st.header("🔑 Postavke")
    api_key = st.text_input("Anthropic API Key", type="password")

geolocator = Nominatim(user_agent="lokacija_generator_hr")
address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

TABLES = [
    {
        "number": "2.1",
        "title": "Opis lokacije",
        "section": "Opis lokacije",
        "rows": [
            ("Opis lokacije",                                               "opis_lokacije"),
            ("Opis okolnih građevina, površina i okoliša",                  "opis_okolnih_gradjevina"),
            ("Načini pristupa",                                             "nacini_pristupa"),
            ("Frekvencija prometa radnim danom, vikendom, noću",            "frekvencija_prometa"),
            ("Stanje kriminaliteta u okolnom prostoru",                     "stanje_kriminaliteta"),
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
            ("Vrsta materijala",                                                          "vrsta_materijala"),
            ("Nagib terena",                                                              "nagib_terena"),
            ("Postojeći elementi javne površine (stepenice, podvožnjaci, objekti) i dr.", "elementi_javne_povrsine"),
        ]
    },
    {
        "number": "2.4",
        "title": "Instalacije",
        "section": "Instalacije",
        "rows": [
            ("Električne instalacije",                        "elektricne_instalacije"),
            ("Ostale instalacije (plin, voda, kanalizacija)",  "ostale_instalacije"),
        ]
    },
    {
        "number": "2.5",
        "title": "Namjena",
        "section": "Namjena",
        "rows": [
            ("Opća namjena",                 "opca_namjena"),
            ("Namjena pojedinih prostora",   "namjena_pojedinih_prostora"),
            ("Radno vrijeme",                "radno_vrijeme"),
            ("Put kretanja osoba i vozila",  "put_kretanja"),
            ("Način zaključavanja prostora", "nacin_zakljucavanja"),
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
            ("Vrste vrijednosti",          "vrste_vrijednosti"),
            ("Visina vrijednosti",         "visina_vrijednosti"),
            ("Način čuvanja vrijednosti",  "nacin_cuvanja_vrijednosti"),
        ]
    },
    {
        "number": "2.8",
        "title": "Organizacija sigurnosti",
        "section": "Organizacija sigurnosti",
        "rows": [
            ("Tjelesna zaštita",     "tjelesna_zastita"),
            ("Tehnička zaštita",     "tehnicka_zastita"),
            ("Organizacijske mjere", "organizacijske_mjere"),
        ]
    },
    {
        "number": "2.9",
        "title": "Stanje dokumentiranosti",
        "section": "Stanje dokumentiranosti",
        "rows": [
            ("Postojeći sustavi zaštite",                           "postojeci_sustavi_zastite"),
            ("Dokumentiranost postojećih sustava tehničke zaštite", "dokumentiranost_sustava"),
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

# 5 malih grupa — max 5 sekcija po pozivu
GROUPS = [
    TABLES[0:2],   # 2.1–2.2  (7 ključeva)
    TABLES[2:4],   # 2.3–2.4  (5 ključeva)
    TABLES[4:6],   # 2.5–2.6  (6 ključeva)
    TABLES[6:9],   # 2.7–2.9  (8 ključeva)
    TABLES[9:11],  # 2.10–2.11 (3 ključa)
]


def call_claude(api_key, location_address, lat, lon, group):
    keys_list = "\n".join(
        f'  "{key}": "..."'
        for table in group
        for (_, key) in table["rows"]
    )
    context = "\n".join(
        f"- {table['title']}: {', '.join(label for label, _ in table['rows'])}"
        for table in group
    )

    prompt = f"""Ti si stručnjak za izradu sigurnosnih elaborata i procjena ugroženosti u Hrvatskoj.
Piši formalno, birokratski i detaljno, kao u službenom elaboratu.

Lokacija: {location_address}
Koordinate: {lat:.6f}, {lon:.6f}

Generiraj tekst za sljedeće sekcije:
{context}

Za svaki ključ napiši 4-6 povezanih rečenica prilagođenih ovoj lokaciji.
Samo čist tekst — bez markdowna, bullet pointova, boldanja ili tablica.

Odgovori SAMO s JSON objektom, bez ikakvih dodatnih znakova:
{{
{keys_list}
}}"""

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    response = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers=headers,
        json={
            "model": "claude-sonnet-4-6",
            "max_tokens": 2500,
            "temperature": 0.6,
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=60
    )

    if response.status_code != 200:
        raise Exception(f"API greška ({response.status_code}): {response.text}")

    raw = response.json()["content"][0]["text"].strip()

    if "```" in raw:
        for part in raw.split("```"):
            part = part.strip().lstrip("json").strip()
            if part.startswith("{"):
                raw = part
                break

    return json.loads(raw)


def add_section_table(doc, table, data):
    heading = doc.add_heading(table["section"], level=3)
    for run in heading.runs:
        run.font.size = Pt(12)

    tbl = doc.add_table(rows=1 + len(table["rows"]), cols=1)
    tbl.style = "Table Grid"

    header_para = tbl.cell(0, 0).paragraphs[0]
    header_para.clear()
    r = header_para.add_run(f"Tablica {table['number']}. {table['title']}")
    r.bold = True
    r.font.size = Pt(10)

    for i, (label, key) in enumerate(table["rows"]):
        para = tbl.cell(i + 1, 0).paragraphs[0]
        para.clear()
        rl = para.add_run(f"{label}: ")
        rl.bold = True
        rl.font.size = Pt(10)
        rt = para.add_run(data.get(key, ""))
        rt.bold = False
        rt.font.size = Pt(10)

    doc.add_paragraph()


if st.button("🚀 Generiraj elaborat", type="primary"):
    if not api_key:
        st.error("Unesi API ključ!")
        st.stop()

    with st.spinner("Geocodiranje adrese..."):
        location = geolocator.geocode(address)
        if not location:
            st.error("Adresa nije pronađena.")
            st.stop()

    st.success(f"📍 {location.address}")

    progress = st.progress(0, text="Pokretanje 5 paralelnih API poziva...")

    with st.spinner("Claude generira tekst..."):
        results = {}
        errors = []
        done_count = 0

        def fetch(group_idx):
            return group_idx, call_claude(
                api_key, location.address,
                location.latitude, location.longitude,
                GROUPS[group_idx]
            )

        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(fetch, i): i for i in range(5)}
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, data = future.result()
                    results.update(data)
                    done_count += 1
                    progress.progress(done_count / 5, text=f"Završeno {done_count}/5 poziva...")
                except Exception as e:
                    errors.append(str(e))
                    done_count += 1
                    progress.progress(done_count / 5, text=f"Završeno {done_count}/5 poziva...")

    if errors:
        st.error("Greške:\n" + "\n".join(errors))
        st.stop()

    progress.progress(1.0, text="Generiranje završeno!")

    st.subheader("📄 Generirani sadržaj")
    for table in TABLES:
        with st.expander(f"Tablica {table['number']} — {table['title']}", expanded=False):
            for label, key in table["rows"]:
                st.markdown(f"**{label}:**")
                st.write(results.get(key, ""))

    doc = Document()
    h = doc.add_heading("Snimka postojećeg stanja", level=1)
    for run in h.runs:
        run.font.size = Pt(14)

    for table in TABLES:
        add_section_table(doc, table, results)

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
