import streamlit as st
import requests
import time
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

geolocator = Nominatim(user_agent="lokacija_generator_hr_v2")
address = st.text_input("Unesite adresu", "Iločka ulica 34, Zagreb")

TABLES = [
    {
        "number": "2.1", "title": "Opis lokacije", "section": "Opis lokacije",
        "rows": [
            ("Opis lokacije",                                    "opis_lokacije"),
            ("Opis okolnih građevina, površina i okoliša",       "opis_okolnih_gradjevina"),
            ("Načini pristupa",                                  "nacini_pristupa"),
            ("Frekvencija prometa radnim danom, vikendom, noću", "frekvencija_prometa"),
            ("Stanje kriminaliteta u okolnom prostoru",          "stanje_kriminaliteta"),
        ]
    },
    {
        "number": "2.2", "title": "Osnovne karakteristike", "section": "Osnovne karakteristike",
        "rows": [
            ("Prostorna organiziranost", "prostorna_organiziranost"),
            ("Veličina i namjena",       "velicina_i_namjena"),
        ]
    },
    {
        "number": "2.3", "title": "Građevinske karakteristike", "section": "Građevinske karakteristike",
        "rows": [
            ("Vrsta materijala",                                                          "vrsta_materijala"),
            ("Nagib terena",                                                              "nagib_terena"),
            ("Postojeći elementi javne površine (stepenice, podvožnjaci, objekti) i dr.", "elementi_javne_povrsine"),
        ]
    },
    {
        "number": "2.4", "title": "Instalacije", "section": "Instalacije",
        "rows": [
            ("Električne instalacije",                       "elektricne_instalacije"),
            ("Ostale instalacije (plin, voda, kanalizacija)", "ostale_instalacije"),
        ]
    },
    {
        "number": "2.5", "title": "Namjena", "section": "Namjena",
        "rows": [
            ("Opća namjena",                 "opca_namjena"),
            ("Namjena pojedinih prostora",   "namjena_pojedinih_prostora"),
            ("Radno vrijeme",                "radno_vrijeme"),
            ("Put kretanja osoba i vozila",  "put_kretanja"),
            ("Način zaključavanja prostora", "nacin_zakljucavanja"),
        ]
    },
    {
        "number": "2.6", "title": "Radni procesi", "section": "Radni procesi",
        "rows": [
            ("Popis i opis procesa i postupaka bitnih za sigurnost", "radni_procesi"),
        ]
    },
    {
        "number": "2.7", "title": "Vrste i visina vrijednosti", "section": "Vrsta i visina vrijednosti",
        "rows": [
            ("Vrste vrijednosti",         "vrste_vrijednosti"),
            ("Visina vrijednosti",        "visina_vrijednosti"),
            ("Način čuvanja vrijednosti", "nacin_cuvanja_vrijednosti"),
        ]
    },
    {
        "number": "2.8", "title": "Organizacija sigurnosti", "section": "Organizacija sigurnosti",
        "rows": [
            ("Tjelesna zaštita",     "tjelesna_zastita"),
            ("Tehnička zaštita",     "tehnicka_zastita"),
            ("Organizacijske mjere", "organizacijske_mjere"),
        ]
    },
    {
        "number": "2.9", "title": "Stanje dokumentiranosti", "section": "Stanje dokumentiranosti",
        "rows": [
            ("Postojeći sustavi zaštite",                           "postojeci_sustavi_zastite"),
            ("Dokumentiranost postojećih sustava tehničke zaštite", "dokumentiranost_sustava"),
        ]
    },
    {
        "number": "2.10", "title": "Uočeni nedostaci", "section": "Uočeni nedostaci",
        "rows": [("Uočeni nedostaci", "uoceni_nedostaci")]
    },
    {
        "number": "2.11", "title": "Kritične točke i ugroženi prostori", "section": "Kritične točke i ugroženi prostori",
        "rows": [
            ("Kritične točke",    "kriticne_tocke"),
            ("Ugroženi prostori", "ugrozeni_prostori"),
        ]
    },
]


def fetch_osm_context(lat, lon, radius=350):
    """Dohvati okolne ulice i POI direktno iz Nominatim search API-a."""
    headers = {"User-Agent": "lokacija_generator_hr_v2"}
    streets, pois = set(), []

    # Reverse geocode — daje adresne detalje (četvrt, grad, poštanski broj…)
    try:
        r = requests.get(
            f"https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "json",
                    "addressdetails": 1, "zoom": 16},
            headers=headers, timeout=10
        )
        if r.ok:
            addr = r.json().get("address", {})
        else:
            addr = {}
    except Exception:
        addr = {}

    # Tražimo ulice u bounding boxu oko lokacije (~350m)
    delta = 0.004
    bbox = f"{lon - delta},{lat - delta},{lon + delta},{lat + delta}"
    for q in ["street", "road"]:
        try:
            r = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": q, "format": "json", "limit": 30,
                        "viewbox": bbox, "bounded": 1, "addressdetails": 1},
                headers=headers, timeout=10
            )
            if r.ok:
                for item in r.json():
                    name = item.get("display_name", "").split(",")[0].strip()
                    if name and len(name) > 3:
                        streets.add(name)
        except Exception:
            pass
        time.sleep(0.5)  # Nominatim rate limit

    # POI pretraga
    for q in ["restoran", "škola", "pošta", "banka", "tramvaj", "bus", "park", "trgovina"]:
        try:
            r = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": q, "format": "json", "limit": 5,
                        "viewbox": bbox, "bounded": 1},
                headers=headers, timeout=8
            )
            if r.ok:
                for item in r.json():
                    name = item.get("display_name", "").split(",")[0].strip()
                    if name and len(name) > 2:
                        pois.append(name)
        except Exception:
            pass
        time.sleep(0.3)

    return {
        "adresa_detalji": addr,
        "ulice": sorted(streets)[:15],
        "pois": list(dict.fromkeys(pois))[:12],
    }



def identify_location(api_key, location_address, lat, lon):
    """
    Korak 0: Claude s web_search-om identificira sto se nalazi na adresi
    i kategorizira okolicu. Vraca strukturirani opis za sve ostale promptove.
    """
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    prompt = f"""Istrazi lokaciju i odgovori SAMO JSON objektom bez komentara i markdown oznaka.

Adresa: {location_address}
Koordinate: {lat:.6f}, {lon:.6f}

Pretrazi web i saznaj:
1. Sto se tocno nalazi na ovoj adresi (zgrada, objekt, institucija, stadion, park, skola, poslovni objekt...)?
2. Koje ulice okruzuju lokaciju (navedi tocna imena)?
3. Koje kategorije sadrzaja postoje u okolici? Koristi generike nazive, bez konkretnih imena:
   npr. ugostiteljski objekti, stambene zgrade, poslovni objekti, trgovacki centri, javne ustanove, zelene povrsine.
4. Kako je lokacija prometno povezana (tramvaj, bus, koje linije/stanice)?

Odgovori ISKLJUCIVO ovim JSON objektom:
{{
  "objekt_na_adresi": "konkretan naziv i tip objekta koji se nalazi na adresi",
  "tip_objekta": "jedna kategorija: stambena zgrada / poslovni objekt / sportski objekt / javna ustanova / skola / bolnica / park / industrijsko postrojenje / misovito",
  "okolne_ulice": "stvarna imena ulica odvojena zarezom",
  "kategorije_okolice": "genericki opis sadrzaja u okolici bez konkretnih naziva objekata",
  "javni_prijevoz": "opis javnog prijevoza s linijama i stanicama",
  "dodatni_kontekst": "ostalo relevantno za sigurnosni elaborat"
}}"""

    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 1000,
                "messages": [{"role": "user", "content": prompt}],
                "tools": [{"type": "web_search_20250305", "name": "web_search"}]
            },
            timeout=60
        )

        if response.status_code != 200:
            return None

        content = response.json().get("content", [])
        raw = ""
        for block in content:
            if block.get("type") == "text":
                raw += block.get("text", "")

        raw = raw.strip()
        if "```" in raw:
            for part in raw.split("```"):
                part = part.strip().lstrip("json").strip()
                if part.startswith("{"):
                    raw = part
                    break

        return json.loads(raw)

    except Exception:
        return None

def build_context_string(osm):
    """Složi čitljiv kontekst za Claude prompt."""
    parts = []
    addr = osm.get("adresa_detalji", {})

    # Administrativni podaci
    admin = []
    for k in ["road", "suburb", "quarter", "city_district", "postcode", "city", "county"]:
        if v := addr.get(k):
            admin.append(v)
    if admin:
        parts.append("Administrativni podaci: " + ", ".join(admin))

    if osm.get("ulice"):
        parts.append("Okolne ulice (u blizini lokacije): " + ", ".join(osm["ulice"]))

    if osm.get("pois"):
        parts.append("Obližnji sadržaji i objekti: " + ", ".join(osm["pois"]))

    return "\n".join(parts) if parts else ""


def call_claude_one_table(api_key, location_address, lat, lon, context_str, table, location_info=None, retries=4):
    """Jedan API poziv po tablici, s retry na 429."""
    keys_list = "\n".join(f'  "{key}": "..."' for (_, key) in table["rows"])
    row_descriptions = "\n".join(
        f'- "{key}": 3-5 recenica o "{label}"' 
        for (label, key) in table["rows"]
    )

    objekt_block = ""
    if location_info:
        objekt_block = (
            "\nOBJEKT NA ADRESI (elaborat se pise za ovaj objekt):\n"
            f"- Naziv i tip: {location_info.get('objekt_na_adresi', '')}\n"
            f"- Kategorija: {location_info.get('tip_objekta', '')}\n"
            f"- Dodatni kontekst: {location_info.get('dodatni_kontekst', '')}\n"
            "\nPODACI O OKOLICI (ne navodi konkretna imena kaficau/restorana, koristi kategorije):\n"
            f"- Okolne ulice: {location_info.get('okolne_ulice', '')}\n"
            f"- Sadrzaji: {location_info.get('kategorije_okolice', '')}\n"
            f"- Javni prijevoz: {location_info.get('javni_prijevoz', '')}\n"
        )
    elif context_str:
        objekt_block = f"\nPodaci o lokaciji:\n{context_str}\n"

    prompt = (
        "Ti si strucnjak za izradu sigurnosnih elaborata u Hrvatskoj. Pisi formalno i birokratski.\n\n"
        f"Lokacija: {location_address}\n"
        f"Koordinate: {lat:.6f}, {lon:.6f}\n"
        f"{objekt_block}\n"
        "UPUTE:\n"
        "- Pisi specificno za tip objekta naveden gore\n"
        "- Navodi prava imena ulica gdje je relevantno\n"
        "- Za sadrzaje u okolici koristi kategorije (ugostiteljski objekti, stambene zgrade...) ne konkretna imena\n"
        "- 3-5 recenica po polju, cist tekst bez formatiranja\n"
        "- Bez navodnika u tekstu (pokvarilo bi JSON)\n\n"
        f"Tablica: {table['number']} - {table['title']}\n"
        f"{row_descriptions}\n\n"
        "Odgovori SAMO ovim JSON objektom:\n"
        "{\n"
        f"{keys_list}\n"
        "}"
    )

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

    with st.spinner("📍 Geocodiranje adrese..."):
        location = geolocator.geocode(address)
        if not location:
            st.error("Adresa nije pronađena.")
            st.stop()
        lat, lon = location.latitude, location.longitude

    st.success(f"📍 {location.address}")

    # Korak 1: identifikacija objekta web searchom
    location_info = None
    with st.spinner("🔍 Identifikacija objekta na adresi (web search)..."):
        location_info = identify_location(api_key, location.address, lat, lon)

    if location_info:
        with st.expander("🏢 Identificirani objekt i okolica", expanded=True):
            st.markdown(f"**Objekt:** {location_info.get('objekt_na_adresi', '—')}")
            st.markdown(f"**Tip:** {location_info.get('tip_objekta', '—')}")
            st.markdown(f"**Okolne ulice:** {location_info.get('okolne_ulice', '—')}")
            st.markdown(f"**Sadržaji u okolici:** {location_info.get('kategorije_okolice', '—')}")
            st.markdown(f"**Javni prijevoz:** {location_info.get('javni_prijevoz', '—')}")
    else:
        st.warning("Identifikacija objekta nije uspjela — elaborat će biti generiran na temelju adrese.")

    # Korak 2: OSM fallback (okolne ulice ako web search nije dao dovoljno)
    context_str = ""
    if not location_info:
        with st.spinner("🗺️ Dohvat podataka o okolici (fallback)..."):
            try:
                osm_data = fetch_osm_context(lat, lon)
                context_str = build_context_string(osm_data)
            except Exception:
                pass

    total = len(TABLES)
    progress = st.progress(0, text="Generiranje elaborata...")
    results = {}
    errors = []
    done_count = 0

    def fetch(table):
        return table["number"], call_claude_one_table(
            api_key, location.address, lat, lon, context_str, table,
            location_info=location_info
        )

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(fetch, t): t for t in TABLES}
        for future in concurrent.futures.as_completed(futures):
            try:
                num, data = future.result()
                results.update(data)
            except Exception as e:
                errors.append(str(e))
            finally:
                done_count += 1
                progress.progress(done_count / total, text=f"Završeno {done_count}/{total} tablica...")

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
