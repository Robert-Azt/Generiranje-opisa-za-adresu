import streamlit as st
import requests
import time
import re
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


# Adresa — text input s live Nominatim prijedlozima
if "address_confirmed" not in st.session_state:
    st.session_state["address_confirmed"] = "Iločka ulica 34, Zagreb"

typed = st.text_input(
    "Unesite adresu ili koordinate",
    value=st.session_state["address_confirmed"],
    placeholder="npr. Savska cesta 18, Zagreb  ili  45.8095, 15.9578",
    help="Koordinate: decimalni (45.8095, 15.9578) ili DMS (45°49\'31.2\"N 16°06\'56.3\"E)"
)

# Prijedlozi — pojavljuju se dok tipkaš (min 3 znaka, nije koordinata)
import re as _re2
def _is_coord(t):
    t2 = t.strip()
    COORD_CHARS = set("0123456789 .,+-NSEWnsew")
    return any(c.isdigit() for c in t2) and all(c in COORD_CHARS or ord(c) in (176, 39, 34, 8217, 8221) for c in t2)

if typed and len(typed) >= 3 and not _is_coord(typed) and typed != st.session_state["address_confirmed"]:
    try:
        _r = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": typed, "format": "json", "limit": 5, "addressdetails": 1},
            headers={"User-Agent": "lokacija_generator_hr_v3", "Accept-Language": "hr"},
            timeout=5
        )
        _suggestions = _r.json() if _r.ok else []
    except Exception:
        _suggestions = []

    if _suggestions:
        st.caption("Prijedlozi — klikni za odabir:")
        for _s in _suggestions:
            _label = _s.get("display_name", "")
            if st.button(_label, key=f"sug_{_s.get('place_id')}", use_container_width=True):
                st.session_state["address_confirmed"] = _label
                st.rerun()

address = typed if typed else st.session_state["address_confirmed"]


if st.button("🚀 Generiraj elaborat", type="primary"):
    if not api_key:
        st.error("Unesi API ključ!")
        st.stop()

    with st.spinner("📍 Geocodiranje adrese..."):
        display_name = None

        # Provjeri jesu li unesene koordinate
        coords = parse_coords(address)
        coords_input = coords is not None  # je li korisnik unio koordinate
        if coords:
            lat, lon = coords
            short, road, suburb, city = reverse_geocode(lat, lon)
            # Za display: kratki naziv + koordinate
            display_name = f"{short} ({lat:.6f}, {lon:.6f})" if short != f"{lat:.6f}, {lon:.6f}" else f"{lat:.6f}, {lon:.6f}"
        else:
            coords_input = False
            result = geocode(address)
            if result:
                lat, lon, display_name = result
                road, suburb, city = "", "", ""
            else:
                st.error("Adresa nije pronađena. Pokušaj s koordinatama (npr. 45.8095, 15.9578).")
                st.stop()

    st.success(f"📍 {display_name}")

    # Korak 1: identifikacija objekta web searchom (samo ako je odabrano)
    location_info = None
    if identify_mode:
        with st.spinner("🔍 Identifikacija objekta na adresi (web search)..."):
            location_info = identify_location(api_key, display_name, lat, lon, coords_input=coords_input)

        if location_info:
            with st.expander("🏢 Identificirani objekt i okolica", expanded=True):
                st.markdown(f"**Objekt:** {location_info.get('objekt_na_adresi', '—')}")
                st.markdown(f"**Tip:** {location_info.get('tip_objekta', '—')}")
                st.markdown(f"**Okolne ulice:** {location_info.get('okolne_ulice', '—')}")
                st.markdown(f"**Sadržaji u okolici:** {location_info.get('kategorije_okolice', '—')}")
                st.markdown(f"**Javni prijevoz:** {location_info.get('javni_prijevoz', '—')}")
        else:
            st.warning(
                "⚠️ Identifikacija objekta nije uspjela (web search nije vratio rezultat). "
                "Elaborat će se nastaviti generirati kao **Samo lokacija** — "
                "na temelju adrese i podataka o okolici bez prepoznavanja konkretnog objekta."
            )
    else:
        st.info("📍 Način rada: samo lokacija — elaborat se generira bez identifikacije objekta.")

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
            api_key, display_name, lat, lon, context_str, table,
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

    progress.progress(1.0, text="Generiranje završeno!")

    # Prikazi upozorenja za neuspjele tablice — ali NE zaustavljaj
    failed_tables = [e.split(" - ")[0].replace("Tablica ", "") for e in errors if "Tablica" in e]
    if errors:
        failed_names = ", ".join(
            f"{t['number']} ({t['title']})"
            for t in TABLES if t["number"] in failed_tables
        )
        st.warning(
            f"⚠️ {len(errors)} tablica nisu generirane ({failed_names}) — "
            f"označene su u dokumentu. Ostale tablice su generirane i dostupne za preuzimanje."
        )

    # Za svaku neuspjelu tablicu popuni prazna polja s placeholder tekstom
    for table in TABLES:
        if table["number"] in failed_tables:
            for (label, key) in table["rows"]:
                if key not in results:
                    results[key] = f"[NIJE GENERIRANO — greška pri generiranju tablice {table['number']}]"

    st.subheader("📄 Generirani sadržaj")
    for table in TABLES:
        failed = table["number"] in failed_tables
        label = f"⚠️ Tablica {table['number']} — {table['title']} (nije generirano)" if failed else f"Tablica {table['number']} — {table['title']}"
        with st.expander(label, expanded=False):
            for label_row, key in table["rows"]:
                val = results.get(key, "")
                st.markdown(f"**{label_row}:**")
                st.write(val)

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
