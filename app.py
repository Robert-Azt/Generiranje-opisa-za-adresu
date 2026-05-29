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

address = st.text_input(
    "Unesite adresu ili koordinate",
    "Iločka ulica 34, Zagreb",
    help="Primjeri: Savska cesta 18, Zagreb  |  45.8095, 15.9578  |  45°46\'04.4\"N 15°59\'27.7\"E"
)


def parse_coords(text):
    # Parsiraj koordinate iz teksta (decimalni ili DMS format)
    # Vraca (lat, lon) ili None
    t = text.strip()

    # Decimalni: 45.8095, 15.9578
    m = re.match(r'^(-?\d+\.?\d*)[,\s]+(-?\d+\.?\d*)$', t)
    if m:
        lat, lon = float(m.group(1)), float(m.group(2))
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon

    # DMS: 45°46'04.4"N 15°59'27.7"E
    pat = (r'(\d+)[° ]\s*(\d+)[\'\u2019 ]\s*([\d.]+)[",]?\s*([NSns])'
           r'[\s,]+(\d+)[° ]\s*(\d+)[\'\u2019 ]\s*([\d.]+)[",]?\s*([EWew])')
    m2 = re.search(pat, t)
    if m2:
        lat = int(m2.group(1)) + int(m2.group(2))/60 + float(m2.group(3))/3600
        if m2.group(4).upper() == 'S': lat = -lat
        lon = int(m2.group(5)) + int(m2.group(6))/60 + float(m2.group(7))/3600
        if m2.group(8).upper() == 'W': lon = -lon
        return lat, lon

    return None

def geocode(query):
    """
    Geocodiranje bez geopy — direktni HTTP na Nominatim.
    Vraca (lat, lon, display_name) ili None.
    """
    headers = {"User-Agent": "lokacija_generator_hr_v3"}
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": query, "format": "json", "limit": 1, "addressdetails": 1},
            headers=headers, timeout=10
        )
        if r.ok and r.json():
            item = r.json()[0]
            return float(item["lat"]), float(item["lon"]), item["display_name"]
    except Exception:
        pass
    return None


def reverse_geocode(lat, lon):
    """Reverse geocodiranje — vraca (display_name, road, suburb, city)."""
    headers = {"User-Agent": "lokacija_generator_hr_v3"}
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "json",
                    "zoom": 18, "addressdetails": 1},
            headers=headers, timeout=10
        )
        if r.ok:
            data = r.json()
            addr = data.get("address", {})
            road    = addr.get("road", "")
            suburb  = addr.get("suburb", addr.get("quarter", ""))
            city    = addr.get("city", addr.get("town", ""))
            display = data.get("display_name", "")
            # Sastavi kratki opis: "Ul. Savezne Republike Njemačke / Vatikanska, Maksimir, Zagreb"
            parts = [p for p in [road, suburb, city] if p]
            short  = ", ".join(parts) if parts else display
            return short or f"{lat:.6f}, {lon:.6f}", road, suburb, city
    except Exception:
        pass
    return f"{lat:.6f}, {lon:.6f}", "", "", ""

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



def identify_location(api_key, location_address, lat, lon, coords_input=False):
    """
    Korak 0: Claude s web_search-om identificira sto se nalazi na lokaciji.
    coords_input=True znaci da korisnik nije unio adresu nego samo koordinate
    (krizanje, neobiljezenaa lokacija) — Claude treba sam iz karte shvatiti sto je tamo.
    """
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    if coords_input:
        # Koordinate bez adrese — Claude mora sam identificirati lokaciju
        lokacija_opis = (
            f"KOORDINATE: {lat:.6f}, {lon:.6f}\n"
            f"Nominatim reverse geocode: {location_address}\n\n"
            "VAZNO: Korisnik je unio iskljucivo koordinate (nema naziva objekta ni adrese). "
            "To znaci da se radi o krizanju, neoznacenoj lokaciji ili teski za adresiranje mjestu. "
            "Pretrazi web koristeci koordinate i saznaj tocno sto se nalazi na toj lokaciji "
            "(npr. pretrazi OpenStreetMap, Google Maps ili slicne izvore za te koordinate). "
            "Ako je krizanje — navedi koje se ulice krizaju. "
            "Ako je park, zelenilo, prometnica — opisi to tocno."
        )
    else:
        lokacija_opis = (
            f"UNESENA ADRESA: {location_address}\n"
            f"Koordinate: {lat:.6f}, {lon:.6f}\n\n"
            "VAZNO: Korisnik je unio adresu s kucnim brojem. "
            "OBAVEZNO pretrazi web za tocnu adresu i saznaj koji OBJEKT se nalazi na toj adresi "
            "(banka, trgovina, skola, stambena zgrada, ured, restoran, hotel...). "
            "Adresa s kucnim brojem gotovo uvijek ukazuje na konkretan objekt — ne pretpostavljaj da je to samo ulica ili krizanje. "
            "Pretrazi npr. naziv firme, institucije ili objekta na toj adresi."
        )

    prompt = "\n".join([
        "Istrazi lokaciju i odgovori SAMO JSON objektom bez komentara i markdown oznaka.",
        "",
        lokacija_opis,
        "",
        "Upute za pretragu:",
        "1. Ako je unesena adresa s kucnim brojem: PRVO pretrazi koji konkretan objekt (firma, institucija, zgrada) se nalazi tocno na toj adresi.",
        "   Primjeri pretrage: naziv ulice + broj + grad, ili adresa + djelatnost.",
        "   Ako je to banka — navedi naziv banke. Ako je trgovina — navedi naziv. Ako je stambena zgrada — navedi to.",
        "2. Ako su unesene samo koordinate ili nema kucnog broja: identificiraj sto je na lokaciji (krizanje, park, otvorena povrsina...)",
        "3. Koje ulice okruzuju lokaciju — navedi tocna imena",
        "4. Kategorije sadrzaja u okolici (genericki: stambene zgrade, poslovni objekti, parkovi... — bez konkretnih imena kaficaa)",
        "5. Javni prijevoz — tramvaj/bus linije i stanice u blizini",
        "",
        "Odgovori ISKLJUCIVO ovim JSON objektom:",
        "{",
        '  "objekt_na_adresi": "puni naziv i tip objekta (npr. PBZ banka, poslovna zgrada; ili NK Lokomotiva, novi stadion u izgradnji; ili krizanje Ozaljske i Nehajske ulice)",',
        '  "tip_objekta": "jedna kategorija: stambena zgrada / poslovni objekt / banka / sportski objekt / javna ustanova / skola / bolnica / park / prometnica-krizanje / ugostiteljski objekt / trgovina / industrijsko postrojenje / misovito",',
        '  "okolne_ulice": "stvarna imena ulica odvojena zarezom",',
        '  "kategorije_okolice": "genericki opis sadrzaja u okolici bez konkretnih naziva objekata",',
        '  "javni_prijevoz": "opis javnog prijevoza s linijama i stanicama",',
        '  "dodatni_kontekst": "ostalo relevantno za sigurnosni elaborat"',
        "}",
    ])

    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 1200,
                "messages": [{"role": "user", "content": prompt}],
                "tools": [{"type": "web_search_20250305", "name": "web_search"}]
            },
            timeout=60
        )

        if response.status_code != 200:
            return None

        blocks = response.json().get("content", [])
        raw = "".join(b.get("text", "") for b in blocks if b.get("type") == "text").strip()

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
    import re as _re

    # Grad JSON predložak za odgovor
    json_keys = ",\n".join(f'  "{key}": "..."' for (_, key) in table["rows"])
    json_template = "{\n" + json_keys + "\n}"

    row_lines = "\n".join(
        f'- {key}: 3-5 recenica o temi "{label}"'
        for (label, key) in table["rows"]
    )

    # Kontekst o objektu
    ctx_lines = []
    if location_info:
        ctx_lines.append("OBJEKT NA ADRESI: " + location_info.get("objekt_na_adresi", ""))
        ctx_lines.append("TIP OBJEKTA: " + location_info.get("tip_objekta", ""))
        ctx_lines.append("OKOLNE ULICE: " + location_info.get("okolne_ulice", ""))
        ctx_lines.append("SADRZAJI U OKOLICI (koristiti kategorije, ne konkretna imena): " + location_info.get("kategorije_okolice", ""))
        ctx_lines.append("JAVNI PRIJEVOZ: " + location_info.get("javni_prijevoz", ""))
        ctx_lines.append("DODATNO: " + location_info.get("dodatni_kontekst", ""))
    elif context_str:
        ctx_lines.append(context_str)
    ctx_block = "\n".join(ctx_lines)

    prompt = "\n".join([
        "Ti si strucnjak za izradu sigurnosnih elaborata u Hrvatskoj. Pisi formalno i birokratski.",
        "",
        "Lokacija: " + location_address,
        "Koordinate: " + f"{lat:.6f}, {lon:.6f}",
        "",
        ctx_block,
        "",
        "UPUTE:",
        "- Elaborat se odnosi na objekt naveden gore - pisi specificno za taj tip objekta",
        "- Navodi prava imena okolnih ulica gdje je relevantno",
        "- Za sadrzaje u okolici koristi opce kategorije, ne konkretna imena ugostiteljskih i slicnih objekata",
        "- 3-5 recenica po polju, cist tekst bez formatiranja",
        "- KRITИЧНО: vrijednosti u JSON-u NE smiju sadrzavati navodnike. Umjesto navodnika koristi zareze ili zagrade.",
        "- Za polja koja nisu primjenjiva za dani tip objekta napisi jednu jasnu konstatacijsku recenicu. Primjeri po tipu:\n  PARK: vrsta materijala = podloge staza i mobilijar (nema gradjevine); ostale instalacije = nema plinskih/vodovodnih; nacin zakljucavanja = javni prostor otvoren 24h; tjelesna zastita = ne provodi se.\n  KRIZANJE/PROMETNICA: vrsta materijala = asfaltni zastor kolnika, betonske plocice nogostupa; instalacije = podzemni komunalni vodovi bez nadzemne gradevine; nacin zakljucavanja = ne zakljucava se; tjelesna zastita = nema, samo prometna signalizacija.\n  BANKA ili POSLOVNI OBJEKT: sva polja su primjenjiva — pisi normalno o zgradi, instalacijama, zastitarima, sustavima videonadzora, sefu, ranom vremenu, nacinu zakljucavanja itd.\n  STAMBENA ZGRADA: radni procesi = nema specificnih radnih procesa; tjelesna zastita = ne provodi se sustavna tjelesna zastita.\n  SKOLA/BOLNICA/USTANOVA: sve tablice primjenjive, naglasi specificnosti ustanove (djeca, pacijenti, radno vrijeme).\n  Opce pravilo: ako objekt ima gradevinu — pisi o njoj. Ako nema (park, krizanje) — jednom recenicom navedi da ne postoji.",
        "",
        "Tablica: " + table["number"] + " - " + table["title"],
        "",
        row_lines,
        "",
        "Odgovori ISKLJUCIVO sljedecim JSON objektom, bez ikakvih dodatnih rijeci prije ili poslije:",
        json_template,
    ])

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    for attempt in range(retries):
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json={
                "model": "claude-haiku-4-5-20251001",
                "max_tokens": 1500,
                "temperature": 0.4,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=45
        )

        if response.status_code == 429:
            time.sleep(5 * (attempt + 1))
            continue

        if response.status_code != 200:
            raise Exception(f"Tablica {table['number']} - greska {response.status_code}: {response.text}")

        raw = response.json()["content"][0]["text"].strip()

        # Ukloni markdown backticks
        if "```" in raw:
            for part in raw.split("```"):
                part = part.strip().lstrip("json").strip()
                if part.startswith("{"):
                    raw = part
                    break

        # Pokusaj 1: direktni json.loads
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass

        # Pokusaj 2: regex po kljucu
        result = {}
        for (_, key) in table["rows"]:
            pattern = rf'"{_re.escape(key)}"\s*:\s*"((?:[^"\\]|\\.)*)"'
            m = _re.search(pattern, raw, _re.DOTALL)
            if m:
                value = m.group(1).replace('\\"', '"').replace('\\n', ' ').replace('\\t', ' ')
                result[key] = value.strip()
            else:
                result[key] = ""

        if any(result.values()):
            return result

        raise Exception(f"Tablica {table['number']} - ne mogu parsirati:\n{raw[:300]}")

    raise Exception(f"Tablica {table['number']} - previse pokusaja (rate limit)")


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
        val = data.get(key, "")
        rt = para.add_run(val)
        rt.bold = False
        rt.font.size = Pt(10)

    doc.add_paragraph()


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

    # Korak 1: identifikacija objekta web searchom
    location_info = None
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

    if errors:
        st.error("Greške:\n" + "\n".join(errors))
        st.stop()

    progress.progress(1.0, text="Generiranje završeno!")

    st.subheader("📄 Generirani sadržaj")
    for table in TABLES:
        with st.expander(f"Tablica {table['number']} — {table['title']}", expanded=False):
            for label, key in table["rows"]:
                val = results.get(key, "")
                st.markdown(f"**{label}:**")
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
