import streamlit as st
import os
import re
import pytesseract
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image
import pandas as pd
from openpyxl import load_workbook
import shutil
from urllib.parse import quote

# ==========================================
# ⚙️ OSNOVNA KONFIGURACIJA
# ==========================================
st.set_page_config(page_title="Radne dozvole – Cloud OCR", layout="centered")

# 📁 Lokalna struktura (radi i na Streamlit Cloud-u)
base_dir = "Radne_Dozvole_Evidencija"
os.makedirs(base_dir, exist_ok=True)

excel_file = os.path.join(base_dir, "Radne_dozvole.xlsx")

# 🌐 OneDrive URL (samo za prikaz linkova)
one_drive_url_root = (
    "https://icars-my.sharepoint.com/personal/hr_mdrauto_intercars_eu/Documents/Documents/Radne_Dozvole_Evidencija"
)

# ==========================================
# 🧭 UI – ODABIR AGENCIJE
# ==========================================
st.title("Automatizirana obrada radnih dozvola (Cloud verzija)")

st.markdown("""
Odaberite agenciju, učitajte PDF datoteku radne dozvole i pokrenite obradu.  
Aplikacija će izvući podatke i spremiti ih u Excel tablicu (spremno za preuzimanje).
""")

agency_folders = {
    "ABILITAS EMPLOYMENT D.O.O.": "Abilitas",
    "AGRAM EMPLOYMENT D.O.O.": "Agram",
    "GLOBAL TEAM NETWORK D.O.O.": "GTM",
    "MAIN PARTNER D.O.O.": "Main Partner",
    "ZAPOSLI STRANCA D.O.O.": "Zaposli Stranca"
}

selected_agency = st.selectbox("Odaberi agenciju:", list(agency_folders.keys()))
if not selected_agency:
    st.stop()

agency_folder_name = agency_folders[selected_agency]
pdf_root = os.path.join(base_dir, agency_folder_name)
processed_dir = os.path.join(pdf_root, "Processed")
os.makedirs(processed_dir, exist_ok=True)

st.success(f"✅ Odabrana agencija: {selected_agency}")

# ==========================================
# 📤 UPLOAD PDF DOKUMENATA
# ==========================================
uploaded_files = st.file_uploader(
    "Odaberite PDF datoteke za obradu:",
    type=["pdf"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("⬆️ Učitajte barem jednu PDF datoteku za nastavak.")
    st.stop()

# ==========================================
# 🧩 REGEX i OCR KONFIGURACIJA
# ==========================================
dpi_values = [150, 200, 300]

sentence_pattern = re.compile(
    r"(?:\d*\.\s*Dozvola\s+za\s+boravak\s+i\s+rad\s+vrijedi\s+od\s+\d{2}\.\d{2}\.\d{4}\.?\s*do\s*\d{2}\.\d{2}\.\d{4})"
    r"|(?:\d*\.\s*Rok\s+važenja\s+dozvole\s+za\s+boravak\s+i\s+rad\s+(?:je\s+)?\d{2}\.\d{2}\.\d{4}\.?\s*[-–]?\s*\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE
)
date_pattern = re.compile(r"\d{2}\.\d{2}\.\d{4}")
name_pattern = re.compile(
    r"(?:Za\s+državljanina\s+treće\s+zemlje:\s*([A-ZČĆŠĐŽ\s]+),\s*rođ)"
    r"|(?:Dozvola\s+za\s+boravak\s+i\s+rad\s+izdaje\s+se\s+([A-ZČĆŠĐŽ\s]+)\s+rođ)",
    re.IGNORECASE
)
position_pattern = re.compile(
    r'(?:za\s+radno\s+mjesto(?:\s+kod\s+korisnika)?|za\s+zanimanje)\s*[:\-–]?\s*'
    r'([A-ZČĆŠĐŽa-zčćšđž\/\-\s]+?)'
    r'(?=\s*(?:,|\.|\bkod\s+poslodavca\b|$|\d+\.\s|za\s+zanimanje))',
    re.IGNORECASE
)

# ==========================================
# 🧹 FUNKCIJE
# ==========================================
def clean_filename_for_name(filename: str) -> str:
    name = re.sub(r"\.pdf$", "", filename, flags=re.IGNORECASE)
    name = re.sub(r"(?i)(dozvola\s+za\s+boravak\s+i\s+rad|radna\s+dozvola)", "", name)
    name = re.sub(r"\d{2}\.\d{2}\.\d{4}", "", name)
    name = re.sub(r"[-–_]", " ", name)
    return re.sub(r"\s+", " ", name).strip(" .-_")

def append_to_excel(data_list):
    df_new = pd.DataFrame(data_list, columns=[
        "Ime i prezime",
        "Poslodavac",
        "Radno mjesto",
        "Vrijedi od",
        "Vrijedi do",
        "Link"
    ])
    if os.path.exists(excel_file):
        with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            book = writer.book
            sheet = book.active
            start_row = sheet.max_row + 1
            df_new.to_excel(writer, index=False, header=False, startrow=start_row - 1)
    else:
        df_new.to_excel(excel_file, index=False)

# ==========================================
# ▶️ OBRADA DOKUMENATA
# ==========================================
if st.button("Pokreni obradu"):
    results = []
    for uploaded_file in uploaded_files:
        # spremi privremenu kopiju
        temp_path = os.path.join(pdf_root, uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.read())

        try:
            info = pdfinfo_from_path(temp_path)
        except Exception as e:
            st.error(f"Ne mogu otvoriti {uploaded_file.name}: {e}")
            continue

        total_pages = info["Pages"]
        ime_prezime = clean_filename_for_name(uploaded_file.name).title()
        poslodavac = selected_agency
        radno_mjesto = "Data not found"
        vrijedi_od = "Date not found"
        vrijedi_do = "Date not found"

        for page_num in range(1, total_pages + 1):
            for dpi in dpi_values:
                try:
                    img_list = convert_from_path(temp_path, dpi=dpi, first_page=page_num, last_page=page_num)
                except Exception:
                    continue

                text = pytesseract.image_to_string(img_list[0], lang='hrv')

                # ime
                m = name_pattern.search(text)
                if m:
                    ime_prezime = (m.group(1) or m.group(2) or ime_prezime).title()

                # pozicija
                if radno_mjesto == "Data not found":
                    t_clean = re.sub(r'[-\n\r]+', ' ', text)
                    p = position_pattern.search(t_clean)
                    if p:
                        radno_mjesto = p.group(1).strip().upper()

                # datumi
                if vrijedi_od == "Date not found":
                    m = sentence_pattern.search(text)
                    if m:
                        dates = date_pattern.findall(m.group())
                        if len(dates) == 2:
                            vrijedi_od, vrijedi_do = dates
                        elif len(dates) == 1:
                            vrijedi_do = dates[0]

        # premjesti PDF u Processed
        new_pdf_path = os.path.join(processed_dir, uploaded_file.name)
        shutil.move(temp_path, new_pdf_path)

        # generiraj OneDrive link
        relative_path = os.path.relpath(new_pdf_path, base_dir)
        link = f"{one_drive_url_root}/{quote(relative_path.replace(os.sep, '/'))}"

        append_to_excel([[ime_prezime, poslodavac, radno_mjesto, vrijedi_od, vrijedi_do, link]])
        results.append(uploaded_file.name)

    # ==========================================
    # 📦 REZULTAT
    # ==========================================
    if results:
        st.success(f"✅ Uspješno obrađeno: {len(results)} datoteka.")
        st.info("Podaci su spremljeni u **Radne_dozvole.xlsx**.")
        
        # 🟢 Gumb za preuzimanje
        with open(excel_file, "rb") as f:
            st.download_button(
                label="⬇️ Preuzmi Excel datoteku",
                data=f,
                file_name="Radne_dozvole.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Nema uspješno obrađenih datoteka.")
