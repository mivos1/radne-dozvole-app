import streamlit as st
import os
import re
import pytesseract
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image
import pandas as pd
from openpyxl import load_workbook
import shutil
import time
import pythoncom
import win32com.client as win32
from urllib.parse import quote

# ==========================================
# ?? OSNOVNA KONFIGURACIJA
# ==========================================
st.set_page_config(page_title="Radne dozvole – Automatizacija", layout="centered")

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_dir = r"C:\Python\poppler\Library\bin"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# ==========================================
# ?? Odabir agencije
# ==========================================
st.title("Automatizirana obrada radnih dozvola")

st.markdown("""
Odaberite svoju agenciju, učitajte PDF datoteku radne dozvole i pokrenite obradu.  
Aplikacija izvlači podatke iz učitane datoteke i sprema iste u excel tablicu.

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

# ?? Postavi dinamicke putanje
agency_folder_name = agency_folders[selected_agency]
base_dir = r"C:\Users\hr.mdrauto.CE\OneDrive - Inter Cars S.A\Documents\Radne_Dozvole_Evidencija"

pdf_root = os.path.join(base_dir, agency_folder_name)
excel_file = os.path.join(base_dir, "Radne_dozvole.xlsm")
one_drive_local_root = base_dir
one_drive_url_root = (
    f"https://icars-my.sharepoint.com/personal/hr_mdrauto_intercars_eu/Documents/Documents/Radne_Dozvole_Evidencija/{agency_folder_name}"
)

st.success(f"Odabrana agencija: {selected_agency}")

# ==========================================
# ?? Upload PDF datoteka
# ==========================================
uploaded_files = st.file_uploader("Odaberite PDF datoteku(e):", type=["pdf"], accept_multiple_files=True)
if not uploaded_files:
    st.info("Odaberite PDF datoteku koju želite obraditi.")
    st.stop()

# ==========================================
# ?? Regex i OCR konfiguracija
# ==========================================
dpi_values = [150, 200, 300]
known_employers = list(agency_folders.keys())

sentence_pattern = re.compile(
    r"(?:\d*\.\s*Dozvola\s+za\s+boravak\s+i\s+rad\s+vrijedi\s+od\s+\d{2}\.\d{2}\.\d{4}\.?\s*do\s*\d{2}\.\d{2}\.\d{4})"
    r"|(?:\d*\.\s*Rok\s+važenja\s+dozvole\s+za\s+boravak\s+i\s+rad\s+(?:je\s+)?\d{2}\.\d{2}\.\d{4}\.?\s*[-–]?\s*\d{2}\.\d{2}\.\d{4})",
    re.IGNORECASE
)
date_pattern = re.compile(r"\d{2}\.\d{2}\.\d{4}")
name_pattern = re.compile(
    r"(?:Za\s+državljanina\s+trece\s+zemlje:\s*([A-ZCCŠÐŽ\s]+),\s*rod)"
    r"|(?:Dozvola\s+za\s+boravak\s+i\s+rad\s+izdaje\s+se\s+([A-ZCCŠÐŽ\s]+)\s+rod)",
    re.IGNORECASE
)
position_pattern = re.compile(
    r'(?:za\s+radno\s+mjesto(?:\s+kod\s+korisnika)?|za\s+zanimanje)\s*[:\-–]?\s*'
    r'([A-ZCCŠÐŽa-zccšdž\/\-\s]+?)'
    r'(?=\s*(?:,|\.|\bkod\s+poslodavca\b|$|\d+\.\s|za\s+zanimanje))',
    re.IGNORECASE
)

def find_employer_in_text(text):
    upper_text = text.upper()
    for employer in known_employers:
        if employer in upper_text:
            return employer
    return selected_agency  # defaultno vraca odabranu agenciju

def clean_filename_for_name(filename: str) -> str:
    name = re.sub(r"\.pdf$", "", filename, flags=re.IGNORECASE)
    name = re.sub(r"(?i)(dozvola\s+za\s+boravak\s+i\s+rad|radna\s+dozvola)", "", name)
    name = re.sub(r"\d{2}\.\d{2}\.\d{4}", "", name)
    name = re.sub(r"[-–_]", " ", name)
    return re.sub(r"\s+", " ", name).strip(" .-_")

def append_to_excel(data_list):
    book = load_workbook(excel_file, keep_vba=True)
    sheet = book.active
    for row_data in data_list:
        sheet.append(row_data)
    book.save(excel_file)
    book.close()

# ==========================================
# ?? Pokretanje obrade
# ==========================================
if st.button("Pokreni obradu"):
    results = []
    processed_dir = os.path.join(pdf_root, "Processed")
    os.makedirs(processed_dir, exist_ok=True)

    for uploaded_file in uploaded_files:
        temp_path = os.path.join(pdf_root, uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.read())

        try:
            info = pdfinfo_from_path(temp_path, poppler_path=poppler_dir)
        except Exception as e:
            st.error(f"Ne mogu otvoriti {uploaded_file.name}: {e}")
            continue

        total_pages = info["Pages"]
        ime_prezime = clean_filename_for_name(uploaded_file.name).title()
        poslodavac = selected_agency
        radno_mjesto = "Data not found"
        vrijedi_od = "Date not found"
        vrijedi_do = "Date not found"
        found_date = found_name = False

        for page_num in range(1, total_pages + 1):
            for dpi in dpi_values:
                try:
                    img_list = convert_from_path(temp_path, dpi=dpi, first_page=page_num, last_page=page_num, poppler_path=poppler_dir)
                except Exception:
                    continue
                text = pytesseract.image_to_string(img_list[0], lang='hrv')

                if not found_name:
                    m = name_pattern.search(text)
                    if m:
                        extracted_name = (m.group(1) or m.group(2)).strip()
                        ime_prezime = extracted_name.title()
                        found_name = True

                if radno_mjesto == "Data not found":
                    txt = re.sub(r'[-\n\r]+', ' ', text)
                    p = position_pattern.search(txt)
                    if p:
                        radno_mjesto = p.group(1).strip().upper()

                if not found_date:
                    m = sentence_pattern.search(text)
                    if m:
                        dates = date_pattern.findall(m.group())
                        if len(dates) == 2:
                            vrijedi_od, vrijedi_do = dates
                        elif len(dates) == 1:
                            vrijedi_do = dates[0]
                        found_date = True

                if found_date and found_name and radno_mjesto != "Data not found":
                    break
            if found_date:
                break

        # premještanje u Processed
        new_pdf_path = os.path.join(processed_dir, uploaded_file.name)
        shutil.move(temp_path, new_pdf_path)
        relative_path = os.path.relpath(new_pdf_path, one_drive_local_root)
        # ukloni naziv agencije iz relacije jer je već u URL rootu
        relative_path = "/".join(relative_path.split(os.sep)[1:])
        link = one_drive_url_root.rstrip('/') + '/' + quote(relative_path.replace("\\", "/"))


        append_to_excel([[ime_prezime, poslodavac, radno_mjesto, vrijedi_od, vrijedi_do, link]])
        results.append(uploaded_file.name)

    if results:
        st.success(f"✅ Uspješno obrađeno: {len(results)} datoteka.")
        try:
            import pythoncom
            pythoncom.CoInitialize()  # 🟢 pokreće COM sesiju
            excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = False
            wb = excel_app.Workbooks.Open(excel_file)
            excel_app.Run("AktivirajLinkove")
            wb.Close(SaveChanges=True)
            excel_app.Quit()
            pythoncom.CoUninitialize()  # 🧹 zatvara COM sesiju
            st.info("Makro **'AktivirajLinkove'** uspješno izvršen. 🔗")
        except Exception as e:
            st.warning(f"⚠️ Makro nije pokrenut: {e}")
    else:
        st.warning("⚠️ Nema uspješno obrađenih datoteka.")

