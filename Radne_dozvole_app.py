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
import win32com.client as win32
from urllib.parse import quote

# ===============================
# 🧩 KONFIGURACIJA
# ===============================

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_dir = r"C:\Python\poppler\Library\bin"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

st.set_page_config(page_title="Radne dozvole – automatizacija", layout="wide")
st.title("📄 Automatizacija radnih dozvola")
st.caption("Obrada PDF-ova, ekstrakcija podataka i automatski unos u Excel (OneDrive).")

# ===============================
# 📂 PUTANJE
# ===============================
pdf_root = r"C:\Users\hr.mdrauto.CE\OneDrive - Inter Cars S.A\Documents\Radne_Dozvole_Evidencija"
excel_file = r"C:\Users\hr.mdrauto.CE\OneDrive - Inter Cars S.A\Documents\Radne_Dozvole_Evidencija\Radne_dozvole.xlsm"
one_drive_local_root = pdf_root
one_drive_url_root = (
    "https://icars-my.sharepoint.com/personal/hr_mdrauto_intercars_eu"
    "/Documents/Documents/Radne_Dozvole_Evidencija"
)

# ===============================
# 📜 DEFINICIJE
# ===============================
dpi_values = [150, 200, 300]
known_employers = [
    "ABILITAS EMPLOYMENT D.O.O.",
    "AGRAM EMPLOYMENT D.O.O.",
    "GLOBAL TEAM NETWORK D.O.O.",
    "MAIN PARTNER D.O.O.",
    "ZAPOSLI STRANCA D.O.O."
]

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

def find_employer_in_text(text):
    upper_text = text.upper()
    for employer in known_employers:
        if employer in upper_text:
            return employer
    return "Data not found"

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

# ===============================
# 🚀 OBRADA PDF-OVA
# ===============================

def process_pdfs():
    st.info("⏳ Pokrećem analizu i ekstrakciju podataka...")

    pdf_files = []
    for root, dirs, files in os.walk(pdf_root):
        if "processed" in [d.lower() for d in dirs]:
            dirs.remove("Processed")
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))

    existing_links = set()
    if os.path.exists(excel_file):
        try:
            df_existing = pd.read_excel(excel_file, dtype=str)
            if "Link" in df_existing.columns:
                existing_links = set(df_existing["Link"].dropna().unique())
        except Exception as e:
            st.warning(f"⚠️ Excel se nije mogao pročitati: {e}")

    rows_to_add = []
    processed_data = []

    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        relative_path_tmp = os.path.relpath(pdf_path, one_drive_local_root)
        link_tmp = one_drive_url_root.rstrip('/') + '/' + quote(relative_path_tmp.replace("\\", "/"))

        if link_tmp in existing_links:
            st.write(f"ℹ️ Preskačem već obrađeni PDF: {filename}")
            continue

        try:
            info = pdfinfo_from_path(pdf_path, poppler_path=poppler_dir)
        except Exception as e:
            st.error(f"❌ Ne mogu otvoriti {filename}: {e}")
            continue

        total_pages = info["Pages"]
        ime_prezime = clean_filename_for_name(filename).title()
        poslodavac = "Data not found"
        radno_mjesto = "Data not found"
        vrijedi_od = "Date not found"
        vrijedi_do = "Date not found"

        found_date = False
        found_name = False

        for page_num in range(1, total_pages + 1):
            for dpi in dpi_values:
                try:
                    img_list = convert_from_path(pdf_path, dpi=dpi, first_page=page_num, last_page=page_num, poppler_path=poppler_dir)
                except Exception:
                    continue

                text = pytesseract.image_to_string(img_list[0], lang='hrv')

                if not found_name:
                    name_match = name_pattern.search(text)
                    if name_match:
                        extracted_name = (name_match.group(1) or name_match.group(2)).strip()
                        ime_prezime = extracted_name.title()
                        found_name = True

                if poslodavac == "Data not found":
                    poslodavac = find_employer_in_text(text)

                if radno_mjesto == "Data not found":
                    text_clean = re.sub(r'[-\n\r]+', ' ', text)
                    position_match = position_pattern.search(text_clean)
                    if position_match:
                        radno_mjesto = position_match.group(1).strip().upper()

                if not found_date:
                    match = sentence_pattern.search(text)
                    if match:
                        dates = date_pattern.findall(match.group())
                        if len(dates) == 2:
                            vrijedi_od, vrijedi_do = dates
                        elif len(dates) == 1:
                            vrijedi_do = dates[0]
                        found_date = True

                if found_date and found_name and poslodavac != "Data not found" and radno_mjesto != "Data not found":
                    break
            if found_date:
                break

        # Premještanje u Processed
        agency_folder = os.path.dirname(pdf_path)
        if os.path.basename(agency_folder).lower() != "processed":
            processed_dir = os.path.join(agency_folder, "Processed")
            os.makedirs(processed_dir, exist_ok=True)
            new_pdf_path = os.path.join(processed_dir, filename)
            shutil.move(pdf_path, new_pdf_path)
            final_path = new_pdf_path
        else:
            final_path = pdf_path

        relative_path = os.path.relpath(final_path, one_drive_local_root)
        link = one_drive_url_root.rstrip('/') + '/' + quote(relative_path.replace("\\", "/"))

        rows_to_add.append([ime_prezime, poslodavac, radno_mjesto, vrijedi_od, vrijedi_do, link])
        processed_data.append({
            "Ime i prezime": ime_prezime,
            "Poslodavac": poslodavac,
            "Radno mjesto": radno_mjesto,
            "Vrijedi od": vrijedi_od,
            "Vrijedi do": vrijedi_do,
            "Link": link
        })

        st.success(f"✅ {filename} — {ime_prezime}")

    # Upis u Excel
    if rows_to_add:
        append_to_excel(rows_to_add)
        st.info("📊 Podaci su dodani u Excel.")
        try:
            excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = False
            wb = excel_app.Workbooks.Open(excel_file)
            excel_app.Run("AktivirajLinkove")
            wb.Close(SaveChanges=True)
            excel_app.Quit()
            st.success("🧩 VBA makro 'AktivirajLinkove' uspješno izvršen.")
        except Exception as e:
            st.warning(f"⚠️ Makro nije pokrenut: {e}")

        st.dataframe(pd.DataFrame(processed_data))
    else:
        st.info("📁 Nema novih PDF-ova za obradu.")

# ===============================
# 🧠 UI GUMB
# ===============================
if st.button("🚀 Pokreni obradu"):
    process_pdfs()
