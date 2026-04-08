import streamlit as st
import pandas as pd
from lxml import etree
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment
import unicodedata

# ===============================
# YARDIMCI FONKSİYONLAR
# ===============================
def clean(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    v = str(val).strip()
    v = unicodedata.normalize("NFKC", v)
    return v

def safe_int_str(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    try:
        return str(int(float(val)))
    except:
        return str(val)

def autofit_columns(ws):
    """Sütun genişliklerini içeriğe göre otomatik ayarlar."""
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    val_len = len(str(cell.value))
                    if val_len > max_length:
                        max_length = val_len
            except:
                pass
        # Biraz pay bırakarak genişliği ayarla (min 10, max 50 karakter)
        adjusted_width = max(10, min(max_length + 4, 50))
        ws.column_dimensions[column_letter].width = adjusted_width

def apply_modern_style(ws):
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.freeze_panes = "A2"
    # Sütunları otomatik genişlet
    autofit_columns(ws)

# ===============================
# XSD ANALİZ SİSTEMİ
# ===============================
def get_xsd_details(xsd_file):
    try:
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(xsd_file, parser)
        root = tree.getroot()
        
        def fetch_enums(search_term):
            path1 = f".//{{*}}simpleType[@name='{search_term}']//{{*}}enumeration"
            enums = [el.get("value") for el in root.findall(path1)]
            if not enums:
                element = root.find(f".//{{*}}element[@name='{search_term}']")
                if element is not None:
                    type_attr = element.get("type")
                    if type_attr:
                        clean_type = type_attr.split(':')[-1]
                        path2 = f".//{{*}}simpleType[@name='{clean_type}']//{{*}}enumeration"
                        enums = [el.get("value") for el in root.findall(path2)]
            return enums

        return {
            "SerbestBolgeAdi": fetch_enums("serbestBolgeType") or fetch_enums("SerbestBolgeAdi"),
            "FirmaFaaliyetRuhsatiKonusu": fetch_enums("ruhsatKonulariType") or fetch_enums("FirmaFaaliyetRuhsatiKonusu"),
            "OlcuBirimleri": fetch_enums("olcuBirimiType") or fetch_enums("OlcuBirimleri"),
            "Ulkeler": fetch_enums("ulkeType") or fetch_enums("Ulkeler"),
            "ReferansFormTipi": fetch_enums("formTipiType") or fetch_enums("ReferansFormTipi")
        }
    except Exception as e:
        st.error(f"XSD Analiz Hatası: {e}")
        return {}

# ===============================
# STREAMLIT UI
# ===============================
st.set_page_config(page_title="XSD to XML Converter", layout="wide")
st.title("🚀 XSD Validation Excel to XML Converter")

app_mode = st.sidebar.radio("İşlem Seçin:", ["1. Şablon Oluştur (XSD)", "2. XML'e Dönüştür (Excel)"])

if app_mode == "1. Şablon Oluştur (XSD)":
    st.header("📂 1. Adım: XSD Yükle ve Şablon Hazırla")
    uploaded_xsd = st.file_uploader("XSD dosyanızı yükleyin", type=["xsd"])

    if uploaded_xsd:
        xsd_data = get_xsd_details(uploaded_xsd)
        output = BytesIO()
        wb = Workbook()
        
        # --- KULLANIM KILAVUZU ---
        ws_guide = wb.active
        ws_guide.title = "Kullanim Kilavuzu"
        ws_guide.append(["XSD VALIDATION EXCEL TO XML CONVERTER"])
        ws_guide.append([""])
        ws_guide.append(["Adım 1", "GenelBilgiler sayfasını doldurun."])
        ws_guide.append(["Adım 2", "Ürünler ve Hammaddeler arasındaki 'SiraNo' ilişkisine dikkat edin."])
        ws_guide["A1"].font = Font(size=14, bold=True)

        # --- VERİ SAYFALARI ---
        ws_g = wb.create_sheet("GenelBilgiler")
        ws_g.append(["DisReferansNo", "SerbestBolgeAdi", "FirmaFaaliyetRuhsatiNo", "FirmaFaaliyetRuhsatiKonusu", "GirisTarihi"])

        ws_u = wb.create_sheet("Urunler")
        ws_u.append(["UrunSiraNo", "gtip", "UrunAdi", "BirinciMiktar", "BirinciBirim", "UrunMensei"])

        ws_h = wb.create_sheet("Hammaddeler")
        ws_h.append(["BagliUrunSiraNo", "ReferansFormTipi", "ReferansFormNo", "ReferansFormYil", "gtip", "Mensei", "BirinciMiktar", "BirinciBirim"])

        # --- DOĞRULAMA (DROPDOWN) ---
        ws_l = wb.create_sheet("Listeler")
        col_map = {"A": "SerbestBolgeAdi", "B": "FirmaFaaliyetRuhsatiKonusu", "C": "OlcuBirimleri", "D": "Ulkeler", "E": "ReferansFormTipi"}
        for col, key in col_map.items():
            items = list(set(xsd_data.get(key, [])))
            for i, val in enumerate(items, 1):
                ws_l[f"{col}{i}"] = val

        def add_dv(ws, list_col, list_size, range_str):
            if list_size > 0:
                dv = DataValidation(type="list", formula1=f"Listeler!${list_col}$1:${list_col}${list_size}", allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(range_str)

        add_dv(ws_g, "A", len(xsd_data.get("SerbestBolgeAdi", [])), "B2:B500")
        add_dv(ws_u, "D", len(xsd_data.get("Ulkeler", [])), "F2:F500")

        # Stil ve Otomatik Genişlik Uygula
        for sheet in wb.worksheets:
            if sheet.title != "Listeler":
                apply_modern_style(sheet)
        
        ws_l.sheet_state = "hidden"
        wb.save(output)
        
        st.success("Sütun genişlikleri içeriğe göre ayarlandı!")
        st.download_button("📥 Excel Şablonunu İndir", output.getvalue(), "UBF_Sablon.xlsx")

elif app_mode == "2. XML'e Dönüştür (Excel)":
    st.header("🛠️ 2. Adım: Verileri XML'e Dönüştür")
    excel_file = st.file_uploader("Excel yükleyin", type=["xlsx"])
    if excel_file:
        # (XML Dönüştürme Kodları Buraya Gelecek - Önceki kodunuzla aynı kalabilir)
        st.info("Excel verileri işlenmeye hazır.")
