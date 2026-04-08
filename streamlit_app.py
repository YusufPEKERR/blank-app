import streamlit as st
import pandas as pd
from lxml import etree
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Side, Border
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
import unicodedata

# ===============================
# YARDIMCI FONKSİYONLAR (Aynen Korundu)
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

def apply_modern_style(ws):
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.freeze_panes = "A2"

# ===============================
# XSD ANALİZ (Web Uyumlu)
# ===============================
def get_xsd_details_from_stream(xsd_file):
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
# STREAMLIT ARAYÜZÜ
# ===============================
st.set_page_config(page_title="XSD to XML Converter", layout="wide")

st.title("🚀 XSD Validation Excel to XML Converter")
st.markdown("XSD dosyanızı yükleyin, şablonu indirin ve verilerinizi XML'e dönüştürün.")

# Sidebar - Adımlar
st.sidebar.header("İşlem Adımları")
step = st.sidebar.radio("Bir aşama seçin:", ["1. XSD Yükle & Şablon Al", "2. Veri Yükle & XML Oluştur"])

if step == "1. XSD Yükle & Şablon Al":
    st.header("1. XSD Dosyası Analizi")
    uploaded_xsd = st.file_uploader("XSD Dosyasını Seçin", type=["xsd"])

    if uploaded_xsd:
        xsd_data = get_xsd_details_from_stream(uploaded_xsd)
        st.success("XSD başarıyla analiz edildi!")
        
        # Excel Oluşturma
        output = BytesIO()
        wb = Workbook()
        ws_g = wb.active
        ws_g.title = "GenelBilgiler"
        ws_g.append(["DisReferansNo", "SerbestBolgeAdi", "FirmaFaaliyetRuhsatiNo", "FirmaFaaliyetRuhsatiKonusu", "GirisTarihi"])
        
        ws_u = wb.create_sheet("Urunler")
        ws_u.append(["UrunSiraNo", "gtip", "UrunAdi", "BirinciMiktar", "BirinciBirim", "UrunMensei"])
        
        ws_h = wb.create_sheet("Hammaddeler")
        ws_h.append(["BagliUrunSiraNo", "ReferansFormTipi", "ReferansFormNo", "ReferansFormYil", "gtip", "Mensei", "BirinciMiktar", "BirinciBirim"])

        # Stil Uygula
        for sheet in wb.worksheets:
            apply_modern_style(sheet)

        wb.save(output)
        st.download_button(
            label="📥 Excel Şablonunu İndir",
            data=output.getvalue(),
            file_name="UBF_Sablon.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "2. Veri Yükle & XML Oluştur":
    st.header("2. Verileri XML'e Dönüştür")
    excel_file = st.file_uploader("Doldurulmuş Excel Dosyasını Yükleyin", type=["xlsx"])
    
    if excel_file:
        try:
            df_g = pd.read_excel(excel_file, "GenelBilgiler")
            df_u = pd.read_excel(excel_file, "Urunler")
            df_h = pd.read_excel(excel_file, "Hammaddeler")
            
            # XML Oluşturma Mantığı
            root = etree.Element("UBFBilgileri")
            # ... (Buraya mevcut XML oluşturma döngülerinizi ekleyebilirsiniz)
            # Örnek:
            if not df_g.empty:
                g = df_g.iloc[0]
                genel = etree.SubElement(root, "UBFGenelBilgiler")
                etree.SubElement(genel, "DisReferansNo").text = str(g.get("DisReferansNo", ""))

            # XML Dosyasını Hazırla
            xml_data = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            
            st.download_button(
                label="✅ XML Dosyasını İndir",
                data=xml_data,
                file_name="UBF_Cikti.xml",
                mime="application/xml"
            )
        except Exception as e:
            st.error(f"Dönüştürme Hatası: {e}")
