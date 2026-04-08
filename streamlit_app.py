import streamlit as st
import pandas as pd
from lxml import etree
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
        # Başlıkların okunabilirliği için min 15, max 50 karakter sınırı
        adjusted_width = max(15, min(max_length + 5, 50))
        ws.column_dimensions[column_letter].width = adjusted_width

def apply_modern_style(ws):
    """Sayfa başlıklarını boyar, dondurur ve sütunları genişletir."""
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    ws.freeze_panes = "A2"
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
        st.error(f"XSD Okuma Hatası: {e}")
        return {}

# ===============================
# STREAMLIT ARAYÜZÜ
# ===============================
st.set_page_config(page_title="XSD to XML Converter", layout="wide")

st.title("🚀 XSD Validation Excel to XML Converter")
st.sidebar.header("İşlem Menüsü")
app_mode = st.sidebar.radio("Bir aşama seçin:", ["1. Şablon Oluştur (XSD)", "2. XML'e Dönüştür (Excel)"])

# -------------------------------
# MOD 1: ŞABLON OLUŞTURMA
# -------------------------------
if app_mode == "1. Şablon Oluştur (XSD)":
    st.header("📂 1. Adım: XSD Yükle ve Şablon Hazırla")
    uploaded_xsd = st.file_uploader("Sistemin XSD dosyasını yükleyin", type=["xsd"])

    if uploaded_xsd:
        xsd_data = get_xsd_details(uploaded_xsd)
        output = BytesIO()
        wb = Workbook()
        
        # --- KULLANIM KILAVUZU ---
        ws_guide = wb.active
        ws_guide.title = "Kullanim Kilavuzu"
        guide_data = [
            ["XSD VALIDATION EXCEL TO XML CONVERTER"],
            [""],
            ["NASIL KULLANILIR?"],
            ["1. GenelBilgiler", "İlgili alanları doldurun (Serbest Bölge adı listeden seçilmelidir)."],
            ["2. Urunler", "Her ürüne 1, 2, 3 gibi benzersiz 'UrunSiraNo' verin."],
            ["3. Hammaddeler", "Hammadde hangi ürüne aitse 'BagliUrunSiraNo' kısmına o ürünün numarasını yazın."],
            ["4. Veri Doğrulama", "Birim ve Ülke gibi alanlarda sadece açılır listedeki kodları kullanın."],
            [""],
            ["⚠️ DİKKAT:", "Sütun başlıklarını veya sayfa isimlerini asla değiştirmeyin."]
        ]
        for row in guide_data:
            ws_guide.append(row)
        ws_guide["A1"].font = Font(size=14, bold=True, color="2C3E50")
        ws_guide.column_dimensions['A'].width = 25
        ws_guide.column_dimensions['B'].width = 60

        # --- VERİ SAYFALARI ---
        ws_g = wb.create_sheet("GenelBilgiler")
        ws_g.append(["DisReferansNo", "SerbestBolgeAdi", "FirmaFaaliyetRuhsatiNo", "FirmaFaaliyetRuhsatiKonusu", "GirisTarihi"])

        ws_u = wb.create_sheet("Urunler")
        ws_u.append(["UrunSiraNo", "gtip", "UrunAdi", "BirinciMiktar", "BirinciBirim", "UrunMensei"])

        ws_h = wb.create_sheet("Hammaddeler")
        ws_h.append(["BagliUrunSiraNo", "ReferansFormTipi", "ReferansFormNo", "ReferansFormYil", "gtip", "Mensei", "BirinciMiktar", "BirinciBirim"])

        # --- LİSTELER VE DROPDOWN ---
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

        add_dv(ws_g, "A", len(xsd_data.get("SerbestBolgeAdi", [])), "B2:B1000")
        add_dv(ws_u, "C", len(xsd_data.get("OlcuBirimleri", [])), "E2:E2000")
        add_dv(ws_u, "D", len(xsd_data.get("Ulkeler", [])), "F2:F2000")
        add_dv(ws_h, "E", len(xsd_data.get("ReferansFormTipi", [])), "B2:B5000")

        # Modern Stil ve Otomatik Genişlik Uygula
        for sheet in wb.worksheets:
            if sheet.title != "Listeler":
                apply_modern_style(sheet)
        
        ws_l.sheet_state = "hidden"
        wb.save(output)
        
        st.success("Şablon başarıyla hazırlandı (Sütunlar otomatik genişletildi)!")
        st.download_button(
            label="📥 Kullanım Kılavuzlu Şablonu İndir",
            data=output.getvalue(),
            file_name="Excel_Sablonu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------------------
# MOD 2: XML DÖNÜŞTÜRME
# -------------------------------
elif app_mode == "2. XML'e Dönüştür (Excel)":
    st.header("🛠️ 2. Adım: Verileri XML'e Dönüştür")
    excel_file = st.file_uploader("Doldurulmuş Excel dosyasını seçin", type=["xlsx"])

    if excel_file:
        try:
            df_g = pd.read_excel(excel_file, sheet_name="GenelBilgiler")
            df_u = pd.read_excel(excel_file, sheet_name="Urunler")
            df_h = pd.read_excel(excel_file, sheet_name="Hammaddeler")

            root = etree.Element("UBFBilgileri")

            # Genel Bilgiler Bölümü
            if not df_g.empty:
                g = df_g.iloc[0]
                genel = etree.SubElement(root, "UBFGenelBilgiler")
                for field in ["DisReferansNo", "SerbestBolgeAdi", "FirmaFaaliyetRuhsatiNo", "FirmaFaaliyetRuhsatiKonusu", "GirisTarihi"]:
                    val = clean(g.get(field))
                    if val: etree.SubElement(genel, field).text = val

            # Ürün ve Hammadde Döngüsü
            for _, u in df_u.iterrows():
                u_el = etree.SubElement(root, "UrunBilgileri")
                urun = etree.SubElement(u_el, "Urun")
                
                gt_u = "".join(filter(str.isdigit, str(clean(u.get("gtip")) or "")))
                etree.SubElement(urun, "gtip").text = gt_u.zfill(12)[:12]
                etree.SubElement(urun, "UrunAdi").text = clean(u.get("UrunAdi")) or ""
                etree.SubElement(urun, "BirinciMiktar").text = str(u.get("BirinciMiktar") or "0")
                etree.SubElement(urun, "BirinciBirim").text = clean(u.get("BirinciBirim")) or ""
                etree.SubElement(urun, "UrunMensei").text = clean(u.get("UrunMensei")) or ""

                u_id = safe_int_str(u.get("UrunSiraNo"))
                if u_id:
                    bagli_hamlar = df_h[df_h["BagliUrunSiraNo"].apply(safe_int_str) == u_id]
                    for _, h in bagli_hamlar.iterrows():
                        ham = etree.SubElement(u_el, "HamMadde")
                        etree.SubElement(ham, "ReferansFormTipi").text = clean(h.get("ReferansFormTipi")) or ""
                        etree.SubElement(ham, "ReferansFormNo").text = clean(h.get("ReferansFormNo")) or ""
                        etree.SubElement(ham, "ReferansFormYil").text = safe_int_str(h.get("ReferansFormYil")) or ""
                        gt_h = "".join(filter(str.isdigit, str(clean(h.get("gtip")) or "")))
                        etree.SubElement(ham, "gtip").text = gt_h.zfill(12)[:12]
                        etree.SubElement(ham, "Mensei").text = clean(h.get("Mensei")) or ""
                        etree.SubElement(ham, "BirinciMiktar").text = str(h.get("BirinciMiktar") or "0")
                        etree.SubElement(ham, "BirinciBirim").text = clean(h.get("BirinciBirim")) or ""

            xml_data = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            
            st.success("XML Dosyası Hazır!")
            st.download_button(
                label="✅ XML Dosyasını İndir",
                data=xml_data,
                file_name="UBF_Cikti.xml",
                mime="application/xml"
            )
        except Exception as e:
            st.error(f"Dönüştürme Hatası: {e}")
