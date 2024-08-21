import os
import cv2
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
from pdf2image import convert_from_path
import pandas as pd
import json
from tkinter import Tk, messagebox, filedialog

def select_pdf_folder_with_message():
    # Tkinter GUI başlatılır
    root = Tk()
    root.withdraw()  # Ana pencereyi gizle
    
    # Yönlendirme mesajı göster
    messagebox.showinfo("Bilgi", "Lütfen PDF dosyalarını içeren klasörü seçin.")
    
    # Dosya seçici penceresi açılır
    folder_selected = filedialog.askdirectory(title="PDF Klasörünü Seçin")
    
    return folder_selected

def show_error_message(message):
    # Hata mesajını bir popup ile göster
    messagebox.showerror("Hata", message)

# PDF klasör yolunu seçici ile belirtin ve mesaj gösterin
pdf_folder = select_pdf_folder_with_message()

# Eğer klasör seçilmediyse hata mesajı göster
if not pdf_folder:
    show_error_message("Hiçbir klasör seçilmedi. Program sonlandırılıyor.")
    exit()

# Geçerli bir klasör olup olmadığını kontrol et
if not os.path.exists(pdf_folder):
    show_error_message(f"{pdf_folder} geçerli bir klasör değil. Program sonlandırılıyor.")
    exit()

print(f"Seçilen klasör: {pdf_folder}")

# PDF klasör yolunu belirtin
# pdf_folder = "Faturalar"
output_excel = "output.xlsx"

# Geçici çözüm
np.float = float

def adaptive_binarization(image):
    # Görüntüyü gri tonlamaya dönüştür
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Adaptive Thresholding ile binarize et (local bölgeye göre ayarlanır)
    binarized_image = cv2.adaptiveThreshold(
        gray_image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
        cv2.THRESH_BINARY, 11, 2
    )
    
    return binarized_image

def detect_qr_with_adaptive(image):
    # Adaptive thresholding ile binarize işlemi yap
    binarized_image = adaptive_binarization(image)
    
    # Binarize edilmiş görüntüde QR kodlarını arıyoruz
    qr_codes = decode(binarized_image, symbols=[ZBarSymbol.QRCODE])
    
    return qr_codes

def extract_qr_from_pdf(pdf_path):
    # PDF'deki sayfaları görüntüye çeviriyoruz
    pages = convert_from_path(pdf_path, 400)  # DPI'yı artırarak netlik sağlıyoruz
    
    qr_data_list = []
    
    for page_num, page in enumerate(pages):
        # Sayfayı numpy dizisine dönüştür
        image = np.array(page)
        
        # QR kodu tespit et ve sonucu yazdır
        qr_codes = detect_qr_with_adaptive(image)
        
        if qr_codes:
            for qr_code in qr_codes:
                qr_data_list.append(qr_code.data.decode('utf-8'))
                print(f"QR Kod Bulundu: Sayfa {page_num + 1}, Veri: {qr_code.data.decode('utf-8')}")
        else:
            print(f"QR kod bulunamadı: Sayfa {page_num + 1}")
    
    return qr_data_list

def process_qr_data(qr_code):
    # QR kod içeriği JSON formatında olduğu için json modülü ile ayrıştırıyoruz
    try:
        qr_json = json.loads(qr_code)
    except json.JSONDecodeError:
        print(f"QR kod ayrıştırılamadı: {qr_code}")
        return []
    
    # Excel'e kaydetmek için satırları hazırlıyoruz
    rows = []

    # Diğer bilgiler
    general_info = {
        "VKN/TCKN": qr_json.get("vkntckn", ""),
        "Alıcı VKN/TCKN": qr_json.get("avkntckn", ""),
        "Senaryo": qr_json.get("senaryo", ""),
        "Tip": qr_json.get("tip", ""),
        "Tarih": qr_json.get("tarih", ""),
        "Fatura No": qr_json.get("no", ""),
        "ETTN": qr_json.get("ettn", ""),
        "Para Birimi": qr_json.get("parabirimi", ""),
        "Mal/Hizmet Toplam": qr_json.get("malhizmettoplam", ""),
        "Vergi Dahil": qr_json.get("vergidahil", ""),
        "Ödenecek Tutar": qr_json.get("odenecek", "")
    }

    # KDV matrahlarını ve hesaplanan KDV'leri satırlara ayırıyoruz
    for i in [1, 10, 20]:
        kdvmatrah_key = f"kdvmatrah({i})"
        hesaplanankdv_key = f"hesaplanankdv({i})"
        
        if kdvmatrah_key in qr_json and hesaplanankdv_key in qr_json:
            row = general_info.copy()  # Diğer bilgiler kopyalanıyor
            row["KDV Oran"] = f"{i}%"
            row["KDV Matrah"] = qr_json[kdvmatrah_key]
            row["Hesaplanan KDV"] = qr_json[hesaplanankdv_key]
            rows.append(row)
    
    return rows

# Klasördeki PDF dosyalarını gezme
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
all_data = []

for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    print(f"İşlenen dosya: {pdf_path}")
    
    # PDF'den QR verilerini çıkar
    qr_data_list = extract_qr_from_pdf(pdf_path)
    
    # QR kodları işleyip KDV bilgilerini ayrıştır
    for qr_code in qr_data_list:
        parsed_rows = process_qr_data(qr_code)
        if parsed_rows:
            all_data.extend(parsed_rows)

# Eğer QR kod verisi bulunmuşsa veriyi Excel'e yaz
if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(output_excel, index=False)
    print(f"Veriler {output_excel} dosyasına başarıyla yazıldı.")
else:
    print("Hiç QR kod verisi bulunamadı.")


# "Tarih": qr_json.get("tarih", ""),
#             "cari kodu": qr_json.get("vkntckn", ""),
#             "FATURA NO": qr_json.get("no", ""),
#             "TUTAR": qr_json.get("vergidahil", ""),
#             # KDV VERGİ
#             # İSKONTO
#             # BELGE TÜRÜ FATURA  1: Z RAPORU 4
#             # EVRAK TİPİ 1 ALIŞ  2 SATIŞ
#             # TİCARET TÜRÜ  1 TOPTAN - 2 PERAKENDE
#             # AÇIK KAPALI
#             # KAPALI FAT CARİ TİPİ  1 CARİ 5 KASA
#             # KAPALI FATURA CARİSİ---BOŞ OLACAK
#             # FATURA CİNSİ 1-TOPTAN 2- PERAKENDE
#             # KDV ORANI
#             # NORMAL İADE
#             # hareket cinsi 1 STOK 2- HİZMET 3 MASRAF
#             # hareket kodu
#             # hareket adı
#             # BABS
#             "Alıcı VKN/TCKN": qr_json.get("avkntckn", ""),
#             "Senaryo": qr_json.get("senaryo", ""),
#             "Tip": qr_json.get("tip", ""),
#             "ETTN": qr_json.get("ettn", ""),
#             "Para Birimi": qr_json.get("parabirimi", ""),
#             "Mal/Hizmet Toplam": qr_json.get("malhizmettoplam", ""),
#             "KDV Matrah (20%)": qr_json.get("kdvmatrah(20)", ""),
#             "Hesaplanan KDV (20%)": qr_json.get("hesaplanankdv(20)", ""),
#             "Ödenecek Tutar": qr_json.get("odenecek", "")
