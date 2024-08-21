import os
import cv2
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
from pdf2image import convert_from_path
import PyPDF2
import pandas as pd
import json
import pypdfium2 as pdfium
from kraken.binarization import nlbin
from PIL import Image
import random

# PDF klasör yolunu belirtin
pdf_folder = "Faturalar"
output_excel = "faturalar.xlsx"

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
    
    # Görseli kaydet ve ekranda göster
    cv2.imwrite(f"adaptive_binarized_result{random.randint(45, 500)}.png", binarized_image)
    # cv2.imshow("Adaptive Binarized Image", binarized_image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()
    
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
        return {
            "Tarih": qr_json.get("tarih", ""),
            "cari kodu": qr_json.get("vkntckn", ""),
            "FATURA NO": qr_json.get("no", ""),
            "TUTAR": qr_json.get("vergidahil", ""),
            # KDV VERGİ
            # İSKONTO
            # BELGE TÜRÜ FATURA  1: Z RAPORU 4
            # EVRAK TİPİ 1 ALIŞ  2 SATIŞ
            # TİCARET TÜRÜ  1 TOPTAN - 2 PERAKENDE
            # AÇIK KAPALI
            # KAPALI FAT CARİ TİPİ  1 CARİ 5 KASA
            # KAPALI FATURA CARİSİ---BOŞ OLACAK
            # FATURA CİNSİ 1-TOPTAN 2- PERAKENDE
            # KDV ORANI
            # NORMAL İADE
            # hareket cinsi 1 STOK 2- HİZMET 3 MASRAF
            # hareket kodu
            # hareket adı
            # BABS
            "Alıcı VKN/TCKN": qr_json.get("avkntckn", ""),
            "Senaryo": qr_json.get("senaryo", ""),
            "Tip": qr_json.get("tip", ""),
            "ETTN": qr_json.get("ettn", ""),
            "Para Birimi": qr_json.get("parabirimi", ""),
            "Mal/Hizmet Toplam": qr_json.get("malhizmettoplam", ""),
            "KDV Matrah (20%)": qr_json.get("kdvmatrah(20)", ""),
            "Hesaplanan KDV (20%)": qr_json.get("hesaplanankdv(20)", ""),
            "Ödenecek Tutar": qr_json.get("odenecek", "")
        }
    except json.JSONDecodeError:
        print(f"QR kod ayrıştırılamadı: {qr_code}")
        return {}

# Klasördeki PDF dosyalarını gezme
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
all_data = []

for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    print(pdf_path)
    qr_data_list = extract_qr_from_pdf(pdf_path)
    
    for qr_code in qr_data_list:
        parsed_data = process_qr_data(qr_code)
        if parsed_data:
            all_data.append(parsed_data)


# QR kod verilerini terminale yazdır
print(qr_data_list)
if qr_data_list:
    print("QR Kod Verileri:")
    for qr_data in qr_data_list:
        print(qr_data)
else:
    print("QR kod içeren görüntü bulunamadı.")

# Veriyi Excel'e yazma
df = pd.DataFrame(all_data)
df.to_excel(output_excel, index=False)

# print(f"Veriler {output_excel} dosyasına başarıyla yazıldı.")
