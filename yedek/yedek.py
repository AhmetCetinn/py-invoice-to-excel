import os
import cv2
import numpy as np
import pyzbar.pyzbar as pyzbar
import pyzbar.pyzbar as pyzbar
from pdf2image import convert_from_path
import PyPDF2
import pandas as pd
import json
import pypdfium2 as pdfium

# PDF klasör yolunu belirtin
pdf_folder = "TEST"
output_excel = "output.xlsx"


def find_qr_code_in_image(image):
    # Görüntüyü gri tona çevirme
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Otsu yöntemi ile görüntüyü eşikleme
    _, thresh_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Konturları bulma
    contours, _ = cv2.findContours(thresh_image, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    # Potansiyel QR kod bölgelerini bulma
    for contour in contours:
        # Konturun çevresinde bir dikdörtgen çizer
        x, y, w, h = cv2.boundingRect(contour)
        
        # Dikdörtgenin en boy oranını kontrol et (QR kodlar genellikle kareye yakın olur)
        aspect_ratio = w / float(h)
        
        if 0.9 <= aspect_ratio <= 1.1:  # Kareye yakınsa
            roi = image[y:y+h, x:x+w]  # Bu bölgeyi çıkar
            
            # QR kodu bu bölgede arıyoruz
            qr_codes = pyzbar.decode(roi)
            if qr_codes:
                return qr_codes  # QR kod bulunduysa geri dönüyoruz
    
    return None  # QR kod bulunamadıysa

def extract_qr_from_pdf(pdf_path):
    # PDF'deki sayfaları görüntüye çeviriyoruz
    pages = convert_from_path(pdf_path, 400)  # DPI'yı artırarak netlik sağlıyoruz
    
    qr_data_list = []
    
    for page_num, page in enumerate(pages):
        # Sayfayı numpy dizisine dönüştür
        image = np.array(page)
        
        # QR kodu bulmaya çalış
        qr_codes = find_qr_code_in_image(image)
        
        if qr_codes:
            for qr_code in qr_codes:
                qr_data_list.append(qr_code.data.decode('utf-8'))
        else:
            print(f"QR kod bulunamadı: Sayfa {page_num + 1}")
    
    return qr_data_list


def process_qr_data(qr_code):
    # QR kod içeriği JSON formatında olduğu için json modülü ile ayrıştırıyoruz
    try:
        qr_json = json.loads(qr_code)
        return {
            "VKN/TCKN": qr_json.get("vkntckn", ""),
            "Alıcı VKN/TCKN": qr_json.get("avkntckn", ""),
            "Senaryo": qr_json.get("senaryo", ""),
            "Tip": qr_json.get("tip", ""),
            "Tarih": qr_json.get("tarih", ""),
            "Fatura No": qr_json.get("no", ""),
            "ETTN": qr_json.get("ettn", ""),
            "Para Birimi": qr_json.get("parabirimi", ""),
            "Mal/Hizmet Toplam": qr_json.get("malhizmettoplam", ""),
            "KDV Matrah (20%)": qr_json.get("kdvmatrah(20)", ""),
            "Hesaplanan KDV (20%)": qr_json.get("hesaplanankdv(20)", ""),
            "Vergi Dahil": qr_json.get("vergidahil", ""),
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

# # Veriyi Excel'e yazma
# df = pd.DataFrame(all_data)
# df.to_excel(output_excel, index=False)

# print(f"Veriler {output_excel} dosyasına başarıyla yazıldı.")
