import os
import numpy as np
from pdf2image import convert_from_path
import pandas as pd
import json
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, StringVar, ttk
from datetime import datetime, timedelta
import qreader
from PIL import Image
import cv2  # OpenCV ekleniyor
import demjson3


def select_pdf_folder():
    folder_selected = filedialog.askdirectory(title="PDF Klasörünü Seçin")
    return folder_selected


def select_output_folder():
    folder_selected = filedialog.askdirectory(title="Çıktı Klasörünü Seçin")
    return folder_selected


def create_output_folder(base_folder):
    now = (datetime.utcnow() + timedelta(hours=3)).strftime('%Y-%m-%d_%H.%M')
    folder_name = f"{now}-Fatura"
    full_output_path = os.path.join(base_folder, folder_name)
    if not os.path.exists(full_output_path):
        os.makedirs(full_output_path)
    return full_output_path

def run_application(input_folder, output_folder, output_file, progress_bar):
    if not input_folder or not os.path.exists(input_folder):
        messagebox.showerror("Hata", "Geçerli bir PDF klasörü seçilmedi.")
        return
    if not output_folder or not os.path.exists(output_folder):
        messagebox.showerror("Hata", "Geçerli bir çıktı klasörü seçilmedi.")
        return

    output_excel = os.path.join(output_folder, output_file)
    print(f"PDF Klasörü: {input_folder}")
    print(f"Çıkış Dosyası: {output_excel}")

    process_pdfs(input_folder, output_folder, output_excel, progress_bar)
    messagebox.showinfo("Başarılı", f"İşlem tamamlandı. Çıktı dosyası '{output_excel}' olarak kaydedildi.")

def start_gui():
    root = Tk()
    root.title("QR İşleme Uygulaması")
    input_folder_path = StringVar()
    output_folder_path = StringVar(value=os.getcwd())

    def choose_input_folder():
        folder = select_pdf_folder()
        if folder:
            input_folder_path.set(folder)

    def choose_output_folder():
        folder = select_output_folder()
        if folder:
            output_folder_path.set(folder)

    Label(root, text="PDF Klasörü Seç").grid(row=0, column=0, padx=10, pady=10)
    Entry(root, textvariable=input_folder_path, width=50).grid(row=0, column=1, padx=10, pady=10)
    Button(root, text="Gözat", command=choose_input_folder).grid(row=0, column=2, padx=10, pady=10)

    Label(root, text="Çıkış Klasörü Seç").grid(row=1, column=0, padx=10, pady=10)
    Entry(root, textvariable=output_folder_path, width=50).grid(row=1, column=1, padx=10, pady=10)
    Button(root, text="Gözat", command=choose_output_folder).grid(row=1, column=2, padx=10, pady=10)

    Label(root, text="Çıkış Dosya Adı (Excel)").grid(row=2, column=0, padx=10, pady=10)
    output_file_name = Entry(root, width=50)
    output_file_name.insert(0, "output.xlsx")
    output_file_name.grid(row=2, column=1, padx=10, pady=10)

    # İlerleme çubuğunu ekliyoruz
    progress = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
    progress.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

    def on_run():
        input_folder = input_folder_path.get()
        output_folder = output_folder_path.get()
        output_file = output_file_name.get()
        run_application(input_folder, output_folder, output_file, progress)

    Button(root, text="Çalıştır", command=on_run).grid(row=4, column=1, padx=10, pady=20)

    root.mainloop()


def extract_qr_from_pdf(pdf_path, output_folder):
    pages = convert_from_path(pdf_path, 800)
    qr_data_list = []
    error_log_path = os.path.join(output_folder, "error-log.txt")

    # Başarılı ve başarısız QR kodlar için klasörler oluşturma
    success_folder = os.path.join(output_folder, "success_qr")
    failed_folder = os.path.join(output_folder, "failed_qr")
    os.makedirs(success_folder, exist_ok=True)
    os.makedirs(failed_folder, exist_ok=True)

    reader = qreader.QReader()

    for page_num, page in enumerate(pages):
        # Görüntüyü numpy dizisine dönüştür
        image = np.array(page)
        image_bgr = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

        # QR kod tespiti
        detection_result = reader.detect(image_bgr)
        if detection_result and len(detection_result) > 0:
            first_detection = detection_result[0]
            decoded_result = reader.decode(image_bgr, first_detection)
            if decoded_result:
                qr_data_list.append(decoded_result)

                # QR kod etrafına yeşil çerçeve çizme
                if 'bounding_box' in first_detection:
                    points = np.array(first_detection['bounding_box']).reshape((-1, 1, 2)).astype(int)
                    cv2.polylines(image_bgr, [points], isClosed=True, color=(0, 255, 0), thickness=3)

                output_image_path = os.path.join(success_folder, f"{os.path.basename(pdf_path).replace('.pdf', '')}_page_{page_num + 1}.png")
                cv2.imwrite(output_image_path, image_bgr)
                print(f"Başarılı QR resmi kaydedildi: {output_image_path}")
            else:
                log_failed_qr(pdf_path, page_num, error_log_path, "QR kod çözümleme başarısız")
        else:
            log_failed_qr(pdf_path, page_num, error_log_path, "QR kod bulunamadı")

    return qr_data_list


def process_qr_data(qr_code, pdf_path, error_log_path):
    rows = []

    try:
        # Eğer qr_code bir liste ise, öncelikle listeyi açıp ilk öğeyi almalıyız
        if isinstance(qr_code, list) and len(qr_code) > 0:
            qr_code = qr_code[0]

        # demjson3 ile JSON'u ayrıştır ve hataları tolere et
        qr_json = demjson3.decode(qr_code, strict=False)

    except demjson3.JSONDecodeError as e:
        with open(error_log_path, 'a') as log_file:
            log_file.write(f"Hatalı JSON ayrıştırması: {pdf_path}, Hata: {e}\n")
        print(f"QR kod ayrıştırılamadı: {qr_code}")
        return []

    except Exception as e:
        with open(error_log_path, 'a') as log_file:
            log_file.write(f"Diğer hata: {pdf_path}, Hata: {e}\n")
        print(f"Bir hata oluştu: {e}")
        return []

    # Genel bilgi alanları
    general_info = {
        "Tarih": qr_json.get("tarih", ""),
        "cari kodu": qr_json.get("vkntckn", ""),
        "cari adı": "",
        "YAZMA BELGE NO": "",
        "FATURA NO": qr_json.get("no", ""),
        "İSKONTO": qr_json.get("iskonto", ""),
        "BELGE TÜRÜ FATURA  1: Z RAPORU 4": "",
        "EVRAK TİPİ 1 ALIŞ  2 SATIŞ": "1",
        "TİCARET TÜRÜ  1 TOPTAN - 2 PERAKENDE": "1",
        "AÇIK KAPALI": "AÇIK",
        "KAPALI FAT CARİ TİPİ  1 CARİ 5 KASA": "1",
        "KAPALI FATURA CARİSİ---BOŞ OLACAK": "",
        "FATURA CİNSİ 1-TOPTAN 2- PERAKENDE": "1",
        "KDV ORANI": "",
        "NORMAL İADE": "NORMAL",
        "hareket cinsi 1 STOK 2- HİZMET 3 MASRAF": "1",
        "hareket kodu": "",
        "hareket adı": ""
    }

    # KDV oranlarını işleyelim, sadece geçerli veriler işlenip Excel'e kaydedilsin
    for i in [1, 10, 20]:
        kdvmatrah_key_float = f"kdvmatrah({i:.2f})"
        kdvmatrah_key_int = f"kdvmatrah({i})"
        hesaplanankdv_key_float = f"hesaplanankdv({i:.2f})"
        hesaplanankdv_key_int = f"hesaplanankdv({i})"

        kdv_matrah_values = qr_json.get(kdvmatrah_key_float) or qr_json.get(kdvmatrah_key_int)
        hesaplanan_kdv_values = qr_json.get(hesaplanankdv_key_float) or qr_json.get(hesaplanankdv_key_int)

        if kdv_matrah_values and hesaplanan_kdv_values:
            if not isinstance(kdv_matrah_values, list):
                kdv_matrah_values = [kdv_matrah_values]
            if not isinstance(hesaplanan_kdv_values, list):
                hesaplanan_kdv_values = [hesaplanan_kdv_values]

            for matrah, kdv in zip(kdv_matrah_values, hesaplanan_kdv_values):
                if isinstance(matrah, (int, float, str)) and isinstance(kdv, (int, float, str)):
                    row = general_info.copy()
                    row["TUTAR"] = matrah
                    row["KDV VERGİ"] = kdv
                    row["KDV ORANI"] = f"{i}%"
                    rows.append(row)
                else:
                    with open(error_log_path, 'a') as log_file:
                        log_file.write(f"Hatalı KDV verisi: {pdf_path}, Matrah: {matrah}, KDV: {kdv}\n")

    return rows


def log_failed_qr(pdf_path, page_num, error_log_path, error_message):
    with open(error_log_path, 'a') as log_file:
        if page_num != -1:
            log_file.write(f"{os.path.basename(pdf_path)} - Sayfa {page_num + 1}: {error_message}\n")
        else:
            log_file.write(f"{os.path.basename(pdf_path)}: {error_message}\n")
    print(f"{error_message}: {os.path.basename(pdf_path)}, Sayfa {page_num + 1}")


def process_pdfs(pdf_folder, output_folder, output_excel, progress_bar):
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
    total_files = len(pdf_files)
    all_data = []
    success_folder = create_output_folder(output_folder)
    error_log_path = os.path.join(output_folder, "error-log.txt")

    # Set up the progress bar
    progress_bar['maximum'] = total_files
    progress_bar['value'] = 0
    progress_bar.update()

    for index, pdf_file in enumerate(pdf_files):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(f"İşlenen dosya: {pdf_path}")

        qr_data_list = extract_qr_from_pdf(pdf_path, output_folder)

        for qr_code in qr_data_list:
            parsed_rows = process_qr_data(qr_code, pdf_path, error_log_path)
            if parsed_rows:
                all_data.extend(parsed_rows)

        # Update the progress bar for each file
        progress_bar['value'] = index + 1
        progress_bar.update()

    # Save to Excel
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(output_excel, index=False)
        print(f"Veriler {output_excel} dosyasına başarıyla yazıldı.")
    else:
        print("Hiç QR kod verisi bulunamadı.")

# Start the application
start_gui()