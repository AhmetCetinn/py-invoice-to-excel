import os
import cv2
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
from pdf2image import convert_from_path
import pandas as pd
import json
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, StringVar
from datetime import datetime,timedelta

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

def run_application(input_folder, output_folder, output_file):
    if not input_folder or not os.path.exists(input_folder):
        messagebox.showerror("Hata", "Geçerli bir PDF klasörü seçilmedi.")
        return
    if not output_folder or not os.path.exists(output_folder):
        messagebox.showerror("Hata", "Geçerli bir çıktı klasörü seçilmedi.")
        return

    output_excel = os.path.join(output_folder, output_file)
    print(f"PDF Klasörü: {input_folder}")
    print(f"Çıkış Dosyası: {output_excel}")

    process_pdfs(input_folder, output_folder, output_excel)
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

    def on_run():
        input_folder = input_folder_path.get()
        output_folder = output_folder_path.get()
        output_file = output_file_name.get()
        run_application(input_folder, output_folder, output_file)

    Button(root, text="Çalıştır", command=on_run).grid(row=3, column=1, padx=10, pady=20)

    root.mainloop()

def adaptive_binarization(image):
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    binarized_image = cv2.adaptiveThreshold(
        gray_image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2
    )
    return binarized_image

def detect_qr_with_adaptive(image):
    binarized_image = adaptive_binarization(image)
    qr_codes = decode(binarized_image, symbols=[ZBarSymbol.QRCODE])
    return qr_codes

def extract_qr_from_pdf(pdf_path, output_folder):
    pages = convert_from_path(pdf_path, 400)
    qr_data_list = []
    error_log_path = os.path.join(output_folder, "error-log.txt")

    for page_num, page in enumerate(pages):
        image = np.array(page)
        qr_codes = detect_qr_with_adaptive(image)

        if qr_codes:
            for qr_code in qr_codes:
                qr_data_list.append(qr_code.data.decode('utf-8'))
                print(f"QR Kod Bulundu: Sayfa {page_num + 1}, Veri: {qr_code.data.decode('utf-8')}")
        else:
            with open(error_log_path, 'a') as log_file:
                log_file.write(f"{os.path.basename(pdf_path)} - Sayfa {page_num + 1}: QR kod bulunamadı.\n")
            print(f"QR kod bulunamadı: Sayfa {page_num + 1}")

    return qr_data_list

def process_qr_data(qr_code):
    try:
        qr_json = json.loads(qr_code)
    except json.JSONDecodeError:
        print(f"QR kod ayrıştırılamadı: {qr_code}")
        return None, None, None

    return qr_json.get("tarih", ""), qr_json.get("no", ""), qr_json.get("vkntckn", "")

def process_pdfs(pdf_folder, output_folder, output_excel):
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
    all_data = []
    success_folder = create_output_folder(output_folder)

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(f"İşlenen dosya: {pdf_path}")

        qr_data_list = extract_qr_from_pdf(pdf_path, output_folder)

        for qr_code in qr_data_list:
            tarih, fatura_no, cari_kodu = process_qr_data(qr_code)
            if tarih and fatura_no and cari_kodu:
                all_data.append({
                    "Tarih": tarih,
                    "Fatura No": fatura_no,
                    "Cari Kodu": cari_kodu
                })

                image_file = f"{fatura_no}-{cari_kodu}.png"
                image_path = os.path.join(success_folder, image_file)
                pages = convert_from_path(pdf_path, 400)
                if pages:
                    pages[0].save(image_path, "PNG")
            else:
                print(f"{pdf_file} dosyasındaki QR kodu işlenemedi.")

    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(output_excel, index=False)
        print(f"Veriler {output_excel} dosyasına başarıyla yazıldı.")
    else:
        print("Hiç QR kod verisi bulunamadı.")

# Uygulamayı başlat
start_gui()
