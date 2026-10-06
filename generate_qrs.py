import json
import os
import argparse
import qrcode
from PIL import Image
import openpyxl

def generate_all(base_domain="https://ilmziyo-cert.vercel.app"):
    with open('students.json', 'r', encoding='utf-8') as f:
        students = json.load(f)

    # Ensure base_domain has no trailing slash
    base_domain = base_domain.rstrip('/')

    out_dir = "qrcodes"
    os.makedirs(out_dir, exist_ok=True)

    # Create summary workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QR Havolalar"
    ws.append([
        "T/R", 
        "F.I.O", 
        "Sertifikat raqami", 
        "Qayd raqami", 
        "JShShIR (PINFL)", 
        "Tug'ilgan sana", 
        "Maxfiy QR Havola (Havolani o'zgartirib bo'lmaydi)", 
        "QR Fayl nomi"
    ])

    print(f"Baza domen: {base_domain}")
    print(f"QR kodlar yaratilmoqda...")

    for s in students:
        idx = s['id']
        token = s['token']
        fio = s['fio']
        cert_no = s['sertifikat_raqami']
        reg_no = s['qayd_raqami']
        pinfl = s['pinfl']
        bdate = s['tugilgan_sana']
        
        # Secret unguessable URL
        qr_url = f"{base_domain}/prof/pt/{token}"
        
        # Safe filename
        safe_name = f"{idx:02d}_{s['familiya']}_{s['ismi']}_{reg_no}.png".replace("'", "").replace(" ", "_")
        qr_path = os.path.join(out_dir, safe_name)

        # Generate QR code
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=12, # ~500x500px, high quality for printing
            border=2
        )
        qr.add_data(qr_url)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        img.save(qr_path)

        ws.append([
            idx,
            fio,
            cert_no,
            reg_no,
            pinfl,
            bdate,
            qr_url,
            safe_name
        ])

    excel_report = "TALABALAR_QR_HAVOLALARI.xlsx"
    wb.save(excel_report)
    print(f"[OK] 30 ta QR kod '{out_dir}/' papkasiga saqlandi!")
    print(f"[OK] Barcha havolalar '{excel_report}' jadvaliga saqlandi!")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Farmatsiya talabalari uchun xavfsiz QR kodlar generatori")
    parser.add_argument("--domain", type=str, default="https://ilmziyo-cert.vercel.app", help="Saytingizning domeni (masalan: https://ilm-ziyo.uz)")
    args = parser.parse_args()

    generate_all(args.domain)
