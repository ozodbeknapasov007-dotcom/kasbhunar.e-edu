import json
import os
import zipfile
import xml.sax.saxutils as saxutils

def generate_all_certificates():
    template_path = os.path.join('qrcodes', 'Abduxakimova Serfitikat namuna.docx')
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Shablon fayli topilmadi: {template_path}")

    with open('students.json', 'r', encoding='utf-8') as f:
        students = json.load(f)

    out_dir = 'sertifikatlar'
    os.makedirs(out_dir, exist_ok=True)

    # Read base template zip entries into memory
    with zipfile.ZipFile(template_path, 'r') as z_in:
        template_entries = {item.filename: z_in.read(item.filename) for item in z_in.infolist()}

    base_doc_xml = template_entries['word/document.xml'].decode('utf-8')

    print(f"Jami {len(students)} ta talaba uchun Word sertifikatlari yaratilmoqda...")

    generated_files = []

    for s in students:
        idx = s['id']
        fio = s['fio']
        reg_no = s['qayd_raqami']
        familiya = s['familiya']
        ismi = s['ismi']
        mutaxassislik = s.get('mutaxassislik', 'Farmatsiya')

        # Find corresponding QR code image
        safe_qr_name = f"{idx:02d}_{familiya}_{ismi}_{reg_no}.png".replace("'", "").replace(" ", "_")
        qr_path = os.path.join('qrcodes', safe_qr_name)

        if not os.path.exists(qr_path):
            raise FileNotFoundError(f"Talaba uchun QR kod topilmadi: {qr_path}")

        with open(qr_path, 'rb') as f_qr:
            qr_bytes = f_qr.read()

        # Clone template files
        file_map = dict(template_entries)

        # Update XML
        doc_xml = base_doc_xml
        # 1. Replace FIO
        doc_xml = doc_xml.replace('ABDIXAKIMOVA FAYOZA KAMOL QIZI', saxutils.escape(fio))

        # 2. Replace registration / certificate number (both places: cert no and reg no)
        doc_xml = doc_xml.replace('0000001', reg_no)

        # 3. Replace specialty if Logopediya
        if mutaxassislik == 'Logopediya':
            doc_xml = doc_xml.replace('FARMATSIYA', 'LOGOPEDIYA')
            doc_xml = doc_xml.replace('PHARMACY', 'LOGOPEDICS')

        file_map['word/document.xml'] = doc_xml.encode('utf-8')
        file_map['word/media/image3.png'] = qr_bytes

        # Output filename
        docx_name = f"{idx:02d}_{familiya}_{ismi}_{reg_no}_sertifikat.docx".replace("'", "").replace(" ", "_")
        docx_path = os.path.join(out_dir, docx_name)

        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for fname, fdata in file_map.items():
                z_out.writestr(fname, fdata)

        generated_files.append((docx_path, docx_name))
        print(f"[{idx:02d}/{len(students)}] Tayyor: {docx_name}")

    # Create ZIP archive containing all docx certificates
    zip_name = "SERTIFIKATLAR_BARCHA_TALABALAR.zip"
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for full_path, arc_name in generated_files:
            zipf.write(full_path, arc_name)

    print("\n" + "="*60)
    print(f"Barcha {len(students)} ta sertifikat muvaffaqiyatli yaratildi!")
    print(f"Sertifikatlar papkasi: '{out_dir}/'")
    print(f"Barcha sertifikatlar arxivi: '{zip_name}'")
    print("="*60)

if __name__ == '__main__':
    generate_all_certificates()
