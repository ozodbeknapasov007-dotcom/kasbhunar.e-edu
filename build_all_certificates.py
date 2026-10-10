import json
import os
import zipfile
import xml.sax.saxutils as saxutils
import win32com.client
import pymupdf

def main():
    template_path = os.path.join('qrcodes', 'Abduxakimova Serfitikat namuna.docx')
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Shablon fayli topilmadi: {template_path}")

    with open('students.json', 'r', encoding='utf-8') as f:
        students = json.load(f)

    docx_dir = 'sertifikatlar'
    pdf_dir = 'sertifikatlar_pdf'
    os.makedirs(docx_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)

    # 1. Read base template
    with zipfile.ZipFile(template_path, 'r') as z_in:
        template_entries = {item.filename: z_in.read(item.filename) for item in z_in.infolist()}

    base_doc_xml = template_entries['word/document.xml'].decode('utf-8')

    # Apply global text fixes to template
    # Fix 1: S E R T I F I KA T -> S E R T I F I K A T
    assert 'S E R T I F I KA T' in base_doc_xml, "Title not found"
    base_doc_xml = base_doc_xml.replace('S E R T I F I KA T', 'S E R T I F I K A T')

    # Fix 2: Director -> Direktor (first occurrence in Director line)
    assert base_doc_xml.count('<w:t>Director</w:t>') == 2, "Expected 2 occurrences of Director"
    base_doc_xml = base_doc_xml.replace('<w:t>Director</w:t>', '<w:t>Direktor</w:t>', 1)

    # Fix 3: keying -> keyingi in Izoh
    assert 'keying turini' in base_doc_xml, "keying turini not found"
    base_doc_xml = base_doc_xml.replace('keying turini', 'keyingi turini')

    print(f"1-qadam: {len(students)} ta talaba uchun to'g'rilangan Word (.docx) sertifikatlari yaratilmoqda...")

    docx_files = []

    for s in students:
        idx = s['id']
        fio = s['fio']
        reg_no = s['qayd_raqami']
        familiya = s['familiya']
        ismi = s['ismi']
        mutaxassislik = s.get('mutaxassislik', 'Farmatsiya')

        safe_qr_name = f"{idx:02d}_{familiya}_{ismi}_{reg_no}.png".replace("'", "").replace(" ", "_")
        qr_path = os.path.join('qrcodes', safe_qr_name)
        if not os.path.exists(qr_path):
            raise FileNotFoundError(f"QR kod topilmadi: {qr_path}")

        with open(qr_path, 'rb') as f_qr:
            qr_bytes = f_qr.read()

        file_map = dict(template_entries)
        doc_xml = base_doc_xml

        # Replace student details
        doc_xml = doc_xml.replace('ABDIXAKIMOVA FAYOZA KAMOL QIZI', saxutils.escape(fio))
        doc_xml = doc_xml.replace('0000001', reg_no)

        if mutaxassislik == 'Logopediya':
            doc_xml = doc_xml.replace('FARMATSIYA', 'LOGOPEDIYA')
            doc_xml = doc_xml.replace('PHARMACY', 'LOGOPEDICS')

        file_map['word/document.xml'] = doc_xml.encode('utf-8')
        file_map['word/media/image3.png'] = qr_bytes

        docx_name = f"{idx:02d}_{familiya}_{ismi}_{reg_no}_sertifikat.docx".replace("'", "").replace(" ", "_")
        docx_path = os.path.join(docx_dir, docx_name)

        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for fname, fdata in file_map.items():
                z_out.writestr(fname, fdata)

        docx_files.append((os.path.abspath(docx_path), docx_name, idx, fio))

    print(f"  [OK] Barcha {len(students)} ta individual Word (.docx) fayllari muvaffaqiyatli saqlandi!")

    # 2. Convert each DOCX to PDF using Word COM & merge DOCX into one
    print("\n2-qadam: Word va PDF hujjatlari yaratilmoqda (Microsoft Word COM orqali)...")
    merged_docx_path = os.path.abspath('BARCHA_31_TALABA_SERTIFIKATLARI.docx')
    merged_pdf_path = os.path.abspath('BARCHA_31_TALABA_SERTIFIKATLARI.pdf')

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False

    pdf_files = []

    try:
        # Convert individual DOCX to PDF
        for full_docx_path, docx_name, idx, fio in docx_files:
            pdf_name = docx_name.replace('.docx', '.pdf')
            full_pdf_path = os.path.abspath(os.path.join(pdf_dir, pdf_name))
            
            doc = word.Documents.Open(full_docx_path)
            doc.SaveAs2(full_pdf_path, FileFormat=17) # 17 = wdFormatPDF
            doc.Close(False)
            
            pdf_files.append((full_pdf_path, pdf_name))

        print(f"  [OK] Barcha {len(pdf_files)} ta individual PDF fayli yaratildi!")

        # Create single combined DOCX
        print("\n3-qadam: Barcha 31 ta sertifikat bitta Word faylga birlashtirilmoqda...")
        merged_doc = word.Documents.Open(docx_files[0][0])
        for full_docx_path, docx_name, idx, fio in docx_files[1:]:
            sel = word.Selection
            sel.EndKey(Unit=6) # wdStory
            sel.InsertBreak(Type=2) # wdSectionBreakNextPage
            sel.InsertFile(FileName=full_docx_path)

        while merged_doc.Paragraphs.Count > 0 and len(merged_doc.Paragraphs.Last.Range.Text.strip()) == 0:
            merged_doc.Paragraphs.Last.Range.Delete()

        merged_doc.SaveAs2(merged_docx_path)
        merged_doc.Close(False)
        print(f"  [OK] Birlashgan Word fayli tayyor: BARCHA_31_TALABA_SERTIFIKATLARI.docx")

    finally:
        word.Quit()

    # 4. Merge individual PDFs into single consolidated PDF
    print("\n4-qadam: Barcha PDF'lar bitta yagona PDF faylga birlashtirilmoqda...")
    merged = pymupdf.open()
    for full_pdf_path, pdf_name in pdf_files:
        with pymupdf.open(full_pdf_path) as sub:
            merged.insert_pdf(sub)

    merged.save(merged_pdf_path)
    merged.close()

    with pymupdf.open(merged_pdf_path) as check:
        pdf_pages = len(check)

    print(f"  [OK] Birlashgan PDF tayyor: BARCHA_31_TALABA_SERTIFIKATLARI.pdf ({pdf_pages} sahifa)")

    # 5. Create ZIP archives
    print("\n5-qadam: ZIP arxivlari yaratilmoqda...")
    # DOCX ZIP
    zip_docx_name = "SERTIFIKATLAR_BARCHA_TALABALAR.zip"
    with zipfile.ZipFile(zip_docx_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for full_path, arc_name, _, _ in docx_files:
            zipf.write(full_path, arc_name)
        zipf.write(merged_docx_path, os.path.basename(merged_docx_path))

    # PDF ZIP
    zip_pdf_name = "SERTIFIKATLAR_PDF_BARCHA_TALABALAR.zip"
    with zipfile.ZipFile(zip_pdf_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for full_path, arc_name in pdf_files:
            zipf.write(full_path, arc_name)
        zipf.write(merged_pdf_path, os.path.basename(merged_pdf_path))

    print("\n" + "="*70)
    print("BARCHA ISHLAR MUVAFFAQIYATLI YAKUNLANDI!")
    print(f"- 31 ta alohida Word sertifikatlar: '{docx_dir}/'")
    print(f"- 31 ta alohida PDF sertifikatlar: '{pdf_dir}/'")
    print(f"- Bitta yig'ilgan Word fayli: 'BARCHA_31_TALABA_SERTIFIKATLARI.docx'")
    print(f"- Bitta yig'ilgan PDF fayli: 'BARCHA_31_TALABA_SERTIFIKATLARI.pdf' ({pdf_pages} sahifa)")
    print(f"- Word arxiv: '{zip_docx_name}'")
    print(f"- PDF arxiv: '{zip_pdf_name}'")
    print("="*70)

if __name__ == '__main__':
    main()
