import json
import os

def build():
    with open('students.json', 'r', encoding='utf-8') as f:
        students = json.load(f)

    os.makedirs('dist', exist_ok=True)

    # 1. 404.html
    html_404 = """<!doctype html>
<html lang="uz">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hujjat topilmadi</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { background-color: #f8f9fa; display: flex; align-items: center; justify-content: center; height: 100vh; font-family: system-ui, -apple-system, sans-serif; }
    .card { max-width: 500px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.06); }
  </style>
</head>
<body>
  <div class="container text-center">
    <div class="card p-4 mx-auto">
      <div class="text-danger mb-3">
        <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" fill="currentColor" viewBox="0 0 16 16">
          <path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z"/>
          <path d="M7.002 11a1 1 0 1 1 2 0 1 1 0 0 1-2 0zM7.1 4.995a.905.905 0 1 1 1.8 0l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 4.995z"/>
        </svg>
      </div>
      <h3 class="fw-bold text-secondary">404 - Hujjat topilmadi</h3>
      <p class="text-muted mt-2">Ushbu havola bo'yicha hech qanday sertifikat yoki diplom ma'lumoti mavjud emas. Iltimos, QR kodni qaytadan to'g'ri skanerlang.</p>
    </div>
  </div>
</body>
</html>"""
    with open('dist/404.html', 'w', encoding='utf-8') as f:
        f.write(html_404)

    # 2. index.html (Root protection)
    html_index = """<!doctype html>
<html lang="uz">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Ta'lim hujjatlarini tekshirish portali</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body { background-color: #f8f9fa; display: flex; align-items: center; justify-content: center; height: 100vh; font-family: system-ui, -apple-system, sans-serif; }
    .card { max-width: 520px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); }
  </style>
</head>
<body>
  <div class="container text-center">
    <div class="card p-4 mx-auto">
      <div class="text-primary mb-3">
        <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" fill="currentColor" viewBox="0 0 16 16">
          <path d="M0 .5A.5.5 0 0 1 .5 0h4a.5.5 0 0 1 .5.5v4a.5.5 0 0 1-.5.5h-4A.5.5 0 0 1 0 4.5v-4zm1 1v2h2v-2H1zm6 0v2h2v-2H7zm6 0v2h2v-2h-2zM0 6.5a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 .5.5v4a.5.5 0 0 1-.5.5h-4a.5.5 0 0 1-.5-.5v-4zm1 1v2h2v-2H1zm9.5-1a.5.5 0 0 0-.5.5v2a.5.5 0 0 0 .5.5h2a.5.5 0 0 0 .5-.5v-2a.5.5 0 0 0-.5-.5h-2zm-6 6a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 .5.5v4a.5.5 0 0 1-.5.5h-4a.5.5 0 0 1-.5-.5v-4zm1 1v2h2v-2h-2zm7.5-.5a.5.5 0 0 0-.5.5v2a.5.5 0 0 0 .5.5h2a.5.5 0 0 0 .5-.5v-2a.5.5 0 0 0-.5-.5h-2z"/>
        </svg>
      </div>
      <h4 class="fw-bold text-dark">Ta'lim hujjatlarini tekshirish tizimi</h4>
      <p class="text-muted mt-2">Sertifikat ma'lumotlarini tasdiqlash uchun rasmiy hujjatdagi QR kodni mobil qurilma orqali skaner qiling.</p>
      <div class="badge bg-secondary p-2 mt-2">Faqat QR kod orqali tekshiriladi</div>
    </div>
  </div>
</body>
</html>"""
    with open('dist/index.html', 'w', encoding='utf-8') as f:
        f.write(html_index)

    # 3. robots.txt
    with open('dist/robots.txt', 'w', encoding='utf-8') as f:
        f.write("User-agent: *\nDisallow: /\n")

    # 4. Generate each student's page
    for s in students:
        token = s['token']
        folder = os.path.join('dist', 'prof', 'pt', token)
        os.makedirs(folder, exist_ok=True)
        
        fam = s['familiya'].title()
        ism = s['ismi'].title()
        title_name = f"{fam} {ism}ning sertifikat ma'lumoti"
        
        html_student = f"""<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{title_name}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-KK94CHFLLe+nY2dmCWGMq91rCGa5gtU4mk92HdvYe+M/SXH301p5ILy+dN9+nJOZ" crossorigin="anonymous">
  </head>
  <body>
    <div class="container my-5">
        <h3 class="text-center text-uppercase mt-4 mb-3">Bitiruvchining shaxsiy ma'lumotlari</h3>
        <table class="table table-striped">
          <tbody>
            <tr>
              <td class="fw-bold" width="25%">Familiya:</td>
              <td colspan="4">{s['familiya']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Ismi:</td>
              <td colspan="4">{s['ismi']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Otasining ismi:</td>
              <td colspan="4">{s['sharifi']}</td>
            </tr>
            <tr>
              <td class="fw-bold">JShShIR(PINFL):</td>
              <td colspan="4">{s['pinfl']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Tug'ilgan sana:</td>
              <td colspan="4">{s['tugilgan_sana']}</td>
            </tr>
          </tbody>
        </table>
        <h3 class="text-center text-uppercase mt-4 mb-3">Bitiruvchining ta'lim ma'lumotlari</h3>
        <table class="table table-striped">
          <tbody>
            <tr>
              <td class="fw-bold" width="25%">TM:</td>
              <td colspan="2">{s['tm']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Mutaxassislik:</td>
              <td colspan="2">{s['mutaxassislik']}</td>
            </tr>
            <tr>
              <td class="fw-bold">O'quv soati:</td>
              <td colspan="2">{s['oqish_soati']}</td>
            </tr>
            <tr>
              <td class="fw-bold">O'qish davri:</td>
              <td colspan="2">{s['oqish_davri']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Sertifikat seriya va raqami:</td>
              <td colspan="2">{s['sertifikat_raqami']}</td>
            </tr>
            <tr>
              <td class="fw-bold">Qayd raqami:</td>
              <td colspan="2">{s['qayd_raqami']}</td>
            </tr>
            <tr>
              <td class="fw-bold">DAK qarori:</td>
              <td colspan="2">{s['dak_qarori']}</td>
            </tr>
            <tr>
              <td class="fw-bold">DAK sanasi:</td>
              <td colspan="2">{s['dak_sanasi']}</td>
            </tr>
          </tbody>
        </table>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ENjdO4Dr2bkBIFxQpeoTz1HIcje39Wm4jDKdf19U8gI4ddQ3GYNS7NTKfAdVQSZe" crossorigin="anonymous"></script>
  </body>
</html>"""
        with open(os.path.join(folder, 'index.html'), 'w', encoding='utf-8') as f:
            f.write(html_student)

    print(f"Barcha {len(students)} ta talaba sahifasi muvaffaqiyatli 'dist/' papkasiga yaratildi!")

if __name__ == '__main__':
    build()
