import streamlit as st
import streamlit.components.v1 as components

# Sahifa sozlamalari
st.set_page_config(
    page_title="O'lchov Birliklari Ta'lim Platformasi",
    page_icon="📐",
    layout="wide"
)

# Streamlit bo'sh joylarini (padding) olib tashlash
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 0rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True
)

# To'liq HTML, CSS va JS kodi
html_code = """
<!DOCTYPE html>
<html lang="uz">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>O'lchov Birliklari Ta'lim Platformasi | Quljanova Shukurjon Zaribovna</title>
  <!-- Excel (.xlsx) kutubxonasi -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --bg-white: #ffffff;
      --blue: #1e88e5;
      --dark-blue: #0d47a1;
      --green: #2e7d32;
      --light-green: #e8f5e9;
      --red: #d32f2f;
      --light-red: #ffebee;
      --light-blue: #e3f2fd;
      --text-dark: #212121;
      --border: #cfd8dc;
      --excel-header: #107c41;
      --shadow: 0 4px 18px rgba(0, 0, 0, 0.08);
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    body {
      background-color: var(--bg-white);
      color: var(--text-dark);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }

    header {
      background: var(--bg-white);
      border-bottom: 3px solid var(--blue);
      box-shadow: 0 2px 10px rgba(0,0,0,0.05);
      padding: 14px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }

    .brand-title {
      font-size: 21px;
      font-weight: 800;
      color: var(--blue);
    }

    .author-badge {
      background: var(--light-blue);
      color: var(--dark-blue);
      border: 1.5px solid var(--blue);
      padding: 5px 14px;
      border-radius: 20px;
      font-size: 14px;
      font-weight: 700;
    }

    .top-user-bar {
      background: #f1f8fe;
      border-bottom: 2px solid #bbdefb;
      padding: 10px 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 15px;
    }

    .top-user-inputs {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }

    .top-user-inputs input {
      padding: 8px 14px;
      border: 1.5px solid var(--border);
      border-radius: 6px;
      font-size: 14px;
      outline: none;
      min-width: 260px;
    }

    .top-user-inputs input:focus {
      border-color: var(--blue);
    }

    main {
      flex: 1;
      max-width: 1200px;
      width: 100%;
      margin: 20px auto;
      padding: 0 16px;
    }

    .step-section {
      display: none;
      background: var(--bg-white);
      border: 2px solid var(--border);
      border-radius: 14px;
      padding: 26px;
      box-shadow: var(--shadow);
    }

    .step-section.active {
      display: block;
      animation: fadeIn 0.3s ease-in;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .section-title {
      font-size: 22px;
      font-weight: 800;
      margin-bottom: 18px;
    }

    .title-blue { color: var(--blue); }
    .title-green { color: var(--green); }
    .title-red { color: var(--red); }

    .form-group {
      margin-bottom: 20px;
    }

    .form-group label {
      display: block;
      font-weight: 700;
      margin-bottom: 8px;
      font-size: 15px;
    }

    .form-control {
      width: 100%;
      max-width: 450px;
      padding: 12px 16px;
      font-size: 16px;
      border: 2px solid var(--border);
      border-radius: 8px;
      outline: none;
    }

    .form-control:focus {
      border-color: var(--blue);
    }

    .btn {
      padding: 10px 22px;
      font-size: 15px;
      font-weight: 700;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: 0.2s;
    }

    .btn-blue { background: var(--blue); color: #fff; }
    .btn-blue:hover { background: var(--dark-blue); }

    .btn-green { background: var(--green); color: #fff; }
    .btn-green:hover { background: #1b5e20; }

    .btn-excel { background: var(--excel-header); color: #fff; }
    .btn-excel:hover { background: #0c5b30; }

    .btn-red { background: var(--red); color: #fff; }
    .btn-red:hover { background: #b71c1c; }

    .btn-outline {
      background: transparent;
      border: 2px solid var(--blue);
      color: var(--blue);
    }
    .btn-outline:hover { background: var(--light-blue); }

    .grid-chapters {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 18px;
      margin-bottom: 25px;
    }

    .chapter-card {
      border: 2px solid;
      border-radius: 12px;
      overflow: hidden;
      background: #fff;
    }

    .chapter-header {
      padding: 12px;
      font-weight: 800;
      text-align: center;
      color: #fff;
      font-size: 15px;
    }

    .card-blue { border-color: var(--blue); }
    .card-blue .chapter-header { background: var(--blue); }

    .card-pink { border-color: #d81b60; }
    .card-pink .chapter-header { background: #d81b60; }

    .card-green { border-color: var(--green); }
    .card-green .chapter-header { background: var(--green); }

    .card-amber { border-color: #ef6c00; }
    .card-amber .chapter-header { background: #ef6c00; }

    .chapter-body {
      padding: 14px 18px;
      line-height: 1.8;
      font-size: 15px;
    }

    .chapter-body ul { list-style: none; }
    .chapter-body li { border-bottom: 1px dashed #eceff1; padding: 3px 0; }
    .chapter-body li:last-child { border-bottom: none; }
    .hl-red { color: var(--red); font-weight: 700; }

    .variant-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(115px, 1fr));
      gap: 10px;
      margin: 18px 0;
    }

    .btn-variant {
      padding: 10px;
      font-weight: 700;
      border: 2px solid var(--blue);
      border-radius: 6px;
      background: #fff;
      color: var(--blue);
      cursor: pointer;
      transition: all 0.2s ease;
    }

    .btn-variant.active {
      background: var(--blue);
      color: #fff;
    }

    .btn-variant.completed {
      background-color: var(--green) !important;
      border-color: var(--green) !important;
      color: #ffffff !important;
    }

    .practice-item {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 14px;
      font-size: 16px;
      font-weight: 600;
      flex-wrap: wrap;
    }

    .practice-input {
      width: 110px;
      padding: 6px 10px;
      border: 2px solid var(--border);
      border-radius: 6px;
      text-align: center;
      font-size: 16px;
    }

    .practice-status {
      font-size: 13px;
      font-weight: 700;
      padding: 3px 8px;
      border-radius: 4px;
    }

    .test-box {
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px;
      margin-bottom: 16px;
      background: #fbfbfb;
    }

    .test-question {
      font-size: 16px;
      font-weight: 700;
      margin-bottom: 12px;
      color: var(--dark-blue);
    }

    .test-options {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }

    .option-label {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 14px;
      border: 1.5px solid var(--border);
      border-radius: 6px;
      background: #fff;
      cursor: pointer;
    }

    .option-label:hover {
      border-color: var(--blue);
      background: var(--light-blue);
    }

    .result-summary-card {
      border: 2px solid var(--green);
      border-radius: 12px;
      padding: 20px;
      background: var(--light-green);
      margin-bottom: 20px;
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 15px;
      text-align: center;
    }

    .summary-item {
      background: #ffffff;
      padding: 12px;
      border-radius: 8px;
      border: 1px solid var(--border);
    }

    .summary-label { font-size: 13px; color: #555; margin-bottom: 4px; font-weight: 600; }
    .summary-val { font-size: 18px; font-weight: 800; color: var(--dark-blue); }

    .excel-table-wrapper {
      border: 2px solid #b0bec5;
      border-radius: 8px;
      overflow-x: auto;
      background: #ffffff;
      box-shadow: 0 2px 8px rgba(0,0,0,0.06);
      margin-top: 15px;
    }

    .table-excel {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
      text-align: left;
      white-space: nowrap;
    }

    .table-excel thead tr {
      background-color: var(--excel-header);
      color: #ffffff;
    }

    .table-excel th {
      padding: 12px 14px;
      font-weight: 700;
      border: 1px solid #0d5c31;
    }

    .table-excel td {
      padding: 10px 14px;
      border: 1px solid #d0d7de;
      color: #24292f;
    }

    .table-excel tbody tr:nth-child(even) {
      background-color: #f6f8fa;
    }

    .table-excel tbody tr:hover {
      background-color: #e8f4fc;
    }

    footer {
      background: #fafafa;
      border-top: 1px solid var(--border);
      text-align: center;
      padding: 14px;
      font-size: 14px;
      color: #666;
      margin-top: auto;
    }
  </style>
</head>
<body>

  <header>
    <div class="brand-title">📐 Matematika: O'lchov Birliklari</div>
    <div class="author-badge">Tuzuvchi: Quljanova Shukurjon Zaribovna</div>
  </header>

  <div class="top-user-bar" id="topUserBar" style="display: none;">
    <div style="font-weight: 700; color: var(--dark-blue);">
      Sinf: <span id="topBarClassDisplay" style="color: var(--green);">-</span>
    </div>
    <div class="top-user-inputs">
      <label for="topStudentName" style="font-weight: 700; font-size: 14px;">O'quvchi Ismi va Familiyasi:</label>
      <input type="text" id="topStudentName" placeholder="Ism va Familiyangizni yozing..." required>
      <button type="button" class="btn btn-blue" style="padding: 7px 16px; font-size: 14px;" onclick="updateStudentNameFromTop()">Saqlash</button>
    </div>
  </div>

  <main>
    <section id="step1" class="step-section active">
      <h2 class="section-title title-blue">🏫 Sinfni kiriting</h2>
      <p style="margin-bottom: 20px; color: #555;">Dasturni boshlash uchun o'z sinfingizni yozing:</p>

      <div class="form-group">
        <label for="inputClass">Sinfingiz (Masalan: 4-A sinf, 3-B sinf):</label>
        <input type="text" id="inputClass" class="form-control" placeholder="Masalan: 4-A sinf" required>
      </div>

      <button type="button" class="btn btn-blue" onclick="submitClassAndOpenApp()">Kirish ➔</button>
    </section>

    <section id="step2" class="step-section">
      <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; margin-bottom: 18px;">
        <h2 class="section-title title-blue">📘 O'lchov Birliklari Qoidalari</h2>
        <span id="studentDisplay" style="font-weight: 700; color: var(--green);"></span>
      </div>

      <div class="grid-chapters">
        <div class="chapter-card card-blue">
          <div class="chapter-header">UZUNLIK O'LCHOV BIRLIKLARI</div>
          <div class="chapter-body">
            <ul>
              <li>1 cm = <span class="hl-red">10</span> mm</li>
              <li>1 dm = <span class="hl-red">10</span> cm = <span class="hl-red">100</span> mm</li>
              <li>1 m = <span class="hl-red">10</span> dm</li>
              <li>1 m = <span class="hl-red">100</span> cm = <span class="hl-red">1 000</span> mm</li>
              <li>1 km = <span class="hl-red">1 000</span> m</li>
              <li>1 km = <span class="hl-red">10 000</span> dm</li>
              <li>1 km = <span class="hl-red">100 000</span> cm</li>
              <li>1 km = <span class="hl-red">1 000 000</span> mm</li>
            </ul>
          </div>
        </div>

        <div class="chapter-card card-pink">
          <div class="chapter-header">MASSA O'LCHOV BIRLIKLARI</div>
          <div class="chapter-body">
            <ul>
              <li>1 kg = <span class="hl-red">1 000</span> g</li>
              <li>1 q (sentner) = <span class="hl-red">100</span> kg</li>
              <li>1 q = <span class="hl-red">100 000</span> g</li>
              <li>1 t (tonna) = <span class="hl-red">10</span> q</li>
              <li>1 t = <span class="hl-red">1 000</span> kg</li>
              <li>1 t = <span class="hl-red">1 000 000</span> g</li>
            </ul>
          </div>
        </div>

        <div class="chapter-card card-green">
          <div class="chapter-header">YUZA O'LCHOV BIRLIKLARI</div>
          <div class="chapter-body">
            <ul>
              <li>1 cm² = <span class="hl-red">100</span> mm²</li>
              <li>1 dm² = <span class="hl-red">100</span> cm²</li>
              <li>1 dm² = <span class="hl-red">10 000</span> mm²</li>
              <li>1 m² = <span class="hl-red">100</span> dm²</li>
              <li>1 m² = <span class="hl-red">10 000</span> cm²</li>
              <li>1 m² = <span class="hl-red">1 000 000</span> mm²</li>
              <li>1 ar (sotix) = <span class="hl-red">100</span> m²</li>
              <li>1 ga (gektar) = <span class="hl-red">100</span> ar = <span class="hl-red">10 000</span> m²</li>
              <li>1 km² = <span class="hl-red">100</span> ga = <span class="hl-red">1 000 000</span> m²</li>
            </ul>
          </div>
        </div>

        <div class="chapter-card card-amber">
          <div class="chapter-header">VAQT O'LCHOV BIRLIKLARI</div>
          <div class="chapter-body">
            <ul>
              <li>1 min = <span class="hl-red">60</span> s</li>
              <li>1 h (soat) = <span class="hl-red">60</span> min = <span class="hl-red">3 600</span> s</li>
              <li>1 d (sutka) = <span class="hl-red">24</span> h</li>
              <li>1 hafta = <span class="hl-red">7</span> kun</li>
              <li>1 oy = <span class="hl-red">30</span> yoki <span class="hl-red">31</span> kun (fevral 28/29)</li>
              <li>1 yil = <span class="hl-red">12</span> oy = <span class="hl-red">365</span> yoki <span class="hl-red">366</span> kun</li>
              <li>1 asr = <span class="hl-red">100</span> yil</li>
            </ul>
          </div>
        </div>
      </div>

      <button type="button" class="btn btn-green" onclick="goToStep(3)">Bilimni Tekshirish Mashqlari (50 ta savol) ➔</button>
    </section>

    <section id="step3" class="step-section">
      <h2 class="section-title title-green">✍️ Bilimni Tekshirish: Jami 50 ta savol (10 ta Variant)</h2>
      <p style="margin-bottom: 12px; color: #555;">Amaliy tekshirish variantini tanlang (har bir variantda 5 tadan savol):</p>

      <div class="variant-grid" id="practiceVariantButtons"></div>

      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; flex-wrap: wrap; gap: 10px;">
        <h3 id="currentPracticeVariantHeading" style="color: var(--green);">1-Variant Amaliy Savollari</h3>
        <span id="practiceGradeBadge" style="font-weight: 800; font-size: 15px; color: var(--dark-blue);"></span>
      </div>
      
      <div id="practiceQuestionsContainer" style="max-width: 650px;"></div>

      <div style="margin-top: 20px; display: flex; gap: 10px; flex-wrap: wrap;">
        <button type="button" class="btn btn-blue" onclick="checkCurrentPractice()">Javoblarni Tekshirish va Baholash</button>
        <button type="button" class="btn btn-green" onclick="goToStep(4)">Yakuniy Testga O'tish ➔</button>
        <button type="button" class="btn btn-outline" onclick="goToStep(2)">⬅ Qoidalarga Qaytish</button>
      </div>
    </section>

    <section id="step4" class="step-section">
      <h2 class="section-title title-red">📝 Yakuniy Test Sinovi (10 ta variant, 100 ta savol)</h2>
      <p style="margin-bottom: 12px; color: #555;">Test variantini tanlang (ishlangan variantning ichi to'liq yashil bo'ladi):</p>

      <div class="variant-grid" id="testVariantButtons"></div>

      <h3 id="currentTestVariantHeading" style="color: var(--dark-blue); margin-bottom: 14px;">1-Variant (10 ta savol)</h3>
      <form id="quizForm">
        <div id="quizContainer"></div>

        <div style="margin-top: 20px; display: flex; gap: 12px; flex-wrap: wrap;">
          <button type="button" class="btn btn-red" onclick="finishQuiz()">Testni Yakunlash va Excelga Saqlash ➔</button>
          <button type="button" class="btn btn-outline" onclick="goToStep(3)">⬅ Mashqlarga Qaytish</button>
        </div>
      </form>
    </section>

    <section id="step5" class="step-section">
      <h2 class="section-title title-green">📊 Test Natijasi va Excel Jadvali</h2>

      <div class="result-summary-card">
        <div class="summary-item">
          <div class="summary-label">O'quvchi va Sinfi</div>
          <div class="summary-val" id="resStudentInfo">-</div>
        </div>
        <div class="summary-item">
          <div class="summary-label">Variant</div>
          <div class="summary-val" id="resVariant">-</div>
        </div>
        <div class="summary-item">
          <div class="summary-label">Bilimni Tekshirish Bahosi</div>
          <div class="summary-val" id="resPracticeGrade" style="color: var(--green);">-</div>
        </div>
        <div class="summary-item">
          <div class="summary-label">Yakuniy Test Bahosi</div>
          <div class="summary-val" id="resTestGrade" style="color: var(--blue);">-</div>
        </div>
      </div>

      <div style="display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px;">
        <button type="button" class="btn btn-excel" onclick="exportToExcel()">
          📥 Excel (.xlsx) Faylga Yuklash
        </button>
        <label class="btn btn-outline" style="cursor: pointer; margin: 0;">
          📤 Excel Faylni Ochish (Import)
          <input type="file" id="excelFileInput" accept=".xlsx, .xls" style="display: none;" onchange="importFromExcel(event)">
        </label>
        <button type="button" class="btn btn-green" onclick="goToStep(4)">🔄 Boshqa Variantni Ishlash</button>
        <button type="button" class="btn btn-outline" onclick="clearResults()">Jurnalni Tozalash</button>
      </div>

      <h3 style="color: var(--excel-header); margin-bottom: 10px;">📋 Tiniq Excel Qaydnomasi: Ikkala Baho Bir Qatorda</h3>
      
      <div class="excel-table-wrapper">
        <table class="table-excel" id="excelReportTable">
          <thead>
            <tr>
              <th>№</th>
              <th>Sinfi</th>
              <th>O'quvchining F.I.SH</th>
              <th>Variant</th>
              <th>Bilimni Tekshirish (To'g'ri / Ball)</th>
              <th>Bilimni Tekshirish Bahosi</th>
              <th>Yakuniy Test (To'g'ri / Ball)</th>
              <th>Yakuniy Test Bahosi</th>
              <th>Topshirilgan Vaqt</th>
            </tr>
          </thead>
          <tbody id="resultsTableBody"></tbody>
        </table>
      </div>
    </section>
  </main>

  <footer>
    <p>O'lchov birliklari ta'limiy dasturi | <b>Tuzuvchi: Quljanova Shukurjon Zaribovna</b></p>
  </footer>

  <script>
    const practiceBank = {
      1: [
        { text: "5 metr", unit: "santimetr (cm)", ans: 500 },
        { text: "3 tonna", unit: "kilogramm (kg)", ans: 3000 },
        { text: "2 sutka", unit: "soat (h)", ans: 48 },
        { text: "4 ar (sotix)", unit: "kvadrat metr (m²)", ans: 400 },
        { text: "2 asr", unit: "yil", ans: 200 }
      ],
      2: [
        { text: "7 detsimetr", unit: "santimetr (cm)", ans: 70 },
        { text: "4 sentner", unit: "kilogramm (kg)", ans: 400 },
        { text: "3 soat", unit: "minut", ans: 180 },
        { text: "2 gektar", unit: "ar (sotix)", ans: 200 },
        { text: "1 sutka", unit: "soat (h)", ans: 24 }
      ],
      3: [
        { text: "6 santimetr", unit: "millimetr (mm)", ans: 60 },
        { text: "5000 gramm", unit: "kilogramm (kg)", ans: 5 },
        { text: "2 minut", unit: "sekund (s)", ans: 120 },
        { text: "3 kvadrat metr", unit: "dm²", ans: 300 },
        { text: "3 yil", unit: "oy", ans: 36 }
      ],
      4: [
        { text: "8 kilometr", unit: "metr (m)", ans: 8000 },
        { text: "2 tonna", unit: "sentner (q)", ans: 20 },
        { text: "4 sutka", unit: "soat (h)", ans: 96 },
        { text: "1 kvadrat detsimetr", unit: "cm²", ans: 100 },
        { text: "4 asr", unit: "yil", ans: 400 }
      ],
      5: [
        { text: "9 metr", unit: "detsimetr (dm)", ans: 90 },
        { text: "6 kilogramm", unit: "gramm (g)", ans: 6000 },
        { text: "300 sekund", unit: "minut", ans: 5 },
        { text: "1 gektar", unit: "kvadrat metr (m²)", ans: 10000 },
        { text: "2 hafta", unit: "kun", ans: 14 }
      ],
      6: [
        { text: "40 millimetr", unit: "santimetr (cm)", ans: 4 },
        { text: "800 sentner", unit: "tonna (t)", ans: 80 },
        { text: "4 soat", unit: "minut", ans: 240 },
        { text: "2 ar (sotix)", unit: "kvadrat metr (m²)", ans: 200 },
        { text: "5 asr", unit: "yil", ans: 500 }
      ],
      7: [
        { text: "300 santimetr", unit: "metr (m)", ans: 3 },
        { text: "7 tonna", unit: "kilogramm (kg)", ans: 7000 },
        { text: "5 sutka", unit: "soat (h)", ans: 120 },
        { text: "4 kvadrat santimetr", unit: "mm²", ans: 400 },
        { text: "1 kvadrat kilometr", unit: "gektar (ga)", ans: 100 }
      ],
      8: [
        { text: "2 kilometr", unit: "detsimetr (dm)", ans: 20000 },
        { text: "9 sentner", unit: "kilogramm (kg)", ans: 900 },
        { text: "180 minut", unit: "soat (h)", ans: 3 },
        { text: "5 gektar", unit: "ar (sotix)", ans: 500 },
        { text: "3 hafta", unit: "kun", ans: 21 }
      ],
      9: [
        { text: "15 detsimetr", unit: "santimetr (cm)", ans: 150 },
        { text: "8000 gramm", unit: "kilogramm (kg)", ans: 8 },
        { text: "6 minut", unit: "sekund (s)", ans: 360 },
        { text: "7 ar", unit: "kvadrat metr (m²)", ans: 700 },
        { text: "60 oy", unit: "yil", ans: 5 }
      ],
      10: [
        { text: "2 metr", unit: "millimetr (mm)", ans: 2000 },
        { text: "10 tonna", unit: "sentner (q)", ans: 100 },
        { text: "1 soat", unit: "sekund (s)", ans: 3600 },
        { text: "2 kvadrat metr", unit: "dm²", ans: 200 },
        { text: "10 asr", unit: "yil", ans: 1000 }
      ]
    };

    const testBank = {
      1: [
        { q: "1 metr necha santimetrga teng?", options: ["10 cm", "100 cm", "1 000 cm", "10 mm"], correct: 1 },
        { q: "1 kilogramm necha gramm bo'ladi?", options: ["100 g", "1 000 g", "10 000 g", "500 g"], correct: 1 },
        { q: "1 sutka (d) necha soatdan iborat?", options: ["12 h", "24 h", "48 h", "60 h"], correct: 1 },
        { q: "1 kvadrat santimetr (cm²) necha mm²?", options: ["10 mm²", "100 mm²", "1 000 mm²", "10 000 mm²"], correct: 1 },
        { q: "1 tonna necha sentnerga (q) teng?", options: ["10 q", "100 q", "1 000 q", "5 q"], correct: 0 },
        { q: "1 soatda necha minut bor?", options: ["30 min", "60 min", "100 min", "360 s"], correct: 1 },
        { q: "1 ar (sotix) necha kvadrat metr?", options: ["10 m²", "100 m²", "1 000 m²", "10 000 m²"], correct: 1 },
        { q: "1 kilometr necha metr bo'ladi?", options: ["100 m", "500 m", "1 000 m", "10 000 m"], correct: 2 },
        { q: "1 asr necha yilga teng?", options: ["10 yil", "50 yil", "100 yil", "1 000 yil"], correct: 2 },
        { q: "1 detsimetr (dm) necha santimetr?", options: ["10 cm", "100 cm", "1 000 cm", "1 mm"], correct: 0 }
      ],
      2: [
        { q: "5 metr necha detsimetr bo'ladi?", options: ["50 dm", "500 dm", "5 dm", "25 dm"], correct: 0 },
        { q: "2 sentner (q) necha kilogramm?", options: ["20 kg", "200 kg", "2 000 kg", "50 kg"], correct: 1 },
        { q: "1 gektar (ga) necha ar (sotix)?", options: ["10 ar", "100 ar", "1 000 ar", "10 000 ar"], correct: 1 },
        { q: "2 soat necha minutga teng?", options: ["60 min", "90 min", "120 min", "180 min"], correct: 2 },
        { q: "30 millimetr necha santimetr bo'ladi?", options: ["3 cm", "30 cm", "300 cm", "0.3 cm"], correct: 0 },
        { q: "1 kvadrat detsimetr (dm²) necha cm²?", options: ["10 cm²", "100 cm²", "1 000 cm²", "10 000 cm²"], correct: 1 },
        { q: "4 kilogramm necha gramm bo'ladi?", options: ["40 g", "400 g", "4 000 g", "40 000 g"], correct: 2 },
        { q: "1 hafta necha kundan iborat?", options: ["5 kun", "6 kun", "7 kun", "8 kun"], correct: 2 },
        { q: "1 kvadrat metr (m²) necha dm²?", options: ["10 dm²", "100 dm²", "1 000 dm²", "10 000 dm²"], correct: 1 },
        { q: "2000 metr necha kilometr bo'ladi?", options: ["1 km", "2 km", "20 km", "200 km"], correct: 1 }
      ],
      3: [
        { q: "1 tonna necha kilogramm?", options: ["100 kg", "500 kg", "1 000 kg", "10 000 kg"], correct: 2 },
        { q: "3 minut necha sekund bo'ladi?", options: ["90 s", "120 s", "180 s", "300 s"], correct: 2 },
        { q: "400 santimetr necha metrga teng?", options: ["4 m", "40 m", "400 m", "4 000 m"], correct: 0 },
        { q: "1 gektar necha kvadrat metr (m²)?", options: ["100 m²", "1 000 m²", "10 000 m²", "100 000 m²"], correct: 2 },
        { q: "1 yil necha oydan iborat?", options: ["10 oy", "11 oy", "12 oy", "14 oy"], correct: 2 },
        { q: "1 sentner necha gramm bo'ladi?", options: ["1 000 g", "10 000 g", "100 000 g", "1 000 000 g"], correct: 2 },
        { q: "1 kvadrat kilometr (km²) necha gektar?", options: ["10 ga", "100 ga", "1 000 ga", "10 000 ga"], correct: 1 },
        { q: "70 detsimetr necha metr bo'ladi?", options: ["7 m", "70 m", "700 m", "0.7 m"], correct: 0 },
        { q: "Kabisa yilida fevral oyi necha kun bo'ladi?", options: ["28 kun", "29 kun", "30 kun", "31 kun"], correct: 1 },
        { q: "5 santimetr necha millimetr?", options: ["5 mm", "50 mm", "500 mm", "5000 mm"], correct: 1 }
      ],
      4: [
        { q: "3 sutka necha soat bo'ladi?", options: ["48 h", "60 h", "72 h", "96 h"], correct: 2 },
        { q: "1 kvadrat metr necha kvadrat santimetr (cm²)?", options: ["100 cm²", "1 000 cm²", "10 000 cm²", "100 000 cm²"], correct: 2 },
        { q: "5000 gramm necha kilogramm?", options: ["5 kg", "50 kg", "500 kg", "0.5 kg"], correct: 0 },
        { q: "6 metr necha santimetr bo'ladi?", options: ["60 cm", "600 cm", "6 000 cm", "60 000 cm"], correct: 1 },
        { q: "2 asr necha yil bo'ladi?", options: ["100 yil", "150 yil", "200 yil", "2 000 yil"], correct: 2 },
        { q: "3 ar necha kvadrat metr?", options: ["30 m²", "300 m²", "3 000 m²", "30 000 m²"], correct: 1 },
        { q: "1 kilometr necha santimetrga teng?", options: ["10 000 cm", "100 000 cm", "1 000 000 cm", "1 000 cm"], correct: 1 },
        { q: "4 sentner necha kilogramm bo'ladi?", options: ["40 kg", "400 kg", "4 000 kg", "40 000 kg"], correct: 1 },
        { q: "300 sekund necha minut bo'ladi?", options: ["3 min", "4 min", "5 min", "6 min"], correct: 2 },
        { q: "8 detsimetr necha millimetr?", options: ["80 mm", "800 mm", "8 000 mm", "8 mm"], correct: 1 }
      ],
      5: [
        { q: "1 tonna necha gramm bo'ladi?", options: ["100 000 g", "1 000 000 g", "10 000 000 g", "10 000 g"], correct: 1 },
        { q: "1 soat necha sekundga teng?", options: ["60 s", "600 s", "3 600 s", "6 000 s"], correct: 2 },
        { q: "8000 metr necha kilometr bo'ladi?", options: ["8 km", "80 km", "800 km", "0.8 km"], correct: 0 },
        { q: "2 gektar necha ar bo'ladi?", options: ["20 ar", "200 ar", "2 000 ar", "20 000 ar"], correct: 1 },
        { q: "1 kvadrat detsimetr necha mm²?", options: ["100 mm²", "1 000 mm²", "10 000 mm²", "100 000 mm²"], correct: 2 },
        { q: "Oddiy yil necha kundan iborat bo'ladi?", options: ["364 kun", "365 kun", "366 kun", "360 kun"], correct: 1 },
        { q: "600 sentner necha tonna bo'ladi?", options: ["6 t", "60 t", "600 t", "6 000 t"], correct: 1 },
        { q: "9 metr necha detsimetr bo'ladi?", options: ["90 dm", "900 dm", "9 dm", "9 000 dm"], correct: 0 },
        { q: "120 minut necha soat bo'ladi?", options: ["1 h", "2 h", "3 h", "4 h"], correct: 1 },
        { q: "40 santimetr necha millimetr?", options: ["4 mm", "40 mm", "400 mm", "4 000 mm"], correct: 2 }
      ],
      6: [
        { q: "1 kilometr necha detsimetr bo'ladi?", options: ["100 dm", "1 000 dm", "10 000 dm", "100 000 dm"], correct: 2 },
        { q: "7 kilogramm necha gramm?", options: ["70 g", "700 g", "7 000 g", "70 000 g"], correct: 2 },
        { q: "4 sutka necha soat bo'ladi?", options: ["48 h", "72 h", "96 h", "100 h"], correct: 2 },
        { q: "5 kvadrat santimetr necha mm²?", options: ["50 mm²", "500 mm²", "5 000 mm²", "50 000 mm²"], correct: 1 },
        { q: "50 sentner necha tonna bo'ladi?", options: ["5 t", "50 t", "500 t", "0.5 t"], correct: 0 },
        { q: "10 ar necha kvadrat metr bo'ladi?", options: ["100 m²", "1 000 m²", "10 000 m²", "100 000 m²"], correct: 1 },
        { q: "3 asr necha yil bo'ladi?", options: ["30 yil", "300 yil", "3 000 yil", "30 000 yil"], correct: 1 },
        { q: "2 metr necha millimetrga teng?", options: ["20 mm", "200 mm", "2 000 mm", "20 000 mm"], correct: 2 },
        { q: "240 sekund necha minut bo'ladi?", options: ["2 min", "3 min", "4 min", "5 min"], correct: 2 },
        { q: "6 detsimetr necha santimetr?", options: ["6 cm", "60 cm", "600 cm", "6 000 cm"], correct: 1 }
      ],
      7: [
        { q: "1 kvadrat metr necha kvadrat millimetr (mm²)?", options: ["10 000 mm²", "100 mm²", "1 000 000 mm²", "1000 mm²"], correct: 2 },
        { q: "8 tonna necha sentner bo'ladi?", options: ["80 q", "800 q", "8 q", "8 000 q"], correct: 0 },
        { q: "1 sutkada necha minut bor?", options: ["720 min", "1 440 min", "2 880 min", "600 min"], correct: 1 },
        { q: "3000 metr necha kilometr bo'ladi?", options: ["3 km", "30 km", "300 km", "0.3 km"], correct: 0 },
        { q: "500 kvadrat detsimetr necha kvadrat metr?", options: ["5 m²", "50 m²", "500 m²", "0.5 m²"], correct: 0 },
        { q: "9 sentner necha kilogramm bo'ladi?", options: ["90 kg", "900 kg", "9 000 kg", "90 000 kg"], correct: 1 },
        { q: "2 hafta necha kun bo'ladi?", options: ["10 kun", "12 kun", "14 kun", "16 kun"], correct: 2 },
        { q: "1 kilometr necha millimetrga teng?", options: ["10 000 mm", "100 000 mm", "1 000 000 mm", "10 000 000 mm"], correct: 2 },
        { q: "4 soat necha minut bo'ladi?", options: ["120 min", "180 min", "240 min", "300 min"], correct: 2 },
        { q: "7 santimetr necha millimetr?", options: ["7 mm", "70 mm", "700 mm", "7 000 mm"], correct: 1 }
      ],
      8: [
        { q: "10 000 kilogramm necha tonna bo'ladi?", options: ["1 t", "10 t", "100 t", "1 000 t"], correct: 1 },
        { q: "1 kvadrat kilometr necha kvadrat metr?", options: ["10 000 m²", "100 000 m²", "1 000 000 m²", "10 000 000 m²"], correct: 2 },
        { q: "5 asr necha yil bo'ladi?", options: ["50 yil", "500 yil", "5 000 yil", "250 yil"], correct: 1 },
        { q: "700 santimetr necha metr bo'ladi?", options: ["7 m", "70 m", "700 m", "0.7 m"], correct: 0 },
        { q: "6 minut necha sekund bo'ladi?", options: ["180 s", "240 s", "360 s", "420 s"], correct: 2 },
        { q: "3 gektar necha kvadrat metr bo'ladi?", options: ["300 m²", "3 000 m²", "30 000 m²", "300 000 m²"], correct: 2 },
        { q: "2 sentner necha gramm bo'ladi?", options: ["2 000 g", "20 000 g", "200 000 g", "2 000 000 g"], correct: 2 },
        { q: "15 detsimetr necha santimetr?", options: ["15 cm", "150 cm", "1 500 cm", "1.5 cm"], correct: 1 },
        { q: "5 sutka necha soat bo'ladi?", options: ["100 h", "120 h", "144 h", "150 h"], correct: 1 },
        { q: "90 millimetr necha santimetr bo'ladi?", options: ["9 cm", "90 cm", "900 cm", "0.9 cm"], correct: 0 }
      ],
      9: [
        { q: "6 kilometr necha metr bo'ladi?", options: ["60 m", "600 m", "6 000 m", "60 000 m"], correct: 2 },
        { q: "8000 gramm necha kilogramm bo'ladi?", options: ["8 kg", "80 kg", "800 kg", "0.8 kg"], correct: 0 },
        { q: "2 kvadrat metr necha kvadrat detsimetr?", options: ["20 dm²", "200 dm²", "2 000 dm²", "20 000 dm²"], correct: 1 },
        { q: "Yarim sutka necha soat bo'ladi?", options: ["6 h", "12 h", "18 h", "24 h"], correct: 1 },
        { q: "7 ar necha kvadrat metr bo'ladi?", options: ["70 m²", "700 m²", "7 000 m²", "70 000 m²"], correct: 1 },
        { q: "4 tonna necha sentner bo'ladi?", options: ["40 q", "400 q", "4 q", "4 000 q"], correct: 0 },
        { q: "3 yil necha oydan iborat bo'ladi?", options: ["24 oy", "30 oy", "36 oy", "48 oy"], correct: 2 },
        { q: "5 metr necha millimetr bo'ladi?", options: ["50 mm", "500 mm", "5 000 mm", "50 000 mm"], correct: 2 },
        { q: "180 minut necha soat bo'ladi?", options: ["2 h", "3 h", "4 h", "5 h"], correct: 1 },
        { q: "20 detsimetr necha metr bo'ladi?", options: ["1 m", "2 m", "20 m", "200 m"], correct: 1 }
      ],
      10: [
        { q: "10 kvadrat santimetr necha mm² bo'ladi?", options: ["100 mm²", "1 000 mm²", "10 000 mm²", "100 000 mm²"], correct: 1 },
        { q: "3000 kilogramm necha tonna bo'ladi?", options: ["3 t", "30 t", "300 t", "0.3 t"], correct: 0 },
        { q: "Soatning to'rtdan bir qismi (1/4 h) necha minut?", options: ["10 min", "15 min", "20 min", "30 min"], correct: 1 },
        { q: "4 kilometr necha detsimetr bo'ladi?", options: ["400 dm", "4 000 dm", "40 000 dm", "400 000 dm"], correct: 2 },
        { q: "5 gektar necha ar bo'ladi?", options: ["50 ar", "500 ar", "5 000 ar", "50 000 ar"], correct: 1 },
        { q: "10 asr necha yil bo'ladi?", options: ["100 yil", "500 yil", "1 000 yil", "10 000 yil"], correct: 2 },
        { q: "6 sentner necha kilogramm bo'ladi?", options: ["60 kg", "600 kg", "6 000 kg", "60 000 kg"], correct: 1 },
        { q: "3 hafta necha kundan iborat bo'ladi?", options: ["14 kun", "18 kun", "21 kun", "24 kun"], correct: 2 },
        { q: "800 millimetr necha santimetr bo'ladi?", options: ["8 cm", "80 cm", "800 cm", "8 000 cm"], correct: 1 },
        { q: "1 kvadrat metr necha kvadrat detsimetr?", options: ["10 dm²", "100 dm²", "1 000 dm²", "10 000 dm²"], correct: 1 }
      ]
    };

    let currentUser = {
      className: "",
      name: "",
      practiceVariant: 1,
      testVariant: 1,
      practiceScore: 0,
      practiceGrade: "-",
      practiceCorrect: "-"
    };

    let completedPracticeVariants = new Set();
    let completedTestVariants = new Set();

    function goToStep(step) {
      document.querySelectorAll('.step-section').forEach(s => s.classList.remove('active'));
      const target = document.getElementById(`step${step}`);
      if (target) {
        target.classList.add('active');
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }
    }

    function submitClassAndOpenApp() {
      const cls = document.getElementById('inputClass').value.trim();
      if (!cls) {
        alert("Iltimos, avval sinfingizni yozing!");
        return;
      }
      currentUser.className = cls;

      document.getElementById('topUserBar').style.display = 'flex';
      document.getElementById('topBarClassDisplay').innerText = cls;

      createPracticeVariantButtons();
      renderPracticeQuestions(1);
      createTestVariantButtons();
      renderTestQuestions(1);
      renderTable();

      goToStep(2);
    }

    function updateStudentNameFromTop() {
      const name = document.getElementById('topStudentName').value.trim();
      if (!name) {
        alert("Iltimos, ism va familiyangizni yozing!");
        return;
      }
      currentUser.name = name;
      document.getElementById('studentDisplay').innerText = `O'quvchi: ${currentUser.name} (${currentUser.className})`;
      alert("O'quvchi ma'lumotlari saqlandi: " + currentUser.name);
    }

    function createPracticeVariantButtons() {
      const container = document.getElementById('practiceVariantButtons');
      container.innerHTML = "";
      for (let i = 1; i <= 10; i++) {
        const btn = document.createElement('button');
        btn.type = "button";
        btn.id = `prac-var-btn-${i}`;
        btn.className = `btn-variant ${i === 1 ? 'active' : ''} ${completedPracticeVariants.has(i) ? 'completed' : ''}`;
        btn.innerText = `${i}-Variant ${completedPracticeVariants.has(i) ? '✓' : ''}`;
        btn.onclick = () => {
          document.querySelectorAll('#practiceVariantButtons .btn-variant').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          currentUser.practiceVariant = i;
          renderPracticeQuestions(i);
        };
        container.appendChild(btn);
      }
    }

    function renderPracticeQuestions(v) {
      document.getElementById('currentPracticeVariantHeading').innerText = `${v}-Variant Amaliy Savollari (5 ta savol)`;
      document.getElementById('practiceGradeBadge').innerText = "";
      const container = document.getElementById('practiceQuestionsContainer');
      container.innerHTML = "";

      const list = practiceBank[v];
      list.forEach((item, idx) => {
        const row = document.createElement('div');
        row.className = 'practice-item';
        row.innerHTML = `
          <span>${idx + 1}) ${item.text} = </span>
          <input type="number" id="prac_input_${idx}" class="practice-input" placeholder="...">
          <span>${item.unit}</span>
          <span id="prac_status_${idx}" class="practice-status"></span>
        `;
        container.appendChild(row);
      });
    }

    function checkCurrentPractice() {
      const v = currentUser.practiceVariant;
      const list = practiceBank[v];
      let correct = 0;

      list.forEach((item, idx) => {
        const val = parseInt(document.getElementById(`prac_input_${idx}`).value);
        const st = document.getElementById(`prac_status_${idx}`);
        if (val === item.ans) {
          st.innerText = "✓ To'g'ri";
          st.style.color = "var(--green)";
          st.style.backgroundColor = "var(--light-green)";
          correct++;
        } else {
          st.innerText = "✗ Noto'g'ri";
          st.style.color = "var(--red)";
          st.style.backgroundColor = "var(--light-red)";
        }
      });

      let pGrade = "2";
      if (correct === 5) pGrade = "5 (A'lo)";
      else if (correct === 4) pGrade = "4 (Yaxshi)";
      else if (correct === 3) pGrade = "3 (Qoniqarli)";
      else pGrade = "2 (Qoniqarsiz)";

      const pScore = correct * 20;
      currentUser.practiceScore = pScore;
      currentUser.practiceGrade = pGrade;
      currentUser.practiceCorrect = `${correct} / 5 (${pScore} ball)`;

      document.getElementById('practiceGradeBadge').innerText = `Natija: ${correct}/5 to'g'ri | Baho: ${pGrade}`;

      completedPracticeVariants.add(v);
      const btn = document.getElementById(`prac-var-btn-${v}`);
      if (btn) {
        btn.classList.add('completed');
        btn.innerText = `${v}-Variant ✓`;
      }

      alert(`Bilimni tekshirish baholandi:\nTo'g'ri javoblar: ${correct}/5\nQo'yilgan Baho: ${pGrade}`);
    }

    function createTestVariantButtons() {
      const container = document.getElementById('testVariantButtons');
      container.innerHTML = "";
      for (let i = 1; i <= 10; i++) {
        const btn = document.createElement('button');
        btn.type = "button";
        btn.id = `test-var-btn-${i}`;
        btn.className = `btn-variant ${i === 1 ? 'active' : ''} ${completedTestVariants.has(i) ? 'completed' : ''}`;
        btn.innerText = `${i}-Variant ${completedTestVariants.has(i) ? '✓' : ''}`;
        btn.onclick = () => {
          document.querySelectorAll('#testVariantButtons .btn-variant').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          currentUser.testVariant = i;
          renderTestQuestions(i);
        };
        container.appendChild(btn);
      }
    }

    function renderTestQuestions(v) {
      document.getElementById('currentTestVariantHeading').innerText = `${v}-Variant (10 ta savol)`;
      const box = document.getElementById('quizContainer');
      box.innerHTML = "";

      const questions = testBank[v];
      questions.forEach((qItem, idx) => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'test-box';

        let opts = '';
        qItem.options.forEach((opt, oIdx) => {
          opts += `
            <label class="option-label">
              <input type="radio" name="q_${idx}" value="${oIdx}">
              <span>${String.fromCharCode(65 + oIdx)}) ${opt}</span>
            </label>
          `;
        });

        itemDiv.innerHTML = `
          <div class="test-question">${idx + 1}. ${qItem.q}</div>
          <div class="test-options">${opts}</div>
        `;
        box.appendChild(itemDiv);
      });
    }

    function finishQuiz() {
      if (!currentUser.name) {
        const topName = document.getElementById('topStudentName').value.trim();
        if (topName) {
          currentUser.name = topName;
        } else {
          alert("Iltimos, avval saytning eng yuqori qismidagi maydonga O'quvchining Ismi va Familiyasini yozing!");
          window.scrollTo({ top: 0, behavior: 'smooth' });
          document.getElementById('topStudentName').focus();
          return;
        }
      }

      const v = currentUser.testVariant;
      const questions = testBank[v];
      let correct = 0;
      let unanswered = 0;

      questions.forEach((qItem, idx) => {
        const checked = document.querySelector(`input[name="q_${idx}"]:checked`);
        if (!checked) {
          unanswered++;
        } else if (parseInt(checked.value) === qItem.correct) {
          correct++;
        }
      });

      if (unanswered > 0) {
        if (!confirm(`${unanswered} ta savol belgilanmadi. Testni baribir yakunlaysizmi?`)) {
          return;
        }
      }

      completedTestVariants.add(v);
      const btn = document.getElementById(`test-var-btn-${v}`);
      if (btn) {
        btn.classList.add('completed');
        btn.innerText = `${v}-Variant ✓`;
      }

      const score = correct * 10;
      let testGrade = "2";
      if (score >= 85) testGrade = "5 (A'lo)";
      else if (score >= 70) testGrade = "4 (Yaxshi)";
      else if (score >= 55) testGrade = "3 (Qoniqarli)";
      else testGrade = "2 (Qoniqarsiz)";

      const now = new Date();
      const dateStr = now.toLocaleDateString() + ' ' + now.toLocaleTimeString([], {hour: '2-digit', minute: '2-digit'});

      document.getElementById('resStudentInfo').innerText = `${currentUser.name} (${currentUser.className})`;
      document.getElementById('resVariant').innerText = `${v}-variant`;
      document.getElementById('resPracticeGrade').innerText = currentUser.practiceGrade !== "-" ? currentUser.practiceGrade : "Baholanmagan";
      document.getElementById('resTestGrade').innerText = `${testGrade} (${score} ball)`;

      saveResult({
        className: currentUser.className,
        name: currentUser.name,
        variant: `${v}-variant`,
        practiceResult: currentUser.practiceCorrect !== "-" ? currentUser.practiceCorrect : "Ishlanmagan",
        practiceGrade: currentUser.practiceGrade !== "-" ? currentUser.practiceGrade : "-",
        testResult: `${correct} / 10 (${score} ball)`,
        testGrade: testGrade,
        date: dateStr
      });

      goToStep(5);
    }

    function saveResult(entry) {
      let data = JSON.parse(localStorage.getItem('dual_graded_quiz_records') || '[]');
      data.unshift(entry);
      localStorage.setItem('dual_graded_quiz_records', JSON.stringify(data));
      renderTable();
    }

    function renderTable() {
      const tbody = document.getElementById('resultsTableBody');
      const data = JSON.parse(localStorage.getItem('dual_graded_quiz_records') || '[]');
      tbody.innerHTML = "";

      if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; color:#777; padding: 18px;">Hozircha natijalar qayd etilmagan</td></tr>`;
        return;
      }

      data.forEach((row, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${i + 1}</td>
          <td><b>${row.className}</b></td>
          <td>${row.name}</td>
          <td>${row.variant}</td>
          <td>${row.practiceResult}</td>
          <td><b style="color: var(--green);">${row.practiceGrade}</b></td>
          <td>${row.testResult}</td>
          <td><b style="color: var(--blue);">${row.testGrade}</b></td>
          <td style="color: #666;">${row.date}</td>
        `;
        tbody.appendChild(tr);
      });
    }

    function exportToExcel() {
      const data = JSON.parse(localStorage.getItem('dual_graded_quiz_records') || '[]');
      if (data.length === 0) {
        alert("Excelga yuklash uchun hali natijalar mavjud emas!");
        return;
      }

      const excelRows = data.map((item, index) => ({
        "№": index + 1,
        "Sinfi": item.className,
        "O'quvchining F.I.SH": item.name,
        "Variant": item.variant,
        "Bilimni Tekshirish (Ball)": item.practiceResult,
        "Bilimni Tekshirish Bahosi": item.practiceGrade,
        "Yakuniy Test (Ball)": item.testResult,
        "Yakuniy Test Bahosi": item.testGrade,
        "Topshirilgan Vaqt": item.date
      }));

      const worksheet = XLSX.utils.json_to_sheet(excelRows);
      worksheet['!cols'] = [
        { wch: 6 },
        { wch: 12 },
        { wch: 28 },
        { wch: 12 },
        { wch: 24 },
        { wch: 24 },
        { wch: 22 },
        { wch: 20 },
        { wch: 20 }
      ];

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "O'zlashtirish Qaydnomasi");

      const fileName = `O'lchov_Birliklari_${currentUser.className || 'Umumiy'}_Baho_Qaydnomasi.xlsx`;
      XLSX.writeFile(workbook, fileName);
    }

    function importFromExcel(event) {
      const file = event.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const importedJson = XLSX.utils.sheet_to_json(firstSheet);

          if (!importedJson || importedJson.length === 0) {
            alert("Yuklangan Excel faylida ma'lumot topilmadi!");
            return;
          }

          let localRecords = JSON.parse(localStorage.getItem('dual_graded_quiz_records') || '[]');
          
          importedJson.forEach(row => {
            localRecords.push({
              className: row["Sinfi"] || "-",
              name: row["O'quvchining F.I.SH"] || "-",
              variant: row["Variant"] || "-",
              practiceResult: row["Bilimni Tekshirish (Ball)"] || "-",
              practiceGrade: row["Bilimni Tekshirish Bahosi"] || "-",
              testResult: row["Yakuniy Test (Ball)"] || "-",
              testGrade: row["Yakuniy Test Bahosi"] || "-",
              date: row["Topshirilgan Vaqt"] || "-"
            });
          });

          localStorage.setItem('dual_graded_quiz_records', JSON.stringify(localRecords));
          renderTable();
          alert("Excel faylidagi ma'lumotlar jadvalga to'liq yuklandi!");
        } catch (err) {
          alert("Excel faylini o'qishda xatolik yuz berdi: " + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function clearResults() {
      if (confirm("Rostdan ham barcha natijalar jadvalini tozalamoqchimisiz?")) {
        localStorage.removeItem('dual_graded_quiz_records');
        renderTable();
      }
    }
  </script>
</body>
</html>
"""

# HTML komponentini to'liq hajmda ishga tushirish
components.html(html_code, height=1350, scrolling=True)
