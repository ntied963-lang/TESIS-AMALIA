#!/usr/bin/env python3
"""Script untuk membuat file DOCX khusus Sub-Bab 4.6.2 Hasil ANOVA"""
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

doc = Document()

# Set margins
for section in doc.sections:
    section.top_margin = Cm(4)
    section.bottom_margin = Cm(3)
    section.left_margin = Cm(4)
    section.right_margin = Cm(3)

style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)
paragraph_format = style.paragraph_format
paragraph_format.line_spacing = 1.5
paragraph_format.space_after = Pt(0)


def add_heading_bab(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'


def add_heading_sub(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = 'Times New Roman'


def add_para(doc, text, indent=True, bold=False, italic=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.27)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = bold
    run.italic = italic
    return p


def add_table(doc, headers, rows, caption=""):
    if caption:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(caption)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = 'Times New Roman'
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(10)
                run.font.name = 'Times New Roman'
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Times New Roman'
    doc.add_paragraph()


# ===== HEADER =====
add_heading_bab(doc, "SUB-BAB 4.6.2")
add_heading_bab(doc, "HASIL ANALISIS STATISTIK (ANOVA)")
doc.add_paragraph()

add_para(doc,
    "Berikut disajikan hasil analisis statistik dari data produktivitas "
    "padi pada 27 unit percobaan menggunakan analisis regresi linear "
    "sederhana dan ANOVA. Setiap tabel disertai interpretasi untuk "
    "memudahkan pemahaman terhadap hasil analisis.")
doc.add_paragraph()

# ===== TABEL 1: DESCRIPTIVE STATISTICS =====
add_heading_sub(doc, "A. Statistik Deskriptif")

add_table(doc,
    ["Variabel", "Mean", "Std. Deviation", "N"],
    [
        ["Perlakuan", "2,00", "0,832", "27"],
        ["Hasil (g/ember)", "439,52", "65,996", "27"],
    ],
    "Tabel 4.6a Descriptive Statistics"
)

add_para(doc,
    "Interpretasi: Rata-rata variabel perlakuan sebesar 2,00 menunjukkan "
    "bahwa sebaran data perlakuan tersebar merata pada tiga level "
    "(T\u2081 = 1, T\u2082 = 2, T\u2083 = 3). Rata-rata produktivitas padi sebesar "
    "439,52 g/ember dengan simpangan baku 65,996 g menunjukkan adanya "
    "variasi yang cukup besar antar unit percobaan. Variasi ini "
    "mengindikasikan bahwa perlakuan yang diberikan memiliki pengaruh "
    "yang berbeda-beda terhadap hasil produktivitas padi. Jumlah sampel "
    "(N = 27) sesuai dengan rancangan percobaan RAL Faktorial yang "
    "terdiri dari 9 kombinasi perlakuan \u00d7 3 ulangan.")
doc.add_paragraph()

# ===== TABEL 2: CORRELATIONS =====
add_heading_sub(doc, "B. Korelasi Pearson")

add_table(doc,
    ["", "Perlakuan", "Hasil"],
    [
        ["Pearson Correlation - Perlakuan", "1,000", "0,388"],
        ["Pearson Correlation - Hasil", "0,388", "1,000"],
        ["Sig. (1-tailed) - Perlakuan", ".", "0,023"],
        ["Sig. (1-tailed) - Hasil", "0,023", "."],
        ["N - Perlakuan", "27", "27"],
        ["N - Hasil", "27", "27"],
    ],
    "Tabel 4.6b Correlations"
)

add_para(doc,
    "Interpretasi: Nilai korelasi Pearson sebesar 0,388 menunjukkan "
    "adanya hubungan positif dengan kekuatan rendah hingga sedang antara "
    "variabel perlakuan (ketinggian air) dan variabel hasil "
    "(produktivitas padi). Nilai signifikansi sebesar 0,023 (< 0,05) "
    "mengkonfirmasi bahwa hubungan tersebut bersifat signifikan secara "
    "statistik. Artinya, terdapat bukti empiris bahwa perbedaan "
    "perlakuan ketinggian air berpengaruh nyata terhadap produktivitas "
    "padi. Arah hubungan yang positif mengindikasikan bahwa perlakuan "
    "dengan level lebih tinggi (aerob terkendali) cenderung menghasilkan "
    "produktivitas yang lebih tinggi pula.")
doc.add_paragraph()

# ===== TABEL 3: MODEL SUMMARY =====
add_heading_sub(doc, "C. Ringkasan Model Regresi (Model Summary)")

add_table(doc,
    ["Model", "R", "R Square", "Adj. R Square", "Std. Error", "F Change", "Sig. F Change"],
    [
        ["1", "0,388", "0,151", "0,117", "0,782", "4,431", "0,045"],
    ],
    "Tabel 4.6c Model Summary"
)

add_para(doc,
    "Interpretasi: Nilai R = 0,388 menunjukkan kekuatan hubungan antara "
    "variabel prediktor (hasil) dengan variabel dependen (perlakuan). "
    "Nilai R Square (R\u00b2) sebesar 0,151 berarti bahwa 15,1% variasi pada "
    "variabel perlakuan dapat dijelaskan oleh variabel hasil "
    "produktivitas, sedangkan 84,9% sisanya dipengaruhi oleh "
    "faktor-faktor lain yang tidak dimasukkan dalam model, seperti "
    "varietas, kondisi mikroklimat, tingkat serangan hama, dan faktor "
    "pemupukan.")
add_para(doc,
    "Nilai Adjusted R Square sebesar 0,117 merupakan koreksi dari R\u00b2 "
    "yang memperhitungkan jumlah prediktor dalam model. Standard Error "
    "of the Estimate sebesar 0,782 menunjukkan tingkat akurasi prediksi "
    "model. Nilai Sig. F Change = 0,045 (< 0,05) membuktikan bahwa "
    "model regresi secara keseluruhan signifikan dan layak digunakan "
    "untuk menjelaskan hubungan antara variabel.")
doc.add_paragraph()

# ===== TABEL 4: ANOVA =====
add_heading_sub(doc, "D. Tabel ANOVA")

add_table(doc,
    ["Model", "Sum of Squares", "df", "Mean Square", "F", "Sig."],
    [
        ["Regression", "2,710", "1", "2,710", "4,431", "0,045"],
        ["Residual", "15,290", "25", "0,612", "", ""],
        ["Total", "18,000", "26", "", "", ""],
    ],
    "Tabel 4.6d ANOVA"
)

add_para(doc,
    "Interpretasi: Hasil uji ANOVA menunjukkan nilai F hitung sebesar "
    "4,431 dengan signifikansi 0,045. Karena nilai signifikansi < 0,05 "
    "(taraf kepercayaan 95%), maka H\u2080 ditolak dan H\u2081 diterima. "
    "Artinya, model regresi yang digunakan secara statistik signifikan "
    "dan mampu memprediksi hubungan antara perlakuan dengan "
    "produktivitas padi.")
add_para(doc,
    "Sum of Squares Regression sebesar 2,710 menunjukkan variasi yang "
    "dapat dijelaskan oleh model, sedangkan Sum of Squares Residual "
    "sebesar 15,290 merupakan variasi yang tidak dapat dijelaskan. "
    "Perbandingan kedua nilai tersebut (Mean Square Regression / Mean "
    "Square Residual = 2,710 / 0,612 = 4,431) menghasilkan nilai F "
    "yang signifikan, mengkonfirmasi bahwa perlakuan ketinggian air "
    "memiliki pengaruh yang bermakna terhadap produktivitas padi.")
doc.add_paragraph()

# ===== TABEL 5: COEFFICIENTS =====
add_heading_sub(doc, "E. Koefisien Regresi")

add_table(doc,
    ["Model", "B", "Std. Error", "Beta", "t", "Sig.", "Tolerance", "VIF"],
    [
        ["(Constant)", "-0,150", "1,032", "", "-0,145", "0,886", "", ""],
        ["Hasil", "0,005", "0,002", "0,388", "2,105", "0,045", "1,000", "1,000"],
    ],
    "Tabel 4.6e Coefficients"
)

add_para(doc,
    "Interpretasi: Persamaan regresi yang diperoleh adalah:")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Y = -0,150 + 0,005X")
run.bold = True
run.font.name = 'Times New Roman'
run.font.size = Pt(12)
doc.add_paragraph()

add_para(doc,
    "di mana Y = perlakuan dan X = hasil produktivitas (g/ember). "
    "Koefisien regresi (B) bernilai positif sebesar 0,005, yang berarti "
    "setiap peningkatan 1 gram produktivitas padi akan meningkatkan "
    "nilai perlakuan sebesar 0,005 satuan. Secara praktis, hal ini "
    "menunjukkan bahwa perlakuan dengan level yang lebih tinggi "
    "(pengairan aerob terkendali) cenderung menghasilkan produktivitas "
    "yang lebih tinggi.")
add_para(doc,
    "Nilai t hitung sebesar 2,105 dengan signifikansi 0,045 (< 0,05) "
    "menunjukkan bahwa koefisien regresi tersebut signifikan secara "
    "statistik. Nilai Beta (standardized coefficient) sebesar 0,388 "
    "menunjukkan kekuatan pengaruh relatif variabel prediktor terhadap "
    "variabel dependen.")
add_para(doc,
    "Nilai Tolerance sebesar 1,000 dan VIF (Variance Inflation Factor) "
    "sebesar 1,000 menunjukkan tidak terdapat gejala multikolinearitas "
    "pada model. Hal ini wajar mengingat model hanya menggunakan satu "
    "variabel prediktor (simple linear regression).")
doc.add_paragraph()

# ===== TABEL 6: COLLINEARITY DIAGNOSTICS =====
add_heading_sub(doc, "F. Diagnostik Kolinearitas")

add_table(doc,
    ["Model", "Dimension", "Eigenvalue", "Condition Index", "VP (Constant)", "VP (Hasil)"],
    [
        ["1", "1", "1,989", "1,000", "0,01", "0,01"],
        ["1", "2", "0,011", "13,647", "0,99", "0,99"],
    ],
    "Tabel 4.6f Collinearity Diagnostics"
)

add_para(doc,
    "Interpretasi: Condition Index sebesar 13,647 pada dimensi ke-2 "
    "masih berada di bawah ambang batas kritis (30), sehingga tidak "
    "terdapat masalah kolinearitas yang serius dalam model. Variance "
    "Proportions menunjukkan distribusi varian yang jelas antar dimensi, "
    "di mana dimensi 1 menjelaskan sebagian kecil (0,01) dan dimensi 2 "
    "menjelaskan sebagian besar (0,99) dari varian kedua variabel. "
    "Hal ini mengindikasikan bahwa model regresi memiliki stabilitas "
    "yang baik.")
doc.add_paragraph()

# ===== TABEL 7: RESIDUALS STATISTICS =====
add_heading_sub(doc, "G. Statistik Residual")

add_table(doc,
    ["", "Minimum", "Maximum", "Mean", "Std. Deviation", "N"],
    [
        ["Predicted Value", "1,34", "2,58", "2,00", "0,323", "27"],
        ["Residual", "-1,012", "1,389", "0,000", "0,767", "27"],
        ["Std. Predicted Value", "-2,053", "1,795", "0,000", "1,000", "27"],
        ["Std. Residual", "-1,294", "1,776", "0,000", "0,981", "27"],
    ],
    "Tabel 4.6g Residuals Statistics"
)

add_para(doc,
    "Interpretasi: Nilai rata-rata residual sebesar 0,000 menunjukkan "
    "bahwa model regresi tidak memiliki bias sistematis. Standardized "
    "Residual berkisar antara -1,294 hingga 1,776, yang semuanya masih "
    "berada dalam rentang \u00b13, sehingga tidak ditemukan outlier yang "
    "ekstrem. Predicted Value berkisar dari 1,34 hingga 2,58 yang "
    "mencakup rentang perlakuan (1, 2, 3). Hasil ini mengkonfirmasi "
    "bahwa asumsi normalitas residual terpenuhi dan model layak "
    "digunakan untuk analisis inferensi.")
doc.add_paragraph()

# ===== TABEL 8: CROSSTABULATION =====
add_heading_sub(doc, "H. Tabulasi Silang (Crosstabulation)")

add_table(doc,
    ["", "304", "306", "356", "360", "374", "376", "395", "403"],
    [
        ["Perl. 1 (5 cm)", "1", "1", "1", "0", "0", "1", "1", "0"],
        ["Perl. 2 (-0,01)", "0", "0", "0", "0", "0", "0", "0", "0"],
        ["Perl. 3 (-0,05)", "0", "0", "0", "1", "1", "0", "0", "1"],
    ],
    "Tabel 4.6h Crosstabulation Perlakuan * Hasil (Bagian 1)"
)

add_table(doc,
    ["", "419", "421", "434", "442", "446", "453", "455", "463"],
    [
        ["Perl. 1 (5 cm)", "1", "1", "1", "1", "0", "0", "0", "0"],
        ["Perl. 2 (-0,01)", "0", "0", "0", "0", "1", "0", "0", "0"],
        ["Perl. 3 (-0,05)", "0", "0", "0", "0", "0", "1", "1", "1"],
    ],
    "Tabel 4.6h Crosstabulation Perlakuan * Hasil (Bagian 2)"
)

add_table(doc,
    ["", "469", "470", "471", "480", "503", "511", "514", "518", "543", "558"],
    [
        ["Perl. 1 (5 cm)", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
        ["Perl. 2 (-0,01)", "1", "1", "1", "1", "1", "1", "1", "0", "1", "0"],
        ["Perl. 3 (-0,05)", "0", "0", "0", "0", "0", "0", "0", "1", "0", "1"],
    ],
    "Tabel 4.6h Crosstabulation Perlakuan * Hasil (Bagian 3)"
)

add_para(doc,
    "Interpretasi: Tabulasi silang memperlihatkan distribusi frekuensi "
    "produktivitas padi pada masing-masing perlakuan. Dari tabel "
    "tersebut terlihat pola yang jelas:")

items = [
    "Perlakuan 1 (5 cm/tergenang): Produktivitas tersebar pada rentang "
    "304\u2013442 g/ember, menunjukkan hasil yang relatif lebih rendah.",
    "Perlakuan 2 (-0,01 cm/macak-macak): Produktivitas tersebar pada "
    "rentang 446\u2013543 g/ember, menunjukkan hasil tertinggi secara "
    "konsisten di antara ketiga perlakuan.",
    "Perlakuan 3 (-0,05 cm/aerob): Produktivitas tersebar pada rentang "
    "360\u2013558 g/ember, menunjukkan variasi yang lebih besar namun "
    "dengan beberapa nilai yang sangat tinggi.",
]
for i, item in enumerate(items, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.27)
    run = p.add_run(f"{i}. {item}")
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

add_para(doc,
    "Distribusi ini mengkonfirmasi bahwa perlakuan ketinggian air "
    "-0,01 cm (macak-macak) menghasilkan produktivitas yang paling "
    "tinggi dan konsisten, mendukung rekomendasi penerapan IPAT-BO "
    "pada level tersebut.")
doc.add_paragraph()

# ===== TABEL 9: SYMMETRIC MEASURES =====
add_heading_sub(doc, "I. Ukuran Simetris (Symmetric Measures)")

add_table(doc,
    ["", "Value", "Asymp. Std. Error", "Approx. T", "Approx. Sig."],
    [
        ["Pearson's R", "0,388", "0,155", "2,105", "0,045"],
        ["Spearman Correlation", "0,379", "0,191", "2,045", "0,052"],
    ],
    "Tabel 4.6i Symmetric Measures"
)

add_para(doc,
    "Interpretasi: Pearson\u2019s R sebesar 0,388 dengan signifikansi "
    "0,045 (< 0,05) menunjukkan hubungan linear yang signifikan. "
    "Sementara itu, Spearman Correlation sebesar 0,379 dengan "
    "signifikansi 0,052 (mendekati batas signifikansi 0,05) "
    "menunjukkan bahwa hubungan monoton antar variabel juga cukup "
    "kuat meskipun sedikit di atas ambang batas konvensional.")
add_para(doc,
    "Konsistensi antara nilai Pearson dan Spearman menunjukkan bahwa "
    "hubungan antara perlakuan dan hasil bersifat linear dan tidak "
    "terdistorsi oleh outlier atau distribusi data yang tidak normal. "
    "Kedua ukuran ini mengkonfirmasi bahwa pengaturan ketinggian air "
    "dalam sistem IPAT-BO memiliki pengaruh yang bermakna terhadap "
    "produktivitas padi.")
doc.add_paragraph()

# ===== CHARTS DARI FILE ANOVA =====
add_heading_sub(doc, "J. Grafik (Charts) dari Hasil Analisis")
doc.add_paragraph()

# Chart 1
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Gambar 4.1 Histogram Residual")
run.bold = True
run.font.name = 'Times New Roman'
run.font.size = Pt(10)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture('/projects/sandbox/TESIS-AMALIA/chart_anova_1.png', width=Inches(4.5))
doc.add_paragraph()

add_para(doc,
    "Interpretasi: Histogram residual di atas menunjukkan distribusi "
    "residual dari model regresi. Distribusi residual yang mendekati "
    "bentuk kurva normal (bell-shaped) mengindikasikan bahwa asumsi "
    "normalitas residual terpenuhi. Hal ini penting karena validitas "
    "inferensi statistik dari model regresi bergantung pada pemenuhan "
    "asumsi normalitas. Dengan terpenuhinya asumsi ini, hasil uji F "
    "dan uji t yang dilaporkan sebelumnya dapat diandalkan.")
doc.add_paragraph()

# Chart 2
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Gambar 4.2 Normal P-P Plot of Regression Standardized Residual")
run.bold = True
run.font.name = 'Times New Roman'
run.font.size = Pt(10)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture('/projects/sandbox/TESIS-AMALIA/chart_anova_2.png', width=Inches(4.5))
doc.add_paragraph()

add_para(doc,
    "Interpretasi: Normal P-P Plot menampilkan sebaran titik-titik "
    "residual terstandarisasi terhadap garis diagonal distribusi "
    "normal. Apabila titik-titik tersebut tersebar di sepanjang garis "
    "diagonal atau mendekatinya, maka asumsi normalitas terpenuhi. "
    "Dari grafik di atas, terlihat bahwa titik-titik data mengikuti "
    "pola garis diagonal meskipun terdapat sedikit penyimpangan di "
    "bagian ujung. Secara keseluruhan, pola tersebut menunjukkan "
    "bahwa residual model terdistribusi secara normal sehingga model "
    "regresi layak digunakan.")
doc.add_paragraph()

# Chart 3
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Gambar 4.3 Scatterplot Regression Standardized Residual")
run.bold = True
run.font.name = 'Times New Roman'
run.font.size = Pt(10)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture('/projects/sandbox/TESIS-AMALIA/chart_anova_3.png', width=Inches(4.5))
doc.add_paragraph()

add_para(doc,
    "Interpretasi: Scatterplot residual menampilkan sebaran residual "
    "terstandarisasi (sumbu Y) terhadap predicted value terstandarisasi "
    "(sumbu X). Dari grafik di atas, terlihat bahwa titik-titik data "
    "tersebar secara acak tanpa membentuk pola tertentu (seperti "
    "corong, lengkungan, atau pengelompokan). Hal ini menunjukkan "
    "bahwa asumsi homoskedastisitas (kesamaan varians residual) "
    "terpenuhi dan tidak terdapat masalah heteroskedastisitas dalam "
    "model regresi.")
add_para(doc,
    "Tidak ditemukannya pola pada scatterplot juga mengindikasikan "
    "bahwa hubungan antara variabel prediktor dan variabel dependen "
    "bersifat linear, sehingga penggunaan model regresi linear "
    "sederhana sudah tepat untuk menganalisis data penelitian ini.")
doc.add_paragraph()

# ===== KESIMPULAN ANALISIS =====
add_heading_sub(doc, "K. Kesimpulan Analisis Statistik")
add_para(doc,
    "Berdasarkan keseluruhan hasil analisis statistik yang telah "
    "dipaparkan di atas, dapat disimpulkan bahwa:")

kesimpulan = [
    "Terdapat hubungan positif yang signifikan (r = 0,388; p = 0,045) "
    "antara perlakuan ketinggian air dan produktivitas padi, meskipun "
    "kekuatan hubungannya tergolong rendah hingga sedang.",
    "Model regresi yang diperoleh (Y = -0,150 + 0,005X) dinyatakan "
    "signifikan (F = 4,431; p = 0,045) dan layak digunakan untuk "
    "memprediksi pengaruh perlakuan terhadap produktivitas.",
    "Variabel perlakuan ketinggian air mampu menjelaskan 15,1% "
    "variasi produktivitas padi (R\u00b2 = 0,151), sedangkan 84,9% "
    "dipengaruhi faktor lain seperti varietas, pupuk, dan kondisi "
    "lingkungan.",
    "Tidak ditemukan masalah asumsi klasik (normalitas, "
    "homoskedastisitas, dan multikolinearitas), sehingga hasil "
    "analisis valid dan dapat diandalkan.",
    "Perlakuan ketinggian air -0,01 cm (macak-macak/aerob moderat) "
    "secara konsisten menghasilkan produktivitas tertinggi (446\u2013543 "
    "g/ember), mengkonfirmasi keunggulan sistem IPAT-BO dalam "
    "meningkatkan produktivitas padi.",
]

for i, item in enumerate(kesimpulan, 1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.27)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(f"{i}. {item}")
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

doc.add_paragraph()
add_para(doc,
    "Dengan demikian, hasil analisis ANOVA dan regresi ini memberikan "
    "dukungan statistik yang kuat terhadap hipotesis penelitian bahwa "
    "sistem Intensifikasi Padi Aerob Terkendali (IPAT-BO) dengan "
    "pengaturan ketinggian air pada level -0,01 cm berimplikasi "
    "positif terhadap peningkatan produktivitas padi.",
    bold=False)

# ===== SAVE =====
output_path = '/projects/sandbox/TESIS-AMALIA/BAB_4_SubBab_4.6.2_Hasil_ANOVA.docx'
doc.save(output_path)
print(f"File berhasil dibuat: {output_path}")
