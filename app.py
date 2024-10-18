from flask import Flask, render_template, request, redirect, url_for
import os
from fuzzywuzzy import fuzz, process
from werkzeug.utils import secure_filename
import mammoth  # Untuk mengonversi file DOCX ke HTML
import pandas as pd  # Untuk membaca file Excel

app = Flask(__name__)

# Konfigurasi lokasi folder penyimpanan dan ekstensi yang diizinkan
IMAGE_FOLDER = os.path.join('static', 'images')
UPLOAD_FOLDER = os.path.join('static', 'uploads')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Data poster dengan kata kunci (tanpa ekstensi)
image_data = {
    'penambalan gigi': ['poster_imunisasi1', 'poster_imunisasi2', 'poster_covid1', 'poster_covid2', '4-revisi', '5-revisi'],
    'covid': ['poster_covid1', 'poster_covid2'],
    'jumpscare': ['abi'],
    'dokumen':['MAKALAH BING KARIR A_Company Research_202131093_Dimas Abi Mesti']
}

# Fungsi untuk memeriksa ekstensi file yang diunggah
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Fungsi untuk mencari gambar berdasarkan kata kunci dengan fuzzy search
def find_images(keyword):
    found_images = []
    # Cari kata kunci yang mirip dalam image_data menggunakan fuzzy matching
    possible_matches = process.extract(keyword, image_data.keys(), limit=2, scorer=fuzz.ratio)
    
    # Ambil hasil terbaik yang memiliki rasio kecocokan lebih dari 60%
    for match, score in possible_matches:
        if score > 50:  # Threshold minimal kecocokan
            for image_name in image_data[match]:
                # Cek file dengan ekstensi .jpg dan .png
                for ext in ['jpg', 'png']:
                    image_path = os.path.join(IMAGE_FOLDER, f"{image_name}.{ext}")
                    if os.path.exists(image_path):
                        found_images.append(f"{image_name}.{ext}")
    return found_images

# Fungsi untuk membaca dan mengonversi file DOCX ke HTML
# def read_docx(file_path):
#     with open(file_path, "rb") as docx_file:
#         result = mammoth.convert_to_html(docx_file)
#         return result.value

# Fungsi untuk membaca file Excel dan mengonversinya menjadi tabel HTML
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df.to_html(classes='table table-striped')

# Rute utama untuk halaman pencarian dan unggah file
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        search_term = request.form['search'].lower()
        return redirect(url_for('search_results', search_term=search_term))
    return render_template('index.html')

# Rute untuk menampilkan hasil pencarian di halaman baru
@app.route('/results/<search_term>')
def search_results(search_term):
    images = find_images(search_term)
    uploaded_files = os.listdir(app.config['UPLOAD_FOLDER'])
    file_previews = []

    # Cek setiap file yang ada di folder upload dan buat tampilan yang sesuai
    for filename in uploaded_files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file_extension = filename.rsplit('.', 1)[1].lower()
        content = None

        if file_extension == 'pdf':
            # Tampilkan PDF dengan <embed> atau <iframe>
            file_previews.append({'filename': filename, 'type': 'pdf'})
        elif file_extension == 'docx':
            # Konversi DOCX ke HTML
            content = read_doxc(file_path)
            file_previews.append({'filename': filename, 'type': 'docx', 'content': content})
        elif file_extension == 'xlsx':
            # Baca Excel dan konversi ke tabel HTML
            content = read_excel(file_path)
            file_previews.append({'filename': filename, 'type': 'excel', 'content': content})

    return render_template('results.html', images=images, file_previews=file_previews, search_term=search_term)

if __name__ == '__main__':
    # Buat folder penyimpanan jika belum ada
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
