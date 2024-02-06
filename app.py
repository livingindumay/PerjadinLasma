from flask import Flask, request, render_template, send_file
import pandas as pd
import os

# Membuat instance aplikasi Flask
app = Flask(__name__)

# Tentukan direktori untuk menyimpan file yang di-upload
UPLOAD_FOLDER = 'uploads/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Membuat direktori jika belum ada, tanpa error jika sudah ada
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Fungsi untuk mendeteksi tumpang tindih dalam data perjalanan
def find_overlaps_inclusive(data):
    overlaps = []
    for index, row in data.iterrows():
        # Pythonic way untuk cek tumpang tindih dengan list comprehension
        overlaps.extend(data[(data['Tanggal Mulai'] <= row['Tanggal Selesai']) &
                             (data['Tanggal Selesai'] >= row['Tanggal Mulai']) &
                             (data['Nama Pelaksana'] == row['Nama Pelaksana']) &
                             (data.index != index)].index)
    return data.loc[list(set(overlaps))]


# Fungsi untuk menyajikan data tumpang tindih secara sebelah-sebelahan
def side_by_side_overlaps(data):
    rows_list = [
        [
            row_i[col] if col < 5 else row_j[col - 5]
            for col in range(10)
        ]
        for pelaksana in data['Nama Pelaksana'].unique()
        for i, row_i in data[data['Nama Pelaksana'] == pelaksana].iterrows()
        for j, row_j in data[data['Nama Pelaksana'] == pelaksana].iterrows()
        if i < j and (row_i['Tanggal Mulai'] <= row_j['Tanggal Selesai']) and (row_i['Tanggal Selesai'] >= row_j['Tanggal Mulai'])
    ]
    columns = ['No 1', 'Nama Pelaksana 1', 'Tanggal Mulai 1', 'Tanggal Selesai 1', 'Detail Perjalanan 1',
               'No 2', 'Nama Pelaksana 2', 'Tanggal Mulai 2', 'Tanggal Selesai 2', 'Detail Perjalanan 2']
    return pd.DataFrame(rows_list, columns=columns)

# Rute utama aplikasi untuk upload dan proses file
@app.route('/', methods=['GET', 'POST'])
def upload_and_process():
    if request.method == 'POST':
        # Dapatkan file dari form POST
        file = request.files['file']
        if file:
            # Tentukan path untuk menyimpan file yang di-upload
            original_filename = file.filename
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
            file.save(filepath)  # Menyimpan file yang di-upload

            try:
                # Proses file yang di-upload dengan pandas
                data = pd.read_excel(filepath, header=4)
                 # Membersihkan data
                data.iloc[:, 1] = data.iloc[:, 1].str.strip()  # Menghapus spasi di awal dan akhir kolom 2 ('Nama Pelaksana')
                data.iloc[:, 2] = data.iloc[:, 2].str.replace(' ', '', regex=True)  # Menghapus semua spasi pada kolom 3

                travel_data = data.iloc[:, [0, 1, 8, 9, 3]]
                # Tentukan nama kolom untuk memudahkan akses data
                travel_data.columns = ['No', 'Nama Pelaksana', 'Tanggal Mulai', 'Tanggal Selesai', 'Detail Perjalanan']
                # Konversi kolom tanggal menjadi tipe datetime
                travel_data.loc[:, 'Tanggal Mulai'] = pd.to_datetime(travel_data['Tanggal Mulai'], errors='coerce')
                travel_data.loc[:, 'Tanggal Selesai'] = pd.to_datetime(travel_data['Tanggal Selesai'], errors='coerce')
                # Buang data yang tidak lengkap
                cleaned_travel_data = travel_data.dropna(subset=['Tanggal Mulai', 'Tanggal Selesai'])
                # Lakukan deteksi tumpang tindih dan susun data sebelah-sebelahan
                inclusive_overlapping_travels = find_overlaps_inclusive(cleaned_travel_data)
                side_by_side_data = side_by_side_overlaps(inclusive_overlapping_travels)
                
                # Menyusun nama file output berdasarkan nama file asli
                name_part = os.path.splitext(original_filename)[0]
                output_filename = f"{name_part}_Processed_Data.xlsx"
                output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                
                # Menyimpan data yang telah diproses ke file baru
                side_by_side_data.to_excel(output_filepath, index=False)
                
                # Kirim file yang telah diproses sebagai file unduhan
                return send_file(output_filepath, as_attachment=True)  

            except Exception as e:
                # Jika terjadi kesalahan, tampilkan pesan error
                return f"Error processing file: {e}"

    # Tampilkan form upload jika method adalah GET
    return render_template('upload.html')  

if __name__ == '__main__':
    app.run(debug=True)  # Jalankan aplikasi dengan mode debug