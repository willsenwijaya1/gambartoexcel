import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image
import os

# Konstanta
IMAGE_FOLDER = 'uploaded_images'
TEMP_EXCEL_FILE = 'temp_data_gambar.xlsx'

# Buat folder gambar jika belum ada
if not os.path.exists(IMAGE_FOLDER):
    os.makedirs(IMAGE_FOLDER)

st.title("📸 Upload Gambar + Keterangan ➜ Simpan ke Excel (1 per 1)")

# 1. Upload file Excel lama (opsional)
st.subheader("1️⃣ Upload File Excel Lama (Jika Ada)")
uploaded_excel = st.file_uploader("Upload file Excel sebelumnya (optional)", type=["xlsx"])

# Simpan sementara jika user upload file Excel lama
if uploaded_excel:
    with open(TEMP_EXCEL_FILE, "wb") as f:
        f.write(uploaded_excel.getbuffer())
    st.success("File Excel lama berhasil dimuat!")
else:
    # Jika tidak upload, cek apakah sudah ada file temp sebelumnya
    if not os.path.exists(TEMP_EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Gambar', 'Keterangan'])
        wb.save(TEMP_EXCEL_FILE)

# 2. Upload 1 gambar + keterangan
st.subheader("2️⃣ Upload Gambar dan Isi Keterangan")
uploaded_file = st.file_uploader("Upload 1 gambar", type=["jpg", "png", "jpeg"])

description = st.text_input("Masukkan keterangan gambar")

# 3. Simpan ke Excel
if st.button("💾 Simpan ke Excel"):
    if uploaded_file and description:
        wb = openpyxl.load_workbook(TEMP_EXCEL_FILE)
        ws = wb.active
        next_row = ws.max_row + 1

        # Simpan gambar ke folder lokal
        image_path = os.path.join(IMAGE_FOLDER, uploaded_file.name)
        with open(image_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Sisipkan gambar
        img = Image(image_path)
        img.width = 100
        img.height = 100
        ws.add_image(img, f'A{next_row}')

        # Tambahkan keterangan
        ws[f'B{next_row}'] = description

        wb.save(TEMP_EXCEL_FILE)
        st.success("Gambar dan keterangan berhasil disimpan ke Excel!")
    else:
        st.error("Silakan upload gambar dan isi keterangan terlebih dahulu.")

# 4. Download file Excel terbaru
st.subheader("3️⃣ Download File Excel Terbaru")
if os.path.exists(TEMP_EXCEL_FILE):
    with open(TEMP_EXCEL_FILE, "rb") as f:
        st.download_button('⬇️ Download Excel', f, file_name='data_gambar_final.xlsx')
