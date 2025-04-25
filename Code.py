#!/usr/bin/env python
# coding: utf-8

# In[1]:


import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image
import os
import io

# Konstanta
IMAGE_FOLDER = 'uploaded_images'
TEMP_EXCEL_FILE = 'temp_data_gambar.xlsx'

# Buat folder gambar jika belum ada
if not os.path.exists(IMAGE_FOLDER):
    os.makedirs(IMAGE_FOLDER)

st.title("📸 Upload Gambar + Keterangan ➜ Simpan ke Excel (dengan Gambar)")

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

# 2. Upload gambar + keterangan
st.subheader("2️⃣ Upload Gambar Baru dan Isi Keterangan")
uploaded_files = st.file_uploader("Upload gambar (bisa multiple)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

descriptions = []
if uploaded_files:
    for img in uploaded_files:
        desc = st.text_input(f"Keterangan untuk {img.name}", key=img.name)
        descriptions.append((img, desc))

# 3. Simpan ke Excel
if st.button("💾 Simpan ke Excel"):
    wb = openpyxl.load_workbook(TEMP_EXCEL_FILE)
    ws = wb.active
    next_row = ws.max_row + 1

    for img_file, desc in descriptions:
        if desc:
            # Simpan gambar ke folder lokal
            image_path = os.path.join(IMAGE_FOLDER, img_file.name)
            with open(image_path, "wb") as f:
                f.write(img_file.getbuffer())

            # Sisipkan gambar
            img = Image(image_path)
            img.width = 100
            img.height = 100
            ws.add_image(img, f'A{next_row}')

            # Tambahkan keterangan
            ws[f'B{next_row}'] = desc

            next_row += 1

    wb.save(TEMP_EXCEL_FILE)
    st.success("Data berhasil ditambahkan ke Excel!")

# 4. Download file Excel terbaru
st.subheader("3️⃣ Download File Excel Terbaru")
if os.path.exists(TEMP_EXCEL_FILE):
    with open(TEMP_EXCEL_FILE, "rb") as f:
        st.download_button('⬇️ Download Excel', f, file_name='data_gambar_final.xlsx')


# In[ ]:




