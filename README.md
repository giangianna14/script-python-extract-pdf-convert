# Extract, Convert, and Merge PDF Utility

Script ini digunakan untuk:
1. Mengekstrak semua file ZIP di folder kerja.
2. Mengonversi file Excel (Lampiran*.xls) hasil ekstrak menjadi PDF, dengan lebar worksheet di-fit ke halaman PDF.
3. Menggabungkan file BASementara_*.pdf dengan pasangannya LampiranBeritaAcaraSementara__*.pdf menjadi satu file PDF baru (out_LampiranBeritaAcaraSementara__*.pdf) di setiap folder.

## Cara Pakai

1. **Pastikan Python 3.x sudah terinstal.**
2. **Install dependensi:**
   - Untuk konversi Excel ke PDF: `pywin32` (hanya Windows, butuh Microsoft Excel terinstal)
   - Untuk merge PDF: `PyPDF2`

   Jalankan perintah berikut di terminal:
   ```bash
   pip install pywin32 PyPDF2
   ```

3. **Letakkan semua file ZIP di folder yang sama dengan script ini.**
4. **Jalankan script:**
   ```bash
   python extract_pdf_convert.py
   ```

## Penjelasan Proses

- Semua file ZIP di folder akan diekstrak ke folder baru sesuai nama file ZIP.
- Semua file Excel dengan nama Lampiran*.xls di setiap folder hasil ekstrak akan dikonversi ke PDF.
- Untuk setiap pasangan file:
  - `BASementara_XXXX.pdf` dan `LampiranBeritaAcaraSementara__XXXX.pdf` (kode unik XXXX sama),
  - Akan digabung menjadi satu file: `out_LampiranBeritaAcaraSementara__XXXX.pdf`.

## Catatan
- Script hanya berjalan di Windows (karena konversi Excel ke PDF butuh Excel/Windows).
- Jika ada error terkait win32com, pastikan Excel dan pywin32 sudah terinstal.
- Jika ada error terkait PyPDF2, pastikan sudah install PyPDF2.

## Kontak
Untuk pertanyaan lebih lanjut, hubungi pengembang script ini.
