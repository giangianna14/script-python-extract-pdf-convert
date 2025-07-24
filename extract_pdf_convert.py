
import os
import zipfile

import fnmatch
import sys
try:
    import win32com.client
except ImportError:
    win32com = None

try:
    from PyPDF2 import PdfMerger
except ImportError:
    PdfMerger = None

def extract_all_zip_files(directory):
    for filename in os.listdir(directory):
        if filename.lower().endswith('.zip'):
            zip_path = os.path.join(directory, filename)
            extract_folder = os.path.join(directory, os.path.splitext(filename)[0])
            os.makedirs(extract_folder, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_folder)
            print(f"Extracted {filename} to {extract_folder}")

def convert_xls_to_pdf_in_folders(directory):
    if win32com is None:
        print("win32com.client is required. Install pywin32 package.")
        return
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if fnmatch.fnmatch(filename, 'Lampiran*.xls'):
                xls_path = os.path.join(root, filename)
                pdf_path = os.path.splitext(xls_path)[0] + ".pdf"
                try:
                    wb = excel.Workbooks.Open(xls_path)
                    for sheet in wb.Worksheets:
                        sheet.PageSetup.Zoom = False
                        sheet.PageSetup.FitToPagesWide = 1
                        sheet.PageSetup.FitToPagesTall = False
                    wb.ExportAsFixedFormat(0, pdf_path)
                    wb.Close(False)
                    print(f"Converted {xls_path} to {pdf_path}")
                except Exception as e:
                    print(f"Failed to convert {xls_path}: {e}")
    excel.Quit()

def merge_pdf_files_in_folders(directory):
    if PdfMerger is None:
        print("PyPDF2 is required. Install with: pip install PyPDF2")
        return
    for root, dirs, files in os.walk(directory):
        ba_files = [f for f in files if fnmatch.fnmatch(f, 'BASementara_*.pdf')]
        lampiran_files = [f for f in files if fnmatch.fnmatch(f, 'LampiranBeritaAcaraSementara__*.pdf')]
        # Buat dict untuk lookup lampiran berdasarkan kode unik
        lampiran_dict = {}
        for lampiran in lampiran_files:
            # Ambil kode unik setelah prefix LampiranBeritaAcaraSementara__
            key = lampiran.replace('LampiranBeritaAcaraSementara__', '').replace('.pdf', '')
            lampiran_dict[key] = lampiran
        for ba in ba_files:
            # Ambil kode unik setelah prefix BASementara_
            key = ba.replace('BASementara_', '').replace('.pdf', '')
            lampiran = lampiran_dict.get(key)
            if lampiran:
                merger = PdfMerger()
                merger.append(os.path.join(root, ba))
                merger.append(os.path.join(root, lampiran))
                out_name = f"out_{lampiran}"
                out_path = os.path.join(root, out_name)
                try:
                    with open(out_path, 'wb') as fout:
                        merger.write(fout)
                    print(f"Merged {ba} + {lampiran} to {out_path}")
                except Exception as e:
                    print(f"Failed to merge {ba} and {lampiran} in {root}: {e}")
                merger.close()

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    extract_all_zip_files(current_dir)
    convert_xls_to_pdf_in_folders(current_dir)
    merge_pdf_files_in_folders(current_dir)
