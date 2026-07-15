import os
import shutil
import subprocess
import glob
import sys

dapur_dir = "Dapur"
required_dapur_files = ["Analisis_AR.py", "Analisis_Audit_1.py", "Analisis_Audit_2.py", "config.conf", "__init__.py"]
req_file_pattern = "PTM*.xlsx"
final_outputs = ["Laporan_Analisis_Prosedur_Audit.xlsx", "Monitoring.xlsx"]

missing_items = []

if not os.path.exists(dapur_dir):
    missing_items.append(dapur_dir)
else:
    for f in required_dapur_files:
        file_path = os.path.join(dapur_dir, f)
        if not os.path.exists(file_path):
            missing_items.append(file_path)

matching_files = glob.glob(req_file_pattern)
if not matching_files:
    missing_items.append(req_file_pattern)
else:
    req_file_1 = matching_files[0]

if missing_items:
    print("--> Proses digagalkan. File atau folder berikut tidak ditemukan:")
    for item in missing_items:
        print(f"--> {item}")
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

shutil.move(req_file_1, os.path.join(dapur_dir, os.path.basename(req_file_1)))

scripts_to_run = ["Analisis_AR.py", "Analisis_Audit_1.py", "Analisis_Audit_2.py"]
current_dir = os.getcwd()
os.chdir(dapur_dir)

try:
    for script in scripts_to_run:
        print(f"--> Menjalankan {script}...")
        subprocess.run([sys.executable, script], check=True)
except subprocess.CalledProcessError:
    print(f"--> Terjadi kesalahan saat menjalankan {script}. Proses dihentikan.")
    os.chdir(current_dir)
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

os.chdir(current_dir)

for output in final_outputs:
    hasil_file = os.path.join(dapur_dir, output)
    if os.path.exists(hasil_file):
        shutil.copy(hasil_file, output)
        print(f"--> File {output} berhasil dibuat dan disalin.")
    else:
        print(f"--> Gagal: File {output} tidak ditemukan setelah proses.")

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

print("--> Semua proses selesai dan folder Dapur telah dibersihkan dari file sampah.")
