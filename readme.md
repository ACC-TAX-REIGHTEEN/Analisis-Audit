# 🔬 Laporan Analisis Prosedur Audit Piutang

> **11 prosedur audit analitik dalam satu laporan Excel — dari statistik klasik hingga Machine Learning, lengkap dengan visualisasi grafik otomatis**

Pipeline Python dua tahap yang membaca data AR bulanan dari `PTM*.xlsx`, lalu menghasilkan **`Laporan_Analisis_Prosedur_Audit.xlsx`** berisi 11 sheet prosedur audit: lima analisis audit klasik bergrafik (Horizontal & Vertikal, Z-Score, Benford, Pareto 80/20, Duplikasi Faktur), dilanjutkan enam deteksi anomali berbasis **Machine Learning** (DBSCAN, Isolation Forest, LOF, One-Class SVM, K-Means, Neural Network Autoencoder) — masing-masing disertai tabel temuan dan scatter plot otomatis.

---

## 📋 Daftar Isi

- [Gambaran Umum & Konteks](#-gambaran-umum--konteks)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Sheet 1–5: Prosedur Audit Klasik](#-sheet-15-prosedur-audit-klasik)
  - [Sheet 1 — Analisis Horisontal & Vertikal](#sheet-1--analisis-horisontal--vertikal)
  - [Sheet 2 — Anomali Z-Score Makro](#sheet-2--anomali-z-score-makro)
  - [Sheet 3 — Forensik Hukum Benford](#sheet-3--forensik-hukum-benford)
  - [Sheet 4 — Risiko Pareto Customer](#sheet-4--risiko-pareto-customer)
  - [Sheet 5 — Uji Duplikasi Faktur](#sheet-5--uji-duplikasi-faktur)
- [Sheet 6–11: Deteksi Anomali Machine Learning](#-sheet-611-deteksi-anomali-machine-learning)
  - [Sheet 6 — DBSCAN](#sheet-6--dbscan-density-based-spatial-clustering)
  - [Sheet 7 — Isolation Forest](#sheet-7--isolation-forest)
  - [Sheet 8 — Local Outlier Factor](#sheet-8--local-outlier-factor-lof)
  - [Sheet 9 — One-Class SVM](#sheet-9--one-class-svm)
  - [Sheet 10 — K-Means Clustering](#sheet-10--k-means-centroid-distance)
  - [Sheet 11 — Neural Network Autoencoder](#sheet-11--neural-network-autoencoder)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum & Konteks

Proyek ini terdiri dari dua lapisan prosedur audit yang berjalan secara berurutan:

| Tahap | Skrip | Isi | Metode |
|---|---|---|---|
| **Tahap 1** | `Analisis_Audit_1.py` | Sheet 1–5 | Statistik klasik + grafik tren/distribusi |
| **Tahap 2** | `Analisis_Audit_2.py` | Sheet 6–11 | Machine Learning unsupervised anomaly detection |

Keduanya dijalankan otomatis oleh `Buat Analisis Audit.py` sebagai orkestrator. Tahap 2 membuka hasil Tahap 1 dan **menambahkan** sheet baru ke dalamnya, sehingga laporan akhir berisi 11 sheet dalam satu file.

---

## ✨ Fitur Utama

- **Pipeline dua tahap terpisah** — Analisis klasik (Sheet 1–5) dan ML (Sheet 6–11) dijalankan sebagai modul terpisah oleh orkestrator, memudahkan debugging dan pengembangan mandiri.
- **4 grafik audit otomatis (Tahap 1)** — Setiap sheet utama menghasilkan chart matplotlib/seaborn yang langsung disisipkan ke Excel: dual-axis tren AR, peta Z-Score dengan garis threshold, perbandingan Benford, dan kurva Pareto.
- **6 metode ML anomaly detection (Tahap 2)** — Menggunakan fitur `Nilai Faktur`, `Sisa Piutang`, dan `Umur Piutang (Hari)` sebagai input tiga dimensi untuk mendeteksi transaksi pencilan dari berbagai sudut pandang algoritmik.
- **Scatter plot per metode ML** — Setiap sheet ML menyertakan scatter plot `Nilai Faktur vs Sisa Piutang` dengan warna berbeda untuk titik WAJAR (biru) dan ANOMALI (merah), langsung tertanam di Excel.
- **Pemrosesan data besar (chunking)** — DBSCAN dan LOF mendukung pemrosesan data hingga lebih dari 20.000 baris dengan mekanisme chunking otomatis.
- **Sampling cerdas One-Class SVM** — Jika data melebihi 10.000 baris, SVM dilatih pada 10.000 sampel acak untuk efisiensi, lalu digunakan untuk memprediksi seluruh dataset.
- **StandardScaler normalisasi** — Semua fitur dinormalisasi sebelum diumpankan ke model ML agar skala IDR tidak mendominasi jarak Euclidean.
- **Auto-cleanup gambar sementara** — File PNG chart yang dihasilkan matplotlib dihapus otomatis setelah disisipkan ke Excel.
- **Validasi kelengkapan file di awal** — Orkestrator memeriksa keberadaan semua file dan folder yang dibutuhkan sebelum memulai, dan menampilkan daftar file yang hilang jika ada.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl numpy matplotlib seaborn scikit-learn
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Keduanya | Baca Excel, transformasi, groupby |
| `openpyxl` | Keduanya | Buat/edit workbook, styling, sisipkan gambar |
| `numpy` | Keduanya | Z-Score, `np.where`, operasi array |
| `matplotlib` | Keduanya | Render grafik ke file PNG |
| `seaborn` | Keduanya | Styling chart (barplot, lineplot, scatterplot) |
| `scikit-learn` | Analisis_Audit_2.py | DBSCAN, Isolation Forest, LOF, One-Class SVM, K-Means, MLPRegressor, StandardScaler |
| `configparser`, `re`, `glob`, `datetime`, `os`, `shutil`, `subprocess` | Keduanya | Utilitas (standard library) |

> **Catatan scikit-learn:** Pastikan versi scikit-learn ≥ 1.0. Gunakan `pip install scikit-learn --upgrade` jika diperlukan.

---

## 📁 Struktur Folder & File

```
📦 Analisis-Audit/
│
├── 📄 Buat Analisis Audit.py              ← Orkestrator utama. Jalankan ini
│
├── 📄 PTM*.xlsx                           ← [INPUT] File data AR bulanan dari Accurate
│                                             (nama bebas asal diawali "PTM")
│
├── 📁 Dapur/                              ← Folder pipeline (jangan diubah strukturnya)
│   ├── 📄 __init__.py
│   ├── 📄 Analisis_Audit_1.py            ← Tahap 1: Sheet 1–5 audit klasik + grafik
│   ├── 📄 Analisis_Audit_2.py            ← Tahap 2: Sheet 6–11 deteksi anomali ML
│   └── 📄 config.conf                    ← Konfigurasi filter produk
│
├── 📁 Contoh Data/                        ← Folder data sampel untuk uji coba
│   └── 📄 PTM - Laporan AR Bulanan Jawa.xlsx
│
└── 📄 Laporan_Analisis_Prosedur_Audit.xlsx ← [OUTPUT] Dihasilkan otomatis
```

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input PTM

Letakkan file ekspor AR dari Accurate di **folder utama** (sejajar dengan `Buat Analisis Audit.py`) dengan nama yang diawali `PTM`:

```
PTM_Jawa_2025.xlsx      ← contoh nama valid
PTM.xlsx                ← contoh nama valid
PTM AR Bulanan.xlsx     ← contoh nama valid
```

File PTM wajib memiliki sheet dengan nama format `MMYY` (mis. `0125` untuk Jan-2025), berisi kolom berikut:

| Kolom wajib | Kolom opsional |
|---|---|
| `Sisa Piutang` | `Negara Pelanggan` |
| `Umur AR base on Tgl Faktur` | `Nilai Faktur` |
| `Nama Penjual` | |
| `Nama Pelanggan` | |
| `No. Faktur` | |

> **Uji coba pertama?** Salin file `Contoh Data/PTM - Laporan AR Bulanan Jawa.xlsx` ke folder utama dan ganti namanya menjadi `PTM - Laporan AR Bulanan Jawa.xlsx` atau nama apapun yang diawali `PTM`.

### Langkah 2 — Jalankan

```bash
python "Buat Analisis Audit.py"
```

atau klik dua kali file tersebut jika Python sudah terasosiasi di sistem.

### Langkah 3 — Pantau progress di terminal

```
--> Menjalankan Analisis_Audit_1.py...
--> Memulai pemrosesan data audit analisis ...
--> Menyusun Sheet 1: Analisis Horisontal & Vertikal...
--> Menyusun Sheet 2: Deteksi Statistik Z-Score...
--> Menyusun Sheet 3: Forensik Digit Angka Hukum Benford...
--> Menyusun Sheet 4: Kosentrasi Risiko Piutang Hukum Pareto...
--> Menyusun Sheet 5: Uji Duplikasi Nomor Faktur...
--> SUKSES! Berkas laporan analitis audit dengan grafik otomatis '...' telah diterbitkan.
--> Menjalankan Analisis_Audit_2.py...
--> Memulai pemrosesan data audit advanced machine learning ...
--> Menyusun Sheet 6: Deteksi Anomali DBSCAN...
--> Menyusun Sheet 7: Deteksi Anomali Isolation Forest...
--> Menyusun Sheet 8: Deteksi Anomali Local Outlier Factor...
--> Menyusun Sheet 9: Deteksi Anomali One-Class SVM...
--> Menyusun Sheet 10: Deteksi Anomali K-Means Distance...
--> Menyusun Sheet 11: Deteksi Anomali Neural Network Autoencoder...
--> SUKSES! Berkas laporan analitis audit '...' telah diperbarui dengan prosedur Machine Learning.
--> File Laporan_Analisis_Prosedur_Audit.xlsx berhasil dibuat dan disalin.
--> Semua proses selesai dan folder Dapur telah dibersihkan dari file sampah.
```

> ⏱️ **Estimasi waktu:** Tahap 1 selesai dalam hitungan detik. Tahap 2 bergantung pada jumlah baris data — untuk data ribuan baris bisa memakan 1–5 menit karena proses training model ML.

### Langkah 4 — Buka laporan

Buka **`Laporan_Analisis_Prosedur_Audit.xlsx`** di folder utama.

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Buat Analisis Audit.py]
   │
   ├─── Validasi kelengkapan file
   │       Cek folder Dapur/ ada
   │       Cek Analisis_Audit_1.py, Analisis_Audit_2.py, config.conf, __init__.py
   │       Cek PTM*.xlsx ada di folder utama
   │       Jika ada yang hilang → tampilkan daftar → berhenti
   │
   ├─── Bersihkan sisa run sebelumnya
   │       Hapus semua *.xls dan *.xlsx dari folder Dapur/
   │
   ├─── Pindahkan PTM*.xlsx → Dapur/
   │
   ├─── Pindah ke direktori Dapur/ → jalankan Analisis_Audit_1.py
   │       [TAHAP 1 — lihat detail di bawah]
   │       Output: Dapur/Laporan_Analisis_Prosedur_Audit.xlsx (Sheet 1–5)
   │       Jika gagal → berhenti + tampilkan pesan error
   │
   ├─── Jalankan Analisis_Audit_2.py (masih di Dapur/)
   │       [TAHAP 2 — lihat detail di bawah]
   │       Output: Dapur/Laporan_Analisis_Prosedur_Audit.xlsx (Sheet 1–11)
   │       Jika gagal → berhenti + tampilkan pesan error
   │
   ├─── Kembali ke direktori utama
   ├─── Salin laporan dari Dapur/ → folder utama
   └─── Bersihkan Dapur/ dari semua file Excel
```

### Detail Tahap 1 (`Analisis_Audit_1.py`)

```
Baca & validasi file PTM → loop semua sheet MMYY
  ├─ Bangun df_macro_master (7 metrik per bulan)
  └─ Bangun df_micro_master (semua transaksi lintas bulan)

[Sheet 1] Horizontal & Vertikal
  └─ Tabel MoM + Rasio → Dual-axis chart (bar AR + line rasio)

[Sheet 2] Z-Score
  └─ Tabel Z-Score per bulan → Bar chart dengan garis threshold ±1.2

[Sheet 3] Benford
  └─ Tabel 9 digit vs distribusi Benford → Grouped bar chart

[Sheet 4] Pareto 80/20
  └─ Tabel top 25 pelanggan → Pareto curve chart (bar + garis kumulatif)

[Sheet 5] Duplikasi Faktur
  └─ Tabel duplikat intra-bulan (tanpa chart)

Hapus file PNG chart sementara → Simpan Laporan_Analisis_Prosedur_Audit.xlsx
```

### Detail Tahap 2 (`Analisis_Audit_2.py`)

```
Buka Laporan_Analisis_Prosedur_Audit.xlsx yang sudah ada (dari Tahap 1)
Baca ulang PTM → bangun df_micro_master
Normalisasi 3 fitur: [Nilai Faktur, Sisa Piutang, Umur Hari] via StandardScaler

Untuk setiap metode ML:
  ├─ Latih model → prediksi label (ANOMALI / WAJAR) per baris transaksi
  ├─ Buat tabel detail baris yang terdeteksi ANOMALI
  ├─ Buat scatter plot (Nilai Faktur vs Sisa Piutang, warna per label)
  └─ Sisipkan tabel + scatter plot ke sheet baru

[Sheet 6]  DBSCAN (cluster = -1 → ANOMALI)
[Sheet 7]  Isolation Forest (contamination = 2%)
[Sheet 8]  Local Outlier Factor (k=20, contamination = 2%)
[Sheet 9]  One-Class SVM (nu=0.02, kernel=rbf)
[Sheet 10] K-Means Distance (k=4, threshold = persentil 98)
[Sheet 11] Autoencoder NN (threshold rekonstruksi = persentil 98)

Hapus file PNG sementara → Simpan Laporan_Analisis_Prosedur_Audit.xlsx (11 sheet)
```

---

## 📊 Sheet 1–5: Prosedur Audit Klasik

### Sheet 1 — Analisis Horisontal & Vertikal

**Tujuan audit:** Menilai volatilitas perubahan saldo bulanan dan komposisi risiko portofolio piutang.

**Analisis Horisontal:** Persentase perubahan Total AR Outstanding bulan ke bulan (MoM) via `pct_change()`. Kenaikan MoM yang konsisten tanpa pembayaran = sinyal penumpukan.

**Analisis Vertikal:** Proporsi AR 60 Hari Up dan Bad Debt terhadap Total AR. Rasio meningkat = penurunan kualitas portofolio.

**Grafik otomatis:** Dual-axis chart — batang biru untuk Total AR Outstanding (sumbu kiri IDR), dua garis untuk Rasio 60+ dan Rasio Bad Debt (sumbu kanan %).

| Kolom | Format |
|---|---|
| Bulan | Teks |
| Total AR Outstanding | `#,##0` |
| Perubahan MoM (%) | `0.00%` |
| AR 60 Hari Up | `#,##0` |
| Rasio Kontribusi 60+ (%) | `0.00%` |
| AR Bad Debt (365+) | `#,##0` |
| Rasio Kontribusi Bad Debt (%) | `0.00%` |

---

### Sheet 2 — Anomali Z-Score Makro

**Tujuan audit:** Menemukan bulan dengan lonjakan saldo ekstrim di luar batas fluktuasi normal secara statistik.

```
Z-Score (bulan X) = (Total AR bulan X − Rata-rata semua bulan) / Std Dev semua bulan
```

| Rentang Z-Score | Status |
|---|---|
| `−1.2` s.d. `+1.2` | **WAJAR** |
| `< −1.2` atau `> +1.2` | **ANOMALI / REVIU MENDALAM** — highlight merah, bold |

**Grafik otomatis:** Bar chart Z-Score per bulan dengan dua garis horizontal putus-putus merah di ±1.2 sebagai batas toleransi.

---

### Sheet 3 — Forensik Hukum Benford

**Tujuan audit:** Mendeteksi indikasi manipulasi atau faktur fiktif via distribusi digit pertama seluruh nilai faktur.

**Distribusi Benford teoritis:**

| Digit | Ekspektasi |
|---|---|
| 1 | 30,1% |
| 2 | 17,6% |
| 3 | 12,5% |
| 4 | 9,7% |
| 5 | 7,9% |
| 6 | 6,7% |
| 7 | 5,8% |
| 8 | 5,1% |
| 9 | 4,6% |

**Threshold:** `|Deviasi| > 5%` → `"REVIU (Deviasi Tinggi)"` + highlight merah.

**Grafik otomatis:** Grouped bar chart membandingkan proporsi riil (oranye) vs ekspektasi Benford (biru) per digit 1–9.

---

### Sheet 4 — Risiko Pareto Customer

**Tujuan audit:** Mengidentifikasi pelanggan yang secara kolektif menguasai 80% saldo piutang (konsentrasi risiko gagal bayar).

**Sumber data:** Hanya bulan terakhir yang tersedia di PTM. Menampilkan **top 25 pelanggan** berdasarkan saldo.

**Logika klasifikasi:**
- Urutkan pelanggan dari saldo terbesar
- Hitung `Kumulatif%` secara kumulatif
- `Kumulatif% ≤ 80%` → `TOP 80% CORE RISK (Kritis)` (highlight merah, bold)

**Grafik otomatis:** Kurva Pareto — batang biru per pelanggan (saldo, sumbu kiri) + garis merah kumulatif (sumbu kanan %) dengan garis putus-putus di 80%.

---

### Sheet 5 — Uji Duplikasi Faktur

**Tujuan audit:** Mendeteksi nomor faktur yang diinput lebih dari sekali dalam bulan yang sama (double input / double journal).

Pengujian bersifat **intra-bulan** — faktur yang muncul di dua bulan berbeda (carry-over) tidak dianggap duplikat. Sheet ini tidak memiliki grafik; hanya tabel temuan.

| Kolom | Format |
|---|---|
| Nomor Faktur | Teks |
| Jumlah Kemunculan Di Bulan Ini | `#,##0` |
| ID/Nama Pelanggan | Teks |
| Bulan Transaksi | Teks |
| Nilai Faktur | `#,##0` |
| Sisa Piutang | `#,##0` |
| Rekomendasi Tindakan | `REVIU JURNAL (Duplikasi Intra-Bulan)` |

---

## 🤖 Sheet 6–11: Deteksi Anomali Machine Learning

Keenam sheet ini menggunakan **tiga fitur numerik** yang telah dinormalisasi via `StandardScaler`:

| Fitur | Kolom sumber |
|---|---|
| `Cleaned_Nilai_Faktur` | `Nilai Faktur` |
| `Cleaned_Sisa_Piutang` | `Sisa Piutang` |
| `Cleaned_Days` | Angka dari `Umur AR base on Tgl Faktur` |

Semua sheet menggunakan format tabel yang sama:

| Kolom | Format | Keterangan |
|---|---|---|
| Bulan | Teks | Bulan transaksi |
| Nomor Faktur | Teks | No. faktur |
| ID/Nama Pelanggan | Teks | Nama atau kode pelanggan |
| Nilai Faktur | `#,##0` | Nilai faktur asli |
| Sisa Piutang | `#,##0` | Sisa piutang |
| Umur Piutang (Hari) | `#,##0` | Nilai `days` |
| Kesimpulan Audit | Teks | `REVIU INDIKASI ANOMALI` (highlight merah) atau `WAJAR` |

Jika tidak ada anomali terdeteksi, tabel berisi satu baris: `"Tidak ada transaksi pencilan terdeteksi"`.

---

### Sheet 6 — DBSCAN (Density-Based Spatial Clustering)

**Prinsip:** Mengelompokkan titik berdasarkan kepadatan. Transaksi di area berkerapatan rendah yang tidak masuk ke cluster manapun diberi label `−1` = **ANOMALI**.

**Parameter:** `eps=0.5`, `min_samples=4`, `algorithm='ball_tree'`

**Kelebihan:** Tidak perlu menentukan jumlah cluster; menemukan anomali struktural secara natural. Cocok untuk data dengan bentuk cluster non-spherical.

**Chunking:** Untuk data > 20.000 baris, DBSCAN dijalankan per chunk dan hasilnya digabungkan.

---

### Sheet 7 — Isolation Forest

**Prinsip:** Membangun pohon isolasi acak. Transaksi yang cepat terisolasi (kedalaman pohon pendek) dianggap pencilan — karena anomali sedikit dan berbeda, lebih mudah dipisahkan.

**Parameter:** `contamination=0.02` (2% data diasumsikan anomali), `random_state=42`

**Kelebihan:** Efisien untuk dataset besar; tidak bergantung pada asumsi distribusi data.

---

### Sheet 8 — Local Outlier Factor (LOF)

**Prinsip:** Membandingkan densitas lokal suatu transaksi dengan densitas k-tetangga terdekatnya. Transaksi dengan densitas jauh lebih rendah dari tetangganya = **ANOMALI**.

**Parameter:** `n_neighbors=20`, `contamination=0.02`, `algorithm='ball_tree'`

**Kelebihan:** Sensitif terhadap anomali lokal — mampu menemukan pencilan di area yang secara global terlihat padat.

**Chunking:** Sama seperti DBSCAN, didukung chunking untuk dataset besar.

---

### Sheet 9 — One-Class SVM

**Prinsip:** Mempelajari batas keputusan non-linear di ruang berdimensi tinggi (via kernel RBF) yang melingkupi mayoritas data normal. Titik di luar batas = **ANOMALI**.

**Parameter:** `nu=0.02` (batas atas fraksi anomali), `kernel='rbf'`, `gamma='scale'`

**Sampling:** Untuk data > 10.000 baris, model dilatih pada 10.000 sampel acak (`random_state=42`) lalu digunakan untuk memprediksi seluruh data.

**Kelebihan:** Sangat fleksibel untuk pola anomali non-linear dan non-konveks.

---

### Sheet 10 — K-Means Centroid Distance

**Prinsip:** Membagi data ke dalam 4 cluster. Transaksi dengan jarak Euclidean terbesar dari centroid clusternya = pencilan potensial.

**Parameter:** `n_clusters=4`, `n_init=3`, `random_state=42`

**Threshold:** Persentil ke-98 dari distribusi jarak seluruh transaksi. Transaksi di atas threshold = **ANOMALI**.

**Kelebihan:** Sederhana dan interpretatif; threshold berbasis persentil membuat jumlah temuan relatif stabil di ~2% data.

---

### Sheet 11 — Neural Network Autoencoder

**Prinsip:** Melatih jaringan syaraf tiruan untuk me-*rekonstruksi* input (encode → decode). Transaksi yang sulit direkonstruksi (error rekonstruksi tinggi) = pola tidak normal yang tidak dipelajari model.

**Arsitektur:** `MLPRegressor` dengan `hidden_layer_sizes=(2,)` — bottleneck dua neuron memaksa kompresi fitur esensial.

**Parameter:** `activation='relu'`, `max_iter=100`, `early_stopping=True`, `random_state=42`

**Threshold:** Persentil ke-98 dari distribusi reconstruction error. Di atas threshold = **ANOMALI**.

**Kelebihan:** Mampu menangkap interaksi kompleks non-linear antar fitur yang tidak terlihat oleh metode geometris/densitas.

---

## ⚙️ Konfigurasi `config.conf`

Terletak di `Dapur/config.conf`:

```ini
[FILTER]
SHELL
IRC
ZN
GT
FILTER
JIMCO
LAIN
TOP 1
```

Dibaca oleh kedua skrip untuk membangun `filter_produk`. Pada versi ini konfigurasi digunakan untuk validasi, namun **tidak secara aktif menghasilkan tabel kontribusi produk** dalam sheet output.

Jika `config.conf` tidak ditemukan, kedua skrip akan membuat versi default secara otomatis.

---

## 🛠️ Troubleshooting

### ❌ `Proses digagalkan. File atau folder berikut tidak ditemukan`
Periksa daftar file yang ditampilkan. Kemungkinan penyebab: folder `Dapur/` tidak ada, salah satu skrip dihapus, atau file `PTM*.xlsx` belum diletakkan di folder utama.

### ❌ `File PTM*.xlsx tidak ditemukan` (dari dalam Dapur/)
File PTM berhasil dipindahkan ke `Dapur/` oleh orkestrator, tapi salah satu skrip tidak menemukannya. Pastikan nama file dimulai dengan `PTM` dan berekstensi `.xlsx` (bukan `.xls`).

### ❌ Sheet 6–11 kosong semua / semua bertuliskan `Tidak ada transaksi pencilan terdeteksi`
Terjadi jika data terlalu sedikit (< 10 baris). Model ML membutuhkan minimum data untuk membangun distribusi normal. Periksa jumlah baris aktual di sheet MMYY file PTM.

### ❌ Error `ModuleNotFoundError: No module named 'sklearn'`
Install scikit-learn:
```bash
pip install scikit-learn
```

### ❌ Error `ModuleNotFoundError: No module named 'seaborn'` atau `'matplotlib'`
```bash
pip install matplotlib seaborn
```

### ❌ Proses sangat lambat di Sheet 9 (One-Class SVM)
SVM memiliki kompleksitas komputasi O(n²) hingga O(n³). Untuk dataset besar (> 10.000 baris), skrip otomatis mengambil 10.000 sampel untuk training. Jika masih lambat, kurangi `nu` menjadi `0.01` atau ganti kernel menjadi `linear` di `Analisis_Audit_2.py`.

### ❌ File PNG chart tidak terhapus setelah proses
Jika skrip dihentikan paksa di tengah jalan, file `temp_chart_*.png` dan `plot_anomaly_*.png` mungkin tersisa di folder `Dapur/`. Hapus secara manual sebelum menjalankan ulang.

### ❌ `Terjadi kesalahan saat menjalankan Analisis_Audit_2.py`
Tahap 1 berhasil, Tahap 2 gagal. Jalankan `Analisis_Audit_2.py` secara manual dari dalam folder `Dapur/` untuk melihat pesan error lengkap:
```bash
cd Dapur
python Analisis_Audit_2.py
```

### ❌ Seluruh kolom AR aging bernilai 0
Kolom `Umur AR base on Tgl Faktur` tidak ditemukan. Cek nama kolom di file PTM — setelah `clean_key()`, teks ini menjadi `"umarbaseontglfaktur"`. Pastikan nama kolom mengandung kata yang sama.

---

## 📌 Catatan Penting

- **Urutan eksekusi wajib dijaga** — Selalu jalankan `Buat Analisis Audit.py`, bukan skrip di dalam `Dapur/` secara langsung. Orkestrator yang memindahkan file PTM ke `Dapur/` sebelum kedua skrip dijalankan.
- **File PTM dipindahkan (bukan disalin)** — `shutil.move()` digunakan, bukan `shutil.copy()`. Setelah proses selesai, file PTM di folder utama akan hilang karena dipindahkan ke `Dapur/` dan kemudian dihapus. **Simpan cadangan file PTM** sebelum menjalankan skrip.
- **File output akan ditimpa setiap run** — `Laporan_Analisis_Prosedur_Audit.xlsx` lama langsung diganti. Ganti nama file lama jika ingin menyimpannya.
- **6 metode ML membaca ulang PTM secara mandiri** — `Analisis_Audit_2.py` tidak membaca output Tahap 1, melainkan membaca ulang file PTM dari `Dapur/`. Ini yang membuatnya bisa dijalankan mandiri untuk debugging.
- **Hasil ML bersifat indikatif, bukan konklusif** — Temuan dari Sheet 6–11 adalah *sinyal awal* yang perlu diverifikasi secara manual. Anomali statistik tidak selalu berarti fraud atau kesalahan.
- **Membandingkan Sheet 6–11** — Transaksi yang ditandai **anomali oleh beberapa metode sekaligus** memiliki prioritas reviu lebih tinggi dibanding yang hanya ditandai satu metode.
- **Z-Score paling efektif dengan data banyak bulan** — Dengan ≤ 3 bulan data, Z-Score kurang bermakna. Tambahkan lebih banyak sheet MMYY ke file PTM untuk hasil yang lebih representatif.
- **Benford paling efektif untuk dataset besar** — Idealnya ≥ 1.000 faktur. Pada dataset kecil, deviasi tinggi bisa muncul hanya karena kebetulan statistik.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
