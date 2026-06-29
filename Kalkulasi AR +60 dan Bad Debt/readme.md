# 📊 Monitoring AR Outstanding

> **Dashboard analitik piutang otomatis — dari data Accurate ke laporan Excel lengkap dengan diagram tren dan kontribusi produk**

Dua skrip Python yang bekerja secara berurutan: **`Buat_Data.py`** membaca file ekspor AR bulanan dari Accurate (`PTM*.xlsx`), menghitung metrik piutang (Total AR, AR 60 Hari Up, Bad Debt), mengelompokkan per bulan dan kuartal, lalu menghasilkan `Monitoring.xlsx` berisi ringkasan rekap dan sheet detail per bulan. Dilanjutkan **`Diagram.py`** yang menyuntikkan diagram batang (tren bulanan) dan diagram kue (kontribusi produk AR-60) secara dinamis ke dalam file yang sama — siap digunakan sebagai laporan monitoring manajemen.

---

## 📋 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Detail](#-alur-kerja-detail)
- [Logika Kalkulasi Metrik](#-logika-kalkulasi-metrik)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Output & Struktur Laporan](#-output--struktur-laporan)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## ✨ Fitur Utama

- **Auto-detect sheet bulanan** — Membaca semua sheet berformat `MMYY` (mis. `0125` untuk Jan-2025) dari file PTM, mengurutkannya secara kronologis, dan memproses semuanya sekaligus.
- **Auto-detect header** — Mencari posisi baris header secara dinamis, sehingga tidak terganggu meski ekspor Accurate menambahkan baris judul di atasnya.
- **Kalkulasi metrik lengkap** — Menghitung 10 metrik per bulan: Total AR Outstanding, AR 60 Hari Up, Bad Debt, persentase berbasis amount, jumlah customer distinct, persentase customer berisiko, hingga kontribusi per produk.
- **Pengelompokan kuartal otomatis** — Menggabungkan bulan-bulan ke dalam kuartal (Q1/Q2/Q3/Q4) secara otomatis berdasarkan data yang tersedia, dan menghitung rata-rata per kuartal.
- **Analisis penurunan (decline tracking)** — Membandingkan rata-rata kuartal berjalan vs kuartal tertua sebagai baseline, untuk memantau tren perbaikan atau penurunan kualitas piutang.
- **Kontribusi produk per filter** — Mengelompokkan nilai piutang AR-60 dan Bad Debt berdasarkan kategori produk dari `config.conf`, dengan baris Total dan Selisih otomatis.
- **Sheet detail per bulan** — Setiap bulan menghasilkan sheet `Detail_MMYY` berisi tiga tabel terpisah: baris AR 60 Hari Up, baris Bad Debt, dan baris di luar kategori analisis.
- **Diagram dinamis injeksi** — Diagram batang (tren 3 metrik bulanan) dan diagram kue (kontribusi produk bulan terakhir) dibuat berdasarkan data aktual, bukan template statis.
- **Auto-fit kolom** — Lebar kolom di semua sheet menyesuaikan otomatis dengan konten terpanjang di setiap kolom.
- **Ekslusi data FRAUD** — Baris dengan kata `FRAUD` di kolom `Nama Penjual` otomatis dikeluarkan dari kalkulasi AR aging.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl numpy
```

| Library | Kegunaan |
|---|---|
| `pandas` | Baca sheet Excel, parsing data, kalkulasi agregat |
| `openpyxl` | Buat/edit workbook, styling sel, sisipkan chart |
| `numpy` | Penanganan NaN dan konversi tipe numerik |
| `configparser` | Baca `config.conf` (sudah termasuk standard library) |
| `re`, `glob`, `datetime`, `os` | Utilitas standar (sudah termasuk standard library) |

---

## 📁 Struktur Folder & File

```
📦 Monitoring-AR-Outstanding/
│
├── 📄 Buat_Data.py          ← Skrip utama. Jalankan PERTAMA
├── 📄 Diagram.py            ← Skrip diagram. Jalankan KEDUA setelah Buat_Data.py
├── 📄 config.conf           ← Konfigurasi filter kategori produk
│
├── 📄 PTM*.xlsx             ← [INPUT] File data AR bulanan dari Accurate
│                                Nama bebas asal diawali "PTM" (mis. PTM_2025.xlsx)
│
└── 📄 Monitoring.xlsx       ← [OUTPUT] Dihasilkan otomatis oleh kedua skrip
```

> Semua file diletakkan dalam **satu folder yang sama**. Tidak ada subfolder.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input PTM

1. Export laporan AR Aging dari **Accurate** ke format `.xlsx`.
2. Simpan file dengan nama yang **diawali `PTM`** (contoh: `PTM_2025.xlsx`, `PTM Jan-Des.xlsx`).
3. Letakkan di folder yang sama dengan kedua skrip.

File Excel tersebut harus memiliki:
- Satu atau lebih sheet dengan nama format **`MMYY`** (4 digit) → contoh: `0125` untuk Januari 2025, `1224` untuk Desember 2024.
- Setiap sheet memiliki kolom: `Sisa Piutang`, `Umur AR Base on Tgl Japo`, `Nama Penjual`, `Nama Pelanggan`, `Negara Pelanggan`, `Kontak Pelanggan`.

### Langkah 2 — Sesuaikan `config.conf`

Buka `config.conf` dan sesuaikan daftar filter produk sesuai kategori yang relevan. Lihat detail di bagian [Konfigurasi](#-konfigurasi-configconf).

### Langkah 3 — Jalankan `Buat_Data.py`

```bash
python Buat_Data.py
```

Skrip akan membaca semua sheet `MMYY`, menghitung semua metrik, dan menghasilkan `Monitoring.xlsx`.

### Langkah 4 — Jalankan `Diagram.py`

```bash
python Diagram.py
```

Skrip akan membuka `Monitoring.xlsx` yang sudah ada, menyuntikkan diagram tren dan diagram kue, lalu menyimpan kembali file yang sama.

### Langkah 5 — Buka hasil

Buka **`Monitoring.xlsx`** — file siap digunakan sebagai laporan monitoring manajemen.

---

## 🔄 Alur Kerja Detail

### `Buat_Data.py`

```
[Mulai]
   │
   ├─── Baca config.conf
   │       Ambil daftar filter_produk dari seksi [FILTER]
   │       Pastikan "LAIN" selalu ada sebagai fallback kategori
   │
   ├─── Deteksi file PTM*.xlsx
   │       Cari file dengan glob("PTM*.xlsx") → ambil yang pertama ditemukan
   │       Jika tidak ada → berhenti dengan pesan error
   │
   ├─── Baca & urutkan sheet MMYY
   │       Filter sheet dengan nama 4 digit (regex ^\d{4}$)
   │       Urutkan secara kronologis berdasarkan parse datetime "%m%y"
   │       Tentukan label kuartal (Q1/Q2/Q3/Q4 + tahun) per bulan
   │
   ├─── Untuk setiap sheet bulan (MMYY):
   │       │
   │       ├─ Deteksi otomatis posisi header
   │       │       Scan baris sampai menemukan kolom kunci
   │       │       (Sisa Piutang atau Nama Penjual)
   │       │
   │       ├─ Hitung kolom "days"
   │       │       Ekstrak angka dari "Umur AR Base on Tgl Japo"
   │       │       (mis. "75 hari" → 75)
   │       │
   │       ├─ Tentukan customer ID per baris
   │       │       Prioritas: Negara Pelanggan (jika ada) → Nama Pelanggan
   │       │
   │       ├─ Hitung 10 metrik bulanan
   │       │       (lihat bagian Logika Kalkulasi Metrik)
   │       │
   │       ├─ Hitung kontribusi produk per filter
   │       │       Cocokkan "Kontak Pelanggan" dengan keyword dari config.conf
   │       │       Jika tidak ada yang cocok → masuk bucket "LAIN"
   │       │
   │       └─ Buat sheet Detail_MMYY
   │               Tiga tabel: AR 60 Hari, Bad Debt, Data Lainnya
   │               Dengan styling header biru + border tipis
   │
   ├─── Bangun sheet "Rekap Monitoring"
   │       │
   │       ├─ Tulis baris header kolom
   │       │       [KETERANGAN | Formula | Jan-25 | Feb-25 | ... | | Q1 2025 | ... | | Penurunan vs Q1]
   │       │
   │       ├─ Tulis 13 baris metrik utama
   │       │       Nilai bulanan + rata-rata kuartal + persentase penurunan
   │       │       Format angka: amount (#,##0), persen (0.00%), count (#,##0)
   │       │
   │       ├─ Tulis tabel kontribusi "Daftar Kontribusi Barang AR-60 HARI"
   │       │       Satu baris per produk + baris Total + baris Selisih
   │       │
   │       └─ Tulis tabel kontribusi "Daftar Kontribusi Barang BADDEBT"
   │               Struktur sama seperti tabel AR-60
   │
   ├─── Auto-fit lebar kolom untuk semua sheet
   │
   └─── Simpan sebagai Monitoring.xlsx ✅
```

### `Diagram.py`

```
[Mulai]
   │
   ├─── Buka Monitoring.xlsx → sheet "Rekap Monitoring"
   │
   ├─── Deteksi posisi baris header ("KETERANGAN")
   │       dan posisi kolom bulan (berhenti di kolom kosong)
   │
   ├─── Deteksi nomor baris tiga metrik kunci:
   │       TOTAL AR OUTSTANDING, AR 60 HARI UP, AR BAD DEBT
   │
   ├─── Buat "Tabel Acuan Diagram" (bridge table)
   │       Salin nilai tiga metrik ke area baru di bawah data
   │       (Dibutuhkan agar chart merujuk ke data flat, bukan data terpencar)
   │
   ├─── Buat BarChart (Diagram Batang)
   │       Judul: "Tren AR Outstanding & Umur Piutang Bulanan"
   │       Series: 3 metrik utama | Kategori: nama bulan
   │       Ukuran: 18 × 10 (lebar × tinggi cm)
   │
   ├─── Deteksi area data pie chart
   │       Cari blok "Daftar Kontribusi Barang AR-60 HARI"
   │       Ambil baris produk hingga baris "Total"
   │       Ambil kolom bulan terakhir sebagai data
   │
   ├─── Buat PieChart (Diagram Kue)
   │       Judul: "Kontribusi Barang pada AR 60 Hari Up ([Bulan Terakhir])"
   │       Label: nama produk | Data: nilai kolom bulan terakhir
   │       Ukuran: 14 × 10 (lebar × tinggi cm)
   │
   ├─── Sisipkan kedua chart ke sheet
   │       BarChart di kolom A | PieChart di kolom L
   │       (di bawah semua data yang ada)
   │
   ├─── Tulis keterangan analisis (teks naratif)
   │       Penjelasan cara membaca Diagram Batang
   │       Penjelasan cara membaca Diagram Kue
   │
   └─── Simpan kembali ke Monitoring.xlsx ✅
```

---

## 🧮 Logika Kalkulasi Metrik

Berikut definisi operasional 10 metrik yang dihitung per bulan dari setiap sheet:

| Kode | Metrik | Definisi |
|---|---|---|
| **A** | TOTAL AR OUTSTANDING (AMOUNT) | Jumlah `Sisa Piutang` seluruh baris (termasuk piutang baru / lancar) |
| **B** | AR 60 HARI UP (AMOUNT) | Jumlah `Sisa Piutang` baris dengan `days` antara 60–364 dan bukan FRAUD |
| **C** | AR BAD DEBT (365 UP) AMOUNT | Jumlah `Sisa Piutang` baris dengan `days` ≥ 365 dan bukan FRAUD |
| **B:A** | AR 60 HARI UP IN AMOUNT (%) | B ÷ A |
| **C:A** | AR BAD DEBT IN AMOUNT (365 UP) % | C ÷ A |
| **D** | JUMLAH CUSTOMER (DISTINCT COUNT) | Distinct count `cust_id` seluruh baris |
| **E** | JUMLAH CUST AR 60 HARI UP | Distinct count `cust_id` dari baris AR 60 Hari Up |
| **F** | JUMLAH CUST AR BAD DEBT | Distinct count `cust_id` dari baris Bad Debt |
| **E:D** | AR 60 HARI UP (%) | E ÷ D |
| **F:D** | AR BAD DEBT (365 UP) % | F ÷ D |

**Catatan definisi:**
- `days` = angka numerik yang diekstrak dari kolom `Umur AR Base on Tgl Japo` (contoh: `"75 hari"` → `75`, `"-10 hari"` → `-–10`)
- `cust_id` = nilai kolom `Negara Pelanggan` (jika terisi), atau fallback ke `Nama Pelanggan` (jika `Negara Pelanggan` kosong/NaN)
- **FRAUD** = baris di mana kolom `Nama Penjual` mengandung kata `"FRAUD"` (case-insensitive) → dikeluarkan dari semua kalkulasi aging
- **Rata-rata kuartal** = rata-rata nilai bulan-bulan yang jatuh dalam kuartal tersebut
- **Penurunan vs Q terlama** = `(nilai_q_berjalan − nilai_q_tertua) / nilai_q_tertua`, hanya dihitung untuk 4 metrik: AR 60 (amount), Bad Debt (amount), jumlah customer AR 60, jumlah customer Bad Debt

**Pencocokan produk (untuk tabel kontribusi):**
- Setiap baris AR (yang masuk kategori 60 Hari atau Bad Debt) dicocokkan ke produk berdasarkan kolom `Kontak Pelanggan`
- Pencocokan dilakukan secara berurutan dari atas ke bawah daftar `filter_produk` di `config.conf` (match pertama yang ditemukan dipakai)
- Baris yang tidak cocok dengan produk manapun masuk ke bucket `"LAIN"`

---

## ⚙️ Konfigurasi `config.conf`

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

### Cara kerja

- Setiap baris di bawah `[FILTER]` adalah satu **kategori produk**.
- Nama kategori akan dicocokkan (substring match, case-insensitive) terhadap nilai kolom **`Kontak Pelanggan`** di setiap baris data.
- Pencocokan dilakukan **berurutan dari atas ke bawah** — kategori yang lebih atas mendapat prioritas lebih tinggi.
- Baris data yang tidak cocok dengan kategori manapun otomatis masuk ke bucket **`LAIN`**.
- Jika `LAIN` tidak dicantumkan di `config.conf`, program tetap menambahkannya secara otomatis sebagai fallback.

### Aturan penulisan

| Aturan | Keterangan |
|---|---|
| Satu kategori per baris | Tidak boleh ada lebih dari satu keyword dalam satu baris |
| Urutan berpengaruh | Tulis kategori yang lebih spesifik lebih atas dari yang lebih umum |
| `LAIN` opsional | Bisa dicantumkan untuk mengatur posisinya di tabel, atau dibiarkan otomatis ditambahkan di akhir |
| Case-insensitive | `SHELL`, `shell`, `Shell` akan terbaca sama |

### Contoh kustomisasi

```ini
[FILTER]
SHELL
CASTROL
PERTAMINA
FEDERAL
LAIN
```

> Jika `config.conf` tidak ditemukan saat `Buat_Data.py` dijalankan, file akan dibuat otomatis dengan isi default: `SHELL, IRC, ZN, GT, FILTER, JIMCO, LAIN, TOP 1, OLI`.

---

## 📤 Output & Struktur Laporan

### File output: `Monitoring.xlsx`

Berisi sheet-sheet berikut:

#### Sheet 1 — `Rekap Monitoring`

Sheet utama dengan layout horizontal (bulan tersusun ke kanan sebagai kolom). Terdiri dari tiga zona:

**Zona 1 — Header kolom (baris 3)**

```
KETERANGAN | Formula | Jan-25 | Feb-25 | Mar-25 | [kosong] | Q1 2025 | Q2 2025 | [kosong] | Penurunan vs Q1 2025 (Q1 2025) | Penurunan vs Q1 2025 (Q2 2025)
```

**Zona 2 — Tabel metrik utama (10 metrik + spacer)**

| Kelompok | Metrik | Format |
|---|---|---|
| Amount | TOTAL AR OUTSTANDING | `#,##0` |
| Amount | AR 60 HARI UP | `#,##0` |
| Amount | AR BAD DEBT (365 UP) | `#,##0` |
| Persentase Amount | AR 60 HARI UP (%) | `0.00%` |
| Persentase Amount | AR BAD DEBT (%) | `0.00%` |
| Customer Count | JUMLAH CUSTOMER | `#,##0` |
| Customer Count | JUMLAH CUST AR 60 | `#,##0` |
| Customer Count | JUMLAH CUST BAD DEBT | `#,##0` |
| Persentase Customer | AR 60 HARI UP (%) | `0.00%` |
| Persentase Customer | AR BAD DEBT (%) | `0.00%` |

Kolom kuartal berisi **rata-rata** nilai bulan dalam kuartal tersebut. Kolom penurunan berisi persentase selisih relatif terhadap kuartal terlama (format `0.00%`).

**Zona 3 — Tabel kontribusi produk**

Dua tabel horizontal (AR-60 dan Bad Debt), masing-masing berisi:
- Satu baris per kategori produk dari `config.conf`
- Baris `Total` (penjumlahan semua produk)
- Baris `Selisih` (Total produk − nilai metrik utama, untuk validasi kelengkapan klasifikasi)

**Di bawah data** (ditambahkan oleh `Diagram.py`):
- Tabel Acuan Diagram (bridge table untuk referensi chart)
- Diagram Batang — tren 3 metrik kunci per bulan
- Diagram Kue — kontribusi produk AR-60 bulan terakhir
- Narasi keterangan dan panduan pembacaan kedua diagram

#### Sheet 2, 3, 4, ... — `Detail_MMYY`

Satu sheet per bulan (mis. `Detail_0125`), berisi tiga tabel:

| Tabel | Isi | Warna Header |
|---|---|---|
| AR 60 HARI UP | Baris dengan `days` 60–364, bukan FRAUD | Biru gelap (#366092), teks putih |
| BAD DEBT (365 UP) | Baris dengan `days` ≥ 365, bukan FRAUD | Biru gelap (#366092), teks putih |
| DATA LAINNYA | Semua baris di luar dua kategori di atas | Biru gelap (#366092), teks putih |

Setiap tabel menyertakan semua kolom asli dari file PTM, ditambah tiga kolom kalkulasi: `Calculated_Days`, `Calculated_Cust_ID`, dan `Is_Fraud`.

---

## 🛠️ Troubleshooting

### ❌ `File PTM*.xlsx tidak ditemukan`
Pastikan file data Accurate ada di folder yang sama dengan `Buat_Data.py` dan namanya diawali huruf `PTM` (case-sensitive di beberapa OS). Coba ganti nama file menjadi `PTM_data.xlsx` atau serupa.

### ❌ `Tidak ada sheet bulanan (format MMYY) yang valid`
Program mencari sheet dengan nama tepat 4 digit angka (`0125`, `1224`, dst). Pastikan nama sheet di file PTM tidak mengandung karakter lain (spasi, huruf, tanda baca). Contoh nama sheet yang valid: `0125`, `0225`, `1224`.

### ❌ Kolom tidak ditemukan / data kosong di metrik
Skrip mendeteksi header dengan mencocokkan nama kolom berikut (case-insensitive, karakter non-alphanumeric diabaikan):
- `Sisa Piutang`
- `Umur AR Base on Tgl Japo`
- `Nama Penjual`
- `Nama Pelanggan`
- `Negara Pelanggan`
- `Kontak Pelanggan`

Pastikan nama kolom di file PTM **mengandung kata-kata tersebut** (tidak harus identik, asalkan setelah strip tanda baca kata kuncinya cocok). Jika ada kolom yang hilang, baris sheet tersebut akan dilewati.

### ❌ `Error: Struktur data tidak dikenali. Baris KETERANGAN tidak ditemukan` (dari `Diagram.py`)
`Diagram.py` harus dijalankan **setelah** `Buat_Data.py` berhasil menghasilkan `Monitoring.xlsx`. Jika `Monitoring.xlsx` belum ada atau rusak, jalankan ulang `Buat_Data.py` terlebih dahulu.

### ❌ Diagram Kue tidak muncul di output
Diagram Kue hanya dibuat jika skrip berhasil menemukan blok tabel `"Daftar Kontribusi Barang AR-60 HARI"` dan baris `"Total"` di dalam sheet Rekap Monitoring. Pastikan `Buat_Data.py` berjalan tanpa error dan file PTM memiliki data di kolom `Kontak Pelanggan`.

### ❌ Kolom `Kontak Pelanggan` tidak ada — semua masuk bucket "LAIN"
Jika kolom `Kontak Pelanggan` tidak ditemukan di sheet PTM, seluruh nilai kontribusi produk akan masuk ke bucket `LAIN`. Tabel kontribusi tetap dibuat, namun hanya baris `LAIN` yang bernilai.

### ❌ Nilai `Selisih` besar di tabel kontribusi produk
Baris `Selisih` menunjukkan perbedaan antara total yang terklafikasikan per produk dengan nilai metrik utama (AR 60 / Bad Debt). Selisih besar menandakan ada data yang tidak terklasifikasi ke produk manapun karena kolom `Kontak Pelanggan`-nya kosong atau formatnya tidak cocok dengan keyword di `config.conf`.

---

## 📌 Catatan Penting

- **Urutan eksekusi wajib dijaga** — `Buat_Data.py` harus selesai dulu sebelum `Diagram.py` dijalankan. Keduanya beroperasi pada file yang sama (`Monitoring.xlsx`).
- **Hanya satu file PTM** — Jika ada lebih dari satu file `PTM*.xlsx` di folder, program mengambil yang pertama ditemukan oleh `glob`. Pastikan hanya ada satu file PTM aktif.
- **Sheet non-MMYY diabaikan** — Sheet di luar format 4 digit angka (mis. `Summary`, `Rekap`, `Template`) tidak akan diproses oleh `Buat_Data.py`.
- **`Monitoring.xlsx` akan ditimpa** — Setiap kali `Buat_Data.py` dijalankan, file output lama akan tertimpa. Simpan salinan jika diperlukan sebelum menjalankan ulang.
- **`Diagram.py` bersifat aditif** — Jika `Diagram.py` dijalankan dua kali pada file yang sama, chart dan narasi akan ditambahkan dua kali. Selalu mulai dari `Buat_Data.py` yang segar jika ingin mengulang proses.
- **Kategori produk bersifat kumulatif** — Penambahan kategori baru di `config.conf` akan langsung berpengaruh pada run berikutnya tanpa perubahan kode.
- **Selalu verifikasi baris Selisih** — Sebelum menyebarkan laporan, pastikan nilai di baris `Selisih` mendekati nol. Selisih yang signifikan mengindikasikan data yang tidak terklasifikasi dan perlu pengecekan kualitas di kolom `Kontak Pelanggan`.

---

## 👤 Author

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
