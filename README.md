# Trace Terminal

Trace Terminal adalah aplikasi terminal **khusus saham Indonesia yang tercatat di Bursa Efek Indonesia (BEI / IDX)**. Aplikasi membaca dokumen Excel kepemilikan saham untuk membandingkan kepemilikan antarperiode, memindai perubahan pada seluruh emiten, dan mencari pemegang saham berdasarkan nama.

Aplikasi mengolah file `.xlsx` di komputer dan belum mengunduh data IDX secara otomatis. Cakupan analisis mengikuti emiten, pemegang saham, dan tanggal yang tersedia dalam dokumen lokal.

## Fungsi aplikasi

| Menu | Fungsi |
| --- | --- |
| 1 — Analisa Pemegang Saham 1% | Menampilkan kepemilikan dan perubahan untuk satu kode emiten pada dokumen 1%. |
| 2 — Analisa Pemegang Saham >5% | Menganalisis kepemilikan satu emiten pada dokumen 5%, termasuk rekening efek dan perubahan kepemilikan. |
| 3 — Cari Klasifikasi Pemegang Saham | Mencari klasifikasi investor berdasarkan kode emiten, seperti pemerintah, perusahaan, dan individu. |
| 4 — Scanner Pemegang Saham 1% | Memindai perubahan kepemilikan seluruh emiten dalam dokumen 1%. |
| 5 — Scanner Pemegang Saham 5% | Memindai perubahan kepemilikan seluruh emiten dalam dokumen 5%. |
| 6 — Scanner Nama Pemegang Saham | Mencari perubahan kepemilikan berdasarkan nama pada seluruh dokumen 1%. Mendukung sebagian nama dan tidak membedakan huruf besar/kecil. |
| 0 — Keluar | Menutup aplikasi. |

Menu analisa kepemilikan dan scanner menyediakan periode **2, 3, 6, atau 12 bulan**, mengikuti data yang tersedia. Pada menu 3, pilihan periode membatasi jumlah tanggal laporan klasifikasi terbaru yang ditampilkan.

Khusus menu 5, pilihan **`1` = 2 Hari** membandingkan dua tanggal kepemilikan dalam laporan terbaru. Satu file 5% sudah dapat memuat kedua tanggal tersebut; keduanya tidak harus berjarak tepat dua hari kalender.

Kolom **Estimasi Lot**, termasuk pada menu 6, dihitung sebagai **perubahan jumlah saham ÷ 100**. Angka positif menunjukkan penambahan kepemilikan, sedangkan angka negatif menunjukkan pengurangan. Hasil analisis dan log sesi disimpan sebagai JSON di `output/logs/`.

## Instal Python di Windows

Panduan berikut menggunakan PowerShell.

1. Buka [halaman unduh resmi Python](https://www.python.org/downloads/).
2. Unduh dan pasang **Python Install Manager** untuk Windows, lalu ikuti petunjuk pemasangannya.
3. Buka PowerShell baru dan pasang Python 3.14:

```powershell
py install 3.14
```

4. Periksa instalasinya:

```powershell
python --version
python -m pip --version
```

Jika menggunakan installer Python klasik yang menyediakan pilihan **Add python.exe to PATH**, aktifkan pilihan tersebut saat instalasi. Jika `python` belum dikenali, buka ulang terminal dan coba `py --version`. Lihat juga [panduan resmi Python untuk Windows](https://docs.python.org/3/using/windows.html).

## Instal dan jalankan terminal

1. Di halaman repositori GitHub ini, pilih **Code → Download ZIP**, lalu ekstrak arsipnya.
2. Buka PowerShell di folder hasil ekstraksi yang berisi `README.md`, `requirements.txt`, dan folder `traceterminal`. Contoh berikut menggunakan `D:\TraceTerminal`; sesuaikan dengan lokasi Anda:

```powershell
cd "D:\TraceTerminal"
```

3. Pasang dependensi dan jalankan aplikasi:

```powershell
python -m pip install -r requirements.txt
python -m traceterminal
```

Dependensinya adalah **Rich** untuk tampilan terminal dan **openpyxl** untuk membaca Excel. Instalasi dependensi membutuhkan internet; analisis dokumen lokal dapat dijalankan tanpa internet.

4. Masukkan nomor menu, tekan **Enter**, lalu isi kode emiten atau nama pemegang saham jika diminta. Pilih periode analisis. Gunakan menu **0** atau **Ctrl+C** untuk keluar.

Untuk pemakaian berikutnya, buka terminal di folder proyek lalu jalankan:

```powershell
python -m traceterminal
```

Selalu jalankan dari folder proyek karena lokasi `documents/` dan `output/` mengikuti direktori kerja terminal.

### Opsional: lingkungan Python terpisah

Untuk memisahkan dependensi dari proyek Python lain, jalankan dari folder proyek:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m traceterminal
```

Cara ini tidak memerlukan aktivasi skrip PowerShell. Gunakan perintah terakhir setiap kali menjalankan aplikasi dengan lingkungan tersebut.

## Unduh data kepemilikan saham 1% dan 5%

Sumber unduhan: [Data Kepemilikan Saham — IDX](https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/).

1. Buka halaman tersebut melalui browser.
2. Cari laporan kepemilikan **1%** dan **5%** untuk tanggal atau periode yang dibutuhkan. Jika tersedia pilihan tanggal/periode, ubah sesuai riwayat yang ingin dianalisis.
3. Unduh file Excel untuk masing-masing jenis laporan. Untuk scanner seluruh emiten, gunakan laporan lengkap seluruh emiten yang tersedia pada tanggal tersebut.
4. Ulangi untuk semua tanggal yang tersedia dalam periode analisis Anda, lalu simpan laporan 1% dan 5% di folder terpisah.

“Seluruh emiten” pada scanner berarti seluruh emiten dalam file yang Anda simpan. Memilih 12 bulan tidak mengunduh atau melengkapi bulan yang belum tersedia. Dokumen ambang kepemilikan 1% dan 5% juga tidak mencakup setiap pemegang saham tanpa batas kepemilikan.

### Folder penyimpanan

```text
TraceTerminal/
├── README.md
├── requirements.txt
├── traceterminal/
├── documents/
│   ├── pemegang_saham_1%/
│   │   ├── 2026-08-31.xlsx
│   │   └── 2026-09-15.xlsx
│   ├── pemegang_saham_5%/
│   │   ├── 2026-09-14.xlsx
│   │   └── 2026-09-15.xlsx
│   └── klasifikasi_pemegang_saham/
│       └── 2026-08-31.xlsx
└── output/
    └── logs/
```

Tanggal di atas adalah contoh penamaan. Simpan file langsung di folder jenis datanya, tanpa subfolder tambahan. Menu 3 memerlukan dokumen klasifikasi investor yang memuat kolom seperti `DATE`, `SHARE CODE`, dan `ISSUER NAME`, beserta kolom klasifikasinya.

### Mengganti nama file dan memperbarui tanggal

Gunakan format **`YYYY-MM-DD.xlsx`** sesuai tanggal laporan. Contoh: laporan 15 September 2026 disimpan sebagai `2026-09-15.xlsx`.

1. Periksa tanggal kepemilikan/laporan di dalam dokumen. Untuk laporan 5% yang memuat dua tanggal, gunakan tanggal kepemilikan terbaru sebagai nama file.
2. Ubah nama hasil unduhan, misalnya `laporan-unduhan.xlsx` menjadi `2026-09-15.xlsx`.
3. Pindahkan ke `documents/pemegang_saham_1%/` atau `documents/pemegang_saham_5%/` sesuai jenisnya.
4. Saat laporan tanggal baru tersedia, tambahkan file baru dan pertahankan file lama untuk perbandingan riwayat.
5. Jika ada revisi untuk tanggal yang sama, simpan cadangan di luar folder data lalu ganti file tanggal tersebut dengan revisinya. Hindari duplikat seperti `2026-09-15 (1).xlsx` dalam folder yang dibaca aplikasi.
6. Jalankan kembali menu analisis untuk membaca data yang diperbarui.

Jangan mengubah tanggal di dalam Excel atau sekadar mengganti nama laporan lama menjadi tanggal baru. Nama file 5% digunakan untuk menentukan urutan dan periode laporan; tanggal serta isinya harus sesuai. Pertahankan struktur kolom dan lembar Excel asli. Mengganti ekstensi PDF atau `.xls` menjadi `.xlsx` tidak mengonversi formatnya.

### Mengganti tanggal atau URL unduhan

Untuk mengambil tanggal lain, kembali ke [halaman data kepemilikan IDX](https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/), pilih laporan tanggal yang dibutuhkan, lalu unduh file yang ditautkan untuk tanggal tersebut.

Jika menyimpan URL unduhan langsung di bookmark atau catatan, salin ulang tautan file dari halaman IDX untuk setiap laporan. Jangan menganggap URL baru dapat diperoleh hanya dengan mengganti angka tanggal pada URL lama. Perubahan URL sumber tidak memerlukan perubahan kode aplikasi karena aplikasi membaca file lokal.

Halaman IDX dapat membatasi akses otomatis; saat panduan ini diperbarui, pemeriksaan otomatis menerima HTTP 403. Gunakan browser untuk membuka halaman sumber dan mengunduh dokumen yang tersedia.

## Daftar kode broker dan Rekening Efek

Referensi utama: [Profil Anggota Bursa — IDX](https://www.idx.co.id/id/anggota-bursa-dan-partisipan/profil-anggota-bursa/). Referensi tambahan: [Daftar Perusahaan Efek — KSEI](https://web.ksei.co.id/services/participants/brokers?setLocale=id-ID).

Pemetaan lokal berada pada `BROKER_CODE_MAP` di [traceterminal/analysis/shareholder5.py](traceterminal/analysis/shareholder5.py). Untuk memperbaruinya, periksa nama dan kode pada referensi, ubah atau tambahkan pasangan nama/kode, lalu mulai ulang aplikasi.

Pada analisa perubahan dan scanner 5%, kolom **Rekening Efek** menampilkan kode hasil pemetaan dari **Nama Pemegang Rekening Efek**. Pencocokan menoleransi variasi `PT`, `PT.`, akhiran `Tbk`, huruf besar/kecil, dan spasi. Nama yang belum dipetakan tetap ditampilkan sesuai sumber.

Pemetaan khusus yang dikonfigurasi atas permintaan pengguna:

| Nama pada dokumen | Kode tampilan |
| --- | --- |
| PT BANK DBS INDONESIA | DP |
| BANK CENTRAL ASIA Tbk, PT | SQ |
| BANK RAKYAT INDONESIA (PERSERO), PT | OD |

Ketiga entri tersebut merupakan alias aplikasi. Nama asli tetap disimpan dalam `NAMA_PEMEGANG_REKENING_EFEK_ASLI` pada hasil analisis. Bank kustodian lain dan rekening dengan keterangan tambahan tetap memakai nama sumber jika belum ada pemetaan yang cocok.

## Screenshot terminal — menu utama

Tempat URL screenshot pertama: **belum diisi**. Unggah gambar ke repositori atau GitHub, kemudian ganti `URL_SCREENSHOT_MENU_UTAMA` pada contoh berikut dan letakkan di luar blok kode agar gambar tampil:

```markdown
![Menu utama Trace Terminal](URL_SCREENSHOT_MENU_UTAMA)
```

## Screenshot terminal — hasil analisis

Tempat URL screenshot kedua: **belum diisi**. Ganti `URL_SCREENSHOT_HASIL_ANALISIS` dengan URL gambar hasil analisa atau scanner, lalu letakkan di luar blok kode:

```markdown
![Hasil analisis Trace Terminal](URL_SCREENSHOT_HASIL_ANALISIS)
```

## Dukungan kopi

[![Buy Me a Coffee](https://img.shields.io/badge/Buy_Me_a_Coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](#pengaturan-tautan-kopi)

### Pengaturan tautan kopi

Tombol di atas masih mengarah ke bagian ini karena URL dukungan pemilik repositori belum diisi. Ganti tujuan `#pengaturan-tautan-kopi` dengan URL akun Buy Me a Coffee, Ko-fi, Saweria, atau layanan dukungan Anda. Setelah tautan asli terpasang, bagian petunjuk ini dapat dihapus.

## Mengatasi kendala umum

| Kendala | Langkah pemeriksaan |
| --- | --- |
| `python` tidak dikenali | Buka ulang terminal setelah instalasi. Coba `py --version` dan, jika berhasil, gunakan `py` sebagai pengganti `python` pada perintah. |
| `No module named traceterminal` | Pastikan terminal berada di folder yang berisi direktori `traceterminal`, bukan di dalam direktori tersebut. |
| Modul `rich` atau `openpyxl` tidak ditemukan | Jalankan kembali `python -m pip install -r requirements.txt` menggunakan Python yang sama dengan yang menjalankan aplikasi. |
| Dokumen tidak ditemukan | Periksa direktori kerja, nama folder, dan pastikan file `.xlsx` berada langsung dalam folder data yang benar. |
| Hasil scanner kosong | Periksa tanggal, kode emiten/nama pemegang saham, serta apakah ada perubahan kepemilikan. Scanner hanya menampilkan perubahan sesuai data yang tersedia. |
| Tanggal 5% tidak sesuai | Cocokkan nama file `YYYY-MM-DD.xlsx` dengan tanggal terbaru pada laporan dan periksa kemungkinan file duplikat. |

Jika muncul `UnicodeEncodeError` saat menampilkan simbol terminal, atur keluaran Python ke UTF-8 pada sesi PowerShell tersebut lalu jalankan kembali:

```powershell
$env:PYTHONIOENCODING = "utf-8"
python -m traceterminal
```
