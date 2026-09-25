## Dukungan kopi

[![Buy Me a Coffee](https://img.shields.io/badge/Buy_Me_a_Coffee-FFDD00?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://sociabuzz.com/basoarfan/tribe)


# Trace Terminal

Trace Terminal adalah Aplikasi membaca dokumen Excel kepemilikan saham untuk membandingkan kepemilikan antarperiode, memindai perubahan pada seluruh emiten, dan mencari pemegang saham berdasarkan nama.

Aplikasi mengolah file `.xlsx` di komputer dan belum mengunduh data IDX secara otomatis. Cakupan analisis mengikuti emiten, pemegang saham, dan tanggal yang tersedia dalam dokumen lokal.

## Fungsi aplikasi

| Menu | Fungsi |
| --- | --- |
| 1 — Analisa Pemegang Saham 1% | Menampilkan kepemilikan dan perubahan untuk satu kode emiten pada dokumen 1%. |
| 2 — Analisa Pemegang Saham >5% | Menganalisis kepemilikan satu emiten pada dokumen 5%, termasuk rekening efek dan perubahan kepemilikan. |
| 3 — Cari Klasifikasi Pemegang Saham | Mencari klasifikasi investor berdasarkan kode emiten, seperti pemerintah, perusahaan, dan individu. |
| 4 — Scanner Pemegang Saham 1% | Memindai perubahan kepemilikan seluruh emiten dalam dokumen 1%. |
| 5 — Scanner Pemegang Saham 5% | Memindai perubahan kepemilikan seluruh emiten dalam dokumen 5%. |
| 6 — Scanner Nama Pemegang Saham | Mencari kepemilikan berdasarkan nama pada dokumen 1%, lalu otomatis memindai dokumen 5% dengan nama dan periode yang sama. Hasil ditampilkan terpisah per sumber. Mendukung sebagian nama dan tidak membedakan huruf besar/kecil. |
| 0 — Keluar | Menutup aplikasi. |

Menu analisa kepemilikan dan scanner menyediakan periode **2, 3, 6, atau 12 bulan**, mengikuti data yang tersedia. Pada menu 3, pilihan periode membatasi jumlah tanggal laporan klasifikasi terbaru yang ditampilkan.

Khusus menu 5, pilihan **`1` = 2 Hari** membandingkan dua tanggal kepemilikan dalam laporan terbaru. Satu file 5% sudah dapat memuat kedua tanggal tersebut; keduanya tidak harus berjarak tepat dua hari kalender.

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

1. Di halaman repositori GitHub ini, pilih **Clone / Download ZIP**, lalu ekstrak arsipnya.
2. Buka PowerShell di folder hasil ekstraksi yang berisi `README.md`, `requirements.txt`, dan folder `traceterminal`. Contoh berikut menggunakan `D:\TraceTerminal`; sesuaikan dengan lokasi Anda:

```powershell
cd "D:\TraceTerminal"
```

3. Pasang dependensi dan jalankan aplikasi:

```powershell
python -m pip install -r requirements.txt
python -m traceterminal
```

Untuk pemakaian berikutnya, buka terminal di folder proyek lalu jalankan:

```powershell
python -m traceterminal
```

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

Tanggal di atas adalah contoh penamaan. Simpan file langsung di folder jenis datanya, tanpa subfolder tambahan.

### Mengganti nama file dan memperbarui tanggal

Gunakan format **`YYYY-MM-DD.xlsx`** sesuai tanggal laporan. Contoh: laporan 15 September 2026 disimpan sebagai `2026-09-15.xlsx`.

1. Periksa tanggal kepemilikan/laporan di dalam dokumen. Untuk laporan 5% yang memuat dua tanggal, gunakan tanggal kepemilikan terbaru sebagai nama file.
2. Ubah nama hasil unduhan, misalnya `peng-2026-09-15-00074-lima-persen.xlsx` menjadi `2026-09-15.xlsx`.
3. Pindahkan ke `documents/pemegang_saham_1%/` atau `documents/pemegang_saham_5%/` sesuai jenisnya.
4. Saat laporan tanggal baru tersedia, tambahkan file baru dan pertahankan file lama untuk perbandingan riwayat.


## Daftar kode broker dan Rekening Efek

Referensi utama: [Profil Anggota Bursa — IDX](https://www.idx.co.id/id/anggota-bursa-dan-partisipan/profil-anggota-bursa/). 

Pada analisa perubahan dan scanner 5%, kolom **Rekening Efek** menampilkan kode hasil pemetaan dari **Nama Pemegang Rekening Efek**. Pencocokan menoleransi variasi `PT`, `PT.`, akhiran `Tbk`, huruf besar/kecil, dan spasi. Nama yang belum dipetakan tetap ditampilkan sesuai sumber.


## Screenshot terminal

<div align="center">
<img src="https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEj_W8tEN9o5OPgMI2E10t6fOKCZmFhgK6eiZd0W0MmSNglHRW0SmgFFauaIj-T1I58gpROWubcekUvW_gP71cgk7ey-IcnTgQHLBZmZYL2u6MGOeit511yjFZCm2iweZ5kba1_IFf0g0R1HSBbUOIhotQ_t5T56BMVuz4OTgEs_G69cOmQ5PayZ44-mWosM/s1919/Screenshot%202026-09-20%20113622.png" alt="Menu utama Trace Terminal" width="100%">



</div>

## Catatan

Trace Terminal hanya digunakan untuk membaca, memproses, dan memfilter file `.xlsx` yang diperoleh dari situs IDX.

Seluruh pemrosesan file dilakukan secara lokal di komputer pengguna. Aplikasi ini tidak mengirimkan isi file ke pihak ketiga.

Jika terdapat kesalahan dalam pembacaan atau hasil analisis, harap periksa kembali file sumber dan hasil yang ditampilkan. Pengguna juga dapat meninjau kode sumber repository ini menggunakan ChatGPT, Claude, atau alat analisis kode lainnya sebelum menggunakan terminal ini



