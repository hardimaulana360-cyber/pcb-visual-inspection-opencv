# Sistem Inspeksi Visual Pemasangan Komponen PCB

Program tugas akhir untuk inspeksi visual pemasangan komponen pada PCB menggunakan **Python** dan **OpenCV**.

## Fitur utama

- Akuisisi citra dari webcam Logitech C920.
- Deteksi komponen menggunakan **multi-scale template matching**.
- Pemeriksaan enam komponen: **C1, C2, CONNECTOR, CRYSTAL, DIODA, dan IC**.
- Deteksi komponen hilang.
- Pengukuran pergeseran posisi terhadap referensi JIG.
- Analisis kemiringan menggunakan `minAreaRect`, `HoughLinesP`, dan PCA.
- Stabilisasi pembacaan antar-frame.
- Dua profil referensi: **JIG 1** dan **JIG 2**.
- Keputusan **GOOD / NOT GOOD**.
- Alarm suara, auto-capture, dan capture manual.
- Penyimpanan hasil ke citra, JSON, TXT, CSV, dataset, dan Excel.

## Struktur repository

```text
pcb-visual-inspection-github/
├── deteksi_pcb.py
├── referensi_jig_1.json
├── referensi_jig_2.json
├── requirements.txt
├── .gitignore
├── README.md
└── gambar_input/
    ├── C1 GOOD.png
    ├── C2 GOOD.png
    ├── Connector GOOD.png
    ├── Connector NOT GOOD.png
    ├── Crystal GOOD.png
    ├── Dioda GOOD.png
    ├── Dioda NOT GOOD.png
    └── IC GOOD.png
```

## Instalasi

Disarankan menggunakan Python 3.10 atau versi yang kompatibel.

```bash
pip install -r requirements.txt
```

## Menjalankan program

Pastikan webcam terhubung dan folder `gambar_input` berada satu folder dengan `deteksi_pcb.py`, kemudian jalankan:

```bash
python deteksi_pcb.py
```

## Kontrol keyboard

| Tombol | Fungsi |
|---|---|
| `1` | Memilih JIG 1 |
| `2` | Memilih JIG 2 |
| `K` | Menyimpan referensi GOOD untuk JIG aktif |
| `Z` | Menghapus referensi JIG aktif |
| `R` | Memuat ulang template dan referensi |
| `C` | Capture manual |
| `M` | Mute / unmute alarm |
| `U` | Membuka kunci auto-capture |
| `Q` | Keluar dari program |
| `↑ / ↓` | Scroll rincian komponen |

## Parameter penting

- Resolusi operasi kamera: **640 × 480 piksel**.
- Target frame rate: **30 FPS**.
- Threshold template matching: **0,58**.
- Skala template: **0,70–1,30**.
- Batas kemiringan: **15°** untuk komponen berarah.
- C1 dan C2: nilai sudut hanya sebagai informasi.
- Minimum toleransi pergeseran: **8 piksel** atau **18% diagonal bounding box referensi**, dipilih nilai yang lebih besar.

## Catatan

Folder hasil seperti `hasil_capture`, `dataset_pembelajaran`, dan `data_excel_inspeksi` tidak disertakan dalam repository ini karena merupakan data keluaran program dan ukurannya besar. Program akan membuat folder tersebut saat digunakan.

Repository ini disusun dari file program aktual yang diberikan untuk keperluan publikasi source code tugas akhir.

## Penulis

**Ardi Maulana**  
NIM: **2211012032**  
Program Studi D4 Teknik Elektronika  
Jurusan Teknik Elektro  
Politeknik Negeri Padang  
2026
