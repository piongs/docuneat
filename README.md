# DocuNeat — Perapih Laporan Word Otomatis

Aplikasi web Python + Flask untuk merapihkan file Word (.docx) yang berantakan menjadi laporan profesional secara otomatis.

## Fitur Utama

- **Upload & Format**: Upload file .docx berantakan, langsung dirapihkan
- **Font Standar**: Terapkan Times New Roman 12pt (atau font lain pilihan)
- **Deteksi Heading Otomatis**: Mendeteksi BAB, PENDAHULUAN, METODOLOGI, dll.
- **Normalkan Spasi**: Hapus baris kosong ganda, atur jarak paragraf
- **Rata Kanan-Kiri**: Terapkan justify pada teks isi secara otomatis
- **Margin Halaman**: Preset narrow/normal/wide/mirror
- **Nomor Halaman**: Tambah otomatis di footer tengah
- **Riwayat**: Lihat, unduh ulang, atau hapus riwayat format
- **Login**: Autentikasi dengan password yang di-hash

## Cara Menjalankan

```bash
# 1. Install dependensi
pip install -r requirements.txt

# 2. Jalankan aplikasi
python app.py

# 3. Buka browser
# http://localhost:5000
```

## Login Default

- **Username**: admin
- **Password**: admin123

## Struktur Folder

```
docuneat/
├── app.py              # Flask app utama
├── docx_formatter.py   # Logika formatting Word
├── requirements.txt
├── templates/
│   ├── login.html
│   └── dashboard.html
├── uploads/            # File sementara (auto-bersih)
└── outputs/            # File hasil yang sudah rapi
```

## Apa yang Dirapihkan Otomatis?

| Masalah | Solusi |
|---------|--------|
| Font campur-campur (Comic Sans, Arial, Courier) | → Semua jadi Times New Roman 12pt |
| Ukuran huruf tidak konsisten | → Disamakan semua |
| Heading tidak dikenali | → Dideteksi & diberi style Heading 1/2/3 |
| Baris kosong berlebihan | → Dihapus, hanya sisakan 1 |
| Teks tidak rata | → Justify otomatis |
| Margin berbeda-beda | → Disamakan sesuai preset |
| Tidak ada nomor halaman | → Ditambah otomatis (opsional) |

## Heading yang Dideteksi Otomatis

- **BAB I, BAB II, ...** → Heading 1 (center)
- **TEKS KAPITAL SEMUA** → Heading 1 (center)
- **Pendahuluan, Metodologi, Kesimpulan, Daftar Pustaka, ...** → Heading 1
- **1.1 Sub Judul** → Heading 2
- **1.1.1 Sub Sub Judul** → Heading 3
