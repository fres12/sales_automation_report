# 📋 PANDUAN VALIDASI TANGGAL OTOMATIS

## 🔍 Apa Itu Fitur Ini?

Sebelum mengambil screenshot dan mengirim report ke WhatsApp, sistem sekarang **otomatis memvalidasi tanggal** pada cell **A1 di sheet "Summary"** untuk memastikan tanggal sudah di-update dengan benar.

---

## 📅 Logika Validasi

### Mekanisme Pengiriman Report:
- **Setiap hari, kirim report yang KEMARIN (H-1)**
- Kecuali untuk tanggal khusus yang memiliki jadwal berbeda

### Jadwal Tanggal:
Berdasarkan angka pada A1 di sheet Summary:

| Tanggal | Jadwal | Expected Hari Ini | Contoh |
|---------|--------|------------------|--------|
| 1-19 (except 1) | H-2 | Tanggal = A1 + 2 | A1=5 → Expected=7 |
| 20-31 | H-1 | Tanggal = A1 + 1 | A1=20 → Expected=21 |
| **1 (special)** | **H-1 (kemarin)** | **Awal bulan** | A1=1 → Hari ini harus 1-3 |

#### 📝 Penjelasan:
- **Tanggal 2-19**: Report dikirim setiap 2 hari
  - Jika A1=3, hari ini harus tanggal 5
  - Jika A1=7, hari ini harus tanggal 9
  
- **Tanggal 20-31**: Report dikirim setiap hari
  - Jika A1=20, hari ini harus tanggal 21
  - Jika A1=31, hari ini harus tanggal 1 (overflow ke bulan depan)

- **Tanggal 1 (special)**: Transisi bulan
  - Jika A1=1 dan kemarin adalah tanggal terakhir bulan lalu (28,29,30,31)
  - Maka hari ini VALID untuk mengirim

---

## ⚡ Apa Yang Terjadi Saat Validasi?

### ✅ Jika Validasi LOLOS:
1. Proses melanjutkan ke step berikutnya
2. Mengambil screenshot dari semua dashboard
3. Mengirim ke WhatsApp grup

### ❌ Jika Validasi GAGAL:
1. **PROSES DIHENTIKAN** ← Tidak ada screenshot, tidak ada pengiriman ke WA grup
2. Sistem mengirimkan pesan error ke **admin/notifikasi grup** (NOTIF_NUMBER di .env):
   ```
   ❌ PROSES GAGAL
   Alasan: connect OCBC gagal
   Detail: [penjelasan spesifik error]
   ```
3. Tidak ada screenshot yang diambil atau dikirim

---

## 🛠️ Cara Mengatasi Jika Gagal Validasi

### 1️⃣ Cek Cell A1 di Sheet "Summary"
- Buka file Excel: `SISO OPA JAVA.xlsx`
- Pergi ke sheet **"Summary"**
- Lihat cell **A1** - berisi angka berapa?

### 2️⃣ Hitung Expected Hari Ini
Berdasarkan jadwal di atas:
- **Jika A1 = 5**: Harusnya hari ini tanggal **7** (5+2)
- **Jika A1 = 20**: Harusnya hari ini tanggal **21** (20+1)
- **Jika A1 = 1**: Harusnya hari ini awal bulan (1-3)

### 3️⃣ Update A1 Jika Salah
- Ubah nilai A1 menjadi angka yang sesuai dengan jadwal
- Contoh: Jika hari ini tanggal 7, A1 harus = 5
- Save file Excel

### 4️⃣ Jalankan Ulang
- Jalankan `main.py` / `run.bat` lagi
- Validasi akan lolos dan proses berjalan normal

---

## 🔧 Implementasi Teknis

### File Yang Dimodifikasi:
1. **main.py**
   - Function `validate_date_before_send(wb)` - melakukan validasi
   - Function `send_error_to_wa(error_message)` - mengirim error ke WA
   - Terintegrasi di step [3.5/5] setelah refresh GCP selesai

2. **wa_bot.js**
   - Added check untuk `is_error` flag di send_config.json
   - Jika error, langsung kirim pesan dan exit

### Flow Eksekusi:
```
[1/5] Buka File
    ↓
[2/5] Refresh Data GCP
    ↓
[3/5] Tunggu Refresh Selesai
    ↓
[3.5/5] 🆕 VALIDASI TANGGAL ← NEW!
    ├─ Lolos → Lanjut ke [4/5]
    └─ Gagal → Kirim Error & STOP
    ↓
[4/5] Screenshot Dashboard
    ↓
[5/5] Kirim ke WhatsApp
```

---

## 📊 Error Messages

### Error: "Sheet 'Summary' tidak ditemukan"
- Sheet Summary tidak ada di file Excel
- Buat sheet baru bernama "Summary"
- Tambahkan nilai di A1

### Error: "Cell A1 kosong"
- Sheet Summary ada tapi A1 kosong
- Isi A1 dengan angka 1-31 sesuai jadwal

### Error: "Cell A1 bukan angka"
- A1 berisi teks atau formula yang salah
- Ganti dengan angka sederhana (1-31)

### Error: "Tanggal di A1 tidak valid"
- Nilai A1 tidak dalam range 1-31
- Pastikan hanya angka 1-31

### Error: "Tanggal tidak sesuai"
- A1 benar tetapi hari ini tidak sesuai jadwal
- Contoh: A1=5 tapi hari ini tanggal 6 (harusnya 7)
- Tunggu hingga jadwal sesuai atau update A1

---

## 📞 Bantuan Troubleshooting

### Q: Berapa lama validasi memakan waktu?
A: Instant (<1 detik), dilakukan langsung setelah refresh GCP selesai

### Q: Apakah validasi bisa di-skip?
A: Tidak, validasi **selalu** dilakukan sebelum screenshot

### Q: Apa yang terlihat di terminal saat validasi?
A: Output akan menampilkan:
```
============================================================
🔍 [3.5/5] VALIDASI TANGGAL SEBELUM KIRIM KE WA
============================================================
   ✓ Sheet 'Summary' ditemukan
   ✓ Tanggal di A1: 5
   📅 Hari ini: 2024-12-07 (Saturday) (Tanggal: 7)
   📋 Expected: Tanggal 7 bulan 12 (H-2 dari 5)
   ✅ VALIDASI BERHASIL!
   📅 Jadwal: Tanggal 2-19 (H-2)
   📊 Report: Mengirim data H-1 dari tanggal 5
```

---

## 🎯 Summary
- ✅ Validasi **otomatis** sebelum kirim report
- ✅ Cegah pengiriman jika tanggal tidak sesuai
- ✅ Error langsung di-notify ke WA dengan pesan "connect OCBC gagal"
- ✅ Memastikan data yang dikirim SELALU tepat waktu sesuai jadwal
