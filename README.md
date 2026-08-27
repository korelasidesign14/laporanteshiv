# Rekap Tes HIV per Layanan

Aplikasi web untuk mengolah data Register/Kunjungan Tes HIV: upload file .xlsx,
langsung dapat rekap jumlah tes per layanan dan capaian SPM per kelompok
populasi, plus unduh hasil rekapnya sebagai .xlsx.

## Menjalankan di komputer sendiri (opsional, untuk coba dulu)

```bash
npm install
npm run dev
```

Buka http://localhost:5173

## Deploy ke Netlify

**Cara termudah (tanpa akun GitHub):**
1. Jalankan `npm install` lalu `npm run build` di folder ini (menghasilkan folder `dist`).
2. Buka https://app.netlify.com/drop
3. Seret folder `dist` ke halaman itu — selesai, langsung dapat URL live.

**Cara lewat GitHub (auto-deploy tiap update):**
1. Push folder ini ke repo GitHub baru.
2. Di Netlify: "Add new site" → "Import an existing project" → pilih repo tsb.
3. Build command: `npm run build`, Publish directory: `dist` (sudah diatur di `netlify.toml`).

## Deploy ke Vercel

**Lewat CLI:**
```bash
npm install -g vercel
vercel
```
Ikuti instruksinya (login, pilih project baru). Konfigurasi build sudah diatur di `vercel.json`.

**Lewat GitHub:**
1. Push folder ini ke repo GitHub.
2. Di https://vercel.com → "Add New Project" → import repo tsb.
3. Vercel otomatis mendeteksi Vite + pengaturan build dari `vercel.json`.

## Catatan
- Semua pemrosesan file terjadi di browser pengguna (client-side) — tidak ada
  data yang dikirim ke server manapun, jadi aman untuk data pasien.
- Format file yang didukung: register/kunjungan dengan kolom `Nama UPK`,
  `Kelompok Populasi`, `Jenis Layanan`, dan `Status ODHIV`.
