# KOL Agency OS

Creator & Campaign Operating System — Google Sheets + GAS + Static Website

## Stack
- **Frontend**: HTML + Vanilla JS (no framework needed)
- **Backend**: Google Apps Script (GAS) Web App
- **Database**: Google Sheets
- **Hosting**: Vercel (auto-deploy from GitHub)

---

## Setup Guide

### 1. Google Sheets + GAS Backend

1. Buka [Google Sheets](https://sheets.google.com) → buat spreadsheet baru (kosong)
2. Klik **Extensions → Apps Script**
3. Hapus semua kode yang ada → paste isi file `api/Code.gs`
4. Klik **Run → setupSystem()** — ini akan auto-buat semua sheet + dummy data
5. Klik **Deploy → New deployment**:
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy **Web App URL** yang muncul
7. Paste URL itu ke variabel `GAS_URL` di semua HTML file (cari `YOUR_SCRIPT_ID`)

### 2. Konfigurasi

Di sheet **CONFIG**, update:
- `admin_password` — ganti dari `admin123` ke password yang kuat
- `wa_community_link` — link WA community creator kamu
- `agency_name` — nama agency kamu

### 3. Deploy ke Vercel

1. Push semua file ke GitHub repo baru
2. Buka [vercel.com](https://vercel.com) → Import repository
3. **Root directory**: pilih folder `public/` (atau taruh semua HTML di root)
4. Deploy → connect domain kamu

### 4. Update URL di HTML

Di setiap file HTML, ganti:
- `GAS_URL` → URL GAS Web App kamu
- `WA_LINK` → link WA community
- `FORM_BASE` → domain website kamu (misal `https://youragency.com/apply`)

---

## File Structure

```
/
├── public/
│   ├── index.html                   → Landing page
│   ├── creator-onboarding-form.html → Form daftar creator (/join)
│   ├── creator-campaign-apply.html  → Form apply campaign (/apply/CAM-xx)
│   └── admin-dashboard.html         → Admin panel (/admin)
├── api/
│   └── Code.gs                      → GAS backend (paste ke Apps Script)
├── vercel.json                      → Routing config
└── README.md
```

---

## URL Structure (setelah deploy)

| URL | Halaman |
|-----|---------|
| `youragency.com` | Landing page |
| `youragency.com/join` | Form daftar creator |
| `youragency.com/apply/CAM-24-001` | Form apply campaign tertentu |
| `youragency.com/admin` | Admin dashboard |

---

## Cara Share Campaign ke WA

Di admin dashboard → halaman Campaigns → klik tombol **Share** di campaign yang ingin dibuka.
Dashboard akan generate:
- Link apply: `youragency.com/apply/CAM-24-001`
- Draft pesan WA yang bisa langsung dicopy

---

## Scoring Formula

Creator score (0–100) dihitung otomatis setiap hari jam 02.00:

```
Score = (ER score × 30%) + (Reliability × 25%) + (Consistency × 20%) + (Cost efficiency × 15%) + (Content quality × 10%)
```

Bobot bisa diubah di sheet CONFIG tanpa coding.

---

## GAS Functions Cheatsheet

| Fungsi | Kapan dijalankan |
|--------|-----------------|
| `setupSystem()` | Sekali saat setup awal |
| `setupTriggers()` | Sekali setelah deploy |
| `recalculateAllScores()` | Auto setiap hari, atau manual |

---

## Admin Password

Default: `admin123` — **WAJIB diganti** sebelum live di sheet CONFIG → key `admin_password`.
