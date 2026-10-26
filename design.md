# Design System Tokens — Komite Keperawatan RSUD Subang

## 1. Color Palette

### Primary / Brand (Teal & Tosca)
- **Primary Base**: `#0d9488` (`teal-600`) — Tombol utama, indikator fokus, status aktif.
- **Primary Hover**: `#0f766e` (`teal-700`) — Hover state tombol, teks judul menu.
- **Primary Active/Dark**: `#115e59` (`teal-800`) — Header dan aksen gelap.
- **Primary Tint / Surface**: `#f0fdfa` (`teal-50`), `#ccfbf1` (`teal-100`) — Background kartu, table header, dan highlight bar.
- **Brand Gradient**: `linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 50%, #99f6e4 100%)` — Background body aplikasi.

### Neutral Colors
- **Text Main**: `#1f2937` (`gray-800`)
- **Text Subdued / Muted**: `#4b5563` (`gray-600`), `#6b7280` (`gray-500`)
- **Border Default**: `#e5e7eb` (`gray-200`)
- **Card Background**: `#ffffff` (`white`)
- **Surface Secondary**: `#f9fafb` (`gray-50`)

### Semantic Colors
- **Success**: `#16a34a` (`green-600`), Tint: `#f0fdf4` (`green-50`) — Status "Verifikasi OK", "Aktif", progress kelengkapan 100%.
- **Warning**: `#ea580c` (`orange-600`), Tint: `#fff7ed` (`orange-50`) — Mutasi, data belum lengkap <50%.
- **Danger**: `#dc2626` (`red-600`), Tint: `#fef2f2` (`red-50`) — "Belum Ter Verifikasi", STR/SIPP kedaluwarsa, error alert.
- **Info / Action**: `#2563eb` (`blue-600`), Tint: `#eff6ff` (`blue-50`) — Progress kelengkapan, tombol tambah/cek data, unduh word.

### Dark Mode Tokens
- **Body Background**: `linear-gradient(135deg, #1f2937 0%, #111827 50%, #030712 100%)`
- **Surface Dark**: `#1f2937` (Cards, Modals)
- **Surface Hover**: `#374151`
- **Text Dark Primary**: `#f3f4f6`
- **Text Dark Secondary**: `#9ca3af`
- **Border Dark**: `#374151` / `#4b5563`

---

## 2. Typography

- **Font Family**: System UI / Tailwind Sans (`ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`)
- **Scale**:
  - `text-2xl` / `text-3xl`: Judul halaman & statistik KPI
  - `text-lg` / `text-xl`: Judul modal & sub-heading seksi
  - `text-sm` (`14px`): Body text, isi data kartu, input form
  - `text-xs` (`12px`): Sub-label, helper text, tombol aksi kecil
  - `text-[10px]` / `text-[11px]`: Badge status & caption uppercase

---

## 3. Spacing & Elevation

- **Radius**:
  - Card & Container: `rounded-2xl` (`1rem`)
  - Input & Tombol: `rounded-lg` (`0.5rem`)
  - Badge & Avatar: `rounded-full` (`9999px`)
- **Shadow**:
  - Default Card: `shadow-sm` / `shadow-md`
  - Modal Dialog: `shadow-2xl`
  - Navbar: `shadow-md`
- **Tap Target (Mobile Accessibility)**:
  - Minimum tap target input: `0.6rem 0.75rem` (`py-2.5 px-3`)
  - Focus Ring: `outline: none; border-color: #0d9488; box-shadow: 0 0 0 3px rgba(13, 148, 136, 0.25)`
  - Active Feedback: `transform: scale(0.97); opacity: 0.9`

---

## 4. Components & UX Rules

1. **SPA Hash Router**:
   - URL hash format: `#/beranda`, `#/biodata`, `#/kepegawaian`, `#/ruangan`, `#/pendidikan`, `#/kompetensi`, `#/dashboard`, `#/datakomper`, `#/users`, `#/slideshow`.
   - Scroll position otomatis kembali ke atas (`window.scrollTo({ top: 0, behavior: 'instant' })`) setiap navigasi view.
2. **Modal Dialogs**:
   - Semua modal wajib dapat ditutup dengan menekan tombol `Escape` atau mengklik area backdrop transparan.
3. **Password Masking**:
   - Password pada card profil disamarkan (`••••••••`) dengan tombol toggle mata (`fa-eye`).
   - Input password form berjenis `type="password"` dengan toggle unmask.
4. **Hero Slideshow**:
   - Memakai `relative z-10` agar tidak menutupi scroll kartu kelengkapan di bawahnya.
5. **Aksesibilitas (WCAG)**:
   - Pinch zoom diaktifkan (`user-scalable` diperbolehkan).
   - Seluruh icon button memiliki `aria-label`.
6. **Running Text (Sticky Marquee di Bawah Header)**:
   - Posisi: Menempel di dalam `<nav>`, sticky bersama navbar di bagian atas (`z-50`).
   - Tinggi bar: `h-8` (`32px`), background tint `bg-teal-50`, border `border-teal-100`.
   - Badge Info: `bg-teal-600 text-white rounded-full text-[10px]`.
   - Perilaku: Animasi berjalan halus (seamless loop) dengan CSS marquee; animasi berhenti sementara saat kursor diarahkan ke teks (`hover:pause`).
   - Offset Konten: Saat bar aktif, `body` menerima kelas `rt-active` yang menyesuaikan padding atas `main` dan `sideMenu` (`pt-[6.25rem]`).
   - Hak Akses: Pengaturan di Sub Menu **Setting → Running Text** hanya dapat diakses oleh Super Admin. Bar otomatis tersembunyi jika teks kosong.
7. **Slideshow Dashboard & Foto Seragam (480 × 360 px)**:
   - Standarisasi ukuran gambar: Wajib rasio 4:3 dengan dimensi tepat **480 × 360 piksel**.
   - Batas ukuran file: Maksimal **3 MB** per file foto (JPG, JPEG, PNG).
   - Fitur Crop Manual: Apabila foto yang diunggah tidak berukuran 480 × 360 px, modal Crop Manual otomatis terbuka dengan kanvas 4:3, grid panduan (*rule of thirds*), pan/drag 1:1, kontrol zoom slider & tombol +/- / scroll mouse wheel, tombol reset posisi, dan tombol eksekusi potong.

