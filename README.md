# cv-ats — Universal ATS CV, Cover Letter & Portfolio Engineer (`ats-cv-engineer`)

Antigravity AI Agent Skill untuk merancang, mengaudit, dan menghasilkan **CV 1 halaman lolos ATS (`.docx` & `.pdf`)**, **Surat Lamaran Kerja formal (`.docx` & `.pdf`)**, serta **Deck Portofolio 16:9 (`.pptx` & `.pdf`)** lintas jurusan dan jenjang karier (*Fresh Graduate*, *1–11 Months Experience*, *Career Switcher*, maupun *Experienced*) dengan menggabungkan standar `cassidoo/getting-a-gig` dan aturan HRD/ATS Indonesia (Kemnaker RI, BNSP, BUMN, Korporat).

---

## Cara Install di Agent Antigravity Lain

Pilih salah satu metode berikut sesuai kebutuhan Anda:

### 1. Install Global (Semua Workspace di Mesin Baru)

**Windows (PowerShell):**
```powershell
git clone https://github.com/imanmukhlisin/cv-ats.git "$env:USERPROFILE\.gemini\config\skills\ats-cv-engineer"
pip install -r "$env:USERPROFILE\.gemini\config\skills\ats-cv-engineer\requirements.txt"
```

**macOS / Linux (Bash/Zsh):**
```bash
git clone https://github.com/imanmukhlisin/cv-ats.git ~/.gemini/config/skills/ats-cv-engineer
pip install -r ~/.gemini/config/skills/ats-cv-engineer/requirements.txt
```

> Setelah terpasang di `~/.gemini/config/skills/ats-cv-engineer/SKILL.md`, skill akan otomatis ditemukan Antigravity dan bisa langsung dipanggil dengan `/ats-cv-engineer` atau lewat prompt natural.

### 2. Install Khusus Workspace / Project Tertentu

Jalankan dari folder root project Anda:
```bash
git clone https://github.com/imanmukhlisin/cv-ats.git .agents/skills/ats-cv-engineer
pip install -r .agents/skills/ats-cv-engineer/requirements.txt
```

### 3. Install via `npx skills` CLI (Jika Tersedia)

```bash
npx skills add imanmukhlisin/cv-ats
```

### 4. Minta Langsung ke Agent Antigravity Lain

Salin dan kirim prompt ini ke chat Agent Antigravity tujuan Anda:

> *"Install skill `ats-cv-engineer` dari repository https://github.com/imanmukhlisin/cv-ats.git ke `~/.gemini/config/skills/ats-cv-engineer`, lalu jalankan `pip install -r requirements.txt`."*

---

## Cara Menggunakan Setelah Terinstall

Di Antigravity chat, aktifkan dengan slash command:
```text
/ats-cv-engineer Buatkan CV ATS 1 halaman untuk posisi [Target Posisi] ...
```

Atau jalankan CLI engine secara langsung dari direktori skill:
```bash
# 1. Generate CV ATS 1 halaman (.docx & .pdf) dengan auto-fit
python scripts/cv_portfolio_cli.py build-cv \
  --input candidate_cv.json \
  --docx-out CV_Candidate.docx \
  --pdf-out CV_Candidate.pdf \
  --output build_report.json

# 2. Audit PDF CV terhadap parser ATS
python scripts/cv_portfolio_cli.py audit-ats \
  --pdf CV_Candidate.pdf \
  --keywords "HIRADC,JSA,Safety Patrol,Kemnaker RI" \
  --output audit_report.json

# 3. Generate Surat Lamaran formal 1 halaman (.docx & .pdf)
python scripts/cv_portfolio_cli.py build-cover-letter \
  --input cover_letter.json \
  --docx-out Surat_Lamaran.docx \
  --pdf-out Surat_Lamaran.pdf \
  --output cl_report.json

# 4. Generate Deck Portofolio 16:9 (.pptx & .pdf)
python scripts/cv_portfolio_cli.py build-portfolio \
  --input portfolio_spec.json \
  --pptx-out Portfolio_Candidate.pptx \
  --pdf-out Portfolio_Candidate.pdf \
  --output portfolio_report.json
```

---

## Struktur Skill

```text
ats-cv-engineer/
├── SKILL.md                                  # Instruksi utama & workflow Antigravity
├── README.md                                 # Panduan instalasi & penggunaan
├── requirements.txt                          # Dependensi Python CLI
├── references/
│   ├── getting_a_gig_rules.md                # Aturan layout 1 halaman & strategi 4 level karier
│   ├── industry_keywords_6_clusters.md       # Bank kata kunci ATS bilingual (6 klaster industri)
│   └── client_order_templates.md             # Template form order & matriks kalibrasi
└── scripts/
    └── cv_portfolio_cli.py                   # CLI generator & ATS PDF auditor
```

## Prasyarat Sistem (`scripts/cv_portfolio_cli.py`)

- Python 3.9+ dengan paket pada `requirements.txt` (`python-docx`, `python-pptx`, `pypdf`, `pywin32`)
- **Windows + Microsoft Word / Microsoft PowerPoint** untuk ekspor `.pdf` via COM serta verifikasi otomatis ketat 1 halaman (`ComputeStatistics(2)`).
