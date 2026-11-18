---
name: ats-cv-engineer
description: >-
  Universal ATS-friendly CV, Cover Letter, and Widescreen Portfolio Deck engineer
  for all majors and career levels (Fresh Graduate, Entry-Level/Months Experience,
  Career Switcher, Experienced) combining cassidoo/getting-a-gig standards with
  Indonesian HRD/ATS screening rules (Kemnaker RI, BNSP, BUMN, Corporate). Use when
  creating, repackaging, auditing, or exporting CVs (.docx & .pdf strictly 1 page),
  cover letters, or project portfolio decks (.pptx & .pdf).
---

# Universal ATS CV, Cover Letter & Portfolio Engineer (`ats-cv-engineer`)

## Overview

End-to-end engineering workflow for producing **100% ATS-compliant 1-page CVs (`.docx` & `.pdf`)**, **Formal Cover Letters (`.docx` & `.pdf`)**, and **Executive 16:9 Portfolio Presentation Decks (`.pptx` & `.pdf`)** for clients from **any academic major** and **any career stage** (Fresh Graduate, 1–11 Months Experience, Career Switcher, or Senior Professional).

This skill enforces the **3 Gates of Screening**:
1. **Gate 1 (Robot ATS Parser):** Single-column layout, zero tables/textboxes, pure black text (`#000000`), standard ASCII hyphens (`-`) for date ranges, and canonical ATS section headings (`PROFESSIONAL SUMMARY`, `WORK EXPERIENCE`, `LICENSES & CERTIFICATIONS`, `SKILLS & COMPETENCIES`, `EDUCATION`).
2. **Gate 2 (6-Second HRD Scan):** Big centered name (`18.5–19.5pt Bold`), high-impact sub-header containing primary licenses/certifications (e.g., *Ahli K3 Umum Kemnaker RI*, *STR Aktif*, *Brevet A/B*, *Full-Stack Engineer*), and active employment positioning (`Present`).
3. **Gate 3 (User / Hiring Manager Technical Scan):** Calibrated to the candidate's actual tenure—"Frontend / On-Site Execution" keywords for entry/junior candidates with months of experience vs. "Backend / Strategic Architecture" keywords for senior candidates.

---

## Dependencies

- **`agy-customizations`**: Follows Antigravity modular skill structure with progressive disclosure (`references/`).
- **`ponytail`**: Enforces zero fluff, concrete metrics, and honest scope (never over-claim corporate architecture for junior/field roles).

---

## Quick Start

1. **Read the relevant reference files** before drafting content:
   - [getting_a_gig_rules.md](./references/getting_a_gig_rules.md) — Core formatting, 1-page budgeting, and 4 career-stage strategies.
   - [industry_keywords_6_clusters.md](./references/industry_keywords_6_clusters.md) — Bilingual ATS keyword banks across 6 major industry clusters (IT, HSE/K3/Engineering, Admin/Finance/Accounting, HR/Legal/Psychology, Sales/Marketing, Healthcare/Education).
   - [client_order_templates.md](./references/client_order_templates.md) — Universal WhatsApp client intake form and "Frontend/Field vs. Backend/Corporate" calibration guide.
2. **Prepare the JSON specification** for the candidate and run the CLI tool [`scripts/cv_portfolio_cli.py`](./scripts/cv_portfolio_cli.py) to generate and verify the `.docx`, `.pptx`, and `.pdf` outputs.

---

## Utility Scripts

All document generation and ATS auditing is handled by [`scripts/cv_portfolio_cli.py`](./scripts/cv_portfolio_cli.py). Every subcommand writes structured results to `--output`.

### 1. `build-cv` — Generate 1-Page ATS CV (`.docx` & `.pdf`) with Auto-Fit

Renders a candidate JSON specification into a single-column `.docx` and exports a `.pdf` via Microsoft Word COM. Automatically shrinks font/spacing if the initial render exceeds 1 page.

```bash
python scripts/cv_portfolio_cli.py build-cv \
  --input candidate_cv.json \
  --docx-out "CV_Candidate_Role.docx" \
  --pdf-out "CV_Candidate_Role.pdf" \
  --output build_report.json
```

### 2. `audit-ats` — Audit Any PDF CV Against ATS Parser Rules

Extracts the raw text stream using `pypdf` and checks for: (a) strictly 1 page, (b) corrupted Unicode characters (`` or En-Dash `–` in date ranges), (c) mandatory ATS section headers (`WORK EXPERIENCE`, `SKILLS`, `EDUCATION`), and (d) keyword match score against target keywords.

```bash
python scripts/cv_portfolio_cli.py audit-ats \
  --pdf "CV_Candidate_Role.pdf" \
  --keywords "HIRADC,JSA,Safety Patrol,Toolbox Meeting,Kemnaker RI" \
  --output audit_report.json
```

### 3. `build-cover-letter` — Generate 1-Page Formal Cover Letter (`.docx` & `.pdf`)

Generates a 1-page Indonesian or English formal cover letter (`Surat Lamaran Kerja`) following the 5W+1H `getting-a-gig` pitch structure.

```bash
python scripts/cv_portfolio_cli.py build-cover-letter \
  --input cover_letter.json \
  --docx-out "Surat_Lamaran_Candidate.docx" \
  --pdf-out "Surat_Lamaran_Candidate.pdf" \
  --output cl_report.json
```

### 4. `build-portfolio` — Generate 16:9 Clean Neutral Portfolio Deck (`.pptx` & `.pdf`)

Builds a widescreen 16:9 presentation deck with a soft neutral executive palette (`#F1F5F9` background, `#FFFFFF` cards, `#0F172A` headings) and exports high-res `.pdf` via PowerPoint COM.

```bash
python scripts/cv_portfolio_cli.py build-portfolio \
  --input portfolio_spec.json \
  --pptx-out "Portfolio_Candidate.pptx" \
  --pdf-out "Portfolio_Candidate.pdf" \
  --output portfolio_report.json
```

---

## Workflow

1. **Classify the Candidate's Career Level & Industry Cluster:**
   - Identify which of the 4 career levels applies (*Fresh Graduate*, *1–11 Months Experience*, *Career Switcher*, or *Experienced 1–5+ Years*) using [getting_a_gig_rules.md](./references/getting_a_gig_rules.md).
   - Select bilingual ATS keywords from the matching cluster in [industry_keywords_6_clusters.md](./references/industry_keywords_6_clusters.md).
2. **Calibrate "Frontend (Field/Execution)" vs. "Backend (Corporate/System)":**
   - For candidates with <1 year of experience, never write explicit month counts in the Summary (e.g., avoid *"3 months experience"*); write *"dengan pengalaman praktis dalam..."* and focus bullets on daily on-site/operational execution rather than high-level corporate policy design.
3. **Generate & Auto-Fit Document:**
   - Write the JSON input file and execute `build-cv` (and `build-cover-letter` / `build-portfolio` if requested).
4. **Run Mandatory ATS Audit:**
   - Run `audit-ats` on the generated PDF and confirm `status == "PASS"` and `pages == 1` before delivering the file link to the user.

---

## Common Mistakes

1. **Using Unicode En-Dash (`–`) in Date Ranges:** Legacy ASCII ATS parsers turn `–` into `` and calculate `0 months` of work experience. Always use standard ASCII hyphen `-` (e.g., `May 2026 - Present`).
2. **Omitting the Word `SKILLS` from the Competency Header:** Naming a section `HSE & REGULATORY COMPETENCIES` without the word `SKILLS` causes strict ATS parsers to skip indexing the candidate's skills. Always use `SKILLS & [DOMAIN] COMPETENCIES`.
3. **Placing Certifications Above Work Experience:** Chronological ATS engines (*Workday, Taleo*) often misparse the first certification line as the candidate's employer name if placed before `WORK EXPERIENCE`. Keep `WORK EXPERIENCE` immediately after `PROFESSIONAL SUMMARY`, and surface key certifications in the bold Sub-Header right under the candidate's name.
