#!/usr/bin/env python3
"""Universal ATS CV, Cover Letter & Widescreen Portfolio Deck CLI Engine."""

import argparse
import json
import os
import sys

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import pypdf
import win32com.client

COLOR_BLACK = RGBColor(0, 0, 0)
HEX_BLACK = "000000"


def sanitize_ascii_hyphens(text: str) -> str:
    """Replace Unicode en-dash/em-dash with standard ASCII hyphen for ATS parsers."""
    if not text:
        return ""
    return text.replace("\u2013", "-").replace("\u2014", "-")


def add_hyperlink(paragraph, url, text, underline=True, font_size_pt=8.2):
    part = paragraph.part
    r_id = part.relate_to(
        url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    c = OxmlElement("w:color")
    c.set(qn("w:val"), HEX_BLACK)
    rPr.append(c)

    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Calibri")
    rFonts.set(qn("w:hAnsi"), "Calibri")
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(font_size_pt * 2)))
    rPr.append(sz)

    new_run.append(rPr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def render_cv_docx(data: dict, output_path: str, scale: float = 1.0):
    """Render candidate JSON into an ATS-compliant single-column DOCX."""
    doc = docx.Document()
    margin_tb = max(0.26, 0.34 * scale)
    margin_lr = max(0.42, 0.48 * scale)
    for section in doc.sections:
        section.top_margin = Inches(margin_tb)
        section.bottom_margin = Inches(margin_tb)
        section.left_margin = Inches(margin_lr)
        section.right_margin = Inches(margin_lr)

    body_pt = round(8.2 * scale, 2)
    header_pt = round(9.5 * scale, 2)
    job_pt = round(8.8 * scale, 2)
    name_pt = round(18.5 * scale, 2)

    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(body_pt)
    style_normal.font.color.rgb = COLOR_BLACK

    def add_sec_hdr(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3.2 * scale)
        p.paragraph_format.space_after = Pt(1.0 * scale)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(header_pt)
        run.font.color.rgb = COLOR_BLACK

        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), HEX_BLACK)
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0.6 * scale)
        p.paragraph_format.line_spacing = max(1.0, 1.03 * scale)
        if bold_prefix:
            rp = p.add_run(sanitize_ascii_hyphens(bold_prefix))
            rp.bold = True
            rp.font.name = "Calibri"
            rp.font.size = Pt(body_pt)
            rp.font.color.rgb = COLOR_BLACK
        rt = p.add_run(sanitize_ascii_hyphens(text))
        rt.font.name = "Calibri"
        rt.font.size = Pt(body_pt)
        rt.font.color.rgb = COLOR_BLACK
        return p

    # 1. HEADER
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(0.5)
    r_name = p_name.add_run(data.get("name", "CANDIDATE NAME").upper())
    r_name.font.name = "Calibri"
    r_name.font.size = Pt(name_pt)
    r_name.bold = True
    r_name.font.color.rgb = COLOR_BLACK

    if data.get("subtitle"):
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(1.0)
        r_sub = p_sub.add_run(sanitize_ascii_hyphens(data["subtitle"]))
        r_sub.font.name = "Calibri"
        r_sub.font.size = Pt(header_pt)
        r_sub.bold = True
        r_sub.font.color.rgb = COLOR_BLACK

    # Contacts
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(2.0 * scale)
    contacts = data.get("contacts", [])
    for idx, item in enumerate(contacts):
        if idx > 0:
            r_sep = p_contact.add_run("  |  ")
            r_sep.font.size = Pt(body_pt)
            r_sep.font.color.rgb = COLOR_BLACK
        if isinstance(item, dict) and item.get("url"):
            add_hyperlink(
                p_contact, item["url"], item["text"], font_size_pt=body_pt
            )
        else:
            txt = item["text"] if isinstance(item, dict) else str(item)
            r_c = p_contact.add_run(txt)
            r_c.font.size = Pt(body_pt)
            r_c.font.color.rgb = COLOR_BLACK

    # 2. PROFESSIONAL SUMMARY
    if data.get("summary"):
        add_sec_hdr("Professional Summary")
        p_sum = doc.add_paragraph()
        p_sum.paragraph_format.space_before = Pt(0)
        p_sum.paragraph_format.space_after = Pt(1.5 * scale)
        p_sum.paragraph_format.line_spacing = max(1.0, 1.03 * scale)
        r_sum = p_sum.add_run(sanitize_ascii_hyphens(data["summary"]))
        r_sum.font.size = Pt(body_pt)
        r_sum.font.color.rgb = COLOR_BLACK

    # 3. WORK EXPERIENCE
    if data.get("experience"):
        add_sec_hdr("Work Experience")
        for exp in data["experience"]:
            p_job = doc.add_paragraph()
            p_job.paragraph_format.space_before = Pt(2.2 * scale)
            p_job.paragraph_format.space_after = Pt(0.3 * scale)
            p_job.paragraph_format.keep_with_next = True
            p_job.paragraph_format.tab_stops.add_tab_stop(
                Inches(8.5 - 2 * margin_lr), WD_TAB_ALIGNMENT.RIGHT
            )
            r_jt = p_job.add_run(
                sanitize_ascii_hyphens(
                    f"{exp.get('title', '')} | {exp.get('company', '')}"
                )
            )
            r_jt.bold = True
            r_jt.font.size = Pt(job_pt)
            r_jt.font.color.rgb = COLOR_BLACK

            r_jd = p_job.add_run(
                f"\t{sanitize_ascii_hyphens(exp.get('date', ''))}"
            )
            r_jd.italic = True
            r_jd.font.size = Pt(body_pt)
            r_jd.font.color.rgb = COLOR_BLACK

            if exp.get("link"):
                p_lnk = doc.add_paragraph()
                p_lnk.paragraph_format.space_before = Pt(0)
                p_lnk.paragraph_format.space_after = Pt(0.3 * scale)
                p_lnk.paragraph_format.keep_with_next = True
                rl = p_lnk.add_run("Project Link: ")
                rl.italic = True
                rl.font.size = Pt(body_pt - 0.3)
                add_hyperlink(
                    p_lnk,
                    exp["link"]["url"],
                    exp["link"].get("text", exp["link"]["url"]),
                    font_size_pt=body_pt - 0.3,
                )

            for b in exp.get("bullets", []):
                if isinstance(b, dict):
                    add_bullet(b.get("prefix", ""), b.get("text", ""))
                else:
                    add_bullet("", str(b))

    # 4. LICENSES & CERTIFICATIONS
    if data.get("certifications"):
        add_sec_hdr("Licenses & Certifications")
        for cert in data["certifications"]:
            if isinstance(cert, dict):
                add_bullet(cert.get("prefix", ""), cert.get("text", ""))
            else:
                add_bullet("", str(cert))

    # 5. SKILLS & COMPETENCIES
    if data.get("skills"):
        skills_title = data.get("skills_header", "Skills & Core Competencies")
        if "SKILL" not in skills_title.upper():
            skills_title = f"Skills & {skills_title}"
        add_sec_hdr(skills_title)
        for sk in data["skills"]:
            p_sk = doc.add_paragraph()
            p_sk.paragraph_format.space_before = Pt(0)
            p_sk.paragraph_format.space_after = Pt(0.5 * scale)
            p_sk.paragraph_format.line_spacing = max(1.0, 1.02 * scale)
            rc = p_sk.add_run(
                f"-  {sanitize_ascii_hyphens(sk.get('category', ''))}: "
            )
            rc.bold = True
            rc.font.size = Pt(body_pt)
            rc.font.color.rgb = COLOR_BLACK
            ri = p_sk.add_run(sanitize_ascii_hyphens(sk.get("items", "")))
            ri.font.size = Pt(body_pt)
            ri.font.color.rgb = COLOR_BLACK

    # 6. EDUCATION
    if data.get("education"):
        add_sec_hdr("Education")
        for edu in data["education"]:
            p_edu = doc.add_paragraph()
            p_edu.paragraph_format.space_before = Pt(1.5 * scale)
            p_edu.paragraph_format.space_after = Pt(0.3 * scale)
            p_edu.paragraph_format.keep_with_next = True
            p_edu.paragraph_format.tab_stops.add_tab_stop(
                Inches(8.5 - 2 * margin_lr), WD_TAB_ALIGNMENT.RIGHT
            )
            ru = p_edu.add_run(
                sanitize_ascii_hyphens(edu.get("institution", ""))
            )
            ru.bold = True
            ru.font.size = Pt(job_pt)
            ru.font.color.rgb = COLOR_BLACK

            rd = p_edu.add_run(
                f"\t{sanitize_ascii_hyphens(edu.get('date', ''))}"
            )
            rd.italic = True
            rd.font.size = Pt(body_pt)
            rd.font.color.rgb = COLOR_BLACK

            p_deg = doc.add_paragraph()
            p_deg.paragraph_format.space_before = Pt(0)
            p_deg.paragraph_format.space_after = Pt(0.5 * scale)
            p_deg.paragraph_format.keep_with_next = True
            rdeg = p_deg.add_run(sanitize_ascii_hyphens(edu.get("degree", "")))
            rdeg.font.size = Pt(body_pt)
            rdeg.font.color.rgb = COLOR_BLACK
            if edu.get("gpa"):
                rgpa = p_deg.add_run(f"  |  GPA: {edu['gpa']}")
                rgpa.bold = True
                rgpa.font.size = Pt(body_pt)
                rgpa.font.color.rgb = COLOR_BLACK

            for b in edu.get("highlights", []):
                if isinstance(b, dict):
                    add_bullet(b.get("prefix", ""), b.get("text", ""))
                else:
                    add_bullet("", str(b))

    doc.save(output_path)


def cmd_build_cv(args):
    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    docx_abs = os.path.abspath(args.docx_out)
    pdf_abs = os.path.abspath(args.pdf_out)

    scales = [1.0, 0.96, 0.92, 0.88]
    final_pages = 1
    used_scale = 1.0

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        for sc in scales:
            used_scale = sc
            render_cv_docx(data, docx_abs, scale=sc)
            doc = word.Documents.Open(docx_abs, False, True)
            final_pages = doc.ComputeStatistics(2)
            if final_pages <= 1 or sc == scales[-1]:
                try:
                    doc.ExportAsFixedFormat(pdf_abs, 17)
                except Exception:
                    base, ext = os.path.splitext(pdf_abs)
                    pdf_abs = f"{base}_Audited{ext}"
                    doc.ExportAsFixedFormat(pdf_abs, 17)
                doc.Close(False)
                break
            doc.Close(False)
    finally:
        word.Quit()

    report = {
        "status": "PASS" if final_pages == 1 else "WARN_MULTI_PAGE",
        "docx_path": docx_abs,
        "pdf_path": pdf_abs,
        "word_pages": final_pages,
        "scale_applied": used_scale,
    }
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    print(f"Success! CV generated ({final_pages} page). Report: {args.output}")


def cmd_audit_ats(args):
    pdf_abs = os.path.abspath(args.pdf)
    reader = pypdf.PdfReader(pdf_abs)
    num_pages = len(reader.pages)
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    required_headers = ["SUMMARY", "EXPERIENCE", "SKILLS", "EDUCATION"]
    missing_headers = [
        h for h in required_headers if h not in full_text.upper()
    ]
    has_en_dash = "\u2013" in full_text or "\u2014" in full_text

    kw_list = [
        k.strip() for k in (args.keywords or "").split(",") if k.strip()
    ]
    matched_kw = [k for k in kw_list if k.lower() in full_text.lower()]
    missing_kw = [k for k in kw_list if k.lower() not in full_text.lower()]
    match_pct = (
        round(len(matched_kw) / len(kw_list) * 100, 1) if kw_list else 100.0
    )

    passed = (
        num_pages == 1 and len(missing_headers) == 0 and not has_en_dash
    )
    report = {
        "status": "PASS" if passed else "NEEDS_FIX",
        "pdf_path": pdf_abs,
        "pages": num_pages,
        "single_page_pass": num_pages == 1,
        "missing_standard_headers": missing_headers,
        "unicode_dash_warning": has_en_dash,
        "keyword_match_percent": match_pct,
        "matched_keywords": matched_kw,
        "missing_keywords": missing_kw,
        "extracted_char_count": len(full_text),
    }
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    print(
        f"ATS Audit complete: {report['status']} (Pages: {num_pages}, Keyword Match: {match_pct}%). Report: {args.output}"
    )


def cmd_build_cover_letter(args):
    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    docx_abs = os.path.abspath(args.docx_out)
    pdf_abs = os.path.abspath(args.pdf_out)

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = COLOR_BLACK

    p_hdr = doc.add_paragraph()
    r_nm = p_hdr.add_run(data.get("name", "CANDIDATE NAME").upper() + "\n")
    r_nm.bold = True
    r_nm.font.size = Pt(14)
    r_ct = p_hdr.add_run(sanitize_ascii_hyphens(data.get("contact_line", "")))
    r_ct.font.size = Pt(9.5)

    p_dt = doc.add_paragraph(sanitize_ascii_hyphens(data.get("date_line", "")))
    p_dt.paragraph_format.space_after = Pt(8)

    p_rc = doc.add_paragraph(sanitize_ascii_hyphens(data.get("recipient", "Yth. Bapak/Ibu HRD")))
    p_rc.paragraph_format.space_after = Pt(8)

    for para in data.get("paragraphs", []):
        p = doc.add_paragraph(sanitize_ascii_hyphens(para))
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.12

    p_sig = doc.add_paragraph(f"\nHormat saya,\n\n{data.get('name', '')}")
    doc.save(docx_abs)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        wdoc = word.Documents.Open(docx_abs, False, True)
        pages = wdoc.ComputeStatistics(2)
        wdoc.ExportAsFixedFormat(pdf_abs, 17)
        wdoc.Close(False)
    finally:
        word.Quit()

    report = {"status": "PASS" if pages == 1 else "WARN", "docx_path": docx_abs, "pdf_path": pdf_abs, "pages": pages}
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    print(f"Success! Cover letter generated ({pages} page). Report: {args.output}")


def cmd_build_portfolio(args):
    from pptx import Presentation
    from pptx.dml.color import RGBColor as PptRGB
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches as PptInches, Pt as PptPt

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    pptx_abs = os.path.abspath(args.pptx_out)
    pdf_abs = os.path.abspath(args.pdf_out)

    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    blank = prs.slide_layouts[6]

    bg_col = PptRGB(241, 245, 249)
    card_bg = PptRGB(255, 255, 255)
    border_col = PptRGB(203, 213, 225)
    dark_col = PptRGB(15, 23, 42)
    body_col = PptRGB(51, 65, 85)
    blue_col = PptRGB(29, 78, 216)

    for slide_spec in data.get("slides", []):
        s = prs.slides.add_slide(blank)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = bg_col
        bg.line.fill.background()

        tb = s.shapes.add_textbox(PptInches(0.8), PptInches(0.5), PptInches(11.7), PptInches(1.0))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = slide_spec.get("title", "PROJECT SHOWCASE")
        r.font.size = PptPt(24)
        r.font.bold = True
        r.font.color.rgb = dark_col

        if slide_spec.get("subtitle"):
            ps = tf.add_paragraph()
            rs = ps.add_run()
            rs.text = slide_spec["subtitle"]
            rs.font.size = PptPt(11)
            rs.font.color.rgb = blue_col

        card = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptInches(0.8), PptInches(1.7), PptInches(11.7), PptInches(5.2))
        card.fill.solid()
        card.fill.fore_color.rgb = card_bg
        card.line.color.rgb = border_col

        tbc = s.shapes.add_textbox(PptInches(1.0), PptInches(1.9), PptInches(11.3), PptInches(4.8))
        tfc = tbc.text_frame
        tfc.word_wrap = True
        for idx, b in enumerate(slide_spec.get("bullets", [])):
            pb = tfc.paragraphs[0] if idx == 0 else tfc.add_paragraph()
            pb.space_before = PptPt(6)
            rb = pb.add_run()
            rb.text = f"• {b}"
            rb.font.size = PptPt(10.5)
            rb.font.color.rgb = body_col

    prs.save(pptx_abs)

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(pptx_abs, WithWindow=False)
        pres.SaveAs(pdf_abs, 32)
        pres.Close()
    finally:
        ppt.Quit()

    report = {"status": "PASS", "pptx_path": pptx_abs, "pdf_path": pdf_abs, "slides": len(data.get("slides", []))}
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    print(f"Success! Portfolio generated. Report: {args.output}")


def main():
    parser = argparse.ArgumentParser(
        description="Universal ATS CV & Portfolio CLI"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_cv = sub.add_parser("build-cv", help="Build 1-page ATS CV DOCX & PDF")
    p_cv.add_argument("--input", required=True, help="Input JSON path")
    p_cv.add_argument("--docx-out", required=True, help="Output DOCX path")
    p_cv.add_argument("--pdf-out", required=True, help="Output PDF path")
    p_cv.add_argument("--output", required=True, help="Output JSON report path")
    p_cv.set_defaults(func=cmd_build_cv)

    p_aud = sub.add_parser("audit-ats", help="Audit PDF against ATS rules")
    p_aud.add_argument("--pdf", required=True, help="PDF file to audit")
    p_aud.add_argument(
        "--keywords", default="", help="Comma-separated target keywords"
    )
    p_aud.add_argument(
        "--output", required=True, help="Output JSON report path"
    )
    p_aud.set_defaults(func=cmd_audit_ats)

    p_cl = sub.add_parser("build-cover-letter", help="Build 1-page Cover Letter DOCX & PDF")
    p_cl.add_argument("--input", required=True, help="Input JSON path")
    p_cl.add_argument("--docx-out", required=True, help="Output DOCX path")
    p_cl.add_argument("--pdf-out", required=True, help="Output PDF path")
    p_cl.add_argument("--output", required=True, help="Output JSON report path")
    p_cl.set_defaults(func=cmd_build_cover_letter)

    p_pf = sub.add_parser("build-portfolio", help="Build 16:9 Portfolio PPTX & PDF")
    p_pf.add_argument("--input", required=True, help="Input JSON path")
    p_pf.add_argument("--pptx-out", required=True, help="Output PPTX path")
    p_pf.add_argument("--pdf-out", required=True, help="Output PDF path")
    p_pf.add_argument("--output", required=True, help="Output JSON report path")
    p_pf.set_defaults(func=cmd_build_portfolio)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
