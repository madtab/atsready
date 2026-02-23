#!/usr/bin/env python3
"""
ATS Resume Converter — "Black Box" Engine
==========================================
Takes a PDF resume, extracts content, scores ATS-readability,
and outputs a clean, ATS-optimised .docx file.

Dependencies:
    pip install pdfplumber python-docx

Usage:
    python ats_converter.py input_resume.pdf output_resume.docx

Returns exit code 0 on success, 1 on failure.
Also prints a JSON summary to stdout with the ATS score and issues found.
"""

import sys
import json
import re
import os
from pathlib import Path

import pdfplumber
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ─── ATS Scoring Rules ────────────────────────────────────────────────────────

ATS_PENALTIES = {
    "no_email": ("No email address found", -15),
    "no_phone": ("No phone number found", -10),
    "uses_tables": ("PDF appears to use tables/columns (ATS parsers struggle with these)", -20),
    "too_short": ("Very little text extracted — may be image-based PDF", -30),
    "special_chars": ("Excessive special characters or symbols detected", -10),
    "no_sections": ("Could not detect standard section headings", -15),
    "too_long": ("Resume exceeds 2 pages of content", -5),
}

SECTION_HEADINGS = [
    "experience", "work experience", "employment", "professional experience",
    "education", "qualifications", "academic",
    "skills", "technical skills", "core competencies", "key skills",
    "summary", "profile", "personal statement", "objective", "about",
    "certifications", "certificates", "training",
    "projects", "portfolio",
    "achievements", "awards", "honours", "honors",
    "references", "referees",
    "languages", "interests", "hobbies", "volunteer", "volunteering",
    "publications", "research",
]


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from a PDF using pdfplumber."""
    full_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text.append(text)
    return "\n".join(full_text)


def detect_sections(text: str) -> list[dict]:
    """
    Detect resume sections by matching known heading patterns.
    Returns a list of {"heading": str, "content": str} dicts.
    """
    lines = text.split("\n")
    sections = []
    current_heading = "HEADER"
    current_lines = []

    for line in lines:
        stripped = line.strip()
        lower = stripped.lower().rstrip(":")

        # Check if this line is a section heading
        is_heading = False
        for heading in SECTION_HEADINGS:
            if lower == heading or lower.startswith(heading + " "):
                is_heading = True
                matched_heading = stripped
                break

        if is_heading:
            # Save previous section
            if current_lines:
                sections.append({
                    "heading": current_heading,
                    "content": "\n".join(current_lines).strip()
                })
            current_heading = matched_heading
            current_lines = []
        else:
            if stripped:
                current_lines.append(stripped)

    # Don't forget last section
    if current_lines:
        sections.append({
            "heading": current_heading,
            "content": "\n".join(current_lines).strip()
        })

    return sections


def score_ats_readability(text: str, sections: list[dict]) -> dict:
    """
    Score the resume for ATS compatibility (0-100).
    Returns {"score": int, "issues": [str], "passed": [str]}.
    """
    score = 100
    issues = []
    passed = []

    # Check for email
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    if re.search(email_pattern, text):
        passed.append("Email address detected")
    else:
        score += ATS_PENALTIES["no_email"][1]
        issues.append(ATS_PENALTIES["no_email"][0])

    # Check for phone
    phone_pattern = r'[\+]?[\d\s\-\(\)]{7,15}'
    if re.search(phone_pattern, text):
        passed.append("Phone number detected")
    else:
        score += ATS_PENALTIES["no_phone"][1]
        issues.append(ATS_PENALTIES["no_phone"][0])

    # Check text length (image-based PDFs extract very little)
    word_count = len(text.split())
    if word_count < 50:
        score += ATS_PENALTIES["too_short"][1]
        issues.append(ATS_PENALTIES["too_short"][0])
    else:
        passed.append(f"Text extraction successful ({word_count} words)")

    # Check for too many special characters (indicates decorative PDFs)
    special_count = len(re.findall(r'[■●►▪◆★☆•→←↑↓▶◀⬛⬜🔹🔸]', text))
    if special_count > 10:
        score += ATS_PENALTIES["special_chars"][1]
        issues.append(ATS_PENALTIES["special_chars"][0])
    else:
        passed.append("Clean character usage")

    # Check for standard section headings
    non_header_sections = [s for s in sections if s["heading"] != "HEADER"]
    if len(non_header_sections) >= 2:
        headings_found = [s["heading"] for s in non_header_sections]
        passed.append(f"Section headings detected: {', '.join(headings_found[:5])}")
    else:
        score += ATS_PENALTIES["no_sections"][1]
        issues.append(ATS_PENALTIES["no_sections"][0])

    # Check length (proxy via word count — ~500 words/page)
    if word_count > 1200:
        score += ATS_PENALTIES["too_long"][1]
        issues.append(ATS_PENALTIES["too_long"][0])
    else:
        passed.append("Appropriate length")

    return {
        "score": max(0, min(100, score)),
        "issues": issues,
        "passed": passed,
        "word_count": word_count,
    }


def build_ats_docx(sections: list[dict], output_path: str, score_data: dict):
    """
    Build a clean, ATS-optimised .docx file.
    Uses only simple formatting that ATS systems can reliably parse:
    - Single column layout
    - Standard fonts (Calibri)
    - Clear section headings
    - No tables, text boxes, or graphics
    - Consistent heading hierarchy
    """
    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # Set narrow margins for more content space
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    for i, sec in enumerate(sections):
        heading_text = sec["heading"]
        content_text = sec["content"]

        if heading_text == "HEADER":
            # First section is the name/contact block
            lines = content_text.split("\n")
            if lines:
                # Name line — large and bold
                name_para = doc.add_paragraph()
                name_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                name_run = name_para.add_run(lines[0])
                name_run.bold = True
                name_run.font.size = Pt(18)
                name_run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

                # Contact details
                for line in lines[1:]:
                    contact_para = doc.add_paragraph()
                    contact_para.paragraph_format.space_before = Pt(1)
                    contact_para.paragraph_format.space_after = Pt(1)
                    contact_run = contact_para.add_run(line)
                    contact_run.font.size = Pt(10)
                    contact_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

                # Add a thin line separator
                separator = doc.add_paragraph()
                separator.paragraph_format.space_before = Pt(6)
                separator.paragraph_format.space_after = Pt(6)
                sep_run = separator.add_run("─" * 70)
                sep_run.font.size = Pt(6)
                sep_run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
        else:
            # Section heading
            heading_para = doc.add_paragraph()
            heading_para.paragraph_format.space_before = Pt(14)
            heading_para.paragraph_format.space_after = Pt(4)
            heading_run = heading_para.add_run(heading_text.upper())
            heading_run.bold = True
            heading_run.font.size = Pt(12)
            heading_run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

            # Subtle underline effect
            underline_para = doc.add_paragraph()
            underline_para.paragraph_format.space_before = Pt(0)
            underline_para.paragraph_format.space_after = Pt(6)
            ul_run = underline_para.add_run("─" * 70)
            ul_run.font.size = Pt(4)
            ul_run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)

            # Section content — preserve line breaks as paragraphs
            content_lines = content_text.split("\n")
            for line in content_lines:
                if line.strip():
                    para = doc.add_paragraph()
                    para.paragraph_format.space_before = Pt(2)
                    para.paragraph_format.space_after = Pt(2)
                    run = para.add_run(line.strip())
                    run.font.size = Pt(11)

    # Add ATS score footer
    doc.add_paragraph()  # spacer
    footer_sep = doc.add_paragraph()
    sep_run = footer_sep.add_run("─" * 70)
    sep_run.font.size = Pt(4)
    sep_run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)

    # Score headline
    score_para = doc.add_paragraph()
    score_para.paragraph_format.space_after = Pt(2)
    score_run = score_para.add_run(
        f"ATS Readability Score: {score_data['score']}/100"
    )
    score_run.font.size = Pt(10)
    score_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    score_run.bold = True

    # Passed checks
    if score_data.get("passed"):
        passed_para = doc.add_paragraph()
        passed_para.paragraph_format.space_before = Pt(4)
        passed_para.paragraph_format.space_after = Pt(1)
        passed_label = passed_para.add_run("✓ Passed: ")
        passed_label.font.size = Pt(9)
        passed_label.font.color.rgb = RGBColor(0x4A, 0x67, 0x41)
        passed_label.bold = True
        passed_text = passed_para.add_run(
            " · ".join(score_data["passed"])
        )
        passed_text.font.size = Pt(9)
        passed_text.font.color.rgb = RGBColor(0x4A, 0x67, 0x41)

    # Issues found
    if score_data.get("issues"):
        issues_para = doc.add_paragraph()
        issues_para.paragraph_format.space_before = Pt(2)
        issues_para.paragraph_format.space_after = Pt(1)
        issues_label = issues_para.add_run("✗ To improve: ")
        issues_label.font.size = Pt(9)
        issues_label.font.color.rgb = RGBColor(0xC4, 0x55, 0x3A)
        issues_label.bold = True
        issues_text = issues_para.add_run(
            " · ".join(score_data["issues"])
        )
        issues_text.font.size = Pt(9)
        issues_text.font.color.rgb = RGBColor(0xC4, 0x55, 0x3A)
    else:
        perfect_para = doc.add_paragraph()
        perfect_para.paragraph_format.space_before = Pt(2)
        perfect_para.paragraph_format.space_after = Pt(1)
        perfect_run = perfect_para.add_run(
            "No issues found — your resume is fully ATS-optimised."
        )
        perfect_run.font.size = Pt(9)
        perfect_run.font.color.rgb = RGBColor(0x4A, 0x67, 0x41)

    # Branding
    brand_para = doc.add_paragraph()
    brand_para.paragraph_format.space_before = Pt(6)
    brand_run = brand_para.add_run("Generated by ATSReady.co.uk")
    brand_run.font.size = Pt(8)
    brand_run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    brand_run.italic = True

    doc.save(output_path)


def convert_resume(input_pdf: str, output_docx: str) -> dict:
    """
    Main conversion pipeline.
    Returns a JSON-serialisable summary dict.
    """
    # 1. Extract text
    raw_text = extract_text_from_pdf(input_pdf)

    if not raw_text.strip():
        return {
            "success": False,
            "error": "Could not extract any text from this PDF. "
                     "It may be an image-only PDF. Please ensure your "
                     "resume contains selectable text.",
            "score": 0,
        }

    # 2. Detect sections
    sections = detect_sections(raw_text)

    # 3. Score ATS readability
    score_data = score_ats_readability(raw_text, sections)

    # 4. Build clean DOCX
    build_ats_docx(sections, output_docx, score_data)

    return {
        "success": True,
        "score": score_data["score"],
        "issues": score_data["issues"],
        "passed": score_data["passed"],
        "word_count": score_data["word_count"],
        "sections_detected": len([s for s in sections if s["heading"] != "HEADER"]),
        "output_file": output_docx,
    }


# ─── CLI Entry Point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python ats_converter.py <input.pdf> <output.docx>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(json.dumps({"success": False, "error": f"File not found: {input_path}"}))
        sys.exit(1)

    result = convert_resume(input_path, output_path)
    print(json.dumps(result, indent=2))

    sys.exit(0 if result["success"] else 1)
