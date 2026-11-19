#!/usr/bin/env python3
"""Prepare truthful, ATS-friendly application packets for LinkedIn roles.

This tool creates local documents only. It does not log into LinkedIn, upload
files, answer screening questions, or submit applications.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
PROFILE_DEFAULT = ROOT / "config" / "candidate_profile.json"
OUTPUT_DEFAULT = ROOT / "output" / "applications"
PYTHON = Path(sys.executable)
RENDERER = None
STOP_WORDS = {
    "and", "the", "for", "with", "from", "into", "your", "you", "our", "their", "will",
    "have", "has", "this", "that", "are", "was", "were", "job", "role", "work", "working",
    "years", "year", "experience", "ability", "strong", "required", "preferred", "including",
    "such", "using", "use", "across", "based", "ensure", "team", "teams", "build", "building",
}


def slug(text: str) -> str:
    text = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return text[:70].strip("-") or "application"


def words(text: str) -> set[str]:
    return {
        x for x in re.findall(r"[a-z0-9+#.]+", text.lower())
        if len(x) > 2 and x not in STOP_WORDS
    }


def score_tags(tags: list[str], job_words: set[str], job_text: str) -> int:
    tag_words = set(re.findall(r"[a-z0-9+#.]+", " ".join(tags).lower())) - STOP_WORDS
    score = len(tag_words & job_words)
    lower = job_text.lower()
    for phrase in tags:
        if " " in phrase and phrase.lower() in lower:
            score += 2
    return score


def role_family(title: str, description: str) -> str:
    text = f"{title} {description}".lower()
    if any(k in text for k in ("forward deployed", "deployment strategist", "field engineer", "solutions engineer", "implementation engineer")):
        return "forward_deployed"
    if any(k in text for k in ("ai engineer", "artificial intelligence engineer", "machine learning engineer", "generative ai")):
        return "ai_engineering"
    if any(k in text for k in ("data engineer", "data platform", "analytics engineer")):
        return "data_engineering"
    return "software_engineering"


def choose_bullets(items: list[dict[str, Any]], job_words: set[str], job_text: str, limit: int, family: str) -> list[str]:
    family_tags = {
        "forward_deployed": {"fde", "customer", "discovery", "workshop", "stakeholder", "requirements", "workflow", "agent", "ai", "implementation", "solution"},
        "ai_engineering": {"ai", "generative", "openai", "agent", "application", "api", "engineering", "python", "testing"},
        "data_engineering": {"data", "etl", "migration", "quality", "modeling", "sql", "python", "pipeline"},
        "software_engineering": {"engineering", "application", "api", "software", "python", "testing", "deployment"},
    }
    priority = family_tags.get(family, set())
    ranked = []
    for order, item in enumerate(items):
        text = item["text"]
        tags = item.get("tags", [])
        tag_set = set(re.findall(r"[a-z0-9+#.]+", " ".join(tags).lower()))
        family_score = 3 * len(tag_set & priority & job_words)
        ranked.append((score_tags(tags, job_words, job_text) + len(words(text) & job_words) + family_score, -order, text))
    ranked.sort(reverse=True)
    return [x[2] for x in ranked[:limit]]


def selected_skills(profile: dict[str, Any], job_words: set[str], job_text: str, title: str) -> list[str]:
    ranked = []
    for index, skill in enumerate(profile["skills"]):
        hits = score_tags(skill.get("tags", []), job_words, job_text)
        hits += len(words(skill["name"]) & job_words)
        ranked.append((hits, -index, skill["name"], skill.get("tags", [])))
    ranked.sort(reverse=True)
    matched = [x for x in ranked if x[0] > 0]
    picked = [x[2] for x in matched[:11]]

    # Keep the candidate's actual prototyping stack visible for engineering roles.
    family = role_family("", job_text)
    if family in {"forward_deployed", "ai_engineering", "data_engineering", "software_engineering"}:
        for must_show in ("Python", "SQL", "Palantir Foundry"):
            if must_show in [s["name"] for s in profile["skills"]] and must_show not in picked:
                picked.append(must_show)

    # SAP is included in the core list only when the posting itself emphasizes it.
    sap_emphasized = "sap" in title.lower() or bool(re.search(r"\bsap\b.{0,30}(required|preferred|experience|skills|expertise)", job_text.lower()))
    if not sap_emphasized:
        picked = [s for s in picked if not re.search(r"\bsap\b|bods|s/4hana|procure-to-pay", s, re.I)]
    return picked[:12]


def set_style(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin = section.bottom_margin = Inches(0.58)
    section.left_margin = section.right_margin = Inches(0.68)
    for name in ("Normal", "Title", "Subtitle", "Heading 1", "Heading 2", "List Bullet"):
        style = doc.styles[name]
        style.font.name = "Arial"
        style.font.color.rgb = RGBColor(0, 0, 0)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(3)
    normal.paragraph_format.line_spacing = 1.02
    title = doc.styles["Title"]
    title.font.size = Pt(20)
    title.font.bold = True
    title.paragraph_format.space_after = Pt(1)
    subtitle = doc.styles["Subtitle"]
    subtitle.font.size = Pt(11)
    subtitle.paragraph_format.space_after = Pt(3)
    for heading, size in (("Heading 1", 10.5), ("Heading 2", 10)):
        style = doc.styles[heading]
        style.font.size = Pt(size)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(7)
        style.paragraph_format.space_after = Pt(2)
        style.paragraph_format.keep_with_next = True
    bullet = doc.styles["List Bullet"]
    bullet.font.size = Pt(9.5)
    bullet.paragraph_format.space_after = Pt(2)
    bullet.paragraph_format.line_spacing = 1.0
    bullet.paragraph_format.left_indent = Inches(0.18)
    bullet.paragraph_format.first_line_indent = Inches(-0.12)
    # Word's built-in Title style can contribute a colored paragraph rule.
    for root in (doc._element, doc.styles.element):
        for border in list(root.iter(qn("w:pBdr"))):
            border.getparent().remove(border)


def add_role_bias(job_words: set[str], family: str) -> set[str]:
    bias = {
        "forward_deployed": {"fde", "customer", "discovery", "workshop", "stakeholder", "requirements", "workflow", "agent", "implementation", "solution"},
        "ai_engineering": {"ai", "generative", "openai", "agent", "application", "api", "engineering", "python"},
        "data_engineering": {"data", "etl", "migration", "quality", "modeling", "sql", "python", "pipeline"},
        "software_engineering": {"engineering", "application", "api", "software", "python", "testing", "deployment"},
    }
    return job_words | bias.get(family, set())


def add_bullet(doc: Document, text: str, keep_with_next: bool = False) -> None:
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.keep_together = True
    p.paragraph_format.widow_control = True
    p.paragraph_format.keep_with_next = keep_with_next


def add_heading(doc: Document, text: str) -> None:
    doc.add_paragraph(text, style="Heading 1")


def make_resume(profile: dict[str, Any], company: str, title: str, description: str, out: Path) -> None:
    job_text = f"{company} {title} {description}"
    family = role_family(title, description)
    job_words = add_role_bias(words(job_text), family)
    headlines = {
        "forward_deployed": "Forward Deployed Engineer | Data and AI Applications",
        "ai_engineering": "AI Application Engineer | Data Systems and Applied AI",
        "data_engineering": "Data Engineer | Integration and Applied Analytics",
        "software_engineering": "Software Engineer | Data and AI Applications",
    }
    openings = {
        "forward_deployed": "Customer-facing data and solutions professional with 9+ years across enterprise consulting, data integration, and application delivery. Built independent data and AI prototypes, including a Foundry procurement application and an OpenAI-powered assistant. Combines Python, SQL, customer workshops, requirements discovery, and hands-on implementation to move from business problem to working software.",
        "ai_engineering": "Software and data professional with 9+ years across enterprise consulting and application delivery. Built and operate an OpenAI-powered assistant integrating Python data processing, Node.js messaging, and live league data, alongside data applications and automated tests. Brings hands-on prototyping, data integration, and customer-facing delivery experience to applied AI engineering.",
        "data_engineering": "Data integration and application professional with 9+ years across enterprise consulting and SAP migration programs. Developed SAP BODS ETL pipelines and built Python data transformations with quality checks and reconciled metrics. Adds customer discovery, data modeling, and application prototyping experience to data engineering work.",
        "software_engineering": "Application and data professional with 9+ years across enterprise consulting, integration, and delivery. Built personal software projects using Python, Node.js, SQLite, Docker, Linux, and OpenAI APIs, and delivered enterprise migration and testing work. Combines practical implementation with customer discovery and production support experience.",
    }
    if company.lower() == "titan" and "product focus" in title.lower():
        headlines[family] = "Forward Deployed Engineer | Applied AI and Product Delivery"
        openings[family] = (
            "Client-facing product and delivery professional with 9+ years across enterprise consulting and workflow implementation. "
            "Facilitated stakeholder workshops, translated procurement needs into functional designs, coordinated integrations and testing, "
            "and built independent user-facing data and AI prototypes. Pair Python and SQL fundamentals with practical discovery, iterative delivery, "
            "and clear communication across business and technical teams; ready to deepen production agent engineering in regulated financial workflows."
        )
    elif company.lower() == "titan":
        headlines[family] = "Forward Deployed Engineer | Applied AI and Customer Delivery"
        openings[family] = (
            "Applied AI and forward-deployed solutions professional with 9+ years in enterprise consulting and delivery. "
            "Built a Python/Node.js/OpenAI assistant and an independent Foundry workflow application with data-quality controls and KPI views. "
            "Combines hands-on integration and testing with client workshops, requirements discovery, iterative delivery, and operational follow-through; "
            "ready to extend this foundation into production agentic systems."
        )
    elif company.lower() == "palantir" and "deployment strategist" in title.lower():
        headlines[family] = "Deployment Strategist | Foundry, Data, and Customer Workflows"
        openings[family] = (
            "Customer-facing data and workflow professional with 9+ years in enterprise consulting and delivery. "
            "Facilitated stakeholder workshops, mapped operations, translated requirements into functional designs, and coordinated implementation and testing across business and technical teams. "
            "Built an independent Palantir Foundry procurement application using Python-transformed data and user-facing Workshop views; it processed 5,825 PO items and surfaced 11 variance leads for review. "
            "Bring analytical rigor, clear customer communication, and an ownership mindset to complex, open-ended data problems."
        )
    elif company.lower() == "makai labs":
        headlines[family] = "Forward Deployed AI Analyst | Workflow Automation"
        openings[family] = (
            "Workflow automation and applied AI candidate with 9+ years in enterprise consulting, process discovery, data integration, and implementation. "
            "Facilitated procurement workshops, translated requirements into functional designs, and built Python/Foundry and OpenAI-powered prototypes. "
            "Strong fit for finding automatable work, prototyping practical solutions, and measuring dependable outcomes, with a clear path toward deeper AI engineering."
        )

    doc = Document()
    set_style(doc)
    p = doc.add_paragraph(profile["name"], style="Title")
    p.paragraph_format.keep_with_next = True
    doc.add_paragraph(headlines[family], style="Subtitle")
    contact = f'{profile["location"]} | {profile["phone"]} | {profile["email"]}'
    doc.add_paragraph(contact)
    doc.add_paragraph(profile["linkedin"])

    add_heading(doc, "Professional Summary")
    doc.add_paragraph(openings[family])

    add_heading(doc, "Technical and Delivery Skills")
    for skill in selected_skills(profile, job_words, job_text, title):
        doc.add_paragraph(skill)

    projects = []
    for index, project in enumerate(profile["projects"]):
        rank = score_tags(project.get("tags", []), job_words, job_text)
        rank += len(words(project["name"] + " " + project["summary"]) & job_words)
        projects.append((rank, -index, project))
    projects.sort(reverse=True, key=lambda x: (x[0], x[1]))
    add_heading(doc, "Professional Experience")
    for role in profile["experience"]:
        heading = doc.add_paragraph()
        heading.paragraph_format.keep_with_next = True
        heading.paragraph_format.space_before = Pt(4)
        # Current Accenture title follows the user's requested presentation.
        heading.add_run(f'{role["company"]} | {role["title"]} | {role["dates"]}').bold = True
        if role.get("summary"):
            summary = doc.add_paragraph(role["summary"])
            summary.paragraph_format.keep_with_next = True
        chosen = choose_bullets(role["bullets"], job_words, job_text, 4 if role["company"] == "Accenture" else 3, family)
        for index, bullet in enumerate(chosen):
            # Keep each employment entry intact across page boundaries.
            add_bullet(doc, bullet, keep_with_next=index < len(chosen) - 1)

    doc.add_page_break()
    add_heading(doc, "Selected Projects")
    for _, _, project in projects[:2]:
        label = doc.add_paragraph()
        label.paragraph_format.keep_with_next = True
        label.paragraph_format.space_before = Pt(3)
        label.add_run(f'{project["name"]} | {project["kind"]}').bold = True
        summary = doc.add_paragraph(project["summary"])
        summary.paragraph_format.keep_with_next = True
        for index, bullet in enumerate(project["bullets"][:3]):
            add_bullet(doc, bullet, keep_with_next=index < min(3, len(project["bullets"][:3])) - 1)

    add_heading(doc, "Education and Certifications")
    for item in profile["education"]:
        doc.add_paragraph(item)
    certs = profile["certifications"]
    if not re.search(r"\bsap\b", title.lower() + " " + description.lower()):
        certs = [cert for cert in certs if not re.search(r"\bsap\b", cert, re.I)]
    # The source profile includes all certifications; keep the list scannable.
    doc.add_paragraph(" | ".join(certs))

    doc.core_properties.title = f'{profile["name"]} - {title} - {company}'
    doc.core_properties.author = profile["name"]
    doc.core_properties.subject = "ATS-friendly tailored resume"
    for root in (doc._element, doc.styles.element):
        for border in list(root.iter(qn("w:pBdr"))):
            border.getparent().remove(border)
    doc.save(out)


def cover_letter_is_needed(description: str, mode: str) -> bool:
    if mode == "required":
        return True
    if mode == "none":
        return False
    text = description.lower()
    explicit = (
        r"cover letter.{0,45}(required|please include|please submit|must include|must submit)",
        r"(required|please include|please submit|must include|must submit).{0,45}cover letter",
        r"include a cover letter",
        r"submit a cover letter",
    )
    return any(re.search(pattern, text) for pattern in explicit)


def make_cover_letter(profile: dict[str, Any], company: str, title: str, description: str, out: Path) -> None:
    family = role_family(title, description)
    project = {
        "forward_deployed": "my Foundry procurement prototype, where I connected Python data transformation, Ontology modeling, and a user-facing Workshop to a measurable business workflow",
        "ai_engineering": "my OpenAI-powered assistant, where I connected Python data processing, Node.js messaging, and grounded responses over live league data",
        "data_engineering": "my SAP BODS migration work and Foundry prototype, which involved transformation logic, validation, data modeling, and reconciled outputs",
        "software_engineering": "my software projects using Python, Node.js, SQLite, Docker, Linux, and OpenAI APIs, alongside enterprise integration and testing work",
    }[family]
    doc = Document()
    set_style(doc)
    doc.add_paragraph(profile["name"], style="Title")
    doc.add_paragraph(f'{profile["location"]} | {profile["phone"]} | {profile["email"]}')
    doc.add_paragraph(profile["linkedin"])
    doc.add_paragraph(date.today().strftime("%B %d, %Y"))
    doc.add_paragraph(f"Hiring Team\n{company}")
    doc.add_paragraph(f"Re: {title}", style="Heading 1")
    doc.add_paragraph("Dear Hiring Team,")
    doc.add_paragraph(
        f"I am applying for the {title} role at {company}. I bring more than nine years of enterprise consulting and data delivery experience, with hands-on projects that connect software implementation to practical user needs."
    )
    doc.add_paragraph(
        f"A relevant example is {project}. In my enterprise work, I have also facilitated requirements workshops, coordinated testing and issue resolution, and worked across business and technical teams to move implementations toward readiness."
    )
    doc.add_paragraph(
        "I would bring a customer-focused approach, practical data integration experience, and a willingness to learn the parts of your stack that are new to me. I would welcome the chance to discuss how my project work and consulting background could support your team."
    )
    doc.add_paragraph(f"Sincerely,\n{profile['name']}")
    doc.core_properties.title = f'{profile["name"]} - Cover Letter - {title} - {company}'
    doc.core_properties.author = profile["name"]
    for root in (doc._element, doc.styles.element):
        for border in list(root.iter(qn("w:pBdr"))):
            border.getparent().remove(border)
    doc.save(out)


def render_pdf(docx_path: Path, render_dir: Path) -> Path:
    renderer = shutil.which("libreoffice") or shutil.which("soffice")
    if not renderer:
        raise RuntimeError("PDF export requires LibreOffice. Install it, or adapt render_pdf() for your preferred document renderer.")
    render_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [renderer, "--headless", "--convert-to", "pdf", "--outdir", str(render_dir), str(docx_path)],
        check=True,
        cwd=str(ROOT),
    )
    pdf = render_dir / f"{docx_path.stem}.pdf"
    if not pdf.exists() or pdf.stat().st_size == 0:
        raise RuntimeError(f"PDF render did not produce a file: {pdf}")
    return pdf


def write_application_review(out_dir: Path, company: str, title: str, url: str, salary: str, remote: str,
                             cover_needed: bool, eligibility: str) -> str:
    letter_line = "A cover letter was generated because the posting asks for one." if cover_needed else "No cover letter was generated because the posting does not explicitly request one."
    content = f"""# Application review

- Company: {company}
- Role: {title}
- Location / work mode: {remote or 'Not provided'}
- Advertised compensation: {salary or 'Not provided'}
- Criteria check: {eligibility}
- Job listing: {url or 'Not provided'}
- Application status: Not submitted

## Review before applying

- Confirm the listing is still active and its location, pay, and work mode meet your criteria.
- Review the résumé facts and the role title used for your Accenture experience.
- {letter_line}
- Answer screening questions only from confirmed personal information. Stop for user input on work authorization, sponsorship, travel, salary expectations, or any other unknown answer.
- Obtain role-specific approval before uploading the résumé or entering personal contact details into the employer's application system.
- Submit through the listing's application portal only after the packet and application answers are reviewed.

The generator creates files locally. It does not upload documents or submit an application.
"""
    review_name = "linkedin-review.md" if url.startswith("https://www.linkedin.com/jobs/") else "application-review.md"
    (out_dir / review_name).write_text(content, encoding="utf-8")
    return review_name


def prepare(args: argparse.Namespace) -> Path:
    profile = json.loads(args.profile.read_text(encoding="utf-8"))
    description = args.description.read_text(encoding="utf-8") if args.description else ""
    if not args.company.strip() or not args.title.strip():
        raise ValueError("Company and role title are required.")
    if not (args.listing_url.startswith("https://www.linkedin.com/jobs/") or args.listing_url.startswith("https://jobs.lever.co/")):
        raise ValueError("Provide a LinkedIn or Lever job listing URL so the packet stays tied to the posting.")
    if not args.description:
        raise ValueError("Provide --description with the job posting text for role-specific tailoring.")
    stamp = date.today().strftime("%Y%m%d")
    key = args.job_id or slug(f"{args.company}-{args.title}")
    folder = args.output_root / f"{stamp}_{slug(args.company)}_{slug(args.title)}_{slug(key)}"
    if folder.exists() and any(folder.iterdir()) and not args.force:
        raise FileExistsError(f"Output folder already contains a packet: {folder}; use --force to replace it.")
    folder.mkdir(parents=True, exist_ok=True)
    if args.force:
        for child in list(folder.iterdir()):
            shutil.rmtree(child) if child.is_dir() else child.unlink()

    resume_label = re.sub(r"[\\/:*?\"<>|]+", " ", f"{args.company} {profile['name']} Resume")
    resume_label = re.sub(r"\s+", " ", resume_label).strip()
    resume_path = folder / f"{resume_label}.docx"
    make_resume(profile, args.company, args.title, description, resume_path)
    render_dir = ROOT / "tmp" / "applications" / folder.name / "render"
    resume_rendered_pdf = render_pdf(resume_path, render_dir / "resume")
    resume_pdf = folder / f"{resume_label}.pdf"
    shutil.copy2(resume_rendered_pdf, resume_pdf)

    cover_needed = cover_letter_is_needed(description, args.cover_letter)
    files = [resume_path.name, resume_pdf.name]
    if cover_needed:
        cover_path = folder / "cover-letter.docx"
        make_cover_letter(profile, args.company, args.title, description, cover_path)
        cover_rendered_pdf = render_pdf(cover_path, render_dir / "cover-letter")
        shutil.copy2(cover_rendered_pdf, folder / "cover-letter.pdf")
        files.extend(["cover-letter.docx", "cover-letter.pdf"])

    if args.description:
        shutil.copy2(args.description, folder / "job-description.txt")
        files.append("job-description.txt")
    salary_numbers = []
    for value in re.findall(r"\$?\s*(\d{1,3}(?:,\d{3})+|\d{2,3}\s*[kK])", args.salary or ""):
        value = value.replace(",", "").strip().lower()
        salary_numbers.append(int(float(value[:-1].strip()) * 1000) if value.endswith("k") else int(value))
    if not args.remote or "remote" not in args.remote.lower():
        eligibility = "Does not meet the remote-only preference or work mode is unverified."
    elif not salary_numbers:
        eligibility = "Pay is unverified; confirm advertised compensation is above $130,000 before applying."
    elif min(salary_numbers) <= 130000:
        eligibility = "Does not meet the strictly-above-$130,000 minimum."
    else:
        eligibility = "Meets the remote and strictly-above-$130,000 filters based on the supplied listing details."
    review_name = write_application_review(folder, args.company, args.title, args.listing_url, args.salary or "", args.remote or "", cover_needed, eligibility)
    files.append(review_name)
    manifest = {
        "company": args.company,
        "role": args.title,
        "listing_url": args.listing_url or None,
        "linkedin_url": args.listing_url if args.listing_url.startswith("https://www.linkedin.com/jobs/") else None,
        "application_url": args.listing_url if args.listing_url.startswith("https://jobs.lever.co/") else None,
        "salary": args.salary or None,
        "remote": args.remote or None,
        "criteria_check": eligibility,
        "cover_letter": "included" if cover_needed else "not_required_by_posting",
        "application_status": "not_submitted",
        "submission_gate": "role-specific user approval required before sending resume or contact details",
        "files": files,
    }
    (folder / "application.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return folder


def main() -> int:
    parser = argparse.ArgumentParser(description="Create ATS-friendly LinkedIn application packets locally.")
    sub = parser.add_subparsers(dest="command", required=True)
    prep = sub.add_parser("prepare", help="Tailor documents and create a review packet for one LinkedIn listing.")
    prep.add_argument("--company", required=True)
    prep.add_argument("--title", required=True)
    prep.add_argument("--description", type=Path, required=True, help="Text file containing the job description.")
    prep.add_argument("--linkedin-url", "--listing-url", dest="listing_url", required=True, help="LinkedIn or Lever job listing URL.")
    prep.add_argument("--job-id", default="")
    prep.add_argument("--salary", default="")
    prep.add_argument("--remote", default="")
    prep.add_argument("--profile", type=Path, default=PROFILE_DEFAULT)
    prep.add_argument("--output-root", type=Path, default=OUTPUT_DEFAULT)
    prep.add_argument("--cover-letter", choices=("auto", "required", "none"), default="auto")
    prep.add_argument("--force", action="store_true", help="Replace contents of this exact packet folder.")
    args = parser.parse_args()
    try:
        folder = prepare(args)
    except Exception as exc:  # noqa: BLE001 - surface actionable CLI errors
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    print(f"Prepared application packet: {folder}")
    print("Application status: not submitted")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
