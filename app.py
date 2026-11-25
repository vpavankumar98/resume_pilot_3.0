import io
import json
import os
import re
import time
import zipfile
from datetime import datetime
from urllib.parse import urlparse, unquote

import pandas as pd
import requests
import streamlit as st
import fitz
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="ResumePilot 3.0", page_icon="📄", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1450px; padding-top: 1.4rem;}
.hero {padding:1.5rem 1.8rem;border-radius:18px;color:white;background:linear-gradient(115deg,#14213d,#245be8);}
.small-muted {color:#667085;font-size:.9rem}
div[data-testid="stMetric"] {background:#f7f9fc;padding:12px;border-radius:12px}
</style>
""", unsafe_allow_html=True)

MAX_JOBS = 500
TIMEOUT = 15
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; ResumePilot/3.0)"}
WEIGHTS = {
    "Role relevance": 40,
    "Experience": 25,
    "Technical skills": 20,
    "Qualifications": 8,
    "Location": 4,
    "Other": 3,
}
STOP = set("""the and for with from that this will your their about into years work team role job using good strong ability experience skills required preferred responsibilities qualifications company position our you are have has not all any can may should must""".split())
SECTION_ALIASES = {
    "summary": ["professional summary", "summary", "profile", "career summary", "professional profile", "objective"],
    "skills": ["technical skills", "skills", "core competencies", "technical competencies", "key skills"],
    "experience": ["professional experience", "work experience", "employment history", "experience"],
    "projects": ["projects", "personal projects", "academic projects", "selected projects"],
    "education": ["education", "academic background", "qualifications"],
    "certifications": ["certifications", "certificates", "licenses", "licenses and certifications"],
}
SECTION_ORDER = ["summary", "skills", "experience", "projects", "education", "certifications"]
SECTION_LABELS = {
    "summary": "PROFESSIONAL SUMMARY",
    "skills": "TECHNICAL SKILLS",
    "experience": "PROFESSIONAL EXPERIENCE",
    "projects": "PROJECTS",
    "education": "EDUCATION",
    "certifications": "CERTIFICATIONS",
}

for k, v in {
    "master": None, "master_name": "", "master_sections": {},
    "jobs": [], "last_import_count": 0
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def words(s):
    return set(re.findall(r"[a-z][a-z0-9+#.-]{1,}", (s or "").lower())) - STOP


def heading_key(line):
    x = re.sub(r"[^a-z ]", "", line.lower()).strip()
    for key, aliases in SECTION_ALIASES.items():
        if x in aliases:
            return key
    return None


def split_into_sections(lines):
    """Preserve paragraph/bullet order. Only recognize section boundaries; never sort lines."""
    sections = {k: [] for k in SECTION_ORDER}
    current = "summary"
    preamble = []
    found_any = False
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        key = heading_key(line)
        if key:
            current = key
            found_any = True
            continue
        if not found_any and current == "summary":
            preamble.append(line)
        else:
            sections[current].append(line)
    if preamble:
        sections["summary"] = preamble + sections["summary"]
    # Keep unclassified content in summary rather than discarding it.
    return sections


def read_resume(upload):
    data = upload.getvalue()
    name = upload.name.lower()
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        lines = []
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                # Preserve list markers where present.
                if p.style and p.style.name and "List" in p.style.name and not txt.startswith(("•", "-", "▪")):
                    txt = "• " + txt
                lines.append(txt)
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    lines.append(" | ".join(cells))
        return split_into_sections(lines)
    if name.endswith(".pdf"):
        pdf = fitz.open(stream=data, filetype="pdf")
        lines = []
        for page in pdf:
            lines.extend([x.strip() for x in page.get_text("text").splitlines() if x.strip()])
        return split_into_sections(lines)
    raise ValueError("Please upload a text-based PDF or DOCX resume.")


def sections_to_text(sections):
    out = []
    for key in SECTION_ORDER:
        items = sections.get(key, [])
        if items:
            out.append(SECTION_LABELS[key])
            out.extend(items)
            out.append("")
    return "\n".join(out).strip()


def extract_urls(upload):
    data, name = upload.getvalue(), upload.name.lower()
    found = []
    if name.endswith(".pdf"):
        pdf = fitz.open(stream=data, filetype="pdf")
        raw = "\n".join(p.get_text() for p in pdf)
        found = re.findall(r'https?://[^\s<>"\']+', raw)
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(data), header=None, dtype=str)
        for cell in df.fillna("").values.flatten():
            found.extend(re.findall(r'https?://[^\s<>"\']+', str(cell)))
    elif name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(data), header=None, dtype=str, on_bad_lines="skip")
        for cell in df.fillna("").values.flatten():
            found.extend(re.findall(r'https?://[^\s<>"\']+', str(cell)))
    elif name.endswith((".txt", ".md")):
        found = re.findall(r'https?://[^\s<>"\']+', data.decode("utf-8", errors="ignore"))
    else:
        raise ValueError("Supported formats: PDF, XLSX, XLS, CSV, TXT.")
    cleaned = []
    for u in found:
        u = u.rstrip(".,);]}")
        if u.startswith(("http://", "https://")) and u not in cleaned:
            cleaned.append(u)
    return cleaned


def walk_jsonld(obj, results):
    if isinstance(obj, dict):
        typ = obj.get("@type", "")
        types = typ if isinstance(typ, list) else [typ]
        if any(str(t).lower() in {"jobposting", "jobposting"} for t in types):
            results.append(obj)
        for v in obj.values():
            walk_jsonld(v, results)
    elif isinstance(obj, list):
        for x in obj:
            walk_jsonld(x, results)


def get_nested(obj, *path):
    cur = obj
    for p in path:
        if not isinstance(cur, dict):
            return ""
        cur = cur.get(p)
    if isinstance(cur, dict):
        return cur.get("name", "")
    return cur or ""


def extract_job_metadata(soup, url):
    title, company, location, description = "", "", "", ""
    json_jobs = []
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or script.get_text())
            walk_jsonld(data, json_jobs)
        except Exception:
            pass
    if json_jobs:
        j = json_jobs[0]
        title = norm(j.get("title"))
        company = norm(get_nested(j, "hiringOrganization", "name"))
        loc = j.get("jobLocation", {})
        if isinstance(loc, list):
            loc = loc[0] if loc else {}
        address = loc.get("address", {}) if isinstance(loc, dict) else {}
        if isinstance(address, list):
            address = address[0] if address else {}
        location = ", ".join(filter(None, [norm(address.get("addressLocality")), norm(address.get("addressRegion")), norm(address.get("addressCountry"))]))
        description = norm(BeautifulSoup(j.get("description", ""), "html.parser").get_text(" ", strip=True))
    # Metadata fallbacks.
    if not title:
        for selector, attr in [('meta[property="og:title"]', "content"), ('meta[name="twitter:title"]', "content")]:
            tag = soup.select_one(selector)
            if tag and tag.get(attr):
                title = norm(tag.get(attr))
                break
    if not company:
        for selector, attr in [('meta[property="og:site_name"]', "content"), ('meta[name="author"]', "content")]:
            tag = soup.select_one(selector)
            if tag and tag.get(attr):
                company = norm(tag.get(attr))
                break
    if not title and soup.title:
        title = norm(soup.title.get_text(" ", strip=True))
    return title, company, location, description


def fetch_job(url):
    try:
        p = urlparse(url)
        if p.scheme not in ("http", "https") or not p.netloc:
            return None, "Invalid URL"
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code >= 400:
            return None, f"HTTP {r.status_code}"
        soup = BeautifulSoup(r.text, "html.parser")
        title, company, location, json_desc = extract_job_metadata(soup, url)
        for tag in soup(["script", "style", "nav", "footer", "header", "noscript", "svg"]):
            tag.decompose()
        desc = json_desc
        if len(desc) < 250:
            selectors = [
                "[data-automation='jobDescription']", "[data-testid*='job-description']",
                ".job-description", "#job-description", ".description", "main", "article"
            ]
            for sel in selectors:
                node = soup.select_one(sel)
                if node:
                    candidate = norm(node.get_text(" ", strip=True))
                    if len(candidate) > len(desc):
                        desc = candidate
                    if len(desc) > 500:
                        break
        if len(desc) < 250:
            desc = norm(soup.get_text(" ", strip=True))
        if not title:
            title = "Job title needs confirmation"
        if not company:
            company = p.netloc.replace("www.", "")
        if len(desc) < 250:
            return {"title": title, "company": company, "location": location, "description": ""}, "Job description not publicly extractable; paste JD manually"
        return {"title": title, "company": company, "location": location, "description": desc[:22000]}, None
    except requests.RequestException as e:
        return None, f"Could not fetch page ({type(e).__name__})"


def score_category(resume, jd, category, job_title="", location=""):
    rt, jt = words(resume), words(jd)
    if not jt:
        return 0
    if category == "Role relevance":
        role_terms = words(job_title)
        if not role_terms:
            return 50
        return round(100 * len(role_terms & rt) / max(1, len(role_terms)))
    if category == "Experience":
        # Compare years requested against explicit years in resume; if JD has no numeric years, use responsibility overlap.
        req = re.findall(r"(\d+)\+?\s*(?:years|yrs)", jd.lower())
        have = re.findall(r"(\d+)\+?\s*(?:years|yrs)", resume.lower())
        if req and have:
            needed, available = max(map(int, req)), max(map(int, have))
            return min(100, round(100 * available / max(1, needed)))
        req_terms = jt & {"develop", "build", "deploy", "manage", "design", "lead", "support", "implement", "analyze", "maintain"}
        return round(100 * len(req_terms & rt) / max(1, len(req_terms))) if req_terms else 50
    if category == "Technical skills":
        skill_terms = jt & rt
        technical = {x for x in jt if any(c in x for c in ["python", "sql", "aws", "java", "spark", "powerbi", "power-bi", "azure", "gcp", "ml", "ai", "docker", "kubernetes", "tensorflow", "pytorch", "excel", "tableau", "snowflake", "databricks", "react", "node", "javascript"])}
        pool = technical or skill_terms
        return round(100 * len(pool & rt) / max(1, len(pool)))
    if category == "Qualifications":
        q = jt & {"bachelor", "master", "degree", "certification", "certified", "phd", "btech", "b.e", "mba"}
        return round(100 * len(q & rt) / max(1, len(q))) if q else 50
    if category == "Location":
        if not location:
            return 50
        return 100 if words(location) & rt else 50
    # Other: remaining job-description keyword coverage.
    relevant = jt - words(job_title)
    return round(100 * len(relevant & rt) / max(1, len(relevant)))


def weighted_match(resume, jd, title, location=""):
    components = {}
    for cat, weight in WEIGHTS.items():
        components[cat] = score_category(resume, jd, cat, title, location)
    total = round(sum(components[k] * WEIGHTS[k] for k in WEIGHTS) / 100)
    return total, components


def ai_rewrite_summary(original_summary, jd, title, company, api_key):
    if not api_key or not original_summary:
        return original_summary, "Original summary preserved (AI summary rewrite not configured)."
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        prompt = f"""
Rewrite ONLY the professional summary for the target role. Preserve all factual claims.
Do not add skills, employers, years, metrics, certifications, education or achievements.
Use only information explicitly present in the original summary. If a JD asks for something
not evidenced in the summary, do not claim it. Return 3-4 concise bullets or sentences only.
Target role: {title}
Company: {company}
Original summary:
{original_summary}
Job description:
{jd[:8000]}
"""
        response = client.models.generate_content(
            model=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
            contents=prompt
        )
        result = (response.text or "").strip()
        if not result:
            return original_summary, "AI returned no summary; original retained."
        return result, "AI rewrote summary only; all other sections remain in original order."
    except Exception as e:
        return original_summary, f"AI unavailable; original summary retained ({type(e).__name__})."


def tailored_sections(master_sections, jd, title, company, api_key):
    # Deep-copy lists; do not sort or delete lines in any section.
    out = {k: list(master_sections.get(k, [])) for k in SECTION_ORDER}
    original_summary = "\n".join(out["summary"])
    new_summary, note = ai_rewrite_summary(original_summary, jd, title, company, api_key)
    if new_summary and new_summary != original_summary:
        out["summary"] = [x.strip() for x in new_summary.splitlines() if x.strip()]
    # Preserve technical skills exactly and in their original order in preserve mode.
    return out, note


def build_docx(sections, title, company):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(9)
    doc.add_heading("RESUME", 0)
    p = doc.add_paragraph()
    p.add_run(f"Target role: {title}").bold = True
    if company:
        p.add_run(f" | {company}")
    for key in SECTION_ORDER:
        items = sections.get(key, [])
        if not items:
            continue
        doc.add_heading(SECTION_LABELS[key], level=1)
        for item in items:
            item = item.strip()
            if not item:
                continue
            # Retain bullet-like lines as bullets; do not invent bullets for all content.
            if item.startswith(("•", "▪", "-", "–")):
                para = doc.add_paragraph(style="List Bullet")
                para.add_run(item.lstrip("•▪-– ").strip())
            else:
                doc.add_paragraph(item)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def make_job(url, idx, api_key):
    data, err = fetch_job(url)
    company = data["company"] if data else urlparse(url).netloc.replace("www.", "")
    title = data["title"] if data else "Job title needs confirmation"
    location = data["location"] if data else ""
    jd = data["description"] if data else ""
    job = {"id": idx, "url": url, "company": company, "title": title, "location": location,
           "jd": jd, "error": err, "score": None, "components": {}, "docx": None,
           "tailor_note": "", "status": "Needs JD review"}
    if jd and st.session_state.master_sections:
        master_text = sections_to_text(st.session_state.master_sections)
        score, comps = weighted_match(master_text, jd, title, location)
        tailored, note = tailored_sections(st.session_state.master_sections, jd, title, company, api_key)
        job.update({"score": score, "components": comps, "docx": build_docx(tailored, title, company),
                    "tailor_note": note, "status": "Resume ready"})
    return job


st.markdown('<div class="hero"><h1>📄 ResumePilot 3.0</h1><p>Structure-preserving resume tailoring · Weighted job matching · Bulk job import</p><b>500 URL import cap · Individual DOCX downloads</b></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("AI configuration")
    api_key = st.text_input("Gemini API key (optional)", value=os.getenv("GEMINI_API_KEY", ""), type="password")
    st.caption("Only the professional summary may be AI-rewritten. All other resume sections, facts, bullets and their order are preserved.")
    st.divider()
    st.caption("This MVP uses temporary Streamlit session memory. It does not yet provide user accounts, persistent private storage or multi-user isolation. Do not use with sensitive personal data until those protections are implemented.")

tab1, tab2, tab3 = st.tabs(["① Master Resume", "② Bulk Job Upload", "③ Job Dashboard"])

with tab1:
    st.subheader("Upload master resume")
    upload = st.file_uploader("PDF or DOCX", type=["pdf", "docx"], key="master_upload")
    if upload:
        try:
            sections = read_resume(upload)
            if not any(sections.values()):
                st.error("No selectable text found. Scanned PDFs need OCR before processing.")
            else:
                st.session_state.master_sections = sections
                st.session_state.master_name = upload.name
                st.success(f"Master resume loaded: {upload.name}")
        except Exception as e:
            st.error(f"Could not read resume: {e}")
    if st.session_state.master_sections:
        st.info(f"Active master resume: {st.session_state.master_name}")
        for key in SECTION_ORDER:
            with st.expander(SECTION_LABELS[key], expanded=(key == "summary")):
                st.text("\n".join(st.session_state.master_sections.get(key, []))[:7000] or "No section detected.")
        if st.button("Clear master resume"):
            st.session_state.master_sections = {}
            st.session_state.master_name = ""
            st.session_state.jobs = []
            st.rerun()

with tab2:
    st.subheader("Import up to 500 job URLs")
    st.caption("PDF, XLSX, XLS, CSV, TXT or paste links, one per line.")
    file = st.file_uploader("Upload job links file", type=["pdf", "xlsx", "xls", "csv", "txt"], key="jobs_file")
    pasted = st.text_area("Or paste multiple job URLs", height=120)
    urls = []
    if file:
        try:
            urls.extend(extract_urls(file))
            st.success(f"Extracted {len(urls)} unique URLs from file.")
        except Exception as e:
            st.error(str(e))
    urls.extend(re.findall(r'https?://[^\s<>"\']+', pasted))
    urls = list(dict.fromkeys([u.rstrip(".,);]}") for u in urls]))
    if len(urls) > MAX_JOBS:
        st.warning("Maximum 500 links per batch; extra links will be ignored.")
        urls = urls[:MAX_JOBS]
    st.metric("Unique job links ready", len(urls))
    if st.button("Automatically process all jobs", type="primary", use_container_width=True):
        if not st.session_state.master_sections:
            st.error("Upload your master resume first.")
        elif not urls:
            st.error("Upload a supported file or paste job URLs.")
        else:
            bar = st.progress(0)
            status = st.empty()
            jobs = []
            for i, url in enumerate(urls, 1):
                status.write(f"Processing {i}/{len(urls)}: {url[:100]}")
                jobs.append(make_job(url, i, api_key))
                bar.progress(i / len(urls))
                time.sleep(0.1)
            st.session_state.jobs = jobs
            status.success(f"Imported {len(jobs)} jobs.")
            st.success("Open Job Dashboard to review results.")

with tab3:
    st.subheader("Job dashboard")
    jobs = st.session_state.jobs
    if not jobs:
        st.info("Upload your master resume and import job URLs first.")
    else:
        ready = sum(j["status"] == "Resume ready" for j in jobs)
        c1, c2, c3 = st.columns(3)
        c1.metric("Total imported", len(jobs))
        c2.metric("Resumes ready", ready)
        c3.metric("Needs JD review", len(jobs) - ready)
        minimum = st.slider("Minimum match score", 0, 100, 0)
        sort = st.selectbox("Sort jobs", ["Highest match first", "Lowest match first", "Import order"])
        visible = [j for j in jobs if j["score"] is None or j["score"] >= minimum]
        if sort == "Highest match first":
            visible.sort(key=lambda x: x["score"] if x["score"] is not None else -1, reverse=True)
        elif sort == "Lowest match first":
            visible.sort(key=lambda x: x["score"] if x["score"] is not None else 101)
        for job in visible:
            with st.container(border=True):
                a, b = st.columns([4, 1])
                with a:
                    st.markdown(f"### {job['title']}")
                    st.caption(f"{job['company']} · {job['location'] or 'Location not detected'} · {job['status']}")
                with b:
                    st.metric("Match score", f"{job['score']}/100" if job["score"] is not None else "N/A")
                x1, x2, x3, x4 = st.columns(4)
                with x1:
                    st.link_button("View Job", job["url"], use_container_width=True)
                with x2:
                    st.link_button("Apply Job", job["url"], use_container_width=True)
                with x3:
                    with st.popover("Master Resume", use_container_width=True):
                        st.write(f"Original: {st.session_state.master_name}")
                        st.text(sections_to_text(st.session_state.master_sections)[:7000])
                with x4:
                    if job["docx"]:
                        st.download_button("Modified Resume", job["docx"], file_name=f"ResumePilot_{job['id']}_Resume.docx",
                                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                           key=f"dl_{job['id']}", use_container_width=True)
                    else:
                        st.button("Modified Resume", disabled=True, key=f"disabled_{job['id']}", use_container_width=True)
                if job["error"]:
                    st.warning(job["error"])
                    with st.expander("Add job description manually"):
                        manual = st.text_area("Paste full job description", key=f"manual_{job['id']}", height=160)
                        if st.button("Analyze and generate resume", key=f"analyze_{job['id']}"):
                            if len(manual.strip()) < 100:
                                st.error("Please paste a fuller job description.")
                            else:
                                score, comps = weighted_match(sections_to_text(st.session_state.master_sections), manual, job["title"], job["location"])
                                tailored, note = tailored_sections(st.session_state.master_sections, manual, job["title"], job["company"], api_key)
                                job.update({"jd": manual, "score": score, "components": comps, "docx": build_docx(tailored, job["title"], job["company"]),
                                            "tailor_note": note, "status": "Resume ready", "error": None})
                                st.rerun()
                with st.expander("Weighted match analysis and resume preview"):
                    if job["components"]:
                        for cat, val in job["components"].items():
                            st.write(f"**{cat} ({WEIGHTS[cat]}%):** {val}/100")
                        st.write(f"**Weighted total:** {job['score']}/100")
                    if job["tailor_note"]:
                        st.caption(job["tailor_note"])
                    st.write("Resume structure is preserved: summary, skills, experience, projects, education and certifications.")
st.divider()
st.caption("ResumePilot 3.0 · Weighted score is an estimate, not an ATS guarantee. Review every generated resume before applying.")
