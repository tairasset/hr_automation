import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import pandas as pd
import json
import os
import io
import time
from datetime import date
from groq import Groq
import groq as groq_lib
from dotenv import load_dotenv
from docx import Document
from striprtf.striprtf import rtf_to_text

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
)
from reportlab.graphics.shapes import Drawing, Circle, String as RLString
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

load_dotenv()


def _setup_pdf_fonts() -> tuple[str, str]:
    """Register a Unicode-capable TTF font (supports Cyrillic, etc.)."""
    candidates = [
        ("C:/Windows/Fonts/arial.ttf",   "C:/Windows/Fonts/arialbd.ttf"),
        ("C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for reg_path, bold_path in candidates:
        if os.path.exists(reg_path):
            try:
                pdfmetrics.registerFont(TTFont("UniFont", reg_path))
                bp = bold_path if os.path.exists(bold_path) else reg_path
                pdfmetrics.registerFont(TTFont("UniFont-Bold", bp))
                return "UniFont", "UniFont-Bold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"   # ASCII-only fallback


PDF_FONT, PDF_FONT_BOLD = _setup_pdf_fonts()


class StoredFile:
    """Wraps file bytes so accumulated uploads survive Streamlit reruns."""
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def read(self) -> bytes:
        return self._data

    def seek(self, _pos: int) -> None:
        pass  # always reads full bytes, seek is a no-op

st.set_page_config(
    page_title="AI Resume Screener",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #f5f7fa; }
    .score-circle-green {
        background: linear-gradient(135deg, #28a745, #20c997);
        color: white; border-radius: 50%; width: 90px; height: 90px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.8rem; font-weight: bold; margin: auto;
        box-shadow: 0 4px 12px rgba(40,167,69,0.4);
    }
    .score-circle-yellow {
        background: linear-gradient(135deg, #ffc107, #fd7e14);
        color: #212529; border-radius: 50%; width: 90px; height: 90px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.8rem; font-weight: bold; margin: auto;
        box-shadow: 0 4px 12px rgba(255,193,7,0.4);
    }
    .score-circle-red {
        background: linear-gradient(135deg, #dc3545, #c82333);
        color: white; border-radius: 50%; width: 90px; height: 90px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.8rem; font-weight: bold; margin: auto;
        box-shadow: 0 4px 12px rgba(220,53,69,0.4);
    }
    .borderline-banner {
        background: #fff3cd; border-left: 4px solid #ffc107;
        border-radius: 4px; padding: 10px 15px; color: #856404;
        font-weight: 500; margin-bottom: 12px;
    }
    .rec-badge {
        display: inline-block; padding: 4px 14px; border-radius: 20px;
        font-weight: 600; font-size: 0.9rem;
    }
    .strength-item { color: #155724; margin-bottom: 6px; }
    .weakness-item { color: #721c24; margin-bottom: 6px; }
    div[data-testid="stMetricValue"] { font-size: 2rem; }
    .iq-card {
        background: #f8f9fa; border: 1px solid #dee2e6;
        border-left: 4px solid #0d6efd;
        border-radius: 6px; padding: 12px 16px; margin-bottom: 10px;
    }
    .iq-num { font-weight: 700; color: #0d6efd; margin-right: 6px; }
    .iq-question { font-weight: 600; color: #212529; margin-bottom: 4px; }
    .iq-reason { color: #6c757d; font-size: 0.88rem; margin-left: 18px; }
    .vq-card {
        background: #fff8e1; border: 1px solid #ffc107;
        border-left: 4px solid #fd7e14;
        border-radius: 6px; padding: 12px 16px; margin-bottom: 10px;
    }
    .vq-claim { color: #856404; font-size: 0.85rem; margin-bottom: 4px; font-style: italic; }
    .vq-question { font-weight: 600; color: #212529; }
    /* Leaderboard */
    .lb-header {
        font-weight: 700; color: #6c757d; font-size: 0.8rem;
        text-transform: uppercase; letter-spacing: 0.05em;
        padding: 6px 0; border-bottom: 2px solid #dee2e6;
    }
    .lb-row {
        padding: 8px 0; border-bottom: 1px solid #f0f0f0;
        align-items: center;
    }
    .lb-row-gold {
        background: linear-gradient(90deg, #fffde7, #ffffff);
        border-left: 3px solid #ffd700; border-radius: 4px;
        padding: 8px 6px;
    }
    .best-match-badge {
        background: linear-gradient(135deg, #ffd700, #ffa000);
        color: #212529; padding: 2px 9px; border-radius: 10px;
        font-size: 0.72rem; font-weight: 700;
        display: inline-block; margin-left: 6px; vertical-align: middle;
    }
    .score-pill-green {
        background: #28a745; color: white; padding: 3px 10px;
        border-radius: 12px; font-weight: 700; font-size: 0.9rem;
        display: inline-block;
    }
    .score-pill-yellow {
        background: #ffc107; color: #212529; padding: 3px 10px;
        border-radius: 12px; font-weight: 700; font-size: 0.9rem;
        display: inline-block;
    }
    .score-pill-red {
        background: #dc3545; color: white; padding: 3px 10px;
        border-radius: 12px; font-weight: 700; font-size: 0.9rem;
        display: inline-block;
    }
    .lb-section {
        background: white; border-radius: 10px; padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 24px;
    }
</style>
""", unsafe_allow_html=True)


# ── Job-requirements constants ────────────────────────────────────────────────

SKILLS_QUICK = [
    # Tech
    "Python", "JavaScript", "SQL", "React", "Node.js",
    "Java", "C++", "TypeScript", "AWS", "Docker",
    "Git", "Machine Learning", "Data Analysis", "REST APIs",
    # Business / Management
    "Agile/Scrum", "Project Management", "Excel",
    # Sales
    "Sales", "B2B Sales", "CRM", "Lead Generation", "Negotiation", "Salesforce",
    # Marketing
    "Digital Marketing", "SEO", "Google Ads", "Content Marketing", "Email Marketing",
    # Design / Soft skills
    "Figma", "Photoshop", "Communication", "Leadership", "Problem-solving",
]

SKILLS_CATEGORIES = {
    "Tech": [
        "Python", "JavaScript", "TypeScript", "Java", "C++", "C#", "Go", "Rust", "PHP", "Ruby",
        "SQL", "NoSQL", "REST APIs", "GraphQL", "Docker", "Kubernetes",
        "AWS", "GCP", "Azure", "Git", "Linux", "Node.js",
        "React", "Vue.js", "Angular", "Django", "FastAPI", "Spring",
    ],
    "Data & Analytics": [
        "Machine Learning", "Deep Learning", "Data Analysis", "Data Science",
        "TensorFlow", "PyTorch", "Pandas", "NumPy", "Spark",
        "Tableau", "Power BI", "Excel", "R", "Statistics", "ETL", "Data Visualization",
    ],
    "Design": [
        "Figma", "Photoshop", "Illustrator", "Adobe XD", "Sketch",
        "UI/UX Design", "Prototyping", "User Research", "Wireframing",
        "InDesign", "After Effects",
    ],
    "Business": [
        "Project Management", "Business Analysis", "Financial Analysis",
        "Excel", "PowerPoint", "CRM", "Salesforce", "SAP", "ERP",
        "Process Improvement", "Risk Management", "Budget Management",
    ],
    "Sales": [
        "Sales", "B2B Sales", "B2C Sales", "Account Management",
        "Lead Generation", "CRM", "Negotiation", "Cold Calling", "Salesforce",
    ],
    "Marketing": [
        "Digital Marketing", "SEO", "SEM", "Google Ads", "Email Marketing",
        "Content Marketing", "Marketing Analytics", "A/B Testing", "HubSpot", "Mailchimp",
    ],
    "SMM & Content": [
        "Social Media Management", "Content Creation", "Copywriting",
        "Instagram", "TikTok", "LinkedIn", "Facebook Ads",
        "Community Management", "Brand Strategy",
    ],
    "Soft Skills": [
        "Communication", "Leadership", "Problem-solving", "Teamwork",
        "Time Management", "Adaptability", "Creativity", "Critical Thinking",
        "Emotional Intelligence", "Conflict Resolution", "Mentoring",
    ],
}


# ── Utilities ─────────────────────────────────────────────────────────────────

def score_tier(score: int) -> str:
    if score >= 75:
        return "green"
    elif score >= 50:
        return "yellow"
    return "red"


def score_pill_html(score: int) -> str:
    tier = score_tier(score)
    return f'<span class="score-pill-{tier}">{score}</span>'


def score_circle_html(score: int) -> str:
    tier = score_tier(score)
    return f'<div class="score-circle-{tier}">{score}</div>'


def recommendation_color(rec: str) -> str:
    return {
        "Strong Hire": "#28a745",
        "Hire": "#20c997",
        "Maybe": "#ffc107",
        "No Hire": "#dc3545",
    }.get(rec, "#6c757d")


def _extract_json(raw: str) -> str:
    """Pull the last complete JSON object from any model response.

    The model sometimes echoes resume text before the JSON — scanning from
    the end finds the real answer even when junk precedes it.
    """
    # 1. Strip markdown fences — prefer fenced block when present
    if "```" in raw:
        blocks = raw.split("```")
        for block in reversed(blocks[1::2]):
            block = block.strip()
            if block.startswith("json"):
                block = block[4:].strip()
            if block.startswith("{"):
                return block

    # 2. Find the LAST complete { … } with balanced braces
    last_close = raw.rfind("}")
    if last_close != -1:
        depth = 0
        for i in range(last_close, -1, -1):
            if raw[i] == "}":
                depth += 1
            elif raw[i] == "{":
                depth -= 1
                if depth == 0:
                    return raw[i:last_close + 1]

    return raw.strip()


# ── File extraction ───────────────────────────────────────────────────────────

def get_groq_client() -> Groq:
    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        st.error("GROQ_API_KEY not found. Copy `.env.example` to `.env` and add your key.")
        st.stop()
    return Groq(api_key=api_key)


def extract_text(uploaded_file) -> str:
    filename = uploaded_file.name.lower()
    raw = uploaded_file.read()
    uploaded_file.seek(0)

    if filename.endswith(".pdf"):
        parts = []
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return "\n".join(parts).strip()

    if filename.endswith(".docx"):
        doc = Document(io.BytesIO(raw))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip()).strip()

    if filename.endswith(".rtf"):
        return rtf_to_text(raw.decode("utf-8", errors="ignore")).strip()

    if filename.endswith(".txt"):
        return raw.decode("utf-8", errors="ignore").strip()

    try:
        return raw.decode("utf-8", errors="ignore").strip()
    except Exception:
        return ""


# ── Groq — Resume Analysis ────────────────────────────────────────────────────

SYSTEM_JSON = """\
You are an HR analysis API. Given a job description (with scoring weights) and a resume, return a JSON object with EXACTLY these fields:
  "score"               integer 0-100 (sum of the 4 category scores below)
  "experience_score"    integer 0 to experience weight — candidate's experience score
  "skills_score"        integer 0 to skills weight — candidate's technical skills score
  "education_score"     integer 0 to education weight — candidate's education score
  "soft_skills_score"   integer 0 to soft skills weight — candidate's soft skills/culture score
  "meets_min_experience" boolean — true if candidate meets minimum experience, or no minimum specified
  "meets_min_education" boolean — true if candidate meets education requirement, or none specified
  "strengths"           array of exactly 3 strings — specific strengths relevant to the job
  "weaknesses"          array of exactly 3 strings — specific gaps vs. the job requirements
  "hire_recommendation" one of: "Strong Hire", "Hire", "Maybe", "No Hire"
  "summary"             2-3 sentence overall assessment

Rules:
- The 4 category scores must sum exactly to "score"
- Each category score must be between 0 and its listed weight maximum
- Base every evaluation point on concrete evidence from the resume vs. the job requirements
- NEVER copy resume text or job description text into your output
- Output ONLY the JSON object — no prose, no markdown, no extra keys\
"""

ANALYSIS_PROMPT = """\
JOB DESCRIPTION:
{job_description}

SCORING WEIGHTS (score each category out of its max):
Experience: max {w_exp} | Technical Skills: max {w_skills} | Education: max {w_edu} | Soft Skills: max {w_soft}

CANDIDATE: {candidate_name}
RESUME:
{resume_text}

Evaluate this candidate. Score each category out of its weight max, then sum for total.\
"""

EMAIL_SYSTEM = ("You are an HR professional. Write only the email body — "
                "no subject line header, no meta-commentary, no placeholders in brackets.")

EMAIL_INVITE_PROMPT = """Write a professional interview invitation email for a strong candidate.

Candidate name: {name}
Position: {position}
Fit score: {score}/100
Top strength: {strength}

Requirements:
- Warm, professional tone
- Reference the position by name
- Naturally mention the candidate's top strength as a reason they stood out
- Invite them for an interview and mention flexible scheduling
- Use "our company" — never write placeholder brackets
- Under 200 words
- Start directly with "Dear {name}," """

EMAIL_REJECTION_PROMPT = """Write a respectful, compassionate rejection email.

Candidate name: {name}
Position: {position}

Requirements:
- Kind, encouraging tone
- Thank them sincerely for their time and application
- Inform them that the role has been filled with another candidate
- Encourage them to apply for future openings
- Use "our company" — never write placeholder brackets
- Under 150 words
- Do NOT mention any weaknesses or gaps
- Start directly with "Dear {name}," """

EMAIL_MORE_INFO_PROMPT = """Write a professional email requesting additional information or portfolio samples from a candidate.

Candidate name: {name}
Position: {position}
Top strength: {strength}

Requirements:
- Professional, encouraging tone
- Acknowledge their application positively
- Reference their top strength as something that caught attention
- Ask them to share portfolio samples or supplementary materials to help evaluate their fit
- Use "our company" — never write placeholder brackets
- Under 180 words
- Start directly with "Dear {name}," """


def _clean_resume(text: str, max_chars: int = 5000) -> str:
    """Strip blank lines, whitespace runs, and JSON-breaking characters, then truncate."""
    import re
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    # These characters cause the model to echo resume text into JSON output
    text = text.replace("\\", "/").replace('"', "'").replace("{", "(").replace("}", ")")
    return text[:max_chars].strip()


DEFAULT_SCORING_WEIGHTS = {"experience": 30, "skills": 30, "education": 20, "soft_skills": 20}
GROQ_MODEL = os.getenv("GROQ_MODEL", "").strip() or "openai/gpt-oss-120b"


def normalize_scoring_weights(weights: dict[str, int]) -> dict[str, int]:
    total = sum(weights.values())
    if total == 0:
        return DEFAULT_SCORING_WEIGHTS.copy()
    quotas = {key: divmod(value * 100, total) for key, value in weights.items()}
    normalized = {key: quota[0] for key, quota in quotas.items()}
    remaining = 100 - sum(normalized.values())
    for key in sorted(quotas, key=lambda key: quotas[key][1], reverse=True)[:remaining]:
        normalized[key] += 1
    return normalized


def _parse_analysis(raw: str, weights: dict | None = None) -> dict:
    w = weights or DEFAULT_SCORING_WEIGHTS
    text = _extract_json(raw)
    data = json.loads(text)
    score_fields = ("experience", "skills", "education", "soft_skills")
    for category in score_fields:
        field = f"{category}_score"
        value = data[field]
        if type(value) is not int:
            raise ValueError(f"{field} must be an integer")
        data[field] = max(0, min(w[category], value))
    data["score"] = sum(data[f"{category}_score"] for category in score_fields)
    data["meets_min_experience"] = bool(data.get("meets_min_experience", True))
    data["meets_min_education"]  = bool(data.get("meets_min_education",  True))
    for key in ("strengths", "weaknesses"):
        lst = [str(x) for x in data.get(key, [])][:3]
        while len(lst) < 3:
            lst.append("—")
        data[key] = lst
    if data.get("hire_recommendation") not in ("Strong Hire", "Hire", "Maybe", "No Hire"):
        raise ValueError("missing or invalid hire_recommendation")
    if not data.get("summary"):
        raise ValueError("missing or empty summary")
    data["weights"] = w
    return data


def _call_with_backoff(client: Groq, messages: list, temperature: float,
                       max_tokens: int, json_mode: bool = False) -> str:
    """Call Groq with exponential backoff on rate-limit errors."""
    delays = [5, 15, 30]
    kwargs: dict = dict(
        model=GROQ_MODEL,
        messages=messages,
        temperature=temperature,
        max_tokens=max_tokens,
    )
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    for wait in delays:
        try:
            response = client.chat.completions.create(**kwargs)
            return response.choices[0].message.content.strip()
        except groq_lib.RateLimitError:
            st.warning(f"⏳ Rate limit hit — waiting {wait}s before retrying…")
            time.sleep(wait)
        except groq_lib.APIError as e:
            raise RuntimeError(f"Groq API error: {e}") from e
    raise RuntimeError("Rate limit persists after retries — please wait a minute and try again.")


class AnalysisUnavailable(RuntimeError):
    pass


def analyze_resume(client: Groq, resume_text: str, job_description: str,
                   candidate_name: str, weights: dict | None = None) -> dict:
    w = weights or DEFAULT_SCORING_WEIGHTS
    cleaned = _clean_resume(resume_text)
    messages = [
        {"role": "system", "content": SYSTEM_JSON},
        {"role": "user", "content": ANALYSIS_PROMPT.format(
            job_description=job_description[:3000],
            candidate_name=candidate_name,
            resume_text=cleaned,
            w_exp=w["experience"],
            w_skills=w["skills"],
            w_edu=w["education"],
            w_soft=w["soft_skills"],
        )},
    ]

    correction = (
        'Output ONLY the JSON object — nothing before it, nothing after it. '
        'Start your response with { and end with }.'
    )

    last_error: Exception | None = None
    for attempt in range(3):
        current_messages = messages if attempt == 0 else [
            *messages, {"role": "user", "content": correction},
        ]
        try:
            raw = _call_with_backoff(client, current_messages, temperature=0.1, max_tokens=2048, json_mode=True)
            return _parse_analysis(raw, w)
        except RuntimeError as exc:
            raise AnalysisUnavailable("Analysis unavailable. Check Groq and try again.") from exc
        except (ValueError, TypeError, KeyError) as exc:
            last_error = exc

    raise AnalysisUnavailable("Analysis unavailable. Groq returned an incomplete assessment.") from last_error


# ── Groq — Interview Questions ────────────────────────────────────────────────

SYSTEM_IQ = (
    'You are a technical interviewer API. Output ONLY a JSON object with exactly two keys:\n'
    '  "questions": array of exactly 5 objects, each {"question": "...", "reason": "..."}\n'
    '  "verification_questions": array of exactly 2 objects, each {"question": "...", "claim": "..."}\n'
    "Rules: questions must target the candidate's specific gaps; verification questions must test "
    "specific resume claims. Never write generic questions. "
    "Output ONLY the JSON object — no prose, no markdown. Start with { and end with }."
)

INTERVIEW_QUESTIONS_PROMPT = """\
Candidate: {candidate_name} (fit score {score}/100)

Identified gaps:
{weaknesses}

Resume excerpt:
{resume_text}

Job description:
{job_description}

Generate 5 targeted interview questions addressing the gaps and 2 verification questions testing specific resume claims.\
"""


def generate_interview_questions(client: Groq, candidate: dict, job_description: str) -> dict:
    weaknesses_text = "\n".join(f"- {w}" for w in candidate["weaknesses"])
    cleaned_resume = _clean_resume(candidate["resume_text"], max_chars=3000)
    messages = [
        {"role": "system", "content": SYSTEM_IQ},
        {"role": "user", "content": INTERVIEW_QUESTIONS_PROMPT.format(
            candidate_name=candidate["name"],
            score=candidate["score"],
            weaknesses=weaknesses_text,
            resume_text=cleaned_resume,
            job_description=job_description[:2000],
        )},
    ]

    correction = 'Output ONLY the JSON object. Start with { and end with }.'
    last_exc: Exception | None = None
    for attempt in range(3):
        try:
            last_raw = _call_with_backoff(client, messages, temperature=0.3, max_tokens=2048)
            data = json.loads(_extract_json(last_raw))
            qs = data.get("questions", [])[:5]
            while len(qs) < 5:
                qs.append({"question": "—", "reason": "—"})
            vqs = data.get("verification_questions", [])[:2]
            while len(vqs) < 2:
                vqs.append({"question": "—", "claim": "—"})
            return {"questions": qs, "verification_questions": vqs}
        except Exception as exc:
            last_exc = exc
            if attempt < 2:
                messages = [
                    {"role": "system", "content": SYSTEM_IQ},
                    {"role": "user", "content": correction},
                ]

    return {
        "questions": [{"question": f"Generation failed: {last_exc}", "reason": "—"}] + [{"question": "—", "reason": "—"}] * 4,
        "verification_questions": [{"question": "—", "claim": "—"}] * 2,
    }


# ── Groq — HR Email ───────────────────────────────────────────────────────────

def generate_hr_email(client: Groq, candidate: dict, position: str) -> str:
    score = candidate["score"]
    name = candidate["name"]
    strength = candidate["strengths"][0] if candidate["strengths"] else "relevant background"

    if score >= 75:
        prompt = EMAIL_INVITE_PROMPT.format(name=name, position=position,
                                            score=score, strength=strength)
    elif score >= 50:
        prompt = EMAIL_MORE_INFO_PROMPT.format(name=name, position=position,
                                               strength=strength)
    else:
        prompt = EMAIL_REJECTION_PROMPT.format(name=name, position=position)

    messages = [
        {"role": "system", "content": EMAIL_SYSTEM},
        {"role": "user", "content": prompt},
    ]
    return _call_with_backoff(client, messages, temperature=0.4, max_tokens=400)


def _clipboard_button(text: str) -> None:
    import hashlib
    uid = "cb" + hashlib.md5(text[:64].encode()).hexdigest()[:8]
    text_js = json.dumps(text)   # properly escaped JS string literal
    components.html(
        f"""<!DOCTYPE html><html><body style="margin:0;padding:4px 0 0 0">
<button id="{uid}"
  style="background:#0d6efd;color:white;border:none;padding:7px 18px;
         border-radius:6px;cursor:pointer;font-size:0.85rem;font-family:sans-serif;">
  📋 Copy to Clipboard
</button>
<script>
(function() {{
  var text = {text_js};
  var btn  = document.getElementById('{uid}');
  btn.addEventListener('click', function() {{
    navigator.clipboard.writeText(text)
      .then(function() {{
        btn.textContent = '✅ Copied!';
        setTimeout(function() {{ btn.textContent = '📋 Copy to Clipboard'; }}, 2000);
      }})
      .catch(function() {{ btn.textContent = '❌ Copy failed'; }});
  }});
}})();
</script>
</body></html>""",
        height=46,
    )


# ── Job-requirements helpers ──────────────────────────────────────────────────

def _skills_multiselect(label: str, sel_key: str, custom_key: str,
                        options: list[str]) -> list[str]:
    """Multiselect with a free-text field for custom/unlisted skills."""
    selected = st.multiselect(label, options=options, key=sel_key)
    custom_raw = st.text_input(
        "Add custom skills (comma-separated)",
        key=custom_key,
        placeholder="e.g. Terraform, GraphQL",
        label_visibility="collapsed",
    )
    if custom_raw.strip():
        extra = [s.strip() for s in custom_raw.split(",") if s.strip()]
        return list(dict.fromkeys(selected + extra))
    return selected


def build_job_description(req: dict) -> str:
    """Convert a structured requirements dict to a plain-text JD string for the LLM."""
    parts: list[str] = []
    if req.get("title"):
        parts.append(f"Position: {req['title']}")
    y = req.get("min_years", 0)
    m = req.get("min_months", 0)
    if not req.get("skip_years") and (y > 0 or m > 0):
        if y > 0 and m > 0:
            exp_str = f"{y} year{'s' if y != 1 else ''} {m} month{'s' if m != 1 else ''}"
        elif y > 0:
            exp_str = f"{y}{'+ ' if y >= 15 else ' '}year{'s' if y != 1 else ''}"
        else:
            exp_str = f"{m} month{'s' if m != 1 else ''}"
        parts.append(f"Experience: Minimum {exp_str} required")
    if not req.get("skip_education") and req.get("education") and req["education"] not in ("Any", ""):
        edu_line = f"Education: {req['education']}"
        if req.get("field_of_study"):
            edu_line += f" in {req['field_of_study']}"
        parts.append(edu_line)
    if req.get("employment_types"):
        parts.append(f"Employment type: {', '.join(req['employment_types'])}")
    if req.get("industry"):
        parts.append(f"Industry/domain: {req['industry']}")
    all_req = list(dict.fromkeys(req.get("key_skills", []) + req.get("required_skills", [])))
    if all_req:
        parts.append(f"Required skills: {', '.join(all_req)}")
    if req.get("nice_to_have_skills"):
        parts.append(f"Nice-to-have: {', '.join(req['nice_to_have_skills'])}")
    if not req.get("skip_soft") and req.get("soft_skills"):
        parts.append(f"Soft skills valued: {', '.join(req['soft_skills'])}")
    if req.get("language"):
        lang = req["language"]
        lvl = req.get("language_level", "")
        if lvl and lvl != "Any":
            lang += f" ({lvl})"
        parts.append(f"Language: {lang}")
    w = req.get("weights", {})
    if w:
        parts.append(
            f"Scoring weights — Experience: {w.get('experience', 30)}%, "
            f"Technical Skills: {w.get('skills', 30)}%, "
            f"Education: {w.get('education', 20)}%, "
            f"Soft Skills: {w.get('soft_skills', 20)}%"
        )
    if req.get("notes"):
        parts.append(f"Additional requirements: {req['notes']}")
    return "\n".join(parts)


def render_jd_section() -> tuple[str, dict]:
    """Render the structured Job Requirements form. Returns (jd_text, req_dict)."""
    st.subheader("2 · Job Requirements")

    # ── Quick Mode ────────────────────────────────────────────────────────────
    title = st.text_input(
        "Job title",
        placeholder="e.g. Senior Software Engineer",
        key="jd_title",
    )
    exp_col1, exp_col2 = st.columns(2)
    min_years  = exp_col1.slider("Years of experience",  0, 15, 0, key="jd_min_years")
    min_months = exp_col2.slider("Additional months",    0, 11, 0, key="jd_min_months")
    key_skills = _skills_multiselect(
        "Key skills",
        "jd_key_skills_sel",
        "jd_key_skills_custom",
        SKILLS_QUICK,
    )

    # Summary preview line
    if title or min_years or min_months or key_skills:
        preview: list[str] = []
        if title:
            preview.append(f"**{title}**")
        if min_years or min_months:
            exp_parts = []
            if min_years:
                exp_parts.append(f"{min_years}{'+ ' if min_years >= 15 else ''}yr{'s' if min_years != 1 else ''}")
            if min_months:
                exp_parts.append(f"{min_months}mo")
            preview.append(" ".join(exp_parts) + " exp")
        if key_skills:
            skill_preview = ", ".join(key_skills[:5])
            if len(key_skills) > 5:
                skill_preview += "…"
            preview.append(f"Skills: {skill_preview}")
        st.caption("Screening for: " + " | ".join(preview))

    # ── Advanced Settings toggle ──────────────────────────────────────────────
    adv_open = st.session_state.get("jd_advanced", False)
    if st.button(
        "⚙️ Advanced Settings " + ("▲" if adv_open else "▼"),
        key="jd_adv_btn",
    ):
        st.session_state["jd_advanced"] = not adv_open
        st.rerun()

    # Defaults returned when advanced section is collapsed
    weights = DEFAULT_SCORING_WEIGHTS.copy()
    edu_level      = "Any"
    field_of_study = ""
    emp_types:  list[str] = []
    industry       = ""
    req_skills: list[str] = []
    nth_skills: list[str] = []
    soft_skills_adv: list[str] = []
    language       = ""
    lang_level     = "Any"
    notes          = ""
    skip_years     = False
    skip_education = False
    skip_soft      = False

    if st.session_state.get("jd_advanced"):
        st.markdown("---")

        # Skip All / Reset buttons
        sa_col, reset_col, _ = st.columns([1, 1, 4])
        if sa_col.button("Skip All", key="jd_skip_all_btn"):
            for k in ("jd_skip_years", "jd_skip_edu", "jd_skip_soft"):
                st.session_state[k] = True
            st.rerun()
        if reset_col.button("Reset Skips", key="jd_reset_btn"):
            for k in ("jd_skip_years", "jd_skip_edu", "jd_skip_soft"):
                st.session_state[k] = False
            st.rerun()

        # Experience (skip checkbox only; sliders are in Quick Mode above)
        skip_years = st.checkbox("Skip experience criterion", key="jd_skip_years")

        # Education
        st.markdown("**Education**")
        skip_education = st.checkbox("Skip education criterion", key="jd_skip_edu")
        if not skip_education:
            edu_level = st.selectbox(
                "Education level",
                ["Any", "High School", "Associate's", "Bachelor's", "Master's", "PhD"],
                key="jd_edu_level",
            )
            field_of_study = st.text_input(
                "Field of study (optional)",
                placeholder="e.g. Computer Science",
                key="jd_field_of_study",
            )

        # Employment type
        st.markdown("**Employment Type**")
        ec = st.columns(5)
        for i, etype in enumerate(["Full-time", "Part-time", "Remote", "Hybrid", "On-site"]):
            if ec[i].checkbox(etype, key=f"jd_emp_{etype.lower().replace('-', '_')}"):
                emp_types.append(etype)

        # Industry
        industry = st.text_input(
            "Industry / domain (optional)",
            placeholder="e.g. FinTech, Healthcare",
            key="jd_industry",
        )

        # Skills (advanced)
        st.markdown("**Skills (Advanced)**")
        all_skills_flat = list(dict.fromkeys(
            s for skills in SKILLS_CATEGORIES.values() for s in skills
        ))
        req_skills = _skills_multiselect(
            "Required skills",
            "jd_req_skills_sel",
            "jd_req_skills_custom",
            all_skills_flat,
        )
        nth_skills = _skills_multiselect(
            "Nice-to-have skills",
            "jd_nth_skills_sel",
            "jd_nth_skills_custom",
            all_skills_flat,
        )

        # Soft skills
        skip_soft = st.checkbox("Skip soft skills criterion", key="jd_skip_soft")
        if not skip_soft:
            soft_skills_adv = _skills_multiselect(
                "Soft skills",
                "jd_soft_skills_sel",
                "jd_soft_skills_custom",
                SKILLS_CATEGORIES["Soft Skills"],
            )

        # Language
        st.markdown("**Language Requirements**")
        lc1, lc2 = st.columns(2)
        language = lc1.text_input("Language", placeholder="e.g. English", key="jd_lang")
        lang_level = lc2.selectbox(
            "Proficiency level",
            ["Any", "Basic", "Intermediate", "Upper-Intermediate", "Advanced", "Native"],
            key="jd_lang_level",
        )

        # Scoring weights
        st.markdown("**Scoring Weights** (should sum to 100)")
        wc1, wc2, wc3, wc4 = st.columns(4)
        w_exp  = wc1.slider("Experience %",  0, 100, 30, step=5, key="jd_w_exp")
        w_skl  = wc2.slider("Tech Skills %", 0, 100, 30, step=5, key="jd_w_skills")
        w_edu  = wc3.slider("Education %",   0, 100, 20, step=5, key="jd_w_edu")
        w_soft = wc4.slider("Soft Skills %", 0, 100, 20, step=5, key="jd_w_soft")
        raw_weights = {"experience": w_exp, "skills": w_skl, "education": w_edu, "soft_skills": w_soft}
        total_w = sum(raw_weights.values())
        weights = normalize_scoring_weights(raw_weights)
        if total_w == 0:
            st.warning("All weights are zero. The default scoring weights are active.")
        elif total_w != 100:
            st.warning(f"Weights sum to {total_w}/100 and were normalized to 100.")
        st.info(
            f"Active weights: Experience {weights['experience']}% | Skills {weights['skills']}% "
            f"| Education {weights['education']}% | Soft Skills {weights['soft_skills']}%"
        )

        # Additional notes
        st.markdown("**Additional Notes**")
        notes = st.text_area(
            "Additional requirements or context (optional)",
            placeholder="Any specific requirements not covered above…",
            key="jd_notes",
            height=80,
        )

    req_dict = {
        "title":             title,
        "min_years":         min_years,
        "min_months":        min_months,
        "key_skills":        key_skills,
        "skip_years":        skip_years,
        "skip_education":    skip_education,
        "skip_soft":         skip_soft,
        "education":         edu_level,
        "field_of_study":    field_of_study,
        "employment_types":  emp_types,
        "industry":          industry,
        "required_skills":   req_skills,
        "nice_to_have_skills": nth_skills,
        "soft_skills":       soft_skills_adv,
        "language":          language,
        "language_level":    lang_level,
        "weights":           weights,
        "notes":             notes,
    }
    return build_job_description(req_dict), req_dict


# ── PDF Report ────────────────────────────────────────────────────────────────

def _score_circle_drawing(score: int, size: float = 72) -> Drawing:
    tier = score_tier(score)
    bg = {"green": HexColor("#28a745"), "yellow": HexColor("#ffc107"), "red": HexColor("#dc3545")}[tier]
    fg = HexColor("#212529") if tier == "yellow" else colors.white
    d = Drawing(size, size)
    d.add(Circle(size / 2, size / 2, size / 2 - 1, fillColor=bg, strokeColor=None))
    d.add(RLString(size / 2, size / 2 - 9, str(score),
                   fontSize=22, fillColor=fg, textAnchor="middle",
                   fontName=PDF_FONT_BOLD))
    return d


def generate_pdf_report(candidate: dict, questions: dict | None = None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2.5 * cm)

    styles = getSampleStyleSheet()
    score = candidate["score"]

    def style(name, **kw):
        kw.setdefault("fontName", PDF_FONT)
        return ParagraphStyle(name, parent=styles["Normal"], **kw)

    header_s  = style("Hdr", fontSize=9,  textColor=HexColor("#6c757d"), alignment=TA_CENTER)
    title_s   = style("Ttl", fontSize=22, textColor=HexColor("#212529"), spaceAfter=4,
                      fontName=PDF_FONT_BOLD)
    section_s = style("Sec", fontSize=13, textColor=HexColor("#343a40"),
                      spaceBefore=14, spaceAfter=6, fontName=PDF_FONT_BOLD)
    body_s    = style("Bod", fontSize=10, leading=15, textColor=HexColor("#212529"))
    green_s   = style("Grn", fontSize=10, leading=15, textColor=HexColor("#155724"))
    red_s     = style("Red", fontSize=10, leading=15, textColor=HexColor("#721c24"))
    iq_q_s    = style("IQQ", fontSize=10, leading=15, fontName=PDF_FONT_BOLD, spaceBefore=6)
    iq_r_s    = style("IQR", fontSize=9,  textColor=HexColor("#6c757d"))
    vq_c_s    = style("VQC", fontSize=9,  textColor=HexColor("#856404"),
                      spaceBefore=8, fontName=PDF_FONT)
    footer_s  = style("Ftr", fontSize=8,  textColor=HexColor("#6c757d"), alignment=TA_CENTER)

    rec = candidate["hire_recommendation"]
    rec_hex = {"Strong Hire": "#28a745", "Hire": "#20c997",
               "Maybe": "#ffc107", "No Hire": "#dc3545"}.get(rec, "#6c757d")

    story = []

    story.append(Paragraph("AI Resume Screening System", header_s))
    story.append(HRFlowable(width="100%", thickness=1, color=HexColor("#dee2e6"), spaceAfter=10))

    rec_s = style("Rec", fontSize=11, fontName=PDF_FONT_BOLD,
                  textColor=HexColor(rec_hex), spaceBefore=0)
    name_cell = [
        Paragraph(candidate["name"], title_s),
        Spacer(1, 8),
        Paragraph(rec, rec_s),
        Spacer(1, 4),
        Paragraph(f'<font color="#6c757d">File: {candidate["filename"]}</font>',
                  style("Cap", fontSize=8, textColor=HexColor("#6c757d"))),
    ]
    circle_cell = [
        _score_circle_drawing(score),
        Paragraph("<font color='#6c757d' size='8'>Score / 100</font>",
                  style("SL", fontSize=8, textColor=HexColor("#6c757d"), alignment=TA_CENTER)),
    ]
    header_tbl = Table([[name_cell, circle_cell]], colWidths=["78%", "22%"])
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
    ]))
    story.append(header_tbl)
    story.append(HRFlowable(width="100%", thickness=0.5,
                            color=HexColor("#dee2e6"), spaceBefore=12, spaceAfter=10))

    story.append(Paragraph("Summary", section_s))
    story.append(Paragraph(candidate["summary"], body_s))

    story.append(Paragraph("Strengths", section_s))
    for s in candidate["strengths"]:
        story.append(Paragraph(f"✓  {s}", green_s))

    story.append(Paragraph("Weaknesses / Gaps", section_s))
    for w in candidate["weaknesses"]:
        story.append(Paragraph(f"✗  {w}", red_s))

    if questions:
        story.append(HRFlowable(width="100%", thickness=0.5,
                                color=HexColor("#dee2e6"), spaceBefore=14, spaceAfter=4))
        story.append(Paragraph("Interview Questions", section_s))
        for i, q in enumerate(questions["questions"], 1):
            story.append(Paragraph(f"{i}.  {q['question']}", iq_q_s))
            story.append(Paragraph(f"Why: {q['reason']}", iq_r_s))
            story.append(Spacer(1, 4))

        story.append(Paragraph("Verification Questions", section_s))
        for i, vq in enumerate(questions["verification_questions"], 1):
            story.append(Paragraph(f"Verifying: \"{vq['claim']}\"", vq_c_s))
            story.append(Paragraph(f"{i}.  {vq['question']}", iq_q_s))
            story.append(Spacer(1, 4))

    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor("#dee2e6")))
    story.append(Paragraph(
        f"Generated on {date.today().strftime('%B %d, %Y')}  ·  Confidential — HR Use Only",
        footer_s,
    ))

    doc.build(story)
    return buf.getvalue()


# ── Leaderboard renderer ──────────────────────────────────────────────────────

def render_leaderboard(label: str, candidates: list[dict], lb_id: str) -> None:
    st.markdown(f"<div class='lb-section'>", unsafe_allow_html=True)
    st.markdown(f"### {label}")

    # Column headers
    hc = st.columns([0.5, 3, 1, 2, 3, 1.5])
    for col, txt in zip(hc, ["Rank", "Candidate", "Score", "Recommendation", "Top Strength", ""]):
        col.markdown(f"<span class='lb-header'>{txt}</span>", unsafe_allow_html=True)

    st.markdown("<hr style='margin:4px 0 8px 0; border-color:#dee2e6'>", unsafe_allow_html=True)

    for rank, cand in enumerate(candidates, 1):
        is_top = rank == 1
        rc = st.columns([0.5, 3, 1, 2, 3, 1.5])

        with rc[0]:
            if is_top:
                st.markdown("🏅 **#1**")
            else:
                st.markdown(f"**#{rank}**")

        with rc[1]:
            badge = '<span class="best-match-badge">★ Best Match</span>' if is_top else ""
            flag = " ⚠️" if 50 <= cand["score"] <= 65 else ""
            dnmr = not cand.get("meets_min_experience", True) or not cand.get("meets_min_education", True)
            dnmr_badge = (
                ' <span style="background:#dc3545;color:white;padding:1px 6px;'
                'border-radius:10px;font-size:0.7rem;font-weight:700">✗ Req</span>'
                if dnmr else ""
            )
            st.markdown(
                f'{"<b>" if is_top else ""}{cand["name"]}{flag}{"</b>" if is_top else ""}{badge}{dnmr_badge}',
                unsafe_allow_html=True,
            )

        with rc[2]:
            st.markdown(score_pill_html(cand["score"]), unsafe_allow_html=True)

        with rc[3]:
            rec = cand["hire_recommendation"]
            rec_color = recommendation_color(rec)
            txt_c = "#212529" if rec == "Maybe" else "white"
            st.markdown(
                f'<span style="background:{rec_color};color:{txt_c};padding:2px 8px;'
                f'border-radius:10px;font-size:0.85rem;font-weight:600">{rec}</span>',
                unsafe_allow_html=True,
            )

        with rc[4]:
            st.caption(cand["strengths"][0])

        with rc[5]:
            if st.button("View Details", key=f"view_{lb_id}_{rank}",
                         use_container_width=True):
                st.session_state["view"] = "detail"
                st.session_state["detail_candidate"] = cand["name"]
                st.rerun()

        if rank < len(candidates):
            st.markdown("<hr style='margin:2px 0; border-color:#f0f0f0'>",
                        unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ── Candidate comparison ──────────────────────────────────────────────────────

def render_comparison_section(results: list[dict]) -> None:
    st.markdown("---")
    st.header("⚖️ Compare Candidates")

    names = [r["name"] for r in results]
    selected = st.multiselect(
        "Select 2 or 3 candidates to compare",
        options=names,
        max_selections=3,
        placeholder="Choose candidates…",
        key="compare_select",
    )

    if len(selected) < 2:
        st.caption("Select at least 2 candidates to see the comparison.")
        return

    candidates = [next(r for r in results if r["name"] == n) for n in selected]
    n = len(candidates)
    col_w = [1.6] + [2.4] * n

    # Use the first candidate's weights as the scoring reference
    w = candidates[0].get(
        "weights", {"experience": 30, "skills": 30, "education": 20, "soft_skills": 20}
    )

    # ── HTML helpers ──────────────────────────────────────────────────────────
    def _bar_cell(value: int, max_val: int, is_best: bool) -> str:
        pct = min(100, round(value / max_val * 100)) if max_val > 0 else 0
        bar_color = "#28a745" if pct >= 75 else ("#ffc107" if pct >= 50 else "#dc3545")
        bg = "#d4f5de" if is_best else "#f8f9fa"
        border = "2px solid #28a745" if is_best else "1px solid #e9ecef"
        return (
            f'<div style="background:{bg};padding:10px 12px;border-radius:8px;border:{border}">'
            f'<div style="font-weight:700;font-size:1.05rem;color:#212529;margin-bottom:6px">'
            f'{value}<span style="font-weight:400;font-size:0.78rem;color:#6c757d"> / {max_val}</span>'
            f'{"  <span style=\'font-size:0.7rem;color:#28a745\'>▲ best</span>" if is_best else ""}'
            f'</div>'
            f'<div style="background:#dee2e6;border-radius:4px;height:7px">'
            f'<div style="background:{bar_color};height:7px;border-radius:4px;width:{pct}%"></div>'
            f'</div></div>'
        )

    def _req_cell(meets: bool) -> str:
        if meets:
            return '<span style="color:#28a745;font-weight:700;font-size:0.9rem">✓ Met</span>'
        return '<span style="color:#dc3545;font-weight:700;font-size:0.9rem">✗ Not met</span>'

    def _rec_cell(rec: str, is_best: bool) -> str:
        color = recommendation_color(rec)
        txt_c = "#212529" if rec == "Maybe" else "white"
        bg = "#d4f5de" if is_best else "#f8f9fa"
        border = "2px solid #28a745" if is_best else "1px solid #e9ecef"
        return (
            f'<div style="background:{bg};padding:8px 12px;border-radius:8px;border:{border}">'
            f'<span style="background:{color};color:{txt_c};padding:3px 10px;'
            f'border-radius:10px;font-size:0.85rem;font-weight:600">{rec}</span>'
            f'</div>'
        )

    # ── Shared row renderer ───────────────────────────────────────────────────
    def _row(label: str, html_cells: list[str], divider: bool = False) -> None:
        cols = st.columns(col_w)
        cols[0].markdown(
            f'<div style="padding:8px 0;color:#495057;font-weight:600;font-size:0.9rem">{label}</div>',
            unsafe_allow_html=True,
        )
        for i, cell in enumerate(html_cells):
            cols[i + 1].markdown(cell, unsafe_allow_html=True)
        if divider:
            st.markdown("<hr style='margin:6px 0;border-color:#e9ecef'>", unsafe_allow_html=True)

    # ── Header row ────────────────────────────────────────────────────────────
    winner = max(candidates, key=lambda c: c["score"])
    hcols = st.columns(col_w)
    hcols[0].markdown("")
    for i, cand in enumerate(candidates):
        is_win = cand["name"] == winner["name"]
        hcols[i + 1].markdown(
            f'<div style="font-weight:700;font-size:1rem;padding:6px 0">'
            f'{cand["name"]}{"  🏆" if is_win else ""}</div>',
            unsafe_allow_html=True,
        )
    st.markdown("<hr style='margin:4px 0 8px;border-color:#dee2e6;border-width:2px'>",
                unsafe_allow_html=True)

    # ── Total score ───────────────────────────────────────────────────────────
    scores = [c["score"] for c in candidates]
    best_s = scores.index(max(scores))
    _row("🎯 Total Score",
         [_bar_cell(c["score"], 100, i == best_s) for i, c in enumerate(candidates)],
         divider=True)

    # ── Experience ────────────────────────────────────────────────────────────
    exp_vals = [c.get("experience_score", 0) for c in candidates]
    best_e = exp_vals.index(max(exp_vals))
    _row(f"💼 Experience  (max {w['experience']})",
         [_bar_cell(v, w["experience"], i == best_e) for i, v in enumerate(exp_vals)])
    _row("Min. experience req.",
         [_req_cell(c.get("meets_min_experience", True)) for c in candidates],
         divider=True)

    # ── Technical skills ──────────────────────────────────────────────────────
    skl_vals = [c.get("skills_score", 0) for c in candidates]
    best_sk = skl_vals.index(max(skl_vals))
    _row(f"⚙️ Tech Skills  (max {w['skills']})",
         [_bar_cell(v, w["skills"], i == best_sk) for i, v in enumerate(skl_vals)],
         divider=True)

    # ── Education ─────────────────────────────────────────────────────────────
    edu_vals = [c.get("education_score", 0) for c in candidates]
    best_ed = edu_vals.index(max(edu_vals))
    _row(f"🎓 Education  (max {w['education']})",
         [_bar_cell(v, w["education"], i == best_ed) for i, v in enumerate(edu_vals)])
    _row("Education req.",
         [_req_cell(c.get("meets_min_education", True)) for c in candidates],
         divider=True)

    # ── Soft skills ───────────────────────────────────────────────────────────
    soft_vals = [c.get("soft_skills_score", 0) for c in candidates]
    best_so = soft_vals.index(max(soft_vals))
    _row(f"🤝 Soft Skills  (max {w['soft_skills']})",
         [_bar_cell(v, w["soft_skills"], i == best_so) for i, v in enumerate(soft_vals)],
         divider=True)

    # ── Hire recommendation ───────────────────────────────────────────────────
    rec_order = {"Strong Hire": 4, "Hire": 3, "Maybe": 2, "No Hire": 1}
    rec_vals = [rec_order.get(c["hire_recommendation"], 0) for c in candidates]
    best_r = rec_vals.index(max(rec_vals))
    _row("📋 Recommendation",
         [_rec_cell(c["hire_recommendation"], i == best_r) for i, c in enumerate(candidates)],
         divider=True)

    # ── Overall winner ────────────────────────────────────────────────────────
    winner_idx = [c["name"] for c in candidates].index(winner["name"])
    _row(
        "🏆 Overall Winner",
        [
            '<div style="background:#d4f5de;padding:10px 12px;border-radius:8px;'
            'border:2px solid #28a745;font-weight:700;color:#155724">✅ Winner</div>'
            if i == winner_idx else
            '<div style="background:#f8f9fa;padding:10px 12px;border-radius:8px;'
            'border:1px solid #e9ecef;color:#adb5bd">—</div>'
            for i in range(n)
        ],
    )


# ── Filter panel ─────────────────────────────────────────────────────────────

def render_filter_panel(results: list[dict]) -> list[dict]:
    """Show filter/sort controls and return the filtered, sorted candidate list."""
    with st.expander("🔧 Filter & Sort Results", expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        min_score  = fc1.slider("Min total score", 0, 100, 0, key="fp_min_score")
        filter_exp = fc2.selectbox(
            "Min experience met", ["All", "Yes only", "No only"], key="fp_exp"
        )
        filter_edu = fc3.selectbox(
            "Education req met", ["All", "Yes only", "No only"], key="fp_edu"
        )
        sort_by = fc4.selectbox(
            "Sort by",
            ["Total Score", "Skills Score", "Experience Score", "Education Score"],
            key="fp_sort",
        )

    filtered = results
    if min_score > 0:
        filtered = [r for r in filtered if r["score"] >= min_score]
    if filter_exp == "Yes only":
        filtered = [r for r in filtered if r.get("meets_min_experience", True)]
    elif filter_exp == "No only":
        filtered = [r for r in filtered if not r.get("meets_min_experience", True)]
    if filter_edu == "Yes only":
        filtered = [r for r in filtered if r.get("meets_min_education", True)]
    elif filter_edu == "No only":
        filtered = [r for r in filtered if not r.get("meets_min_education", True)]

    sort_key = {
        "Total Score":      "score",
        "Skills Score":     "skills_score",
        "Experience Score": "experience_score",
        "Education Score":  "education_score",
    }[sort_by]
    filtered = sorted(filtered, key=lambda r: r.get(sort_key, 0), reverse=True)

    if len(filtered) < len(results):
        st.info(f"Showing {len(filtered)} of {len(results)} candidates after filtering.")
    return filtered


# ── Candidate detail page ─────────────────────────────────────────────────────

def render_detail_page(candidate: dict, saved_jd: str) -> None:
    # Back button
    if st.button("← Back to Rankings", type="secondary"):
        st.session_state["view"] = "rankings"
        st.rerun()

    st.divider()
    score = candidate["score"]

    if 50 <= score <= 65:
        st.markdown(
            '<div class="borderline-banner">⚠️ <b>Borderline score — manual review recommended.</b></div>',
            unsafe_allow_html=True,
        )

    # "Does Not Meet Requirements" warning
    dnmr_flags = []
    if not candidate.get("meets_min_experience", True):
        dnmr_flags.append("minimum experience")
    if not candidate.get("meets_min_education", True):
        dnmr_flags.append("education requirement")
    if dnmr_flags:
        st.error(f"⛔ Does Not Meet Requirements: {', '.join(dnmr_flags)}")

    # Header row
    d1, d2 = st.columns([4, 1])
    with d1:
        st.subheader(candidate["name"])
        st.caption(f"File: {candidate['filename']}")
        rec = candidate["hire_recommendation"]
        rec_color = recommendation_color(rec)
        st.markdown(
            f'**Recommendation:** <span class="rec-badge" style="background:{rec_color};'
            f'color:{"white" if rec != "Maybe" else "#212529"}">{rec}</span>',
            unsafe_allow_html=True,
        )
        # Score breakdown
        w = candidate.get("weights", {"experience": 30, "skills": 30, "education": 20, "soft_skills": 20})
        breakdown = (
            f"Education: **{candidate.get('education_score', '?')}/{w['education']}** | "
            f"Experience: **{candidate.get('experience_score', '?')}/{w['experience']}** | "
            f"Skills: **{candidate.get('skills_score', '?')}/{w['skills']}** | "
            f"Soft Skills: **{candidate.get('soft_skills_score', '?')}/{w['soft_skills']}** | "
            f"Total: **{score}/100**"
        )
        st.markdown(breakdown)
        st.markdown(f"**Summary:** {candidate['summary']}")
    with d2:
        st.markdown(score_circle_html(score), unsafe_allow_html=True)
        st.markdown("<p style='text-align:center;color:#6c757d;margin-top:6px;'>out of 100</p>",
                    unsafe_allow_html=True)

    st.divider()

    sw1, sw2 = st.columns(2)
    with sw1:
        st.markdown("#### ✅ Strengths")
        for s in candidate["strengths"]:
            st.markdown(f'<p class="strength-item">✔ {s}</p>', unsafe_allow_html=True)
    with sw2:
        st.markdown("#### ❌ Weaknesses / Gaps")
        for w in candidate["weaknesses"]:
            st.markdown(f'<p class="weakness-item">✖ {w}</p>', unsafe_allow_html=True)

    # ── Interview Questions ───────────────────────────────────────────────────
    st.divider()
    st.markdown("#### 🎯 Interview Questions")

    if "questions" not in st.session_state:
        st.session_state["questions"] = {}

    cname = candidate["name"]
    existing_qs = st.session_state["questions"].get(cname)

    gc, rc_ = st.columns([3, 1])
    with gc:
        gen_btn = st.button(
            "Generate Interview Questions" if not existing_qs else "Questions generated ✓",
            key=f"det_gen_{cname}", disabled=bool(existing_qs),
        )
    with rc_:
        regen_btn = st.button("🔄 Regenerate", key=f"det_regen_{cname}",
                              disabled=not bool(existing_qs))

    if gen_btn or regen_btn:
        with st.spinner("Generating questions…"):
            client = get_groq_client()
            qs = generate_interview_questions(client, candidate, saved_jd)
            st.session_state["questions"][cname] = qs
        existing_qs = qs

    if existing_qs:
        with st.expander("📋 Interview Questions", expanded=True):
            st.markdown("**5 Targeted Questions**")
            for i, q in enumerate(existing_qs["questions"], 1):
                st.markdown(
                    f'<div class="iq-card">'
                    f'<div class="iq-question"><span class="iq-num">{i}.</span>{q["question"]}</div>'
                    f'<div class="iq-reason">💡 {q["reason"]}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            st.markdown("---")
            st.markdown("**🔍 Verification Questions**")
            for i, vq in enumerate(existing_qs["verification_questions"], 1):
                st.markdown(
                    f'<div class="vq-card">'
                    f'<div class="vq-claim">🔍 Verifying: "{vq["claim"]}"</div>'
                    f'<div class="vq-question">{i}. {vq["question"]}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    # Resume text
    with st.expander("📄 View extracted resume text"):
        st.text_area("Resume text", value=candidate["resume_text"], height=300,
                     disabled=True, label_visibility="collapsed")

    # ── HR Email Generator ────────────────────────────────────────────────────
    st.divider()
    st.markdown("#### ✉️ Generate HR Email")

    score = candidate["score"]
    if score >= 75:
        email_type_label, email_type_bg, email_type_fg = "Interview Invitation", "#28a745", "white"
    elif score >= 50:
        email_type_label, email_type_bg, email_type_fg = "Additional Info Request", "#ffc107", "#212529"
    else:
        email_type_label, email_type_bg, email_type_fg = "Rejection Email", "#dc3545", "white"

    st.markdown(
        f'Will generate: <span style="background:{email_type_bg};color:{email_type_fg};'
        f'padding:2px 10px;border-radius:10px;font-size:0.85rem;font-weight:600">'
        f'{email_type_label}</span>',
        unsafe_allow_html=True,
    )

    email_pos_key = f"email_pos_{cname}"
    if email_pos_key not in st.session_state:
        # Pre-fill with the job title used during screening; user can edit freely
        st.session_state[email_pos_key] = st.session_state.get("req_dict", {}).get("title", "")

    position_input = st.text_input(
        "Position title",
        key=email_pos_key,
    )

    email_body_key = f"email_body_{cname}"
    if email_body_key not in st.session_state:
        st.session_state[email_body_key] = ""

    gen_email_col, _ = st.columns([1, 3])
    with gen_email_col:
        gen_email_btn = st.button("✉️ Generate Email", key=f"gen_email_{cname}",
                                  use_container_width=True)

    if gen_email_btn:
        if not position_input.strip():
            st.warning("Enter the position title to continue.")
        else:
            with st.spinner("Generating email…"):
                client = get_groq_client()
                generated = generate_hr_email(client, candidate, position_input.strip())
                st.session_state[email_body_key] = generated

    if st.session_state[email_body_key]:
        edited_email = st.text_area(
            "Generated email — edit before sending",
            value=st.session_state[email_body_key],
            height=260,
            key=f"email_ta_{cname}",
        )
        _clipboard_button(edited_email)

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.divider()
    st.markdown("#### ⬇️ Export Report")

    qs_data = st.session_state.get("questions", {}).get(cname)
    txt_lines = [
        f"RESUME SCREENING REPORT — {candidate['name']}",
        "=" * 55,
        f"Score:          {score}/100",
        f"Recommendation: {candidate['hire_recommendation']}",
        "",
        "SUMMARY:",
        candidate["summary"],
        "",
        "STRENGTHS:",
        *[f"  ✓  {s}" for s in candidate["strengths"]],
        "",
        "WEAKNESSES:",
        *[f"  ✗  {w}" for w in candidate["weaknesses"]],
    ]
    if qs_data:
        txt_lines += ["", "INTERVIEW QUESTIONS:", "─" * 40]
        for i, q in enumerate(qs_data["questions"], 1):
            txt_lines += [f"{i}. {q['question']}", f"   Why: {q['reason']}", ""]
        txt_lines += ["VERIFICATION QUESTIONS:", "─" * 40]
        for i, vq in enumerate(qs_data["verification_questions"], 1):
            txt_lines += [f"Verifying: \"{vq['claim']}\"", f"{i}. {vq['question']}", ""]
    txt_lines += ["", "─" * 55,
                  f"Generated: {date.today().strftime('%B %d, %Y')}",
                  "Confidential — HR Use Only"]

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📄 Download TXT",
            data="\n".join(txt_lines),
            file_name=f"report_{cname.replace(' ', '_')}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with dl2:
        try:
            pdf_bytes = generate_pdf_report(candidate, qs_data)
            st.download_button(
                "📑 Export PDF Report",
                data=pdf_bytes,
                file_name=f"report_{cname.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as pdf_err:
            st.button("📑 Export PDF Report", disabled=True, use_container_width=True)
            st.caption(f"PDF error: {pdf_err}")

    st.divider()
    if st.button("← Back to Rankings", key="back_bottom", type="secondary"):
        st.session_state["view"] = "rankings"
        st.rerun()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("📄 Resume Screener")
    st.markdown(f"*Powered by Groq · {GROQ_MODEL}*")
    st.divider()
    st.markdown("### Score Legend")
    st.markdown("""
| Color | Range | Meaning |
|-------|-------|---------|
| 🟢 Green | 75 – 100 | Strong candidate |
| 🟡 Yellow | 50 – 74 | Consider |
| 🔴 Red | 0 – 49 | Weak fit |
| ⚠️ Flag | 50 – 65 | Borderline — human review |
""")
    st.divider()
    st.markdown("### How to use")
    st.markdown("""
1. Upload up to 100 resumes (PDF/DOCX/TXT/RTF) **or** a CSV
2. Paste the job description
3. Click **Analyze Resumes**
4. Browse Top 5 / 10 / 15 leaderboards
5. Click **View Details** on any candidate
6. Generate interview questions & export PDF
""")

    if st.session_state.get("results") or st.session_state.get("analysis_errors"):
        st.divider()
        n = len(st.session_state.get("results", []))
        st.markdown(f"**{n} candidate(s) analyzed**")
        if st.button("🔄 New Analysis", use_container_width=True):
            for key in ("results", "questions", "view", "detail_candidate",
                        "job_description", "stored_files", "analysis_errors"):
                st.session_state.pop(key, None)
            st.rerun()


# ── Session state init ────────────────────────────────────────────────────────

if "view" not in st.session_state:
    st.session_state["view"] = "rankings"
if "questions" not in st.session_state:
    st.session_state["questions"] = {}
if "stored_files" not in st.session_state:
    st.session_state["stored_files"] = {}   # filename -> StoredFile


# ── Main UI (always visible) ──────────────────────────────────────────────────

st.title("🔍 AI-Powered Resume Screener")

# Hide upload UI when viewing candidate detail to keep layout clean
if st.session_state["view"] == "rankings" or not st.session_state.get("results"):

    col_upload, col_jd = st.columns([1, 1], gap="large")

    with col_upload:
        st.subheader("1 · Upload Resumes")
        st.caption("Supports up to 100 files per batch — PDF, DOCX, TXT, RTF")
        tab_files, tab_csv = st.tabs(["📁 Individual Files", "📊 CSV Dataset"])

        uploaded_files = []
        csv_candidates: list[dict] = []

        with tab_files:
            raw_files = st.file_uploader(
                "Drop files here — upload in multiple batches to accumulate",
                type=["pdf", "docx", "txt", "rtf"],
                accept_multiple_files=True,
                label_visibility="collapsed",
                key="file_uploader",
            )

            # Merge new uploads into the persistent store
            if raw_files:
                added = 0
                for f in raw_files:
                    if f.name not in st.session_state["stored_files"]:
                        st.session_state["stored_files"][f.name] = StoredFile(f.name, f.read())
                        added += 1
                if added:
                    st.toast(f"✅ {added} new file(s) added")

            store = st.session_state["stored_files"]

            if store:
                # Cap at 100
                if len(store) > 100:
                    trimmed = dict(list(store.items())[:100])
                    st.session_state["stored_files"] = trimmed
                    store = trimmed
                    st.warning("⚠️ Capped at 100 files — extras removed.")

                st.success(f"{len(store)} file(s) accumulated across all batches")

                # Show file list with individual remove buttons
                shown = list(store.keys())[:10]
                for fname in shown:
                    fc1, fc2 = st.columns([5, 1])
                    fc1.caption(f"📎 {fname}")
                    if fc2.button("✕", key=f"rm_{fname}", help="Remove this file"):
                        del st.session_state["stored_files"][fname]
                        st.rerun()
                if len(store) > 10:
                    st.caption(f"… and {len(store) - 10} more file(s)")

                if st.button("🗑️ Clear all files", use_container_width=True):
                    st.session_state["stored_files"] = {}
                    st.rerun()

            uploaded_files = list(store.values())

        with tab_csv:
            csv_file = st.file_uploader(
                "Upload a CSV where each row is one candidate",
                type=["csv"],
                label_visibility="collapsed",
                key="csv_uploader",
            )
            if csv_file:
                try:
                    csv_df = pd.read_csv(csv_file)
                    st.success(f"CSV loaded — {len(csv_df)} rows, {len(csv_df.columns)} columns")
                    st.dataframe(csv_df.head(3), use_container_width=True)
                    all_cols = list(csv_df.columns)
                    col_a, col_b = st.columns(2)
                    with col_a:
                        name_col = st.selectbox("Candidate name column", all_cols, key="name_col")
                    with col_b:
                        text_col = st.selectbox("Resume text column", all_cols,
                                                index=min(1, len(all_cols) - 1), key="text_col")
                    if name_col and text_col:
                        csv_candidates = [
                            {"name": str(row[name_col]).strip().title(),
                             "resume_text": str(row[text_col]).strip()}
                            for _, row in csv_df.iterrows()
                            if str(row[text_col]).strip()
                        ]
                        if len(csv_candidates) > 100:
                            st.warning(f"⚠️ {len(csv_candidates)} rows — only the first 100 will be analyzed.")
                            csv_candidates = csv_candidates[:100]
                        st.info(f"✅ {len(csv_candidates)} candidate(s) ready from CSV")
                except Exception as e:
                    st.error(f"Could not read CSV: {e}")

    with col_jd:
        jd_text, req_dict = render_jd_section()

    st.divider()

    has_input = bool((uploaded_files or csv_candidates) and jd_text.strip())
    analyze_btn = st.button("🚀  Analyze Resumes", type="primary",
                            disabled=not has_input, use_container_width=True)
    if not has_input:
        st.caption("Upload files or a CSV dataset, then enter a job description to continue.")

    # ── Run Analysis ──────────────────────────────────────────────────────────

    if analyze_btn:
        client = get_groq_client()
        results: list[dict] = []
        errors: list[str] = []
        weights = req_dict.get("weights", DEFAULT_SCORING_WEIGHTS)

        candidates: list[dict] = []
        for f in uploaded_files:
            ext = f.name.rsplit(".", 1)[-1].lower()
            name = f.name[:-(len(ext) + 1)].replace("_", " ").replace("-", " ").title()
            candidates.append({"name": name, "filename": f.name,
                               "file_obj": f, "resume_text": None})
        for c in csv_candidates:
            candidates.append({"name": c["name"], "filename": "CSV",
                               "file_obj": None, "resume_text": c["resume_text"]})

        total = len(candidates)
        progress = st.progress(0.0, text=f"Starting analysis of {total} candidate(s)…")
        status = st.empty()

        for idx, cand in enumerate(candidates):
            name = cand["name"]
            pct = (idx + 1) / total
            progress.progress(pct, text=f"Analyzing {idx + 1}/{total} — {name}")
            status.caption(f"⏳ {total - idx - 1} remaining after this…")

            try:
                resume_text = (extract_text(cand["file_obj"])
                               if cand["file_obj"] else cand["resume_text"])
                if not resume_text:
                    errors.append(f"{cand['filename']}: no extractable text.")
                    continue
                analysis = analyze_resume(client, resume_text, jd_text, name, weights=weights)
                results.append({
                    "name":                 name,
                    "filename":             cand["filename"],
                    "score":                analysis["score"],
                    "experience_score":     analysis.get("experience_score", 0),
                    "skills_score":         analysis.get("skills_score", 0),
                    "education_score":      analysis.get("education_score", 0),
                    "soft_skills_score":    analysis.get("soft_skills_score", 0),
                    "meets_min_experience": analysis.get("meets_min_experience", True),
                    "meets_min_education":  analysis.get("meets_min_education", True),
                    "weights":              analysis.get("weights", weights),
                    "strengths":            analysis["strengths"],
                    "weaknesses":           analysis["weaknesses"],
                    "hire_recommendation":  analysis["hire_recommendation"],
                    "summary":              analysis["summary"],
                    "resume_text":          resume_text,
                })
            except AnalysisUnavailable as exc:
                errors.append(f"{name}: {exc}")
            except Exception:
                errors.append(f"{name}: Could not process the resume file.")

            # Rate-limit guard between calls (skip after the last one)
            if idx < total - 1:
                time.sleep(4)

        progress.progress(1.0, text=f"✅ Done! Analyzed {len(results)}/{total} candidates.")
        status.empty()

        st.session_state["analysis_errors"]  = errors
        st.session_state["results"]          = results
        st.session_state["job_description"]  = jd_text
        st.session_state["req_dict"]         = req_dict
        st.session_state["questions"]        = {}
        st.session_state["view"]             = "rankings"
        st.rerun()


if st.session_state["view"] == "rankings":
    for message in st.session_state.get("analysis_errors", []):
        st.warning(message)


# ── Rankings page ─────────────────────────────────────────────────────────────

if st.session_state.get("results") and st.session_state["view"] == "rankings":
    all_results: list[dict] = sorted(st.session_state["results"],
                                     key=lambda r: r["score"], reverse=True)
    saved_jd = st.session_state.get("job_description", "")
    total = len(all_results)

    st.divider()
    st.header("📊 Bulk Analysis Results")

    # Summary metrics (always from full unfiltered set)
    scores = [r["score"] for r in all_results]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Analyzed", total)
    m2.metric("🟢 Strong (75+)", sum(s >= 75 for s in scores))
    m3.metric("🟡 Consider (50–74)", sum(50 <= s < 75 for s in scores))
    m4.metric("🔴 Weak (<50)", sum(s < 50 for s in scores))

    borderline = [r for r in all_results if 50 <= r["score"] <= 65]
    if borderline:
        names = ", ".join(r["name"] for r in borderline)
        st.markdown(
            f'<div class="borderline-banner">⚠️ <b>Borderline — human review recommended:</b> {names}</div>',
            unsafe_allow_html=True,
        )

    dnmr = [r for r in all_results
            if not r.get("meets_min_experience", True) or not r.get("meets_min_education", True)]
    if dnmr:
        names = ", ".join(r["name"] for r in dnmr)
        st.error(f"⛔ **Does Not Meet Requirements:** {names}")

    st.divider()

    # Filter panel — returns filtered+sorted subset
    results = render_filter_panel(all_results)
    ftotal = len(results)

    # Leaderboards
    top5 = results[:min(5, ftotal)]
    render_leaderboard("🏆 Top 5 Candidates", top5, "top5")

    if ftotal > 5:
        top10 = results[:min(10, ftotal)]
        render_leaderboard("🥈 Top 10 Candidates", top10, "top10")

    if ftotal > 10:
        top15 = results[:min(15, ftotal)]
        render_leaderboard("📋 Top 15 Candidates", top15, "top15")

    render_comparison_section(results)


# ── Detail page ───────────────────────────────────────────────────────────────

elif st.session_state.get("results") and st.session_state["view"] == "detail":
    results = st.session_state["results"]
    saved_jd = st.session_state.get("job_description", "")
    detail_name = st.session_state.get("detail_candidate")
    candidate = next((r for r in results if r["name"] == detail_name), None)

    if candidate:
        st.header(f"🔎 Candidate Profile — {candidate['name']}")
        render_detail_page(candidate, saved_jd)
    else:
        st.error("Candidate not found. Returning to rankings.")
        st.session_state["view"] = "rankings"
        st.rerun()
