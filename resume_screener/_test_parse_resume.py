"""
Drop a resume PDF path as the first argument to test parsing end-to-end.
Usage: python _test_parse_resume.py "path/to/resume.pdf" "Job title here"
"""
import sys, os, io, json
from dotenv import load_dotenv
import pdfplumber
from groq import Groq

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
job_desc  = sys.argv[2] if len(sys.argv) > 2 else "Software Engineer"

if not pdf_path:
    print("Usage: python _test_parse_resume.py resume.pdf 'Job description'")
    sys.exit(1)

# Extract text
with pdfplumber.open(pdf_path) as pdf:
    pages = [p.extract_text() or "" for p in pdf.pages]
resume_text = "\n".join(pages).strip()
print(f"--- Extracted {len(resume_text)} chars ---")
print(resume_text[:400])
print("...")

# Call Groq
SYSTEM_JSON = ("You are an HR analysis API. "
               "You ONLY output a single valid JSON object — no prose, no markdown, "
               "no explanations, no code fences. Never repeat the resume text.")
PROMPT = f"""You are a senior HR recruiter evaluating a resume against a job description.

=== JOB DESCRIPTION ===
{job_desc}

=== RESUME ===
{resume_text[:5000]}

Return ONLY a valid JSON object:
{{
  "score": <integer 0-100>,
  "strengths": ["...", "...", "..."],
  "weaknesses": ["...", "...", "..."],
  "hire_recommendation": "<Strong Hire | Hire | Maybe | No Hire>",
  "summary": "<2-3 sentence assessment>"
}}"""

r = client.chat.completions.create(
    model="llama-3.3-70b-versatile",
    messages=[
        {"role": "system", "content": SYSTEM_JSON},
        {"role": "user",   "content": PROMPT},
    ],
    max_tokens=1024,
    temperature=0.1,
)
raw = r.choices[0].message.content
print(f"\n--- Model output ({len(raw)} chars, finish={r.choices[0].finish_reason}) ---")
print(raw)

try:
    data = json.loads(raw)
    print("\n✅ JSON parsed OK:", list(data.keys()))
except json.JSONDecodeError as e:
    print(f"\n❌ JSON parse failed: {e}")
