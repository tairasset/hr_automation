import os, json
from dotenv import load_dotenv
from groq import Groq
import groq as groq_lib

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

SYSTEM_JSON = ("You are an HR analysis API. "
               "You ONLY output a single valid JSON object — no prose, no markdown, "
               "no explanations, no code fences. Never repeat the resume text.")

PROMPT = """You are a senior HR recruiter evaluating a candidate's resume against a job description.

=== JOB DESCRIPTION ===
Software Engineer with 3+ years Python experience

=== RESUME: John Smith ===
John Smith | john@email.com
Experience: 2 years Python developer at TechCorp. Built REST APIs with Django.
Education: BS Computer Science 2021.
Skills: Python, Django, SQL, Git.

Return ONLY a valid JSON object:
{
  "score": <integer 0-100>,
  "strengths": ["...", "...", "..."],
  "weaknesses": ["...", "...", "..."],
  "hire_recommendation": "<Strong Hire | Hire | Maybe | No Hire>",
  "summary": "<2-3 sentence assessment>"
}"""

try:
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": SYSTEM_JSON},
            {"role": "user", "content": PROMPT},
        ],
        temperature=0.1,
        max_tokens=1024,
        response_format={"type": "json_object"},
    )
    raw = r.choices[0].message.content
    print("RAW:", raw[:300])
    data = json.loads(raw)
    print("PARSED OK:", list(data.keys()))
except groq_lib.BadRequestError as e:
    print("BadRequestError:", e)
except groq_lib.RateLimitError as e:
    print("RateLimitError:", e)
except groq_lib.APIError as e:
    print("APIError:", type(e).__name__, e)
except Exception as e:
    print("Other error:", type(e).__name__, e)
