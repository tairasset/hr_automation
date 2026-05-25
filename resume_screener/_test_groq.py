import os
from dotenv import load_dotenv
from groq import Groq

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

r = client.chat.completions.create(
    model="llama-3.3-70b-versatile",
    messages=[
        {"role": "system", "content": "You ONLY output a single valid JSON object — no prose, no markdown, no code fences."},
        {"role": "user", "content": 'Respond with exactly: {"score": 80, "hire_recommendation": "Hire", "strengths": ["a","b","c"], "weaknesses": ["x","y","z"], "summary": "Good."}'},
    ],
    max_tokens=200,
)
print("CONTENT:", repr(r.choices[0].message.content))
print("FINISH :", r.choices[0].finish_reason)
