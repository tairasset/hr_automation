# AI Resume Screener

A Python and Streamlit prototype that compares resumes against a job description and shows the criteria behind each assessment.

## What you can try

- Upload PDF, DOCX, RTF, or TXT resumes, or import candidate text from a CSV.
- Set weights for experience, skills, education, and soft skills.
- Compare completed assessments, draft interview questions, and export PDF reports. Failed assessments appear as unavailable and stay out of candidate rankings.

## Run locally

Use Python 3.10 or later.

```sh
cd resume_screener
python -m pip install -r requirements.txt
```

Copy `.env.example` to `.env` inside `resume_screener` and set `GROQ_API_KEY`. The default model is `openai/gpt-oss-120b`. Set `GROQ_MODEL` if your Groq project uses a different compatible model. Start the app:

```sh
python -m streamlit run app.py
```

## Demo

Use the [fictional resume](examples/Sample_Candidate.txt) and try a role that requires Python, API integration, and basic documentation. The app sends resume text and job requirements to Groq for analysis. Review each result against the original resume and job description before making hiring decisions.
