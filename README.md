# ResumePilot 3.0

## Main fixes
- Removes line-by-line keyword sorting completely.
- Keeps original resume section order and paragraph/bullet order.
- Only the Professional Summary may be rewritten by Gemini, and only using facts already in that summary.
- Without Gemini, the original summary and all resume content are retained.
- Adds weighted match score:
  - Role relevance 40%
  - Experience 25%
  - Technical skills 20%
  - Qualifications 8%
  - Location 4%
  - Other 3%
- Extracts job title/company/location from Schema.org JobPosting JSON-LD and common metadata before fallback.
- Keeps individual job cards and individual DOCX downloads.
- Supports up to 500 URLs per import batch.

## Deploy over the existing app
1. Download and extract this ZIP.
2. Open the GitHub repository currently connected to the Streamlit app you want to update.
3. Replace `app.py` and `requirements.txt` with these files.
4. Commit to the same branch used by that Streamlit app (usually `main`).
5. Wait for Streamlit to redeploy. Do not delete the app or create a new Streamlit deployment.
6. Open the existing app URL and press Ctrl+Shift+R.

Important: If you currently have two Streamlit app URLs, update the repository connected to the URL you intend to keep. The code commit does not automatically merge two separate Streamlit apps.

## Gemini
The API key can be entered in the app sidebar. For Streamlit Secrets, configure:
GEMINI_API_KEY = "your-key"
GEMINI_MODEL = "gemini-2.5-flash"

The model only receives the existing summary and JD, and is instructed not to add unsupported claims. Review generated content.

## Limitations
- PDF extraction works best with text-based PDFs; scanned PDFs require OCR.
- Some ATS/job pages block bots or render descriptions in JavaScript. Paste the JD manually when extraction fails.
- Match score is a heuristic estimate, not a real ATS score or hiring probability.
- Experience and location scoring are approximate when a JD lacks explicit requirements.
- 500 is an import cap, not a guarantee that 500 pages can be fetched/processed in one free hosting session.
- This app still uses temporary Streamlit session state. It does not yet include authentication, persistent database storage, private per-user files, or production-grade multi-user security. Do not invite public users to upload personal resumes until those features are implemented and tested.
- Apply Job opens the employer's job URL; it does not submit applications automatically.
