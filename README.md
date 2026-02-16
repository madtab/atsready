# ATSReady

**Convert PDF resumes into ATS-optimised formats in 60 seconds.**

Most resumes are automatically rejected by Applicant Tracking Systems before a human ever reads them. ATSReady takes a PDF resume, extracts the content, scores it for ATS readability, and outputs a clean `.docx` file that ATS parsers can reliably read.

## Project Structure

```
atsready/
├── ats_converter.py      # Core conversion engine
├── app.py                # Flask API (endpoints: /convert, /score, /health)
├── requirements.txt      # Python dependencies
├── Procfile              # Railway deployment config
├── index.html            # Holding / coming soon page (Netlify)
├── app.html              # Full converter site (Netlify — swap to index.html at launch)
└── AUTOMATION_GUIDE.md   # Deployment & automation reference
```

## Backend (Python / Flask)

### Local Development

```bash
pip install -r requirements.txt
python -m flask --app app run
```

API runs at `http://127.0.0.1:5000` with three endpoints:

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/convert` | POST | Accepts a PDF, returns an ATS-optimised `.docx` |
| `/score` | POST | Accepts a PDF, returns JSON with ATS score and issues |
| `/health` | GET | Health check |

Both POST endpoints expect multipart form data with a `resume` field (PDF file) and an `X-API-Key` header.

### CLI Usage

```bash
python ats_converter.py input.pdf output.docx
```

Prints a JSON score report to stdout and saves the converted `.docx`.

### Deployment

**Railway**: Connect this repo → auto-deploys from `main` branch. Uses `Procfile` and `requirements.txt` automatically.

**PythonAnywhere**: Upload `ats_converter.py` and `app.py`, install dependencies via console, configure as Flask web app.

## Frontend (Static HTML)

### Local Testing

Open `app.html?test=true` in a browser with the Flask API running locally. Test mode bypasses Stripe and sends files directly to the API.

### Deployment

Deploy `index.html` (and `app.html` when ready) to **Netlify** via drag-and-drop or connect this repo.

At launch, rename `app.html` → `index.html` and redeploy.

## Automation

See `AUTOMATION_GUIDE.md` for full setup instructions covering:

- Stripe payment integration (£7 per conversion)
- Make.com webhook automation (Stripe → convert → email)
- Cost breakdown and scaling notes

## Tech Stack

- **Backend**: Python, Flask, pdfplumber, python-docx
- **Frontend**: HTML, vanilla JS, Google Fonts (Fraunces + DM Sans)
- **Payments**: Stripe Payment Links
- **Automation**: Make.com
- **Hosting**: Netlify (frontend), Railway or PythonAnywhere (backend)
