# Certificate Generator - Backend (FastAPI)

## Quick Start

```bash
# 1. Create virtual environment
python -m venv venv
source venv/bin/activate          # macOS/Linux
# venv\Scripts\activate           # Windows

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure SMTP
cp .env.example .env
# Edit .env with your Gmail address + App Password

# 4. (Optional) Drop a custom TTF in ./fonts/  e.g. Poppins-Bold.ttf
#    The renderer will pick it up automatically.

# 5. Run
uvicorn main:app --reload --port 8000
```

API will be live at `http://localhost:8000`. Interactive docs: `http://localhost:8000/docs`.

## Gmail Setup (Important)

Gmail will not accept your regular password. You need an **App Password**:
1. Enable 2-Factor Authentication on your Google account.
2. Visit https://myaccount.google.com/apppasswords
3. Create an app password labeled "Certificate Generator".
4. Paste the 16-character password into `.env` as `SMTP_PASSWORD`.

For other providers (Outlook, custom SMTP), update `SMTP_HOST` and `SMTP_PORT` accordingly.

## API Endpoints

| Method | Path | Purpose |
| --- | --- | --- |
| POST | `/api/upload` | Upload template + recipient list (CSV/XLSX) |
| GET  | `/api/template/{job_id}` | Get the uploaded template image |
| POST | `/api/preview` | Render a single sample certificate |
| POST | `/api/generate` | Generate all certificates |
| POST | `/api/send` | Email all certificates to recipients |
| GET  | `/api/status/{job_id}` | Poll progress (sending, sent, failed) |
| GET  | `/api/download/{job_id}` | Download all certs as a zip |

## CSV/XLSX format

Two columns are required (case-insensitive): **name** and **email**.
Accepted aliases: `full name`, `student name`, `participant`, `email id`, `mail`.

```csv
name,email
Aarav Sharma,aarav@example.com
Priya Patel,priya@example.com
```
