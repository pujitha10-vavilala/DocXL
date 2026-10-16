# DocXL Backend


## Windows PowerShell

From your existing backend directory (back up `app\main.py` first):

```powershell
# Activate your existing environment, or create one with: python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:DOCXL_AUTH_SECRET = (python -c "import secrets; print(secrets.token_urlsafe(48))")
python -m uvicorn app.main:app --reload
```

API docs: http://127.0.0.1:8000/docs
Health check: http://127.0.0.1:8000/health


**Development limitation:** document upload/extract/export/download routes currently do not enforce per-user ownership. Do not expose publicly or use for sensitive multi-user production data until ownership checks and deployment security are added.

Authentication uses DocXL email/password accounts. Google Sign-In has been removed.
