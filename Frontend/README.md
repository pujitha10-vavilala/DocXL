# DocXL Frontend


## Windows PowerShell

```powershell
Copy-Item .env.example .env
npm install
npm run dev
```

Open http://127.0.0.1:5173. Backend should be running at http://127.0.0.1:8000. Set `VITE_API_BASE_URL` if the backend uses a different local URL.


Authentication uses DocXL email/password accounts. Google Sign-In has been removed.
