# Resume Intake

A small login-gated web app: upload resumes (PDF/DOCX), each one is parsed and stored
in a database, and you can search or export everything to Excel anytime.

## Run it locally

```bash
pip install -r requirements.txt
export ADMIN_USERNAME=admin
export ADMIN_PASSWORD=choose-a-real-password
export SECRET_KEY=any-random-string
python app.py
```

Open http://127.0.0.1:5000 and sign in with the admin username/password you set above.

## Deploy to Render

1. Push this folder to a new GitHub repository.
2. In Render: **New → Web Service**, connect the repo.
3. Render will detect `render.yaml` automatically (or set manually):
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app`
4. Set these environment variables in Render's dashboard:
   - `ADMIN_USERNAME` — your login username
   - `ADMIN_PASSWORD` — your login password (**required**, or a weak default is used)
   - `SECRET_KEY` — any random string (Render can auto-generate this)
   - `ANTHROPIC_API_KEY` — *(optional)* add this to turn on AI-enhanced parsing for
     every upload, server-side. Without it, the app uses a built-in offline
     pattern-matching parser — no key needed to run at all.
5. Deploy. Render gives you a live URL — that's your app.

### Important: persistent storage

Render's free web services use an **ephemeral filesystem** — anything written to disk
(including the SQLite database) is wiped on every redeploy or restart. The included
`render.yaml` attaches a small **persistent disk** mounted at `data/`, which keeps your
database across deploys. This requires a paid Render plan that supports disks (the
free tier does not support persistent disks). If you want to stay on the free tier,
swap SQLite for Render's free PostgreSQL database instead — ask me and I can wire
that up.

## What gets stored

Each upload is parsed into: name, email, phone, LinkedIn, location, current title,
years of experience, education, skills, a short summary, and up to 5 recent roles
(title/company/dates/highlights). The original PDF/DOCX file itself is deleted after
parsing — only the extracted data is kept in the database.

## Adding teammates

Once logged in, expand "Add a teammate" at the bottom of the dashboard to create more
login accounts. Everyone shares the same candidate database.
