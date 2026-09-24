# Progress Log

A personal workout tracker built with Flask. It tracks lifts, templates, PRs, progress charts, body weight, cardio (with splits), sleep, rest days, weekly notes and stats, and can export your data to Excel.

## Run it locally

Requires Python 3.12+.

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open http://localhost:5000. The first run creates an empty `workout.db` (SQLite) next to `app.py` and fills the exercise bank with a starter list. You can add your own exercises on the Exercises page.

## Make it yours

Open **Settings** (the gear icon in the top bar, or the sidebar) to set:

- **Weekly schedule:** what you train each day, shown as a badge on the Dashboard and Lifts pages
- **Goals:** weekly cardio minutes, nightly sleep hours, streak goal, and whether weekends break streaks
- **Body:** goal weight, height, age and sex (used for BMI and body fat estimates)

## Environment variables (all optional)

| Variable       | What it does                                                                 |
|----------------|------------------------------------------------------------------------------|
| `APP_PASSWORD` | If set, the site asks for this password before you can use it. Set this if you deploy it publicly. |
| `SECRET_KEY`   | Signs login sessions. Set it to a long random string when deployed, or you'll get logged out on every restart. |
| `DATABASE_URL` | A Postgres connection URL. If set, the app uses Postgres instead of SQLite.  |
| `API_KEY`      | Protects the `/api/export/*` JSON endpoints.                                 |
| `PORT`         | Port to run on (default 5000).                                               |

## Deploy (Railway)

1. Push this folder to a GitHub repo.
2. On [Railway](https://railway.app), create a new project from that repo.
3. Add a **Postgres** database to the project. Railway sets `DATABASE_URL` for you.
4. Set `APP_PASSWORD` and `SECRET_KEY` in the service's variables.
5. Railway uses the `Procfile` (`gunicorn app:app`) to start the app. Tables are created automatically on startup.

## Add it to your phone

Open the site in Safari, tap Share, then **Add to Home Screen**. It runs like an app.
