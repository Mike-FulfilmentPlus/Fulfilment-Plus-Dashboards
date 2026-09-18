# Adding the Desk365 report to the GitHub dashboards repo

These files add the Desk365 customer-support visibility report to the same
repo as the other KPI dashboards, refreshed automatically by GitHub Actions
(no Claude session involved in the routine refresh).

## Files in this package

- `desk365_report.py` — fetches tickets from Desk365, computes the metrics,
  writes `desk365_visibility_report.html`. Reads the API key from the
  `DESK365_API_KEY` environment variable — nothing is hardcoded.
- `.github/workflows/desk365-report.yml` — runs the script 3x/day and
  commits the refreshed HTML back to the repo automatically.
- `requirements.txt` — just `requests` (the script only needs that plus the
  Python standard library).

## Steps to add this (all on github.com, no coding needed)

1. **Rotate the Desk365 API key first.** The key used to build the first
   version of this report was typed directly into a chat conversation, so
   treat it as compromised — the same rule the CartonCloud/StarShipIt
   credentials are following in the migration checklist. Generate a fresh
   key in Desk365 (Settings → API), and only use the fresh one below.

2. **Add the fresh key as a GitHub secret** (never paste it into a chat,
   issue, or commit):
   - Go to the repo on github.com → **Settings** → **Secrets and variables**
     → **Actions** → **New repository secret**
   - Name: `DESK365_API_KEY`
   - Value: the fresh key from step 1
   - Save.

3. **Upload the files to the repo**, preserving the folder structure:
   - `desk365_report.py` → repo root
   - `requirements.txt` → repo root (if the repo doesn't already have one —
     if it does, just add the `requests` line to the existing file instead
     of overwriting it)
   - `.github/workflows/desk365-report.yml` → exactly that path (GitHub
     Actions only picks up workflows from `.github/workflows/`)

   Easiest way on github.com: use **Add file → Upload files**, drag in
   `desk365_report.py` and `requirements.txt` at the repo root, then
   separately navigate into (or create) the `.github/workflows/` folder and
   upload `desk365-report.yml` there.

4. **Test it manually before trusting the schedule**: go to the repo's
   **Actions** tab, click into the "Desk365 visibility report" workflow, and
   use **Run workflow** (this is what `workflow_dispatch` in the YAML
   enables). Check the run's log for errors, then confirm
   `desk365_visibility_report.html` appeared/updated in the repo.

5. **Enable GitHub Pages** for the repo, if it isn't already on for the
   other dashboards (Settings → Pages → deploy from the branch the workflow
   commits to). Once enabled, the report will be reachable at your Pages
   URL + `/desk365_visibility_report.html`, alongside `customer_dashboard.html`
   and `business_dashboard.html`.

6. Once you've seen a clean scheduled run or two, this report no longer
   depends on any Claude scheduled task or artifact — you can point people
   at the GitHub Pages link instead of the claude.ai artifact link.

## Schedule

The workflow runs at roughly 8:15am, 1:15pm and 5:15pm NZ time (UTC cron
times, so they'll drift by an hour during NZ daylight saving — not worth
fixing precisely for an internal report, but worth knowing). Edit the
`cron:` lines in the YAML any time to change the cadence, or just use
**Run workflow** on the Actions tab for an on-demand refresh.
