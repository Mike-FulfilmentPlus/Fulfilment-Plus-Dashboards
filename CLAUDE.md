# CLAUDE.md

## Git workflow
- Commit and push changes directly to `main`. Don't create feature branches or PRs unless asked.
- GitHub Actions push auto-refresh commits to `main` frequently, so pull (`git pull --ff-only`) before committing and pushing.

## Logging issues and exceptions (requests from GitHub issues)

When a GitHub issue asks you to log an issue or exception, edit `pipeline/kpi_dashboard.xlsx` with openpyxl and open a **pull request** for review (do not push to main for these requests; this overrides the "commit directly to main" rule above). Always start from the latest `main`. Never edit any sheet other than the one named below, and never touch formulas or the notes rows. After saving, reload the file and print the new row(s) back, and put them in the PR description. Use today's date in NZ time (`Pacific/Auckland`) for Date Logged unless the issue says otherwise. If a required field is missing or ambiguous, do NOT guess: reply on the issue asking for it and make no edit.

### Issue Log tab (feeds Order Accuracy by Unit)
Header is row 4: `Date Logged | Customer | Order Reference | Quantity Wrong | Issue Details`. Data starts at row 5. Write the new row in the first row where column A is empty. Required: customer (check against the Customers tab for spelling), order reference, Quantity Wrong, issue details.
- Quantity Wrong is a whole number of UNITS affected by this specific issue, not the whole order. If the issue doesn't state a number, ask. Never assume 1.
- Date Logged is a real date (not text). The month an issue counts against comes from Date Logged.

### Exceptions tab - Order-Level Exceptions table
Columns: `Date Logged | Customer | Order Reference | Month (YYYY-MM) | KPI | Action | Reason | Date From | Date To`.
- Insert new rows directly ABOVE the grey note row that starts `Action must be either...` (use `ws.insert_rows`, which shifts everything below down). Find that row by its text, don't hard-code a row number. Copy font, fill, border, alignment and number format from the existing data row above it.
- Date Logged: a date (yyyy-mm-dd). Month: text `YYYY-MM` (the month the order counts against). Month and Date From/To may be blank for open-ended rules.
- KPI must match exactly: `Order Cut Off`, `Dock to Stock`, `Freight DIFOT`, etc. Action is exactly one of:
  - `Count as met` - the order passes whatever time it ships.
  - `Exclude from calculation` - the order is removed from the KPI.
  - `Fixed (YYYY-MM-DD HH:MM)` - **Order Cut Off only.** Moves THIS order's deadline to that NZ time (e.g. `Fixed (2026-10-16 13:00)`). The order must ship by then to count as met, and shows that deadline on the warehouse dashboard instead of showing late. Use it when a request says an order "isn't required until <date/time>" or "has an issue and needs to be pushed out to <date/time>". If only a date is given with no time, use 20:00 and say so in the PR description. Put the reason in the Reason column. Do NOT use `Count as met` for these unless the issue asks for it.
- Order Reference `*` is a wildcard meaning every order for that customer (optionally limited by Date From/Date To). It is how ongoing rules are written, e.g. Flo & Frankie Order Cut Off, Made Group NZ Order Cut Off. Only use `*` when the issue clearly asks for all orders; otherwise one row per order number.
- Reason: one short operational sentence on every row.
- A deadline change for ONE specific order (or a few named orders) is a `Fixed (...)` row here, one row per order. A one-off deadline change for a group of orders (e.g. "due by 13 Oct instead of next Tuesday") belongs in the "Order Content-Based Overrides" table, not here, using a Fixed deadline such as `Fixed (2026-10-13 20:00)` with Effective From/To limited to the creation date. If asked for one of these, say what you propose in the PR description.
