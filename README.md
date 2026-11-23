# 💍 M&E BIG DAY

The wedding planner app for Ed & Touta, installed on the iPhone home screen and linked live to
`_Wedding_Budget_Planner.xlsx` in OneDrive.

**👉 First time? Follow [SETUP.md](SETUP.md).**

## What's inside
- **Home:** countdown to **Sept 25, 2027**, budget KPIs in USD first (EGP underneath),
  where the money goes, payments due, RSVP and vendor progress, and the next tasks.
- **Money:** every budget category (allocation, forecast, paid, remaining, status) and the payments log.
  You can add payments, mark them paid and set due dates.
- **Vendors:** all quote tabs (Venue like-for-like ranking, Decor, Catering, Dress…). Add, edit, compare and select quotes.
- **Guests:** the guest list (side, group, RSVP, plus-ones, phone, travel and hotel) and a drag-and-drop **seating chart**.
  It starts with round tables of 10, and you can change the number of tables and their size and shape anytime.
- **More:** Timeline tasks, the wedding trip days, and Settings (live exchange rate, budget, guest count, wedding date, sharing).

## How it works
Plain HTML, CSS and JavaScript with no build step and no third-party libraries.
- `js/graph.js`: Microsoft sign-in (OAuth PKCE) and OneDrive access through Microsoft Graph
- `js/zip.js`, `js/xlsx.js`: read and patch the .xlsx at the cell level, so formatting and formulas are preserved
- `js/model.js`: recalculates the workbook's formulas (budget, quote totals and ranks, venue like-for-like, payment status)
- `js/store.js`: offline cache, the edit queue, and sync with merge on save
- `sw.js`: lets the app open offline

The workbook itself is never committed here.
