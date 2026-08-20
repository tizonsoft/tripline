# Trip Timeline

A tiny static website that turns a trip-planning spreadsheet into a beautiful,
phone-friendly, printable day-by-day timeline. No backend, no accounts,
nothing stored — the Excel file is parsed entirely in your browser and never
leaves your device.

## Try it

Open `index.html` in a browser (or visit the GitHub Pages URL once deployed),
then either:

- Upload your own `.xlsx` file, or
- Click **"Try it with a sample itinerary"** (needs the page served over
  http/https — see "Run locally" below; it won't fetch the sample if you just
  double-click `index.html`, though uploading your own file always works even
  then).

Use the **Save as PDF** button (top bar or the round button bottom-right) to
open the browser's print dialog and export a clean PDF — same result as a
printable itinerary, but it's a normal webpage the rest of the time.

## Expected Excel format

**Sheet 1 — `Itinerary`** (sheet name can vary, first sheet is used as a
fallback). Columns, in this order: `Date`, `Day`, `Accommodations`,
`Activities`, `Notes`. The column to the left of `Date` can optionally hold a
label like `Day 1`.

- One row starts a new day (must have a `Date` value).
- Extra rows directly underneath, with `Date` left blank, add more
  activities/notes to that same day.
- A blank row separates one day from the next.

| (label) | Date  | Day | Accommodations      | Activities                | Notes                     |
|---------|-------|-----|----------------------|----------------------------|----------------------------|
| Day 1   | 26-08 | Wed | Hotel Rias Baixas    | Arrive, drive south        | Flight lands evening       |
|         |       |     | Meis                 | Stop in Padrón for dinner  |                             |
|         |       |     |                      |                            |                             |
| Day 2   | 27-08 | Thu | Hotel Rias Baixas    | Illa de Arousa             | Kayak around the island    |

The first `Accommodations` value of each day is shown as a "🛏️ staying at"
badge on the day card; any later value in the same day is shown as a small
📍 location tag on that specific activity.

**Sheet 2 — `Resources`** (optional). A single column: a topic/title row,
followed by one URL per row, blank row, next topic, and so on. Rendered as
grouped, tappable link cards (hotel booking pages, maps, tour sites, etc).

The parser matches headers loosely (e.g. `Activites`/`Activities` both work,
sheet names containing "itinerary" or "resources" are auto-detected), so
small variations in your file are fine.

## Deploy to GitHub Pages

1. Create a new GitHub repo and push this folder's contents to it (root of
   the repo, or a `/docs` folder — either works with GitHub Pages).
2. In the repo: **Settings → Pages → Build and deployment → Deploy from a
   branch**, pick your branch and the folder you used, save.
3. GitHub gives you a URL like `https://<user>.github.io/<repo>/` — open it
   on your phone before the trip and bookmark it. Because everything (the
   spreadsheet library, styles, code) is bundled locally in this repo, the
   page keeps working with no signal once it's loaded once — handy on the
   road. Uploading your itinerary file works fully offline too.

## Run locally

```
cd trip-timeline
python3 -m http.server 8080
```

Then open http://localhost:8080 — this also makes the sample-itinerary button
work (it needs `fetch`, which browsers block on `file://` pages).

## Files

- `index.html` — page structure (upload screen + timeline/resources view)
- `style.css` — design, including a `@media print` stylesheet for clean PDFs
- `app.js` — spreadsheet parsing + rendering, all client-side
- `vendor/xlsx.full.min.js` — [SheetJS](https://sheetjs.com) bundled locally
  (so nothing needs the network at trip time)
- `sample/Galicia.xlsx` — example itinerary used by the "try a sample" button
