# Fleet Operations — Control Center

A classical, responsive web app for your truck/transport company, built from
`Fleet_Operations_Master_Final.xlsx`. Every sheet from the workbook appears as a
section, and you can **view, search, add, edit and remove** records.

## How to open

Just **double-click `index.html`** — it runs in any modern browser (Chrome, Edge,
Firefox). No internet or install needed.

> Tip: for the smoothest experience you can also serve the folder locally by
> running `.server\serve.ps1` and opening `http://localhost:8777/`, but this is
> optional.

## Files

| File / folder        | Purpose                                                      |
|----------------------|--------------------------------------------------------------|
| `index.html`         | The app — open this.                                         |
| `styles.css`         | Theme & layout (navy + amber, classical serif headings).    |
| `app.js`             | All the logic (tables, edit, add/remove, search, export).   |
| `data.js`            | All workbook data (27 sheets) — this is your database.      |
| `.tools\import_excel.ps1` | Re-generate `data.js` from a fresh Excel export.       |
| `.server\serve.ps1`  | Optional tiny local web server.                              |

## Using it

- **Overview** shows a card per section with row/column counts. Click any card.
- **Left sidebar** lists every sheet; the filter box narrows the list.
- Inside a section:
  - **Search** filters rows live.
  - **Click any cell** to edit it; press *Enter* or click away to save.
  - **+ Add Row** appends a new blank row (jumps you to it).
  - **✖** on the left of a row deletes it.
  - **CSV** downloads the current section as a spreadsheet file.
  - Pager + "rows / page" handle large sheets (e.g. Anomaly Log has 100k+ rows).

## Two ways to run it

### A) Live — connected to Google Sheets (recommended)
Refresh pulls the latest data from your Google Sheet, it auto-updates, **many
computers can use it at once**, and adding/editing/deleting on the website writes
straight back into the sheet.

**One-time setup (just edit one line in code):**
1. Open `.tools/GoogleAppsScript.gs`, follow the steps at the top (paste it into
   your sheet's **Extensions → Apps Script**, then **Deploy → Web app**, access
   = *Anyone*). Copy the Web App URL (ends with `/exec`).
2. Open **`config.js`** and paste the URL into `webAppUrl: "...."`. Save. **Done.**
   Every computer that opens this folder now connects **automatically** — no
   pasting anything in the website itself.
3. On open the chip top-right turns green **● Live**.

Then:
- **↻ Refresh** pulls the current section fresh from the sheet (and re-connects
  automatically if needed). No URL entry required, ever.
- It **auto-refreshes** every 30s (toggle in the ⚙ gear). 
- **+ Add Row / edit a cell / ✖ delete** writes directly into the Google Sheet,
  safely (the script uses a lock so two people can't clash).
- The **Dashboard** tab is read-only (it's auto-calculated by formulas).
- Huge tabs (e.g. Anomaly Log) load the newest 5,000 rows for speed
  (change `bigSheetLimit` in `config.js`).

### B) Offline — bundled data.js (no internet)
Leave `config.js` `webAppUrl` empty. The site runs from `data.js`. Edits are saved
**in your browser** (localStorage). To make them permanent/shareable, click
**💾 Save File** → it downloads an updated `data.js` → replace the one in this folder.

## Re-importing an updated Excel file

If you get a new version of the Excel workbook, regenerate the data:

```powershell
powershell -ExecutionPolicy Bypass -File ".tools\import_excel.ps1" -Source "C:\path\to\NewWorkbook.xlsx"
```

It reads every sheet (in workbook order), converts dates, fixes Hindi/emoji text,
and writes a fresh `data.js` next to `index.html`.
