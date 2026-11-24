# Rack Power Report

A web page that turns the raw export from the monitoring software (CSV or Excel) into the same report as the Custom Report Excel template: one grey row per rack with its total, and one row per PDU with Last Reading, Current Reading and Usage.

The client opens the page, loads the raw data, clicks their company, chooses a date range, and gets their report. They can download it as a PDF or CSV. Everything runs in the browser. The data file is never uploaded anywhere.

## Files

| File | What it is |
|---|---|
| `index.html` | The page |
| `css/styles.css` | All the styling |
| `js/lookup.js` | **Racks and companies.** The only file to edit when racks or companies change |
| `js/core.js` | Reading the files and working out usage |
| `js/ui.js` | The buttons, tables and the PDF and CSV downloads |
| `assets/logo.png` | **Optional.** Your logo, shown at the top right of the PDF (see below) |

## The PDF report

"Download PDF" makes an A4 report with the same header as the Excel template: `<COMPANY> POWER CONSUMPTION REPORT`, Company Name, Location, Reading Date (today), and From Date and To Date from the range the client chose. The table below it repeats its column titles on every page and ends with a Total row.

To show your logo, save it as `assets/logo.png` (a PNG, ideally with a transparent or white background). It appears at the top right. If the file is missing, the PDF is made without a logo. The logo only appears when the page is served from a web address (GitHub Pages, or `python3 -m http.server`), not when `index.html` is opened straight from a folder.

The location line comes from `BUILT_IN_LOCATION` at the bottom of `js/lookup.js`.

## Run it on your computer

Open `index.html` in a browser. Or, from this folder, run `python3 -m http.server` and open http://localhost:8000.

The page loads three libraries from the internet (jsPDF and its table plugin for the PDF, and SheetJS for reading Excel files), so it needs a connection. If the PDF button says the library did not load, the connection or a network filter is blocking `cdnjs.cloudflare.com`.


**Privacy.** `js/lookup.js` lists every rack and company, and anyone who can open the page can read it. The page also has no login, so any visitor can choose any company's button. A public repository makes the lookup public too. Use a private repository, or host the page where only your clients can reach it. GitHub Pages from a private repository needs a paid GitHub plan.

## Change the racks or companies

Open `js/lookup.js` and edit the lines between the backticks. One line per meter:

```
meter name, rack name, company
PDU-S1-FN4-R01,NR1,AGIC2
```

- A meter listed under two racks is counted in both, as in the Excel template.
- Keep the backticks around the list.
- Easy way: on the page, open "Import racks and companies from your Excel report", choose the updated Custom Report workbook and click "Use these racks and companies". Copy the lines from the list-1 box in the settings and paste them into `js/lookup.js`.
- Uploading the Custom Report workbook itself also works: the page then reads the racks and companies from its sheets for that visit.

## How usage is worked out

This follows the Excel template. With the date range chosen on the page:

- **Last Reading** is the first reading inside the range.
- **Current Reading** is the latest reading inside the range.
- **Usage** is Current minus Last. Both the start and end dates are included.

## Known problems in the Excel template

The page reproduces the template's meters but not its mistakes:

- `S20-1`, rack `COLR012`: `PDU-S2-FN1-B10` has a stray line break, so Excel shows 0 for it.
- A typed total on the AGCI rack `VSP`, and `PVPA_001`, `PVPA_006` and `PVPA_007` sum only part of their meters.
- `PDU-S1-FN2-B08` is listed twice in `COLR016` and `PDU-S2-FN2-B08` is in no sheet.
- `PDU-S1/S2-FN2-Y16` and `Y17` are in both AGIC and COLO-S20, and `PDU-S1/S2-FN1-Y14` is under both `COLR014` and `COLR015`.
