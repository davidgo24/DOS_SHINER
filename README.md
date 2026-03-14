# DOS Report Formatter

A web app that transforms raw DOS spreadsheet reports into a formatted layout matching the standard output format. Drop in your Excel file, configure the section order if needed, and export to Excel or PDF.

## How to Use

1. **Open the app** – Run a local server (see below) and open `index.html` in your browser.

2. **Drop your file** – Drag and drop your raw DOS Report Excel file (`.xlsx` or `.xls`) onto the upload area, or click to browse.

3. **Configure (optional)** – Adjust the **Section Order** using the structured blocks: pick from starter templates in the dropdown, add custom sections as needed, and use the ↑/↓ buttons to reorder. Use `*paddle` for numeric paddle blocks (10001–10077).

4. **Export** – Use **Export Excel** to download the formatted spreadsheet, or **Print / Save as PDF** to print or save as PDF from your browser.

## Running Locally

From the project folder:

```bash
# Python 3
python3 -m http.server 8000

# Or with npx
npx serve .
```

Then open http://localhost:8000 (or the port shown) in your browser.

## Configuration

- **Section Order** – One section name per line. Rows are grouped by their "Paddle" value and placed in the order you define. Use `*paddle` for standard block numbers.
- **Report Title** – Shown at the bottom (e.g., "SUPERVISORS ABSENT").
- **Report Date** – Shown next to the title (e.g., "WEDNESDAY - 3/11/2026").

## Expected Raw Columns

The app expects these column names (case-insensitive) in your Excel file:

- Paddle, Block  
- Planned Shift Start Time, Planned Shift End Time, Hrs (Planned Duration)  
- Vehicle  
- Actual Start Time, Actual End Time, Trim (Actual Duration, Hrs.)  
- Primary Driver Name, Primary Driver ID  
- Alternative Driver Name, Alternative Driver ID  
- Labels, Driver Notes, Internal Notes, Was Cancelled  

If your export uses different names, the column mapping in `app.js` (DEFAULT_COLUMN_MAP) can be extended.
