# SEC Financials

Pull a public company's **income statement, balance sheet, and cash flow**
straight from its latest SEC 10-K and get a clean Excel workbook — from a
command line, or from a tiny double-click app.

Data comes from [SEC EDGAR](https://www.sec.gov/edgar) XBRL structured facts via
[edgartools](https://github.com/dgunning/edgartools). Free, no API key.

Each statement is written twice:

- **Clean model view** — primary face lines only, values in **$ millions**, the
  latest 3 fiscal years (2 year-ends for the balance sheet).
- **Detail view** — every line item as reported (including product/geographic
  breakdowns), all periods, in full dollars.

Everything ties back to what the company actually reported — this tool assembles
the factual statements; it does not model, forecast, or give advice.

## Install

```bash
python -m pip install edgartools openpyxl pandas
```

Requires Python 3.9+.

## Use it from the command line

```bash
python sec_financials.py GTM
```

This writes `GTM_financials.xlsx` in the current folder. SEC requires you to
identify your requests — set your name and email once:

```bash
# macOS/Linux
export EDGAR_IDENTITY="Jane Smith jane@company.com"
# Windows (PowerShell)
setx EDGAR_IDENTITY "Jane Smith jane@company.com"
```

## Use it as a desktop app

```bash
python sec_financials_app.py
```

A small window opens: type a ticker, enter your name/email, click **Generate**,
and the workbook opens automatically.

## Build a standalone .exe (Windows)

So non-technical users need nothing installed:

```bat
build_exe.bat
```

This produces `dist/SEC Financials.exe` — a single file you can share. Tip:
attach it to a GitHub **Release** rather than committing it (it's large and
platform-specific). Unsigned executables trigger a one-time Windows SmartScreen
prompt ("More info → Run anyway").

## Open the result in Google Sheets

`File → Import → Upload`, choose the `.xlsx`, and pick "Insert new sheet(s)".

## Notes & limitations

- The clean view keeps the first line per XBRL concept to stay tidy; the detail
  tab always has the complete as-reported picture.
- Line coverage depends on how each company tags its filing; unusual filers may
  show blanks in the clean view — check the detail tab.
- Not investment advice. Verify against the original filing before relying on any
  figure.

## Author

Kyle Gilbride

## License

MIT — see [LICENSE](LICENSE).
