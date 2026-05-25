# uncorrupt-excel

An Excel-native TypeScript scanner for the patterns documented in [`shitcoinsherpa/UnCorrupt`](https://github.com/shitcoinsherpa/UnCorrupt). Runs inside the Excel Web Automate sandbox, no install required.

Catches the four most common surfaces:

1. Gene symbol coerced to date (`SEPT2` becomes `2024-09-02`)
2. Gene symbol stored as date text (`Sep-07`, `9-Sep`, `Oct-3/4`)
3. Excel-serial integer that decodes to a vulnerable date in a column with other date corruption
4. RIKEN or accession identifier coerced to scientific-notation float (`2310009E13` becomes `2.31e+19`)

Doesn't catch (these need the full Python package at [`shitcoinsherpa/UnCorrupt`](https://github.com/shitcoinsherpa/UnCorrupt)):

- Row-xref boost against the 1.8 M cross-species symbol pool
- Heavy column-uniformity gates
- Unicode-homoglyph normalisation (TR39)

For full-precision scans, export the file and run `uncorrupt detect` locally, or use the Pyodide build at [`shitcoinsherpa/uncorrupt-pyodide`](https://github.com/shitcoinsherpa/uncorrupt-pyodide) (browser-only, file never leaves the page).

## Install

1. Open the file in Excel on the web (or Excel desktop with Office Scripts enabled).
2. Click the **Automate** tab.
3. Click **New Script**.
4. Paste the entire contents of [`uncorrupt-scan.ts`](uncorrupt-scan.ts).
5. Click **Run**.

A `UnCorrupt-Report` sheet appears with one row per flagged cell, colour-coded by post-detection confidence band:

- Red (>= 0.95): high confidence, column-corroborated by other corruption in the same column
- Yellow (0.60 to 0.95): moderate confidence
- Grey (< 0.60): low confidence, isolated cells the column-corroboration gate could not lift

## Compatibility

Office Scripts is generally available on:

- Excel for the web
- Excel for Windows (with a Microsoft 365 Apps for business plan that includes Office Scripts)
- Excel for Mac (same prerequisite)

## License

Apache-2.0, matching the upstream package.

## See also

- [`shitcoinsherpa/UnCorrupt`](https://github.com/shitcoinsherpa/UnCorrupt) - the canonical Python package
- [`shitcoinsherpa/uncorrupt-pyodide`](https://github.com/shitcoinsherpa/uncorrupt-pyodide) - browser-only WASM build
