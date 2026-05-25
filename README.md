# uncorrupt-excel

**Find Excel-mangled gene names without leaving Excel.**

A single TypeScript file you paste into Excel's Automate tab. Click Run. A report sheet appears showing every flagged cell colour-coded by confidence. No install. No CLI. Works on Excel for the web, Windows, and Mac (Office Scripts must be enabled by your tenant admin).

## How to use

1. Open your spreadsheet in Excel for the web (or Excel desktop with Office Scripts enabled).
2. Click the **Automate** tab in the ribbon.
3. Click **New Script**.
4. Paste the entire contents of [`uncorrupt-scan.ts`](uncorrupt-scan.ts).
5. Click **Run**.

A new sheet appears titled `UnCorrupt-Report` with one row per flagged cell:

- 🔴 Red rows: high-confidence corruption (column has multiple similar mangled cells)
- 🟡 Yellow rows: medium-confidence (pattern matches but only one example in this column)
- ⚪ Grey rows: low-confidence (worth a glance; often false alarms)

## What this catches

The four most common Excel-to-gene-name corruptions:

| Excel did this | UnCorrupt-excel flags it as |
|---|---|
| `SEPT2` becoming `2024-09-02` | Date mistaken for gene name |
| `Sep-07`, `9-Sep`, `Oct-3/4` | Date text in gene column |
| Excel-serial integer in a column where dates were also mangled | Excel-serial integer in gene column |
| `2310009E13` becoming `2.31e+19` | Float (gene ID lost to scientific notation) |

## What this does NOT catch

The full [Python package](https://github.com/shitcoinsherpa/UnCorrupt) covers these; the Office Scripts version cannot because the sandbox does not allow downloading the required databases:

- Cross-reference boost using the 1.8 M cross-species symbol pool (HGNC + MGI + ZFIN + FlyBase + WormBase + RGD)
- Unicode homoglyph normalisation (Cyrillic `А` for Latin `A`, etc.)
- Heavy column-uniformity gates that the Python detector uses to suppress false positives

If you need any of that, export the file and run the Python CLI locally, or use the [browser-only Pyodide build](https://github.com/shitcoinsherpa/uncorrupt-pyodide) (no install, file never leaves your machine).

## Compatibility

Office Scripts is generally available on:

- Excel for the web
- Excel for Windows (Microsoft 365 Apps for business with Office Scripts enabled)
- Excel for Mac (same prerequisite)

## See also

- [`shitcoinsherpa/UnCorrupt`](https://github.com/shitcoinsherpa/UnCorrupt) : the canonical Python package, CLI, and Gradio UI
- [`shitcoinsherpa/uncorrupt-pyodide`](https://github.com/shitcoinsherpa/uncorrupt-pyodide) : browser-only build, no install
- [HuggingFace Space](https://huggingface.co/spaces/Sherpa/uncorrupt) : hosted Gradio version

## License

Apache-2.0.
