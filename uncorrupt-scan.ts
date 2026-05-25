/*
 * UnCorrupt for Office Scripts (Excel-native TypeScript)
 *
 * Drop this into the Excel Web Automate tab to scan the active workbook
 * for Excel-style gene-symbol corruption. Self-contained, no install,
 * runs in the browser sandbox.
 *
 * What this catches:
 *   1. Gene to date corruption (cells like 2024-09-02 in gene-symbol columns)
 *   2. Gene to date-text (cells like "Sep-07", "9-Sep", "Oct-3/4")
 *   3. Gene to Excel-serial integer in columns that also have date corruption
 *   4. RIKEN / accession to huge-float coercion (cells >= 1e10 in id columns)
 *
 * What this does NOT catch (needs the full Python package):
 *   - row-xref boost (uses 1.8 M cross-species symbol pool, too large for OS)
 *   - column-uniformity gates (heavy column scans across the whole sheet)
 *   - homoglyph TR39 normalisation (no Unicode tables in the OS runtime)
 *
 * For the full-precision scan, export the file and run `uncorrupt detect`
 * locally, or visit the Pyodide page (file never leaves your browser).
 *
 * Run:
 *   Excel -> Automate tab -> New Script -> paste this -> Run
 *
 * Output: a "UnCorrupt-Report" sheet appended with one row per flag.
 */

// ---------- HGNC vulnerable-family subset ----------
// We embed only the date-corruption-vulnerable subset of HGNC: the
// SEPT, MARCH, DEC, OCT, NOV, APR, FEB, AGO families and their modern
// HGNC renames. Full HGNC coverage requires the Python package.
const VULNERABLE_FAMILIES: { [month: number]: { [day: number]: string[] } } = {
  // March (MARCHF1..MARCHF12 family)
  3: {
    1:  ["MARCH1",  "MARCHF1",  "MARC1"],
    2:  ["MARCH2",  "MARCHF2"],
    3:  ["MARCH3",  "MARCHF3"],
    4:  ["MARCH4",  "MARCHF4"],
    5:  ["MARCH5",  "MARCHF5"],
    6:  ["MARCH6",  "MARCHF6"],
    7:  ["MARCH7",  "MARCHF7"],
    8:  ["MARCH8",  "MARCHF8"],
    9:  ["MARCH9",  "MARCHF9"],
    10: ["MARCH10", "MARCHF10"],
    11: ["MARCH11", "MARCHF11"],
    12: ["MARCH12", "MARCHF12"],
  },
  // April (APR family; APR2 is fungal AGO/Argonaute alias documented in
  // Ziemann 2021 S1)
  4: {
    1: ["APR1"],
    2: ["APR2"],
    3: ["APR3"],
  },
  // August: locale-specific (Italian / Spanish / Portuguese "Ago" =
  // August). AGO2 typed in those locales -> "Aug-02", etc.
  8: {
    2: ["AGO2"],
    3: ["AGO3"],
    4: ["AGO4"],
  },
  // September (SEPT1..SEPT15 family, modern HGNC: SEPTIN1..SEPTIN15)
  9: {
    1:  ["SEPT1",  "SEP1",  "SEPTIN1"],
    2:  ["SEPT2",  "SEP2",  "SEPTIN2"],
    3:  ["SEPT3",  "SEP3",  "SEPTIN3"],
    4:  ["SEPT4",  "SEP4",  "SEPTIN4"],
    5:  ["SEPT5",  "SEP5",  "SEPTIN5"],
    6:  ["SEPT6",  "SEP6",  "SEPTIN6"],
    7:  ["SEPT7",  "SEP7",  "SEPTIN7"],
    8:  ["SEPT8",  "SEP8",  "SEPTIN8"],
    9:  ["SEPT9",  "SEP9",  "SEPTIN9"],
    10: ["SEPT10", "SEP10", "SEPTIN10"],
    11: ["SEPT11", "SEP11", "SEPTIN11"],
    12: ["SEPT12", "SEP12", "SEPTIN12"],
    13: ["SEPT13", "SEP13", "SEPTIN13"],
    14: ["SEPT14", "SEP14", "SEPTIN14"],
    15: ["SEPT15", "SEP15", "SEPTIN15"],
  },
  // October (OCT1..OCT11 family; OCT4 = POU5F1 alias)
  10: {
    1:  ["OCT1",  "POU2F1"],
    2:  ["OCT2",  "POU2F2"],
    3:  ["OCT3",  "POU5F1"],
    4:  ["OCT4",  "POU5F1"],
    5:  ["OCT5"],
    6:  ["OCT6",  "POU3F1"],
    7:  ["OCT7",  "POU3F2"],
    8:  ["OCT8"],
    9:  ["OCT9"],
    10: ["OCT10"],
    11: ["OCT11"],
  },
  // November
  11: {
    1: ["NOV1"],
  },
  // December
  12: {
    1: ["DEC1"],
    2: ["DEC2"],
  },
  // February
  2: {
    3: ["FEB3"],
    4: ["FEB4"],
  },
};

// Identifier-shaped column headers (gene-symbol columns expected to
// contain HGNC-style symbols, so any date / float here is suspicious)
const IDENTIFIER_HEADER_TOKENS: string[] = [
  "gene", "symbol", "gene_symbol", "genesymbol",
  "id", "ids", "identifier", "accession", "name",
  "probe", "probeid", "transcript", "feature",
  "hgnc", "ensembl", "entrez", "refseq", "uniprot",
];

// Measurement / quantitative column headers (numeric content is
// expected here, so a "float" in this column is not corruption)
const MEASUREMENT_HEADER_TOKENS: string[] = [
  "count", "counts", "rpkm", "fpkm", "tpm", "cpm",
  "log2", "log2fc", "logfc", "fc", "fold",
  "pvalue", "p-value", "p_value", "padj", "fdr", "qvalue", "q-value",
  "mean", "median", "stdev", "stddev", "std", "variance", "sd", "sem",
  "expression", "intensity", "signal", "score",
  // Cohort / publication metadata (also numeric, not gene-symbol)
  "donor", "subject", "patient", "participant",
  "pages", "page", "volume", "issue", "doi", "pubmed",
  "cas#", "cas no", "cas-no", "casno",
];

function lcase(s: string): string {
  return (s || "").toString().toLowerCase().trim();
}

function isIdentifierColumn(header: string): boolean {
  const lower = lcase(header);
  if (!lower) return false;
  const isMeas = MEASUREMENT_HEADER_TOKENS.some(t => lower.includes(t));
  if (isMeas) return false;
  return IDENTIFIER_HEADER_TOKENS.some(t => lower.includes(t));
}

function reverseGeneDate(d: Date): string[] {
  const year = d.getFullYear();
  if (year < 1900 || year > 2100) return [];  // year-sanity gate
  const m = d.getMonth() + 1;
  const day = d.getDate();
  return (VULNERABLE_FAMILIES[m] && VULNERABLE_FAMILIES[m][day]) || [];
}

// Excel serial number to JS Date (Excel epoch is 1899-12-30, matches
// the off-by-one for the 1900-leap-year bug)
function excelSerialToDate(n: number): Date | null {
  if (!isFinite(n) || n <= 0 || n > 2958466) return null;  // > 9999-12-31
  const ms = (n - 25569) * 86400 * 1000;
  return new Date(ms);
}

// RIKEN-shape ID pattern: 7 digits then a letter then 2 digits. If a
// cell of that shape was coerced to a float, the magnitude is >= 1e10
// (e.g. 2310009E13 -> 2.31e+19). We flag any sufficiently-large float
// in an identifier column.
const RIKEN_FLOAT_THRESHOLD = 1e10;

function looksLikeRikenFloat(v: number): boolean {
  return isFinite(v) && Math.abs(v) >= RIKEN_FLOAT_THRESHOLD;
}

// Date-text shapes: "Sep-07", "9-Sep", "Oct-3/4", "2-Oct-14"
const DATE_TEXT_RX = /^([0-9]{1,2})[-/ ]([A-Za-z]{3,4})(?:[-/ ]([0-9]{2,4}))?$|^([A-Za-z]{3,4})[-/ ]([0-9]{1,2})(?:[-/ ]([0-9]{2,4}))?$|^Oct-?3\/4$|^OCT-?3\/4$/;

const MONTH_NAME_TO_NUM: { [name: string]: number } = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
  // Locale variants (Italian / Spanish / Portuguese "Ago" = August)
  ago: 8,
};

function parseDateText(text: string): { month: number; day: number; year: number } | null {
  if (!text) return null;
  const t = text.trim();
  if (t.toLowerCase() === "oct-3/4" || t.toLowerCase() === "oct3/4") {
    return { month: 10, day: 4, year: new Date().getFullYear() };
  }
  const m = DATE_TEXT_RX.exec(t);
  if (!m) return null;
  let day: number, mon: number, year: number;
  if (m[1] && m[2]) {
    day = parseInt(m[1], 10);
    mon = MONTH_NAME_TO_NUM[m[2].toLowerCase()] || 0;
    year = m[3] ? (parseInt(m[3], 10) < 100 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10)) : new Date().getFullYear();
  } else if (m[4] && m[5]) {
    mon = MONTH_NAME_TO_NUM[m[4].toLowerCase()] || 0;
    day = parseInt(m[5], 10);
    year = m[6] ? (parseInt(m[6], 10) < 100 ? 2000 + parseInt(m[6], 10) : parseInt(m[6], 10)) : new Date().getFullYear();
  } else {
    return null;
  }
  if (!mon || !day) return null;
  return { month: mon, day, year };
}

// ---------- Suspicion record ----------

interface Suspicion {
  sheet: string;
  column: string;
  row: number;
  value: string;
  kind: string;
  suggestion: string;
  confidence: number;
  reason: string;
}

// ---------- Main entry point ----------

function main(workbook: ExcelScript.Workbook): void {
  const suspicions: Suspicion[] = [];
  const sheets = workbook.getWorksheets();

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName.toLowerCase() === "uncorrupt-report") continue;  // skip our own output

    const usedRange = sheet.getUsedRange();
    if (!usedRange) continue;

    const values = usedRange.getValues();        // any[][]
    const numberFormats = usedRange.getNumberFormats();
    if (values.length < 2) continue;

    const headers = values[0].map(v => (v == null ? "" : String(v)));
    const idColumnIndices: number[] = [];
    for (let c = 0; c < headers.length; c++) {
      if (isIdentifierColumn(headers[c])) idColumnIndices.push(c);
    }
    if (idColumnIndices.length === 0) continue;

    // First pass: find columns with >= 2 date-corruption cells to enable
    // the column-corroboration band. Single dates in isolation are
    // demoted to low confidence; clusters lift to high confidence.
    const dateHitCount: { [col: number]: number } = {};
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      for (const c of idColumnIndices) {
        const v = row[c];
        if (v instanceof Date && reverseGeneDate(v).length > 0) {
          dateHitCount[c] = (dateHitCount[c] || 0) + 1;
        } else if (typeof v === "string" && parseDateText(v) !== null) {
          const pd = parseDateText(v)!;
          const suggs = reverseGeneDate(new Date(pd.year, pd.month - 1, pd.day));
          if (suggs.length > 0) dateHitCount[c] = (dateHitCount[c] || 0) + 1;
        } else if (typeof v === "number" && Number.isInteger(v) && v > 1 && v < 2958466) {
          const d = excelSerialToDate(v);
          if (d && reverseGeneDate(d).length > 0) {
            dateHitCount[c] = (dateHitCount[c] || 0) + 1;
          }
        }
      }
    }

    // Second pass: emit suspicions
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      for (const c of idColumnIndices) {
        const v = row[c];
        const colHeader = headers[c];
        const corroborated = (dateHitCount[c] || 0) >= 2;

        // Case 1: real datetime cell
        if (v instanceof Date) {
          const suggs = reverseGeneDate(v);
          if (suggs.length > 0) {
            suspicions.push({
              sheet: sheetName,
              column: colHeader,
              row: r + 1,
              value: v.toISOString().slice(0, 10),
              kind: "gene-date",
              suggestion: suggs.join("|"),
              confidence: corroborated ? 0.95 : 0.30,
              reason: corroborated
                ? "datetime cell in identifier column with >= 2 such cells (column-corroborated)"
                : "datetime cell in identifier column (isolated; verify before acting)",
            });
          }
        }

        // Case 2: text-form date
        else if (typeof v === "string" && v.trim()) {
          const pd = parseDateText(v);
          if (pd) {
            const suggs = reverseGeneDate(new Date(pd.year, pd.month - 1, pd.day));
            if (suggs.length > 0) {
              suspicions.push({
                sheet: sheetName,
                column: colHeader,
                row: r + 1,
                value: v,
                kind: "gene-date-string",
                suggestion: suggs.join("|"),
                confidence: corroborated ? 0.95 : 0.30,
                reason: corroborated
                  ? "date-text in identifier column with >= 2 such cells"
                  : "date-text in identifier column (isolated)",
              });
            }
          }
        }

        // Case 3: integer that decodes to a vulnerable date (only flag
        // when the column has other date corruption: otherwise the
        // base rate of false positives on incidental ints is too high)
        else if (typeof v === "number" && Number.isInteger(v) && v > 1 && v < 2958466 && corroborated) {
          const d = excelSerialToDate(v);
          if (d) {
            const suggs = reverseGeneDate(d);
            if (suggs.length > 0) {
              suspicions.push({
                sheet: sheetName,
                column: colHeader,
                row: r + 1,
                value: String(v),
                kind: "gene-date-serial",
                suggestion: suggs.join("|"),
                confidence: 0.60,
                reason: "Excel-serial integer decoding to a vulnerable date in a date-corrupted column",
              });
            }
          }
        }

        // Case 4: RIKEN-shape float (huge magnitude in an identifier column)
        else if (typeof v === "number" && looksLikeRikenFloat(v)) {
          suspicions.push({
            sheet: sheetName,
            column: colHeader,
            row: r + 1,
            value: String(v),
            kind: "id-float",
            suggestion: "(RIKEN-shape ID coerced to float; original symbol lost)",
            confidence: 0.85,
            reason: "magnitude >= 1e10 in identifier column; likely \\d+E\\d+ string coerced to scientific-notation float",
          });
        }
      }
    }
  }

  // ---------- Emit "UnCorrupt-Report" sheet ----------

  let report = workbook.getWorksheet("UnCorrupt-Report");
  if (report) {
    report.delete();
  }
  report = workbook.addWorksheet("UnCorrupt-Report");

  const headerRow = ["sheet", "column", "row", "value", "kind", "suggestion", "confidence", "reason"];
  report.getRange("A1:H1").setValues([headerRow]);

  if (suspicions.length === 0) {
    report.getRange("A2").setValue("No corruption flags found.");
    report.getRange("A2:H2").getFormat().getFill().setColor("#d4edda");
    return;
  }

  const rows = suspicions.map(s => [
    s.sheet, s.column, s.row, s.value, s.kind, s.suggestion, s.confidence, s.reason
  ]);
  const lastRow = rows.length + 1;
  const range = report.getRange(`A2:H${lastRow}`);
  range.setValues(rows);

  // Colour-code confidence
  for (let i = 0; i < suspicions.length; i++) {
    const r = i + 2;
    const c = suspicions[i].confidence;
    const colour = c >= 0.95 ? "#f8d7da"     // high: red
                  : c >= 0.60 ? "#fff3cd"   // moderate: yellow
                  : "#e2e3e5";              // low: grey
    report.getRange(`A${r}:H${r}`).getFormat().getFill().setColor(colour);
  }

  report.getRange("A1:H1").getFormat().getFont().setBold(true);
  report.getRange("A1:H1").getFormat().getFill().setColor("#d1ecf1");
  report.getUsedRange().getFormat().autofitColumns();
}
