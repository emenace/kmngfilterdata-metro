# kmngfilterdata-metro

Lightweight Node.js web application for cleaning, filtering, optionally merging DUK HTML support data, and exporting monthly employee datasets for **Kantor Kementerian Agama Kota Metro**.

## Overview

Every month, provincial SIMPEG Excel reports (`.xls` or `.xlsx`) list employees across all regencies/cities in Lampung Province. In addition, some newly appointed employees (such as PPPK Paruh Waktu) may appear in the local **DUK-SISDM HTML** report before appearing in the provincial Excel sheet. This tool automates the full cleanup and merge pipeline:

1. **Reads** the uploaded base Excel file (supporting both legacy BIFF8 `.xls` and modern `.xlsx` workbooks).
2. **Locates** column `X` titled `SATKER_3`.
3. **Filters rows** keeping only rows where `SATKER_3` contains `"Kantor Kementerian Agama Kota Metro"`.
4. **Removes columns** whose header starts with `KODE_` (e.g., `KODE_LEVEL_JABATAN`, `KODE_PANGKAT`, `KODE_JABATAN`, `KODE_SATUAN_KERJA`, `KODE_SATKER_2`–`KODE_SATKER_5`, `KODE_GRUP_SATUAN_KERJA`).
5. **Optional DUK HTML Support File Merge**: When the **"Add DUK HTML support file"** checkbox is enabled, parses the uploaded DUK-SISDM `.html`/`.htm` file, checks each person by `NIP_BARU` (with `NIP` fallback), and appends anyone not yet present in the filtered base Excel rows without altering any base Excel row.
6. **Adds Sequential Column A (`NO`) in XLSX & CSV**: Prepends column `A` titled `NO` numbered `1` to `N` on generated `.xlsx` and `.csv` outputs.
7. **Exports 3 separate files** for immediate individual download:
   - **XLSX** (`.xlsx`, `NO` + 27 data columns)
   - **CSV** (`.csv`, RFC 4180 + UTF-8 BOM, `NO` + 27 data columns)
   - **SQLite** (`.sqlite`, single-file database using Node's built-in `node:sqlite`, 27 data columns)
8. **Verifies data integrity** automatically before returning results: all 3 files are read back from disk and compared cell-by-cell (`0 cells altered`, identical SHA-256 canonical digest, plus `NO` = `1..N` verified in XLSX & CSV).

## Data Integrity Guarantees

- **Zero mutation of base Excel rows**: Existing rows retain exact original values even when DUK merge is enabled.
- **No numeric rounding or scientific notation**: 18-digit identifiers like `NIP_BARU` (e.g. `199708312025212042`) stay exact strings across XLSX (`t: 's'`), CSV, and SQLite (`TEXT NOT NULL`).
- **No date reformatting**: Mixed source date strings (e.g. `01-10-2025` in `TMT_CPNS` and `2025-10-01` in `TMT_JABATAN`) are preserved verbatim.
- **No whitespace, sign, or punctuation mutation**: Trailing spaces, negative values (`-2`, `-3`), empty strings, and commas (`SHOLEHA, S.Pd.`) remain unchanged.
- **Read-back SHA-256 verification**: Any cell mismatch or missing row aborts the job instead of serving corrupted output.

## Prerequisites

- **Node.js `>= 22.5.0`** (tested on `v25.9.0`; required for the built-in `node:sqlite` module).
- **npm** (included with Node.js).

## Quick Start

```bash
# Install dependencies
npm install

# Start the web application (default port: 3000)
npm start
```

Then open [http://localhost:3000](http://localhost:3000) in your browser:
- **Excel only**: Drop or select your `.xls`/`.xlsx` file — processing starts automatically.
- **Excel + DUK HTML support file**: Check **"Add DUK HTML support file"**, upload both the base Excel file and the DUK `.html`/`.htm` file, and click **"Process & Combine Files"**.

### Development & Testing

```bash
# Run with file-watch auto-restart
npm run dev

# Run the automated unit & HTTP integration tests
npm test
```

### Environment Variables

| Variable | Default | Description |
| --- | --- | --- |
| `PORT` | `3000` | HTTP server port |
| `MAX_FILE_SIZE_MB` | `100` | Maximum allowed upload size per file in MB |
| `JOB_TTL_MINUTES` | `60` | Temporary export file retention period before auto-cleanup |

## API Endpoints

| Method | Path | Description |
| --- | --- | --- |
| `GET` | `/api/health` | Health check returning `{ status: "ok", nodeVersion }` |
| `POST` | `/api/process` | Multipart upload (`file`, optional `includeDuk=true`, optional `dukFile`), filters/merges data, generates and verifies all 3 exports, returns stats, DUK merge summary, integrity checksum, download URLs, and preview rows |
| `GET` | `/api/download/:jobId/:format` | Downloads the verified `xlsx`, `csv`, or `sqlite` export for a given `jobId` |

## Project Structure

```text
kmngfilterdata-metro/
├── docs/
│   ├── ARCHITECTURE.md   # Technical architecture, pipeline diagrams & integrity design
│   └── PRD.md            # Product Requirements Document & exact column schema
├── public/
│   ├── app.js            # Vanilla JS upload, DUK toggle, preview & safe DOM rendering
│   ├── index.html        # Single-page UI with 3 separate download buttons
│   └── styles.css        # Responsive, zero-CDN styling
├── src/
│   ├── app.js            # Express app wiring, dual-file upload & HTTP endpoints
│   ├── config.js         # Runtime configuration & filter rule constants
│   ├── server.js         # HTTP entrypoint & graceful shutdown
│   ├── services/
│   │   ├── dukHtmlProcessor.js   # DUK-SISDM HTML parser & NIP_BARU deduplication merger
│   │   ├── excelProcessor.js     # Raw Excel parsing, SATKER_3 filter & KODE_* pruning
│   │   ├── exporters.js          # XLSX (NO + data), RFC-4180 CSV (NO + data) & SQLite
│   │   ├── integrityVerifier.js  # Read-back verification & canonical SHA-256 digest
│   │   └── jobStore.js           # UUID job storage & TTL directory cleanup
│   └── utils/
│       └── csv.js        # Zero-dependency RFC-4180 CSV serializer & parser
├── tests/
│   ├── processor.test.js # Real Excel + DUK HTML & synthetic edge-case tests
│   └── server.test.js    # End-to-end HTTP upload & 3-format download tests
└── package.json
```

## Documentation

- [Product Requirements Document (PRD)](docs/PRD.md)
- [System Architecture & Integrity Design](docs/ARCHITECTURE.md)
