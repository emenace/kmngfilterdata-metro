# kmngfilterdata-metro

Lightweight Node.js web application for storing, viewing, cleaning, and exporting employee datasets for **Kantor Kementerian Agama Kota Metro**.

## Overview

The application includes password-only authentication, a mobile-friendly **Employee Data Dashboard** backed by a persistent server-side SQLite database, multi-format data import, and the automated Excel + DUK HTML filter/generator tool:

1. **Password-Only Login (`/login`)**: Protects all pages and API endpoints behind a single configured password (`APP_PASSWORD`) and persistent `HttpOnly` sessions stored in SQLite.
2. **Main Dashboard (`/`)**:
   - Displays all records saved in the server's local SQLite database (`data/kmng_metro.sqlite`, table `pegawai_data`) inside a scrollable, non-resizeable, word-wrapping table designed for both desktop and mobile screens.
   - **Click / tap any cell** to copy its exact value to the clipboard.
   - Top-right **Import** button uploads `.xls`, `.xlsx`, `.csv`, or `.sqlite` files (including files exported by the generator) and persists all rows on the server.
   - Top-right **Generate** button navigates to `/generate`.
3. **Excel + DUK HTML Generator (`/generate`)**:
   - Reads base provincial SIMPEG Excel files (`.xls` BIFF8 or `.xlsx`).
   - Filters Column `X` (`SATKER_3`) for `"Kantor Kementerian Agama Kota Metro"`.
   - Removes all `KODE_*` columns (`KODE_LEVEL_JABATAN`, `KODE_PANGKAT`, `KODE_JABATAN`, `KODE_SATUAN_KERJA`, `KODE_SATKER_2`–`KODE_SATKER_5`, `KODE_GRUP_SATUAN_KERJA`).
   - Optional **"Add DUK HTML support file"** checkbox parses a DUK-SISDM `.html`/`.htm` export, deduplicates by `NIP_BARU` (`NIP` fallback), and appends missing employees without mutating base Excel rows.
   - Prepends sequential Column `A` (`NO`: `1..N`) on `.xlsx` and `.csv` outputs.
   - Generates **3 separate downloadable files** (`.xlsx`, `.csv`, `.sqlite`) verified cell-by-cell with canonical SHA-256 read-back checks (`0 cells altered`).

## Data Integrity Guarantees

- **Zero mutation of base Excel rows**: Existing rows retain exact original values even when DUK merge or dashboard import is used.
- **No numeric rounding or scientific notation**: 18-digit identifiers like `NIP_BARU` (e.g. `199708312025212042`) stay exact strings across XLSX (`t: 's'`), CSV, exported SQLite, and the server's persistent SQLite database (`TEXT NOT NULL`).
- **No date reformatting**: Mixed source date strings (e.g. `01-10-2025` in `TMT_CPNS` and `2025-10-01` in `TMT_JABATAN`) are preserved verbatim.
- **Read-back SHA-256 verification**: Generated `.xlsx`, `.csv`, and `.sqlite` exports are re-read from disk and compared cell-by-cell before serving.

## Prerequisites

- **Node.js `>= 22.5.0`** (tested on `v25.9.0`; required for the built-in `node:sqlite` module).
- **npm** (included with Node.js).

## Quick Start

```bash
# Install dependencies
npm install

# Start the web application (default port: 3000, default password: kemenagmetro)
APP_PASSWORD="your-strong-password" npm start
```

Or run with PM2 using the included script:

```bash
APP_PASSWORD="your-strong-password" ./start-pm2.sh
```

Then open [http://localhost:3000](http://localhost:3000), sign in with the password, and either click **Import** to populate the local SQLite table or **Generate** to filter and export files.

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
| `APP_PASSWORD` | `kemenagmetro` | Password required on the login screen |
| `DB_PATH` | `data/kmng_metro.sqlite` | Path to persistent server-side SQLite database |
| `SESSION_TTL_HOURS` | `5` | Login session cookie/DB lifetime in hours |
| `COOKIE_SECURE` | `false` | Set to `true` when serving over HTTPS |
| `MAX_FILE_SIZE_MB` | `100` | Maximum allowed upload size per file in MB |
| `JOB_TTL_MINUTES` | `60` | Temporary export download retention period before auto-cleanup |

## API Endpoints

| Method | Path | Auth | Description |
| --- | --- | --- | --- |
| `GET` | `/api/health` | Public | Health check returning `{ status: "ok", nodeVersion }` |
| `GET` | `/login` | Public | Password-only login page |
| `GET` | `/api/auth/status` | Public | Returns `{ authenticated: boolean }` |
| `POST` | `/api/auth/login` | Public | Validates `{ password }`, creates SQLite session, sets `HttpOnly` cookie |
| `POST` | `/api/auth/logout` | Public | Revokes current session and clears cookie |
| `GET` | `/` | Session | Main Dashboard page |
| `GET` | `/generate` | Session | Excel + DUK HTML Generator page |
| `GET` | `/api/data?q=...` | Session | Reads persisted SQLite rows (with optional search filter) |
| `POST` | `/api/import` | Session | Imports `.xls`, `.xlsx`, `.csv`, or `.sqlite`/`.db` into `pegawai_data` |
| `POST` | `/api/process` | Session | Filters Excel + optional DUK HTML and generates verified XLSX/CSV/SQLite exports |
| `GET` | `/api/download/:jobId/:format` | Session | Downloads verified `xlsx`, `csv`, or `sqlite` file |

## Project Structure

```text
kmngfilterdata-metro/
├── docs/
│   ├── ARCHITECTURE.md   # Technical architecture, auth, SQLite DB & integrity design
│   └── PRD.md            # Product Requirements Document & exact column schema
├── public/
│   ├── dashboard.js      # Dashboard search, import modal, word-wrap table & click-to-copy
│   ├── generate.js       # Excel + DUK upload, preview pagination & download handlers
│   ├── login.js          # Password-only login form & show/hide toggle
│   └── styles.css        # Mobile-first responsive styling
├── views/
│   ├── dashboard.html    # Main table view with top-right Import, Generate & Logout
│   ├── generate.html     # Excel filter & multi-format exporter view
│   └── login.html        # Password-only login view
├── src/
│   ├── app.js            # Express app wiring, auth routes, import & export endpoints
│   ├── config.js         # Runtime configuration, password, DB path & filter constants
│   ├── server.js         # HTTP entrypoint & graceful shutdown
│   ├── middleware/
│   │   └── auth.js       # Constant-time password check, rate limit & session guards
│   ├── services/
│   │   ├── dataImporter.js       # Multi-format importer (.xls, .xlsx, .csv, .sqlite)
│   │   ├── dukHtmlProcessor.js   # DUK-SISDM HTML parser & NIP_BARU deduplication merger
│   │   ├── excelProcessor.js     # Raw Excel parsing, SATKER_3 filter & KODE_* pruning
│   │   ├── exporters.js          # XLSX (NO + data), RFC-4180 CSV (NO + data) & SQLite
│   │   ├── integrityVerifier.js  # Read-back verification & canonical SHA-256 digest
│   │   ├── jobStore.js           # UUID job storage & TTL directory cleanup
│   │   └── localDb.js            # Persistent server SQLite DB (_sessions, pegawai_data)
│   └── utils/
│       └── csv.js        # Zero-dependency RFC-4180 CSV serializer & parser
├── tests/
│   ├── dbAndImporter.test.js # Persistent SQLite & multi-format importer tests
│   ├── processor.test.js     # Excel + DUK HTML & synthetic edge-case tests
│   └── server.test.js        # Auth, dashboard, import, generate & download HTTP tests
├── start-pm2.sh          # Single-script PM2 launcher
└── package.json
```

## Documentation

- [Product Requirements Document (PRD)](docs/PRD.md)
- [System Architecture & Integrity Design](docs/ARCHITECTURE.md)
