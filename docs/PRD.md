# Product Requirements Document (PRD): Kemenag Kota Metro Data Portal

## 1. Product Overview & Background

Every month, **Kantor Kementerian Agama Kota Metro** receives a province-wide employee Excel workbook exported from SIMPEG (e.g., `02090000000000-3.xls`, containing 8,066 employee rows across 16 work units in Provinsi Lampung and 36 columns). Local **DUK-SISDM HTML** reports sometimes also list newly appointed staff (such as PPPK Paruh Waktu) who are not yet included in the provincial Excel export.

**`kmngfilterdata-metro`** is a lightweight, self-hosted Node.js web application that:
1. Protects employee data behind a **password-only login screen**.
2. Provides a **Main Dashboard** backed by a **persistent server-side SQLite database**, allowing staff to search, read, and copy any employee cell easily on desktop or mobile phones.
3. Lets staff **Import** generated `.xlsx`, `.xls`, `.csv`, or `.sqlite` files directly into the persistent SQLite database, or click **Generate** to access the automated Excel + DUK HTML filtering & export tool.

## 2. Target Users & Workflows

- **Primary Users**: Kepegawaian / IT / administrative staff at Kantor Kementerian Agama Kota Metro using desktop browsers or mobile phones.
- **Core Workflows**:
  1. **Sign In (`/login`)**: Enter the shared application password (no username required) to unlock the app.
  2. **Browse & Copy (`/`)**: View all persisted rows from the local SQLite database in a word-wrapped, scrollable table; search across any column; click/tap any cell to copy its value to the clipboard.
  3. **Generate Filtered Files (`/generate`)**: Click **Generate** on the top right of the dashboard, upload the monthly provincial `.xls`/`.xlsx` (plus optional DUK-SISDM `.html`), verify integrity (`0 cells altered`), and download **XLSX**, **CSV**, or **SQLite**.
  4. **Import to Dashboard (`/`)**: Click **Import** on the top right of the dashboard, upload an `.xls`, `.xlsx`, `.csv`, or `.sqlite` file, and save all rows into the server's local SQLite database (`data/kmng_metro.sqlite`).

## 3. Functional Requirements

| ID | Requirement | Details |
| --- | --- | --- |
| **FR-1** | **Excel File Upload** | Accept `.xls` (BIFF8) and `.xlsx` files via drag-and-drop or file picker on `/generate`. Reject unsupported extensions with HTTP `400`. |
| **FR-2** | **Automatic `SATKER_3` Row Filtering** | Locate header column `SATKER_3` (Column `X` in standard exports) and retain only data rows where `SATKER_3` contains `"Kantor Kementerian Agama Kota Metro"` (case-insensitive substring match). Preserve original row order. |
| **FR-3** | **Automatic `KODE_*` Column Removal** | Remove every column whose header starts with `KODE_` (case-insensitive `/^KODE_/i`) while preserving the exact column order of all remaining columns. |
| **FR-4** | **Three Output Formats** | Generate 3 separate output files from the exact same filtered/combined dataset: `.xlsx` (Excel OpenXML), `.csv` (RFC 4180 UTF-8 with BOM), and `.sqlite` (single-file SQLite 3 database). |
| **FR-5** | **Three Separate Download Buttons** | Provide 3 independent download buttons on `/generate` (one each for XLSX, CSV, and SQLite) displaying filename and byte size, served via dedicated download URLs. |
| **FR-6** | **Automated Data Integrity Verification** | Read back all 3 generated files from disk after export, confirm `0 cells altered` and matching canonical SHA-256 checksums across all formats, and surface the verification badge and digest in the UI. |
| **FR-7** | **Optional DUK HTML Support File Merge** | Provide a UI checkbox to enable a secondary `.html`/`.htm` upload field for DUK-SISDM HTML exports. Deduplicate against filtered base Excel rows by `NIP_BARU` (`NIP` fallback) and append missing people mapped to the 27 retained columns without modifying any base Excel row. |
| **FR-8** | **Sequential Row Numbering Column (`NO`) in XLSX & CSV** | Prepend Column `A` titled `NO` containing sequential iteration numbers `1` through `N` in generated `.xlsx` (numeric cells) and `.csv` files. SQLite keeps the 27 data columns. |
| **FR-9** | **Password-Only Login & Session Auth** | Protect `/`, `/generate`, and all data/import/process/download APIs behind a password-only login form (`/login`). Verify passwords in constant time (`crypto.timingSafeEqual`), issue an `HttpOnly; SameSite=Lax` cookie backed by SQLite `_sessions`, rate-limit failed attempts, and support `/api/auth/logout`. |
| **FR-10** | **Main Dashboard & Persistent Server SQLite Storage** | After login, render the Main Dashboard (`/`) showing top-right **Import**, **Generate**, and **Logout** buttons and all rows stored in the server's persistent SQLite database (`data/kmng_metro.sqlite`). Data and sessions survive server restarts. |
| **FR-11** | **Multi-Format Dashboard Import (XLS/XLSX/CSV/SQLite)** | Allow users to upload `.xls`, `.xlsx`, `.csv`, `.sqlite`, or `.db` files via the dashboard **Import** modal. Strip any leading `NO` column automatically so generated `.xlsx`/`.csv`/`.sqlite` files populate the same normalized table (`pegawai_data`) atomically. |
| **FR-12** | **Scrollable Non-Resizeable Word-Wrapped Click-to-Copy Table & Mobile UX** | Render dashboard table with fixed column widths (`resize: none`), smooth horizontal/vertical scrolling, multi-line word wrapping (`white-space: normal; overflow-wrap: anywhere`), sticky headers, click/tap-to-copy on every non-empty cell with clipboard feedback toast, and responsive mobile phone sizing (`16px` search/password inputs, `44px+` touch targets). |

## 4. Non-Functional Requirements

- **Zero Data Mutation**: No cell value in any retained base Excel row or imported row may be rounded, trimmed, type-coerced, or reformatted:
  - 18-digit `NIP_BARU` strings remain exact strings (`TEXT NOT NULL` in SQLite, `t: 's'` in XLSX).
  - Mixed date formats (`DD-MM-YYYY` and `YYYY-MM-DD`) remain verbatim strings.
  - Negative integer strings (`-2`, `-3`), empty cells (`''`), leading zeros, and commas remain unchanged.
- **Lightweight Footprint**: Vanilla HTML/CSS/JS frontend (zero build step, zero external CDNs), Node.js `>= 22.5.0` using built-in `node:sqlite`, `node:crypto`, and `node:test`, and only 3 npm runtime dependencies (`express`, `multer`, `xlsx`).
- **Security & Isolation**: HTML pages served from `views/` only after session validation; `HttpOnly` session cookies; constant-time password verification; login rate limiting; UUID-validated download job IDs; XSS-safe DOM rendering (`textContent`).
- **Offline & Mobile Usability**: Works completely offline on local LAN with responsive layouts and clipboard fallback over HTTP.

## 5. Data Schema Specification

Verified against the provincial sample workbook (`02090000000000-3.xls`, worksheet `Pegawai`, 8,066 data rows total, 360 rows for `Kantor Kementerian Agama Kota Metro`).

### 5.1 Base Excel Columns & Disposition (36 Source Columns -> 27 Kept, 9 Removed)

| Col | Column Header | Disposition | Notes |
| --- | --- | --- | --- |
| `A` | `NIP` | **Kept** | Legacy employee ID string |
| `B` | `NIP_BARU` | **Kept** | 18-digit identifier stored strictly as text |
| `C` | `NAMA_LENGKAP` | **Kept** | Full name with academic titles |
| `D` | `AGAMA` | **Kept** | Religion |
| `E` | `PENDIDIKAN` | **Kept** | Full education description |
| `F` | `JENJANG_PENDIDIKAN` | **Kept** | Education level (may be empty `''`) |
| `G` | `KODE_LEVEL_JABATAN` | **Removed** | Matches `KODE_*` |
| `H` | `LEVEL_JABATAN` | **Kept** | Position level label |
| `I` | `KODE_PANGKAT` | **Removed** | Matches `KODE_*` |
| `J` | `PANGKAT` | **Kept** | Rank title (empty string `''` for PPPK rows) |
| `K` | `GOL_RUANG` | **Kept** | Rank grade |
| `L` | `TMT_CPNS` | **Kept** | `DD-MM-YYYY` string format |
| `M` | `TMT_PANGKAT` | **Kept** | `DD-MM-YYYY` string format |
| `N` | `TIPE_JABATAN` | **Kept** | Position type |
| `O` | `KODE_JABATAN` | **Removed** | Matches `KODE_*` |
| `P` | `TAMPIL_JABATAN` | **Kept** | Display position title |
| `Q` | `TMT_JABATAN` | **Kept** | `YYYY-MM-DD` string format |
| `R` | `KODE_SATUAN_KERJA` | **Removed** | Matches `KODE_*` |
| `S` | `SATKER_1` | **Kept** | Unit level 1 |
| `T` | `KETERANGAN_SATUAN_KERJA` | **Kept** | Full work unit description |
| `U` | `KODE_SATKER_2` | **Removed** | Matches `KODE_*` |
| `V` | `SATKER_2` | **Kept** | Unit level 2 |
| `W` | `KODE_SATKER_3` | **Removed** | Matches `KODE_*` |
| `X` | `SATKER_3` | **Kept (Filter Key)** | Must contain `"Kantor Kementerian Agama Kota Metro"` |
| `Y` | `KODE_SATKER_4` | **Removed** | Matches `KODE_*` |
| `Z` | `SATKER_4` | **Kept** | Provincial office (`Kanwil Kementerian Agama Provinsi Lampung`) |
| `AA` | `KODE_SATKER_5` | **Removed** | Matches `KODE_*` |
| `AB` | `SATKER_5` | **Kept** | Unit level 5 |
| `AC` | `KODE_GRUP_SATUAN_KERJA` | **Removed** | Matches `KODE_*` |
| `AD` | `GRUP_SATUAN_KERJA` | **Kept** | Work unit group |
| `AE` | `STATUS_KAWIN` | **Kept** | Marital status |
| `AF` | `STATUS_PEGAWAI` | **Kept** | Employment status (e.g. `PNS`, `PPPK`, `PPPK PARUH WAKTU`) |
| `AG` | `USIA_PENSIUN` | **Kept** | Retirement age |
| `AH` | `TMT_PENSIUN` | **Kept** | `YYYY-MM-DD` string format |
| `AI` | `MASA_KERJA_TAHUN` | **Kept** | Years of service (may be negative e.g. `-2`) |
| `AJ` | `MASA_KERJA_BULAN` | **Kept** | Months of service (may be negative e.g. `-3`) |

> **Output Column A (`NO`)**: On generated `.xlsx` and `.csv` exports (and in the Dashboard table view), Column `A` is `NO` (`1`..`N`).

### 5.2 Server Persistent SQLite Schema (`data/kmng_metro.sqlite`)

- `_sessions`: `token TEXT PRIMARY KEY, created_at INTEGER NOT NULL, expires_at INTEGER NOT NULL`
- `_dataset_meta`: `id INTEGER PRIMARY KEY CHECK (id = 1), source_filename TEXT NOT NULL, source_format TEXT NOT NULL, imported_at TEXT NOT NULL, row_count INTEGER NOT NULL, columns_json TEXT NOT NULL`
- `pegawai_data`: `_row_id INTEGER PRIMARY KEY AUTOINCREMENT`, followed by one `TEXT NOT NULL` column per imported header in exact original order.

## 6. Acceptance Criteria

- Unauthenticated requests to `/` or `/generate` redirect `302` to `/login`; unauthenticated API calls return `401`.
- Submitting the configured password creates a persistent SQLite session and redirects to `/` where **Import** and **Generate** buttons appear at the top right.
- Importing a generated `.xlsx`, `.csv`, or `.sqlite` file (or `.xls`) populates `pegawai_data` on disk, persists across server restarts, and displays all rows in a scrollable, non-resizeable, word-wrapped, click-to-copy table usable on desktop and mobile phones.
- `/generate` preserves full Excel filtering (`360` Kota Metro rows from `02090000000000-3.xls`), optional DUK HTML merge, `NO` (`1..N`) in XLSX/CSV, and `0 cells altered` SHA-256 integrity verification across all 3 download formats.

## 7. Out of Scope

- Multi-user accounts / per-user role permissions (single shared password per `APP_PASSWORD`).
- Inline cell editing inside the dashboard table (all cell contents are read-only and copy-only to preserve source data integrity).
