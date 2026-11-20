# Product Requirements Document (PRD): Kemenag Kota Metro Excel Data Filter

## 1. Product Overview & Background

Every month, **Kantor Kementerian Agama Kota Metro** receives a province-wide employee Excel workbook exported from SIMPEG (e.g., `02090000000000-3.xls`, containing 8,066 employee rows across 16 work units in Provinsi Lampung and 36 columns). In addition, local **DUK-SISDM HTML** reports (`November DUK-SISDM Kementerian Agama.html`, 373 rows) sometimes include newly appointed staff (such as PPPK Paruh Waktu) who are not yet present in the provincial Excel export.

Manually opening the `.xls` workbook, filtering `SATKER_3`, deleting 9 internal `KODE_*` columns, cross-referencing DUK HTML rows to append missing personnel, numbering rows, and converting into multiple formats is error-prone and introduces severe spreadsheet data-corruption risks (notably Excel truncating 18-digit `NIP_BARU` identifiers into scientific notation or altering date formats).

**`kmngfilterdata-metro`** is a lightweight, self-hosted Node.js web application that automates the entire cleanup, filtering, optional DUK HTML merge, sequential numbering, multi-format export, and integrity verification workflow.

## 2. Target Users & Workflow

- **Primary Users**: Kepegawaian / IT / administrative staff at Kantor Kementerian Agama Kota Metro.
- **Monthly Workflows**:
  1. **Base Excel Only**: Open `http://localhost:3000`, drop or select the monthly `.xls`/`.xlsx` file; processing starts immediately.
  2. **Base Excel + DUK HTML Support File**: Check **"Add DUK HTML support file"**, select both the base Excel file and the DUK-SISDM `.html` file, and click **"Process & Combine Files"**.
  3. Review KPI metrics, newly added DUK personnel (if enabled), the SHA-256 verification banner, and the searchable preview table.
  4. Click any of the **3 separate Download buttons** (**XLSX**, **CSV**, or **SQLite**) as needed.

## 3. Functional Requirements

| ID | Requirement | Details |
| --- | --- | --- |
| **FR-1** | **Excel File Upload** | Accept `.xls` (BIFF8) and `.xlsx` files via drag-and-drop or file picker. Reject unsupported extensions with HTTP `400`. |
| **FR-2** | **Automatic `SATKER_3` Row Filtering** | Locate header column `SATKER_3` (Column `X` in standard exports) and retain only data rows where `SATKER_3` contains `"Kantor Kementerian Agama Kota Metro"` (case-insensitive substring match). Preserve original row order. |
| **FR-3** | **Automatic `KODE_*` Column Removal** | Remove every column whose header starts with `KODE_` (case-insensitive `/^KODE_/i`) while preserving the exact column order of all remaining columns. |
| **FR-4** | **Three Output Formats** | Generate 3 separate output files from the exact same filtered/combined dataset: `.xlsx` (Excel OpenXML), `.csv` (RFC 4180 UTF-8 with BOM), and `.sqlite` (single-file SQLite 3 database). |
| **FR-5** | **Three Separate Download Buttons** | Provide 3 independent download buttons in the UI (one each for XLSX, CSV, and SQLite) displaying filename and byte size, served via dedicated download URLs. |
| **FR-6** | **Automated Data Integrity Verification** | Read back all 3 generated files from disk after export, confirm `0 cells altered` and matching canonical SHA-256 checksums across all formats, and surface the verification badge and digest in the UI. |
| **FR-7** | **Optional DUK HTML Support File Merge** | Provide a UI checkbox to enable a secondary `.html`/`.htm` upload field for DUK-SISDM HTML exports. When checked, parse all employee records from the DUK HTML table, deduplicate against filtered base Excel rows by `NIP_BARU` (with `NIP` fallback), and append only missing people mapped to the 27 retained columns without modifying any base Excel row. |
| **FR-8** | **Sequential Row Numbering Column (`NO`) in XLSX & CSV** | Prepend Column `A` titled `NO` containing sequential iteration numbers `1` through `N` (where `N` is the total filtered/combined row count) in generated `.xlsx` (numeric cells) and `.csv` files. SQLite keeps the 27 data columns (`rowid` is native). |

## 4. Non-Functional Requirements

- **Zero Data Mutation**: No cell value in any retained base Excel row may be rounded, trimmed, type-coerced, or reformatted during filtering, merging, or export. Specifically:
  - 18-digit `NIP_BARU` strings (exceeding IEEE-754 `Number.MAX_SAFE_INTEGER`) remain exact strings.
  - Mixed date formats (`DD-MM-YYYY` and `YYYY-MM-DD`) remain verbatim strings.
  - Negative integer strings (`-2`, `-3`), empty cells (`''`), leading zeros, and commas remain unchanged.
- **Lightweight Footprint**: Vanilla HTML/CSS/JS frontend (zero build step, zero external CDNs), Node.js `>= 22.5.0` using built-in `node:sqlite`, `node:crypto`, and `node:test`, and only 3 npm runtime dependencies (`express`, `multer`, `xlsx`).
- **Performance**: Filtering an 8,000+ row `.xls` workbook, merging DUK HTML rows, exporting all 3 formats, and verifying every cell completes in under a few seconds on commodity hardware.
- **Security & Isolation**: UUID-validated job IDs (`crypto.randomUUID()`), strict file extension and size checks (`100 MB` default limit), automatic TTL cleanup of temporary job files, and XSS-safe DOM rendering (`textContent`).
- **Offline Capability**: Runs completely locally with no external network calls required at runtime.

## 5. Data Schema Specification

Verified against the attached provincial sample workbook (`02090000000000-3.xls`, worksheet `Pegawai`, 8,066 data rows total, 360 rows for `Kantor Kementerian Agama Kota Metro`) and DUK support file (`November DUK-SISDM Kementerian Agama.html`, 373 rows total, 13 missing from base Excel):

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

> **Output Column A (`NO`)**: On generated `.xlsx` and `.csv` exports, Column `A` is `NO` (`1`..`N`), shifting the 27 data columns above to Columns `B`..`AB` (28 total columns).

### 5.2 DUK HTML -> Retained Excel Column Mapping (For Newly Added Rows)

| Target Column | Source in DUK-SISDM HTML | Notes |
| --- | --- | --- |
| `NIP` / `NIP_BARU` | Col 2 (`NAMA`) line 2 (`NIP / NIP_BARU`) | Deduplication key (`NIP_BARU` primary, `NIP` fallback) |
| `NAMA_LENGKAP` | Col 2 (`NAMA`) line 1 | Exact text |
| `PENDIDIKAN` / `JENJANG_PENDIDIKAN` | Cols 9, 10, 11 | Combined format `<JENJANG> <NAMA> Tahun <THN>` and `JENJANG` |
| `PANGKAT` / `GOL_RUANG` | Col 3 (`GOL/ RUANG`) | Lookup for PNS grades (`I/a`–`IV/e`); `''` for `LAIN-LAIN` / PPPK |
| `TMT_PANGKAT` | Col 4 (`PANGKAT TMT`) | Preserved as `DD-MM-YYYY` |
| `TIPE_JABATAN`, `TAMPIL_JABATAN`, `KETERANGAN_SATUAN_KERJA`, `SATKER_1`, `SATKER_2`, `STATUS_PEGAWAI` | Col 5 (`JABATAN NAMA`) + `GOL/ RUANG` | Extracts `(Jabatan ...)` type, position title, satker unit after `pada`, and detects `(PPPK PARUH WAKTU)` |
| `TMT_JABATAN` / `TMT_PENSIUN` | Col 6 / Col 13 | Normalized from `DD-MM-YYYY` to `YYYY-MM-DD` to match base Excel |
| `SATKER_3`, `SATKER_4`, `SATKER_5` | Target Satker & Base Excel uniform values | `"Kantor Kementerian Agama Kota Metro"` + provincial Kanwil values |
| `USIA_PENSIUN` | Derived from Col 12 (`TGL LAHIR`) & Col 13 (`TMT PENSIUN`) | `pensionYear - birthYear - (pensionMonth < birthMonth ? 1 : 0)` |
| `MASA_KERJA_TAHUN` / `MASA_KERJA_BULAN` | Col 7 / Col 8 | Exact strings |
| `AGAMA`, `LEVEL_JABATAN`, `TMT_CPNS`, `GRUP_SATUAN_KERJA`, `STATUS_KAWIN` | Not in DUK HTML | Left as `''` |

### 5.3 SQLite Table Schema

The SQLite output uses a sanitized lowercase table name derived from the worksheet name (e.g. `pegawai` for sheet `Pegawai`), `PRAGMA journal_mode = DELETE` (single portable `.sqlite` file with no `-wal`/`-shm` sidecars), and `TEXT NOT NULL` columns for the 27 retained data headers in original order.

## 6. Acceptance Criteria

- **Base Excel Only**: Uploading `02090000000000-3.xls` returns HTTP `200` with `totalSourceRows = 8066`, `filteredRowCount = 360`, `removedColumnCount = 9`, `outputColumnCount = 27`, `NO` (`1..360`) in XLSX/CSV (28 cols), 27 cols in SQLite, and `integrity.verified = true` (`alteredCells = 0`).
- **Base Excel + DUK HTML Combined**: Uploading `02090000000000-3.xls` with `November DUK-SISDM Kementerian Agama.html` returns HTTP `200` with `baseExcelFilteredRows = 360`, `addedFromDukCount = 13`, `filteredRowCount = 373`, untouched first 360 rows, `NO` (`1..373`) in XLSX/CSV (28 cols), 373 rows in SQLite (27 cols), and `integrity.verified = true` (`alteredCells = 0`).
- Invalid extensions, missing `SATKER_3`, missing `dukFile` when `includeDuk=true`, or non-DUK HTML content return clear HTTP `400` error messages.

## 7. Out of Scope

- User authentication / multi-tenant roles (intended for local or trusted internal office deployment).
- Persistent database storage of uploaded employee PII beyond temporary TTL download jobs (`60` minutes default).
- In-browser editing or manual cell overrides.
