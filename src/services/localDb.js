const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');
const { quoteIdent } = require('./exporters');

/**
 * Persistent server-side SQLite storage for authentication sessions
 * and imported employee table data.
 */
class LocalDatabase {
  constructor({ dbPath, sessionTtlMs }) {
    this.dbPath = path.resolve(dbPath);
    this.sessionTtlMs = sessionTtlMs;
    fs.mkdirSync(path.dirname(this.dbPath), { recursive: true });

    this.db = new DatabaseSync(this.dbPath);
    this.db.exec('PRAGMA journal_mode = WAL;');
    this.db.exec('PRAGMA synchronous = NORMAL;');
    this.initSchema();
  }

  initSchema() {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS _sessions (
        token TEXT PRIMARY KEY,
        created_at INTEGER NOT NULL,
        expires_at INTEGER NOT NULL
      );

      CREATE TABLE IF NOT EXISTS _dataset_meta (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        source_filename TEXT NOT NULL,
        source_format TEXT NOT NULL,
        imported_at TEXT NOT NULL,
        row_count INTEGER NOT NULL,
        columns_json TEXT NOT NULL
      );
    `);
  }

  createSession() {
    const now = Date.now();
    const expiresAt = now + this.sessionTtlMs;
    const token = crypto.randomBytes(32).toString('hex');

    this.pruneExpiredSessions(now);
    this.db
      .prepare('INSERT INTO _sessions (token, created_at, expires_at) VALUES (?, ?, ?);')
      .run(token, now, expiresAt);

    return { token, expiresAt };
  }

  validateSession(token) {
    if (typeof token !== 'string' || token.length < 32) {
      return false;
    }

    const row = this.db
      .prepare('SELECT token, expires_at FROM _sessions WHERE token = ?;')
      .get(token);

    if (!row) {
      return false;
    }

    if (row.expires_at <= Date.now()) {
      this.deleteSession(token);
      return false;
    }

    return true;
  }

  deleteSession(token) {
    if (typeof token === 'string' && token.length > 0) {
      this.db.prepare('DELETE FROM _sessions WHERE token = ?;').run(token);
    }
  }

  pruneExpiredSessions(now = Date.now()) {
    this.db.prepare('DELETE FROM _sessions WHERE expires_at <= ?;').run(now);
  }

  /**
   * Atomically replace the stored employee table and metadata with newly imported data.
   */
  replaceDataset({ sourceFilename, sourceFormat, columns, rows }) {
    if (!Array.isArray(columns) || columns.length === 0) {
      throw new Error('Cannot save dataset without columns.');
    }

    const importedAt = new Date().toISOString();
    const colDefs = columns.map((col) => `${quoteIdent(col)} TEXT NOT NULL`).join(', ');
    const quotedCols = columns.map((col) => quoteIdent(col)).join(', ');
    const placeholders = columns.map(() => '?').join(', ');

    this.db.exec('BEGIN IMMEDIATE;');
    try {
      this.db.exec('DROP TABLE IF EXISTS pegawai_data;');
      this.db.exec(
        `CREATE TABLE pegawai_data (_row_id INTEGER PRIMARY KEY AUTOINCREMENT, ${colDefs});`
      );

      const insertStmt = this.db.prepare(
        `INSERT INTO pegawai_data (${quotedCols}) VALUES (${placeholders});`
      );

      for (let r = 0; r < rows.length; r += 1) {
        const srcRow = rows[r];
        const values = new Array(columns.length);
        for (let c = 0; c < columns.length; c += 1) {
          const val = srcRow[c];
          values[c] = val === undefined || val === null ? '' : String(val);
        }
        insertStmt.run(...values);
      }

      this.db
        .prepare(
          `INSERT OR REPLACE INTO _dataset_meta
            (id, source_filename, source_format, imported_at, row_count, columns_json)
           VALUES (1, ?, ?, ?, ?, ?);`
        )
        .run(
          String(sourceFilename || 'imported'),
          String(sourceFormat || 'unknown'),
          importedAt,
          rows.length,
          JSON.stringify(columns)
        );

      this.db.exec('COMMIT;');
    } catch (err) {
      this.db.exec('ROLLBACK;');
      throw err;
    }

    return {
      sourceFilename,
      sourceFormat,
      importedAt,
      rowCount: rows.length,
      columns,
    };
  }

  /**
   * Retrieve all stored rows (or rows matching an optional case-insensitive substring query).
   */
  getDataset({ search = '' } = {}) {
    const metaRow = this.db
      .prepare(
        `SELECT source_filename, source_format, imported_at, row_count, columns_json
         FROM _dataset_meta WHERE id = 1;`
      )
      .get();

    if (!metaRow) {
      return {
        hasData: false,
        meta: null,
        columns: [],
        rows: [],
        totalRows: 0,
        filteredRows: 0,
      };
    }

    const columns = JSON.parse(metaRow.columns_json);
    const quotedCols = columns.map((col) => quoteIdent(col)).join(', ');
    const query = String(search || '').trim();

    let records;
    if (query.length > 0 && columns.length > 0) {
      const escaped = query.replace(/[\\%_]/g, (ch) => `\\${ch}`);
      const likeParam = `%${escaped}%`;
      const whereClause = columns
        .map((col) => `${quoteIdent(col)} LIKE ? ESCAPE '\\'`)
        .join(' OR ');
      const params = columns.map(() => likeParam);

      records = this.db
        .prepare(
          `SELECT ${quotedCols} FROM pegawai_data WHERE (${whereClause}) ORDER BY _row_id ASC;`
        )
        .all(...params);
    } else {
      records = this.db
        .prepare(`SELECT ${quotedCols} FROM pegawai_data ORDER BY _row_id ASC;`)
        .all();
    }

    const rows = records.map((rec) =>
      columns.map((col) => (rec[col] === undefined || rec[col] === null ? '' : String(rec[col])))
    );

    return {
      hasData: true,
      meta: {
        sourceFilename: metaRow.source_filename,
        sourceFormat: metaRow.source_format,
        importedAt: metaRow.imported_at,
        rowCount: metaRow.row_count,
      },
      columns,
      rows,
      totalRows: metaRow.row_count,
      filteredRows: rows.length,
    };
  }

  close() {
    if (this.db) {
      this.db.close();
      this.db = null;
    }
  }
}

module.exports = {
  LocalDatabase,
};

