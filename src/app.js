const fs = require('node:fs');
const path = require('node:path');
const express = require('express');
const multer = require('multer');
const { resolveConfig } = require('./config');
const { ValidationError, processExcelBuffer } = require('./services/excelProcessor');
const { parseDukHtmlBuffer, mergeDukIntoExcelResult } = require('./services/dukHtmlProcessor');
const { parseImportBuffer } = require('./services/dataImporter');
const {
  sanitizeFileBase,
  sanitizeTableName,
  exportXlsx,
  exportCsv,
  exportSqlite,
} = require('./services/exporters');
const { verifyOutputs } = require('./services/integrityVerifier');
const { JobStore } = require('./services/jobStore');
const { LocalDatabase } = require('./services/localDb');
const {
  verifyPassword,
  buildSessionCookie,
  createAuthHelpers,
} = require('./middleware/auth');

const FORMAT_MIME_TYPES = Object.freeze({
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  csv: 'text/csv; charset=utf-8',
  sqlite: 'application/vnd.sqlite3',
});

const PREVIEW_ROW_LIMIT = 500;

/**
 * Create and configure the Express application with injected config.
 */
function createApp(customConfig = {}) {
  const cfg = resolveConfig(customConfig);
  const jobStore = new JobStore({
    tmpDir: cfg.TMP_DIR,
    ttlMs: cfg.JOB_TTL_MS,
  });
  const localDb = new LocalDatabase({
    dbPath: cfg.DB_PATH,
    sessionTtlMs: cfg.SESSION_TTL_MS,
  });
  const auth = createAuthHelpers({ config: cfg, localDb });

  const app = express();
  app.locals.config = cfg;
  app.locals.jobStore = jobStore;
  app.locals.localDb = localDb;

  const publicDir = path.resolve(__dirname, '..', 'public');
  const viewsDir = path.resolve(__dirname, '..', 'views');

  app.use(express.static(publicDir, { index: false }));
  app.use(express.json({ limit: '1mb' }));

  const uploadProcess = multer({
    storage: multer.memoryStorage(),
    limits: {
      fileSize: cfg.MAX_FILE_SIZE_BYTES,
      files: 2,
    },
    fileFilter: (_req, file, cb) => {
      const ext = path.extname(file.originalname || '').toLowerCase();
      if (file.fieldname === 'file') {
        if (!cfg.ALLOWED_EXTENSIONS.includes(ext)) {
          cb(
            new ValidationError(
              `Unsupported base Excel extension "${ext || '(none)'}". Please upload a .xls or .xlsx file.`
            )
          );
          return;
        }
        cb(null, true);
        return;
      }

      if (file.fieldname === 'dukFile') {
        if (!cfg.ALLOWED_HTML_EXTENSIONS.includes(ext)) {
          cb(
            new ValidationError(
              `Unsupported DUK file extension "${ext || '(none)'}". Please upload an .html or .htm file.`
            )
          );
          return;
        }
        cb(null, true);
        return;
      }

      cb(new ValidationError(`Unexpected upload field "${file.fieldname}".`));
    },
  });

  const uploadImport = multer({
    storage: multer.memoryStorage(),
    limits: {
      fileSize: cfg.MAX_FILE_SIZE_BYTES,
      files: 1,
    },
    fileFilter: (_req, file, cb) => {
      const ext = path.extname(file.originalname || '').toLowerCase();
      if (!cfg.ALLOWED_IMPORT_EXTENSIONS.includes(ext)) {
        cb(
          new ValidationError(
            `Unsupported import file extension "${ext || '(none)'}". Allowed formats: .xls, .xlsx, .csv, .sqlite, .db.`
          )
        );
        return;
      }
      cb(null, true);
    },
  });

  // Public health & authentication routes
  app.get('/api/health', (_req, res) => {
    res.json({
      status: 'ok',
      nodeVersion: process.version,
    });
  });

  app.get('/login', (req, res) => {
    if (auth.isAuthenticated(req)) {
      res.redirect(302, '/');
      return;
    }
    res.sendFile(path.join(viewsDir, 'login.html'));
  });

  app.get('/api/auth/status', (req, res) => {
    res.json({
      authenticated: auth.isAuthenticated(req),
    });
  });

  app.post('/api/auth/login', (req, res) => {
    if (auth.isRateLimited(req)) {
      res.status(429).json({ error: 'Too many failed login attempts. Please wait a few minutes and try again.' });
      return;
    }

    const password = req.body && typeof req.body.password === 'string' ? req.body.password : '';
    if (!password) {
      res.status(400).json({ error: 'Please enter your password.' });
      return;
    }

    if (!verifyPassword(password, cfg.APP_PASSWORD)) {
      auth.recordFailedAttempt(req);
      res.status(401).json({ error: 'Invalid password. Please try again.' });
      return;
    }

    auth.clearFailedAttempts(req);
    const { token } = localDb.createSession();
    res.setHeader(
      'Set-Cookie',
      buildSessionCookie({
        name: cfg.SESSION_COOKIE_NAME,
        value: token,
        maxAgeSeconds: cfg.SESSION_TTL_MS / 1000,
        secure: cfg.COOKIE_SECURE,
      })
    );
    res.json({ ok: true });
  });

  app.post('/api/auth/logout', (req, res) => {
    const token = auth.getSessionToken(req);
    if (token) {
      localDb.deleteSession(token);
    }
    res.setHeader(
      'Set-Cookie',
      buildSessionCookie({
        name: cfg.SESSION_COOKIE_NAME,
        value: '',
        maxAgeSeconds: 0,
        secure: cfg.COOKIE_SECURE,
      })
    );
    res.json({ ok: true });
  });

  // Protected HTML pages
  app.get('/', auth.requirePageAuth, (_req, res) => {
    res.sendFile(path.join(viewsDir, 'dashboard.html'));
  });

  app.get('/generate', auth.requirePageAuth, (_req, res) => {
    res.sendFile(path.join(viewsDir, 'generate.html'));
  });

  // Protected dashboard SQLite read & multi-format import endpoints
  app.get('/api/data', auth.requireApiAuth, (req, res) => {
    const search = typeof req.query.q === 'string' ? req.query.q : '';
    const dataset = localDb.getDataset({ search });
    res.json(dataset);
  });

  app.post(
    '/api/import',
    auth.requireApiAuth,
    uploadImport.single('file'),
    (req, res, next) => {
      if (!req.file) {
        next(new ValidationError('Please choose a file (.xls, .xlsx, .csv, or .sqlite) to import.'));
        return;
      }

      try {
        const parsed = parseImportBuffer(req.file.buffer, req.file.originalname);
        const savedMeta = localDb.replaceDataset({
          sourceFilename: req.file.originalname,
          sourceFormat: parsed.sourceFormat,
          columns: parsed.columns,
          rows: parsed.rows,
        });

        res.json({
          ok: true,
          meta: {
            sourceFilename: savedMeta.sourceFilename,
            sourceFormat: savedMeta.sourceFormat,
            importedAt: savedMeta.importedAt,
            rowCount: savedMeta.rowCount,
          },
          columns: savedMeta.columns,
          totalRows: savedMeta.rowCount,
        });
      } catch (err) {
        next(err);
      }
    }
  );

  // Protected Excel + DUK HTML generator & export download endpoints
  const uploadFields = uploadProcess.fields([
    { name: 'file', maxCount: 1 },
    { name: 'dukFile', maxCount: 1 },
  ]);

  app.post('/api/process', auth.requireApiAuth, uploadFields, (req, res, next) => {
    const baseFile = req.files && req.files.file && req.files.file[0];
    if (!baseFile) {
      next(new ValidationError('Please upload a base Excel file (.xls or .xlsx) in the "file" field.'));
      return;
    }

    const dukFile = req.files && req.files.dukFile && req.files.dukFile[0];
    const includeDukFlag = String((req.body && req.body.includeDuk) || '').toLowerCase() === 'true';
    const includeDuk = includeDukFlag || Boolean(dukFile);

    if (includeDukFlag && !dukFile) {
      next(new ValidationError('DUK support option is enabled, please upload the DUK HTML file (.html or .htm).'));
      return;
    }

    const startTime = process.hrtime.bigint();
    let createdJobId = null;

    try {
      const baseResult = processExcelBuffer(baseFile.buffer, baseFile.originalname, {
        filterColumn: cfg.FILTER_COLUMN,
        targetSatker: cfg.TARGET_SATKER,
        removeColumnPattern: cfg.REMOVE_COLUMN_PATTERN,
      });

      let finalResult = {
        ...baseResult,
        baseExcelFilteredRows: baseResult.filteredRowCount,
        duk: { enabled: false },
      };

      if (includeDuk && dukFile) {
        const dukParsed = parseDukHtmlBuffer(dukFile.buffer, dukFile.originalname);
        finalResult = mergeDukIntoExcelResult(baseResult, dukParsed, {
          dukFilename: dukFile.originalname,
        });
      }

      const fileSuffix = finalResult.duk.enabled
        ? 'kota_metro_combined'
        : 'kota_metro_filtered';
      const fileBase = sanitizeFileBase(baseFile.originalname, fileSuffix);
      const tableName = sanitizeTableName(finalResult.sheetName);

      const { jobId, jobDir } = jobStore.createJobDir();
      createdJobId = jobId;

      const filenames = {
        xlsx: `${fileBase}.xlsx`,
        csv: `${fileBase}.csv`,
        sqlite: `${fileBase}.sqlite`,
      };

      const filePaths = {
        xlsx: path.join(jobDir, filenames.xlsx),
        csv: path.join(jobDir, filenames.csv),
        sqlite: path.join(jobDir, filenames.sqlite),
      };

      exportXlsx({
        filePath: filePaths.xlsx,
        sheetName: finalResult.sheetName,
        headers: finalResult.keptColumns,
        rows: finalResult.rows,
        rawRows: finalResult.rawRows,
      });

      exportCsv({
        filePath: filePaths.csv,
        headers: finalResult.keptColumns,
        rows: finalResult.rows,
      });

      exportSqlite({
        filePath: filePaths.sqlite,
        tableName,
        headers: finalResult.keptColumns,
        rows: finalResult.rows,
      });

      const integrity = verifyOutputs({
        expectedHeaders: finalResult.keptColumns,
        expectedRows: finalResult.rows,
        filterColumnName: finalResult.filterColumn.name,
        targetSatker: finalResult.targetSatker,
        removeColumnPattern: cfg.REMOVE_COLUMN_PATTERN,
        files: filePaths,
        tableName,
      });

      jobStore.saveJob(jobId, {
        jobDir,
        files: {
          xlsx: {
            filePath: filePaths.xlsx,
            downloadName: filenames.xlsx,
            mimeType: FORMAT_MIME_TYPES.xlsx,
            bytes: integrity.formats.xlsx.bytes,
          },
          csv: {
            filePath: filePaths.csv,
            downloadName: filenames.csv,
            mimeType: FORMAT_MIME_TYPES.csv,
            bytes: integrity.formats.csv.bytes,
          },
          sqlite: {
            filePath: filePaths.sqlite,
            downloadName: filenames.sqlite,
            mimeType: FORMAT_MIME_TYPES.sqlite,
            bytes: integrity.formats.sqlite.bytes,
          },
        },
      });

      const elapsedMs = Number(process.hrtime.bigint() - startTime) / 1e6;
      const addedFromDukCount = finalResult.duk.enabled
        ? finalResult.duk.addedFromDukCount
        : 0;

      res.json({
        jobId,
        sourceFile: baseFile.originalname,
        sheetName: finalResult.sheetName,
        tableName,
        filter: {
          column: finalResult.filterColumn.name,
          columnLetter: finalResult.filterColumn.letter,
          query: finalResult.targetSatker,
        },
        stats: {
          totalSourceRows: finalResult.totalSourceRows,
          baseExcelFilteredRows: finalResult.baseExcelFilteredRows,
          addedFromDukCount,
          filteredRowCount: finalResult.filteredRowCount,
          excludedRowCount: finalResult.totalSourceRows - finalResult.baseExcelFilteredRows,
          originalColumnCount: finalResult.originalColumnCount,
          outputColumnCount: finalResult.keptColumns.length,
          removedColumnCount: finalResult.removedColumns.length,
          processingTimeMs: Math.round(elapsedMs * 10) / 10,
        },
        duk: finalResult.duk,
        columns: {
          kept: finalResult.keptColumns,
          removed: finalResult.removedColumns,
        },
        integrity,
        downloads: {
          xlsx: {
            url: `/api/download/${jobId}/xlsx`,
            filename: filenames.xlsx,
            bytes: integrity.formats.xlsx.bytes,
          },
          csv: {
            url: `/api/download/${jobId}/csv`,
            filename: filenames.csv,
            bytes: integrity.formats.csv.bytes,
          },
          sqlite: {
            url: `/api/download/${jobId}/sqlite`,
            filename: filenames.sqlite,
            bytes: integrity.formats.sqlite.bytes,
          },
        },
        preview: {
          headers: finalResult.keptColumns,
          rows: finalResult.rows.slice(0, PREVIEW_ROW_LIMIT),
          totalRows: finalResult.rows.length,
        },
      });
    } catch (err) {
      if (createdJobId) {
        jobStore.removeJobDir(createdJobId);
      }
      next(err);
    }
  });

  app.get('/api/download/:jobId/:format', auth.requireApiAuth, (req, res, next) => {
    const { jobId, format } = req.params;
    const normalizedFormat = String(format || '').toLowerCase();

    if (!jobStore.isValidJobId(jobId)) {
      res.status(400).json({ error: 'Invalid job identifier.' });
      return;
    }

    if (!Object.prototype.hasOwnProperty.call(FORMAT_MIME_TYPES, normalizedFormat)) {
      res.status(400).json({ error: 'Invalid download format. Allowed formats: xlsx, csv, sqlite.' });
      return;
    }

    const fileEntry = jobStore.getJobFile(jobId, normalizedFormat);
    if (!fileEntry) {
      res.status(404).json({ error: 'Export file not found or expired. Please re-upload your Excel file.' });
      return;
    }

    res.setHeader('Content-Type', fileEntry.mimeType);
    res.setHeader('Content-Length', String(fileEntry.bytes));
    res.setHeader('Content-Disposition', `attachment; filename="${fileEntry.downloadName}"`);
    res.setHeader('Cache-Control', 'no-store');

    const stream = fs.createReadStream(fileEntry.filePath);
    stream.on('error', (err) => {
      if (!res.headersSent) {
        next(err);
      } else {
        res.destroy(err);
      }
    });
    stream.pipe(res);
  });

  // Centralized JSON error handler.
  app.use((err, _req, res, _next) => {
    if (err instanceof multer.MulterError) {
      const status = err.code === 'LIMIT_FILE_SIZE' ? 413 : 400;
      const message =
        err.code === 'LIMIT_FILE_SIZE'
          ? `Uploaded file exceeds maximum allowed size (${Math.round(cfg.MAX_FILE_SIZE_BYTES / (1024 * 1024))} MB).`
          : err.message;
      res.status(status).json({ error: message });
      return;
    }

    if (err instanceof ValidationError || err.statusCode === 400) {
      res.status(400).json({ error: err.message });
      return;
    }

    console.error('[server error]', err);
    res.status(500).json({ error: 'An unexpected error occurred while processing the request.' });
  });

  return app;
}

module.exports = {
  createApp,
  FORMAT_MIME_TYPES,
  PREVIEW_ROW_LIMIT,
};
