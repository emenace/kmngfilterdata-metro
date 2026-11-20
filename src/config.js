const fs = require('node:fs');
const path = require('node:path');

const ENV_FILE_PATH = path.resolve(__dirname, '..', '.env');
if (typeof process.loadEnvFile === 'function' && fs.existsSync(ENV_FILE_PATH)) {
  process.loadEnvFile(ENV_FILE_PATH);
}

const DEFAULT_CONFIG = Object.freeze({
  PORT: Number(process.env.PORT) || 3000,
  APP_PASSWORD: process.env.APP_PASSWORD || 'kemenagmetro',
  SESSION_COOKIE_NAME: 'kmng_session',
  SESSION_TTL_MS: (Number(process.env.SESSION_TTL_HOURS) || 5) * 60 * 60 * 1000,
  COOKIE_SECURE: String(process.env.COOKIE_SECURE || '').toLowerCase() === 'true',
  DB_PATH: process.env.DB_PATH
    ? path.resolve(process.env.DB_PATH)
    : path.resolve(__dirname, '..', 'data', 'kmng_metro.sqlite'),
  FILTER_COLUMN: 'SATKER_3',
  TARGET_SATKER: 'Kantor Kementerian Agama Kota Metro',
  REMOVE_COLUMN_PATTERN: /^KODE_/i,
  MAX_FILE_SIZE_BYTES: (Number(process.env.MAX_FILE_SIZE_MB) || 100) * 1024 * 1024,
  JOB_TTL_MS: (Number(process.env.JOB_TTL_MINUTES) || 60) * 60 * 1000,
  TMP_DIR: path.resolve(__dirname, '..', 'tmp', 'jobs'),
  ALLOWED_EXTENSIONS: ['.xls', '.xlsx'],
  ALLOWED_HTML_EXTENSIONS: ['.html', '.htm'],
  ALLOWED_IMPORT_EXTENSIONS: ['.xls', '.xlsx', '.csv', '.sqlite', '.db'],
});

/**
 * Merge default configuration with optional runtime overrides (e.g. for tests).
 */
function resolveConfig(overrides = {}) {
  const resolvedTmpDir = overrides.tmpDir
    ? path.resolve(overrides.tmpDir)
    : (overrides.TMP_DIR ? path.resolve(overrides.TMP_DIR) : DEFAULT_CONFIG.TMP_DIR);

  const resolvedDbPath = overrides.dbPath
    ? path.resolve(overrides.dbPath)
    : (overrides.DB_PATH ? path.resolve(overrides.DB_PATH) : DEFAULT_CONFIG.DB_PATH);

  return {
    ...DEFAULT_CONFIG,
    ...overrides,
    TMP_DIR: resolvedTmpDir,
    DB_PATH: resolvedDbPath,
  };
}

module.exports = {
  DEFAULT_CONFIG,
  resolveConfig,
};
