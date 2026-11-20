const path = require('node:path');

const DEFAULT_CONFIG = Object.freeze({
  PORT: Number(process.env.PORT) || 3000,
  FILTER_COLUMN: 'SATKER_3',
  TARGET_SATKER: 'Kantor Kementerian Agama Kota Metro',
  REMOVE_COLUMN_PATTERN: /^KODE_/i,
  MAX_FILE_SIZE_BYTES: (Number(process.env.MAX_FILE_SIZE_MB) || 100) * 1024 * 1024,
  JOB_TTL_MS: (Number(process.env.JOB_TTL_MINUTES) || 60) * 60 * 1000,
  TMP_DIR: path.resolve(__dirname, '..', 'tmp', 'jobs'),
  ALLOWED_EXTENSIONS: ['.xls', '.xlsx'],
  ALLOWED_HTML_EXTENSIONS: ['.html', '.htm'],
});

/**
 * Merge default configuration with optional runtime overrides (e.g. for tests).
 */
function resolveConfig(overrides = {}) {
  return {
    ...DEFAULT_CONFIG,
    ...overrides,
    TMP_DIR: overrides.tmpDir
      ? path.resolve(overrides.tmpDir)
      : (overrides.TMP_DIR ? path.resolve(overrides.TMP_DIR) : DEFAULT_CONFIG.TMP_DIR),
  };
}

module.exports = {
  DEFAULT_CONFIG,
  resolveConfig,
};

