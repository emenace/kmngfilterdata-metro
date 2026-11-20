const UTF8_BOM = '\uFEFF';

/**
 * Escape a single field according to RFC 4180 without altering cell content
 * (preserves leading zeros, spaces, signs, and mixed date formats verbatim).
 */
function escapeCsvField(value) {
  const str = value === null || value === undefined ? '' : String(value);
  if (
    str.includes(',') ||
    str.includes('"') ||
    str.includes('\r') ||
    str.includes('\n')
  ) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * Serialize headers and data rows into an RFC 4180 CSV string with UTF-8 BOM.
 */
function serializeCsv(headers, rows) {
  const allRows = [headers, ...rows];
  const lines = allRows.map((row) => row.map(escapeCsvField).join(','));
  return UTF8_BOM + lines.join('\r\n') + '\r\n';
}

/**
 * Parse an RFC 4180 CSV string (with optional leading UTF-8 BOM) into string[][].
 */
function parseCsv(text) {
  const input = text.startsWith(UTF8_BOM) ? text.slice(1) : text;
  if (input.length === 0) {
    return [];
  }

  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];

    if (inQuotes) {
      if (ch === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"' && field.length === 0) {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field);
      field = '';
    } else if (ch === '\r') {
      if (input[i + 1] === '\n') {
        i += 1;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }

  if (inQuotes) {
    throw new Error('Malformed CSV: unterminated quoted field.');
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

module.exports = {
  UTF8_BOM,
  escapeCsvField,
  serializeCsv,
  parseCsv,
};

