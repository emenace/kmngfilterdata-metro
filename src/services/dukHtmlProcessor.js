const { ValidationError } = require('./excelProcessor');

const GOL_RUANG_TO_PANGKAT = Object.freeze({
  'i/a': 'Juru Muda',
  'i/b': 'Juru Muda Tk. I',
  'i/c': 'Juru',
  'i/d': 'Juru Tk.I',
  'ii/a': 'Pengatur Muda',
  'ii/b': 'Pengatur Muda Tk.I',
  'ii/c': 'Pengatur',
  'ii/d': 'Pengatur Tk.I',
  'iii/a': 'Penata Muda',
  'iii/b': 'Penata Muda Tk. I',
  'iii/c': 'Penata',
  'iii/d': 'Penata Tk. I',
  'iv/a': 'Pembina',
  'iv/b': 'Pembina Tk. I',
  'iv/c': 'Pembina Utama Muda',
  'iv/d': 'Pembina Utama Madya',
  'iv/e': 'Pembina Utama',
});

/**
 * Decode common HTML entities into plain text characters.
 */
function decodeHtmlEntities(str) {
  return String(str || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#0*39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

/**
 * Convert HTML cell fragment into clean text while preserving <br> line breaks.
 */
function cleanHtmlCellText(fragment) {
  return decodeHtmlEntities(
    String(fragment || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, '')
  )
    .split('\n')
    .map((line) => line.replace(/[ \t\r]+/g, ' ').trim())
    .join('\n')
    .trim();
}

/**
 * Convert DD-MM-YYYY date strings to YYYY-MM-DD; leave other formats untouched.
 */
function dmyToYmd(dateStr) {
  const val = String(dateStr || '').trim();
  const match = val.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) {
    return val;
  }
  return `${match[3]}-${match[2]}-${match[1]}`;
}

/**
 * Derive retirement age in years from birth date and pension TMT (both DD-MM-YYYY).
 */
function calculateRetirementAge(birthDmy, pensionDmy) {
  const birth = String(birthDmy || '').trim().match(/^(\d{2})-(\d{2})-(\d{4})$/);
  const pension = String(pensionDmy || '').trim().match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!birth || !pension) {
    return '';
  }

  let age = Number(pension[3]) - Number(birth[3]);
  if (Number(pension[2]) < Number(birth[2])) {
    age -= 1;
  }
  return Number.isFinite(age) && age > 0 ? String(age) : '';
}

/**
 * Extract TIPE_JABATAN, TAMPIL_JABATAN, SATKER_1/2, KETERANGAN_SATUAN_KERJA, and STATUS_PEGAWAI
 * from a DUK JABATAN description string.
 */
function parseJabatanDetails(jabatan, golRuang, targetSatker) {
  const rawJabatan = String(jabatan || '').trim();
  const isPppkParuhWaktu = /\(PPPK\s+PARUH\s+WAKTU\)/i.test(rawJabatan);
  const withoutPppkTag = rawJabatan.replace(/\s*\(PPPK\s+PARUH\s+WAKTU\)\s*$/i, '').trim();

  const padaMatch = withoutPppkTag.match(/^([\s\S]*?)\s+pada\s+([\s\S]+)$/i);
  const posPart = padaMatch ? padaMatch[1].trim() : withoutPppkTag;
  const keteranganSatker = padaMatch ? padaMatch[2].trim() : '';

  const tipeMatch = posPart.match(/\((Jabatan[^)]+)\)/i);
  const tipeJabatan = tipeMatch ? tipeMatch[1].trim() : '';
  const tampilJabatan = posPart
    .replace(/\s*\(Jabatan[^)]+\)/gi, '')
    .replace(/\s+/g, ' ')
    .trim();

  let satker1 = targetSatker;
  let satker2 = targetSatker;
  if (keteranganSatker) {
    if (keteranganSatker.toLowerCase().startsWith(targetSatker.toLowerCase())) {
      satker1 = targetSatker;
      satker2 = targetSatker;
    } else {
      const strippedUnit = keteranganSatker
        .replace(/,?\s*(?:Kota Metro\s+)?Provinsi Lampung$/i, '')
        .trim();
      if (strippedUnit) {
        satker1 = strippedUnit;
        satker2 = strippedUnit;
      }
    }
  }

  const normalizedGol = String(golRuang || '').trim();
  let statusPegawai = '';
  if (isPppkParuhWaktu) {
    statusPegawai = 'PPPK PARUH WAKTU';
  } else if (/^[IVX]+\/[a-e]$/i.test(normalizedGol)) {
    statusPegawai = 'PNS';
  } else if (/^(?:X{0,1}(?:IX|IV|V?I{0,3}))$/i.test(normalizedGol)) {
    statusPegawai = 'PPPK';
  }

  return {
    tipeJabatan,
    tampilJabatan,
    keteranganSatker: keteranganSatker || targetSatker,
    satker1,
    satker2,
    statusPegawai,
  };
}

/**
 * Format combined education text consistent with the provincial Excel PENDIDIKAN column.
 */
function formatPendidikan(jenjang, pendNama, pendTahun) {
  const parts = [];
  const cleanJenjang = String(jenjang || '').trim();
  const cleanNama = String(pendNama || '').trim();
  const cleanTahun = String(pendTahun || '').trim();

  if (cleanJenjang && cleanJenjang !== '-') {
    parts.push(cleanJenjang);
  }
  if (cleanNama && cleanNama !== '-') {
    parts.push(cleanNama);
  }
  if (cleanTahun && cleanTahun !== '-') {
    parts.push(`Tahun ${cleanTahun}`);
  }
  return parts.join(' ');
}

/**
 * Parse a DUK-SISDM HTML buffer and extract structured employee records.
 */
function parseDukHtmlBuffer(buffer, originalFilename = 'duk.html') {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new ValidationError('Uploaded DUK HTML file is empty.');
  }

  const html = buffer.toString('utf8');
  if (!/<table\b/i.test(html)) {
    throw new ValidationError(
      `Uploaded file "${originalFilename}" does not contain a DUK-SISDM HTML table.`
    );
  }

  const lblMatch = html.match(/<span[^>]*id=["']lblSatker["'][^>]*>([\s\S]*?)<\/span>/i);
  const satkerTitle = lblMatch ? cleanHtmlCellText(lblMatch[1]) : '';

  const trMatches = html.match(/<tr\b[\s\S]*?<\/tr>/gi) || [];
  let hasExpectedHeader = false;
  const records = [];

  for (const tr of trMatches) {
    const plainTr = cleanHtmlCellText(tr).toUpperCase();
    if (
      plainTr.includes('PANGKAT') ||
      plainTr.includes('TMT PENSIUN') ||
      plainTr.includes('PENDIDIKAN TERAKHIR')
    ) {
      hasExpectedHeader = true;
    }

    const tdMatches = tr.match(/<t[hd]\b[\s\S]*?<\/t[hd]>/gi) || [];
    if (tdMatches.length !== 14) {
      continue;
    }

    const cells = tdMatches.map(cleanHtmlCellText);
    // Skip column numbering row ("1".."14")
    if (cells[0] === '1' && cells[1] === '2' && cells[13] === '14') {
      continue;
    }

    const nameLines = cells[1]
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);
    const nama = nameLines[0] || '';
    let nip = '';
    let nipBaru = '';

    if (nameLines.length >= 2) {
      const idLine = nameLines[nameLines.length - 1];
      const parts = idLine.split('/').map((part) => part.trim());
      if (parts.length >= 2) {
        nip = parts[0];
        nipBaru = parts[1];
      } else if (/^\d{18}$/.test(idLine)) {
        nipBaru = idLine;
      } else {
        nip = idLine;
      }
    }

    if (!nama && !nip && !nipBaru) {
      continue;
    }

    records.push({
      no: cells[0],
      nama,
      nip,
      nipBaru,
      golRuang: cells[2],
      tmtPangkat: cells[3],
      jabatan: cells[4],
      tmtJabatan: cells[5],
      mkTahun: cells[6],
      mkBulan: cells[7],
      pendNama: cells[8],
      pendTahun: cells[9],
      jenjang: cells[10],
      tglLahir: cells[11],
      tmtPensiun: cells[12],
      ket: cells[13],
    });
  }

  if (!hasExpectedHeader || records.length === 0) {
    throw new ValidationError(
      `Uploaded HTML file "${originalFilename}" does not match the expected DUK-SISDM table format.`
    );
  }

  return {
    satkerTitle,
    totalDukRows: records.length,
    records,
  };
}

/**
 * Find the single common non-empty value in a column across filtered base rows, if uniform.
 */
function getUniformColumnValue(rows, colIdx) {
  if (colIdx === -1) {
    return '';
  }
  let candidate = '';
  for (const row of rows) {
    const val = String(row[colIdx] || '');
    if (!val) {
      continue;
    }
    if (!candidate) {
      candidate = val;
    } else if (candidate !== val) {
      return '';
    }
  }
  return candidate;
}

/**
 * Merge missing people from a parsed DUK HTML dataset into the filtered base Excel result
 * without modifying any existing base Excel row.
 */
function mergeDukIntoExcelResult(excelResult, dukResult, { dukFilename }) {
  const keptColumns = excelResult.keptColumns;
  const combinedRows = excelResult.rows.map((row) => [...row]);
  const combinedRawRows = excelResult.rawRows.map((row) => [...row]);

  const nipBaruIdx = keptColumns.indexOf('NIP_BARU');
  const nipIdx = keptColumns.indexOf('NIP');
  const satker4Idx = keptColumns.indexOf('SATKER_4');
  const satker5Idx = keptColumns.indexOf('SATKER_5');

  const seenNipBaru = new Set();
  const seenNip = new Set();

  for (const row of excelResult.rows) {
    if (nipBaruIdx !== -1) {
      const nb = String(row[nipBaruIdx] || '').trim();
      if (nb) {
        seenNipBaru.add(nb);
      }
    }
    if (nipIdx !== -1) {
      const n = String(row[nipIdx] || '').trim();
      if (n) {
        seenNip.add(n);
      }
    }
  }

  const uniformSatker4 = getUniformColumnValue(excelResult.rows, satker4Idx);
  const uniformSatker5 = getUniformColumnValue(excelResult.rows, satker5Idx);

  let alreadyInExcelCount = 0;
  const addedPeople = [];

  for (const record of dukResult.records) {
    const cleanNipBaru = record.nipBaru.trim();
    const cleanNip = record.nip.trim();

    const existsByNipBaru = Boolean(cleanNipBaru) && seenNipBaru.has(cleanNipBaru);
    const existsByNip = Boolean(cleanNip) && seenNip.has(cleanNip);
    if (existsByNipBaru || existsByNip) {
      alreadyInExcelCount += 1;
      continue;
    }

    const jabatanDetails = parseJabatanDetails(
      record.jabatan,
      record.golRuang,
      excelResult.targetSatker
    );
    const pangkat =
      GOL_RUANG_TO_PANGKAT[String(record.golRuang || '').trim().toLowerCase()] || '';

    const mappedByHeader = {
      NIP: record.nip,
      NIP_BARU: record.nipBaru,
      NAMA_LENGKAP: record.nama,
      AGAMA: '',
      PENDIDIKAN: formatPendidikan(record.jenjang, record.pendNama, record.pendTahun),
      JENJANG_PENDIDIKAN: record.jenjang,
      LEVEL_JABATAN: '',
      PANGKAT: pangkat,
      GOL_RUANG: record.golRuang,
      TMT_CPNS: '',
      TMT_PANGKAT: record.tmtPangkat,
      TIPE_JABATAN: jabatanDetails.tipeJabatan,
      TAMPIL_JABATAN: jabatanDetails.tampilJabatan,
      TMT_JABATAN: dmyToYmd(record.tmtJabatan),
      SATKER_1: jabatanDetails.satker1,
      KETERANGAN_SATUAN_KERJA: jabatanDetails.keteranganSatker,
      SATKER_2: jabatanDetails.satker2,
      SATKER_3: excelResult.targetSatker,
      SATKER_4: uniformSatker4,
      SATKER_5: uniformSatker5,
      GRUP_SATUAN_KERJA: '',
      STATUS_KAWIN: '',
      STATUS_PEGAWAI: jabatanDetails.statusPegawai,
      USIA_PENSIUN: calculateRetirementAge(record.tglLahir, record.tmtPensiun),
      TMT_PENSIUN: dmyToYmd(record.tmtPensiun),
      MASA_KERJA_TAHUN: record.mkTahun,
      MASA_KERJA_BULAN: record.mkBulan,
    };

    const canonicalRow = keptColumns.map((header) =>
      Object.prototype.hasOwnProperty.call(mappedByHeader, header) ? mappedByHeader[header] : ''
    );
    const rawRow = canonicalRow.map((val) => ({ t: 's', v: val }));

    combinedRows.push(canonicalRow);
    combinedRawRows.push(rawRow);

    if (cleanNipBaru) {
      seenNipBaru.add(cleanNipBaru);
    }
    if (cleanNip) {
      seenNip.add(cleanNip);
    }

    addedPeople.push({
      nip: record.nip,
      nipBaru: record.nipBaru,
      nama: record.nama,
      statusPegawai: jabatanDetails.statusPegawai,
    });
  }

  return {
    ...excelResult,
    baseExcelFilteredRows: excelResult.filteredRowCount,
    filteredRowCount: combinedRows.length,
    rows: combinedRows,
    rawRows: combinedRawRows,
    duk: {
      enabled: true,
      sourceFile: dukFilename,
      satkerTitle: dukResult.satkerTitle,
      totalDukRows: dukResult.totalDukRows,
      alreadyInExcelCount,
      addedFromDukCount: addedPeople.length,
      addedPeople,
    },
  };
}

module.exports = {
  GOL_RUANG_TO_PANGKAT,
  decodeHtmlEntities,
  dmyToYmd,
  calculateRetirementAge,
  parseJabatanDetails,
  formatPendidikan,
  parseDukHtmlBuffer,
  mergeDukIntoExcelResult,
};

