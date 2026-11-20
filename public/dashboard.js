(() => {
  const ALLOWED_IMPORT_EXTS = ['.xls', '.xlsx', '.csv', '.sqlite', '.db'];

  const datasetMetaSummary = document.getElementById('datasetMetaSummary');
  const openImportBtn = document.getElementById('openImportBtn');
  const emptyImportBtn = document.getElementById('emptyImportBtn');
  const logoutBtn = document.getElementById('logoutBtn');

  const dashboardError = document.getElementById('dashboardError');
  const dashboardErrorMessage = document.getElementById('dashboardErrorMessage');
  const emptyState = document.getElementById('emptyState');
  const tableCard = document.getElementById('tableCard');

  const tableSearch = document.getElementById('tableSearch');
  const clearSearchBtn = document.getElementById('clearSearchBtn');
  const recordCounter = document.getElementById('recordCounter');
  const dashboardColgroup = document.getElementById('dashboardColgroup');
  const dashboardThead = document.getElementById('dashboardThead');
  const dashboardTbody = document.getElementById('dashboardTbody');

  const importModal = document.getElementById('importModal');
  const closeImportModalBtn = document.getElementById('closeImportModalBtn');
  const cancelImportBtn = document.getElementById('cancelImportBtn');
  const importDropzone = document.getElementById('importDropzone');
  const importFileInput = document.getElementById('importFileInput');
  const importSelectedBadge = document.getElementById('importSelectedBadge');
  const importProcessing = document.getElementById('importProcessing');
  const importError = document.getElementById('importError');
  const importErrorMessage = document.getElementById('importErrorMessage');
  const submitImportBtn = document.getElementById('submitImportBtn');

  const copyToast = document.getElementById('copyToast');

  let stagedImportFile = null;
  let searchDebounceTimer = null;
  let toastHideTimer = null;
  let activeCopiedCell = null;

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes < 1024) {
      return `${bytes || 0} B`;
    }
    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  }

  function formatTimestamp(isoString) {
    if (!isoString) {
      return '';
    }
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) {
      return isoString;
    }
    return date.toLocaleString();
  }

  function getColumnWidth(colName) {
    const upper = String(colName || '').toUpperCase();
    if (upper === 'NO') return '58px';
    if (upper === 'NIP') return '130px';
    if (upper === 'NIP_BARU') return '182px';
    if (upper === 'NAMA_LENGKAP') return '215px';
    if (upper === 'PENDIDIKAN' || upper === 'KETERANGAN_SATUAN_KERJA') return '250px';
    if (upper === 'TAMPIL_JABATAN' || upper === 'LEVEL_JABATAN' || upper === 'TIPE_JABATAN') {
      return '205px';
    }
    if (upper.startsWith('SATKER_') || upper === 'GRUP_SATUAN_KERJA') return '205px';
    if (
      upper === 'AGAMA' ||
      upper === 'GOL_RUANG' ||
      upper === 'USIA_PENSIUN' ||
      upper === 'MASA_KERJA_TAHUN' ||
      upper === 'MASA_KERJA_BULAN' ||
      upper === 'STATUS_KAWIN' ||
      upper === 'STATUS_PEGAWAI' ||
      upper.startsWith('TMT_')
    ) {
      return '122px';
    }
    return '170px';
  }

  function showToast(message) {
    copyToast.textContent = message;
    copyToast.classList.remove('hidden');
    if (toastHideTimer) {
      clearTimeout(toastHideTimer);
    }
    toastHideTimer = setTimeout(() => {
      copyToast.classList.add('hidden');
    }, 2200);
  }

  async function copyTextToClipboard(text) {
    if (navigator.clipboard && window.isSecureContext) {
      try {
        await navigator.clipboard.writeText(text);
        return true;
      } catch {
        // Fallback below for restricted contexts
      }
    }

    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.top = '-9999px';
    textarea.style.left = '-9999px';
    document.body.appendChild(textarea);
    textarea.select();
    textarea.setSelectionRange(0, text.length);

    let copied = false;
    try {
      copied = document.execCommand('copy');
    } catch {
      copied = false;
    }
    document.body.removeChild(textarea);
    return copied;
  }

  function showDashboardError(message) {
    dashboardErrorMessage.textContent = message;
    dashboardError.classList.remove('hidden');
  }

  function hideDashboardError() {
    dashboardError.classList.add('hidden');
    dashboardErrorMessage.textContent = '';
  }

  function showImportError(message) {
    importErrorMessage.textContent = message;
    importError.classList.remove('hidden');
  }

  function hideImportError() {
    importError.classList.add('hidden');
    importErrorMessage.textContent = '';
  }

  function resetImportModal() {
    stagedImportFile = null;
    importFileInput.value = '';
    importSelectedBadge.textContent = '';
    importSelectedBadge.classList.add('hidden');
    importDropzone.classList.remove('has-file');
    importProcessing.classList.add('hidden');
    submitImportBtn.disabled = true;
    hideImportError();
  }

  function openImportModal() {
    resetImportModal();
    importModal.classList.remove('hidden');
    importDropzone.focus();
  }

  function closeImportModal() {
    importModal.classList.add('hidden');
    resetImportModal();
  }

  function handleImportFileChoice(file) {
    if (!file) {
      return;
    }
    const lowerName = file.name.toLowerCase();
    const isValidExt = ALLOWED_IMPORT_EXTS.some((ext) => lowerName.endsWith(ext));
    if (!isValidExt) {
      showImportError('Invalid file type. Allowed formats: .xls, .xlsx, .csv, .sqlite, .db.');
      importFileInput.value = '';
      return;
    }

    hideImportError();
    stagedImportFile = file;
    importSelectedBadge.textContent = `Selected: ${file.name} (${formatBytes(file.size)})`;
    importSelectedBadge.classList.remove('hidden');
    importDropzone.classList.add('has-file');
    submitImportBtn.disabled = false;
  }

  function renderTable(dataset) {
    if (!dataset.hasData) {
      datasetMetaSummary.textContent = 'No dataset stored in local SQLite yet.';
      tableCard.classList.add('hidden');
      emptyState.classList.remove('hidden');
      return;
    }

    emptyState.classList.add('hidden');
    tableCard.classList.remove('hidden');

    const { meta, columns, rows, totalRows, filteredRows } = dataset;
    const formatUpper = String(meta.sourceFormat || '').toUpperCase();
    datasetMetaSummary.textContent =
      `${totalRows.toLocaleString()} records • Source: ${meta.sourceFilename} (${formatUpper}) • Imported: ${formatTimestamp(meta.importedAt)}`;

    const isSearching = tableSearch.value.trim().length > 0;
    clearSearchBtn.classList.toggle('hidden', !isSearching);

    if (isSearching) {
      recordCounter.textContent = `Showing ${filteredRows.toLocaleString()} of ${totalRows.toLocaleString()} records`;
    } else {
      recordCounter.textContent = `${totalRows.toLocaleString()} records`;
    }

    const displayHeaders = ['NO', ...columns];
    const colgroupFrag = document.createDocumentFragment();
    const headTr = document.createElement('tr');

    for (let c = 0; c < displayHeaders.length; c += 1) {
      const colEl = document.createElement('col');
      colEl.style.width = getColumnWidth(displayHeaders[c]);
      colgroupFrag.appendChild(colEl);

      const th = document.createElement('th');
      th.textContent = displayHeaders[c];
      headTr.appendChild(th);
    }

    dashboardColgroup.replaceChildren(colgroupFrag);
    dashboardThead.replaceChildren(headTr);

    const bodyFrag = document.createDocumentFragment();
    if (rows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = displayHeaders.length;
      td.className = 'table-empty-row';
      td.textContent = isSearching
        ? 'No matching records found for your search.'
        : 'The imported dataset contains 0 data rows.';
      tr.appendChild(td);
      bodyFrag.appendChild(tr);
    } else {
      for (let r = 0; r < rows.length; r += 1) {
        const tr = document.createElement('tr');

        const noTd = document.createElement('td');
        const noStr = String(r + 1);
        noTd.className = 'copyable-cell cell-no';
        noTd.tabIndex = 0;
        noTd.setAttribute('role', 'button');
        noTd.title = 'Click to copy';
        noTd.dataset.copyValue = noStr;
        noTd.textContent = noStr;
        tr.appendChild(noTd);

        const row = rows[r];
        for (let c = 0; c < columns.length; c += 1) {
          const td = document.createElement('td');
          const cellVal = row[c] === undefined || row[c] === null ? '' : String(row[c]);

          if (cellVal === '') {
            const emptySpan = document.createElement('span');
            emptySpan.className = 'cell-empty';
            emptySpan.setAttribute('aria-label', 'Empty');
            emptySpan.textContent = '—';
            td.appendChild(emptySpan);
          } else {
            td.className = 'copyable-cell';
            td.tabIndex = 0;
            td.setAttribute('role', 'button');
            td.title = 'Click to copy';
            td.dataset.copyValue = cellVal;
            td.textContent = cellVal;
          }
          tr.appendChild(td);
        }

        bodyFrag.appendChild(tr);
      }
    }

    dashboardTbody.replaceChildren(bodyFrag);
  }

  async function loadDashboardData(searchQuery = '') {
    hideDashboardError();
    try {
      const url = searchQuery
        ? `/api/data?q=${encodeURIComponent(searchQuery)}`
        : '/api/data';
      const response = await fetch(url);
      if (response.status === 401) {
        window.location.assign('/login');
        return;
      }

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error || 'Failed to load data from SQLite.');
      }

      renderTable(payload);
    } catch (err) {
      showDashboardError(err.message || 'Unable to load saved employee data.');
    }
  }

  async function submitImport() {
    if (!stagedImportFile) {
      showImportError('Please select a file to import.');
      return;
    }

    hideImportError();
    importProcessing.classList.remove('hidden');
    submitImportBtn.disabled = true;
    cancelImportBtn.disabled = true;

    try {
      const formData = new FormData();
      formData.append('file', stagedImportFile);

      const response = await fetch('/api/import', {
        method: 'POST',
        body: formData,
      });

      if (response.status === 401) {
        window.location.assign('/login');
        return;
      }

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error || 'Failed to import file into SQLite.');
      }

      closeImportModal();
      tableSearch.value = '';
      await loadDashboardData('');
      showToast(`Imported ${payload.totalRows.toLocaleString()} records into local SQLite.`);
    } catch (err) {
      importProcessing.classList.add('hidden');
      submitImportBtn.disabled = false;
      showImportError(err.message || 'Import failed.');
    } finally {
      cancelImportBtn.disabled = false;
    }
  }

  async function handleCellCopy(td) {
    if (!td || !td.classList.contains('copyable-cell')) {
      return;
    }
    const text = td.dataset.copyValue !== undefined ? td.dataset.copyValue : td.textContent;
    if (!text) {
      return;
    }

    const copied = await copyTextToClipboard(text);
    if (!copied) {
      showToast('Unable to copy automatically — please copy manually.');
      return;
    }

    if (activeCopiedCell && activeCopiedCell !== td) {
      activeCopiedCell.classList.remove('cell-copied');
    }
    td.classList.remove('cell-copied');
    // Force reflow so rapid re-clicks replay highlight cleanly
    void td.offsetWidth;
    td.classList.add('cell-copied');
    activeCopiedCell = td;

    const preview = text.length > 48 ? `${text.slice(0, 48)}…` : text;
    showToast(`Copied: ${preview}`);
  }

  dashboardTbody.addEventListener('click', (event) => {
    const td = event.target.closest('td.copyable-cell');
    if (td) {
      handleCellCopy(td);
    }
  });

  dashboardTbody.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      const td = event.target.closest('td.copyable-cell');
      if (td) {
        event.preventDefault();
        handleCellCopy(td);
      }
    }
  });

  tableSearch.addEventListener('input', () => {
    const query = tableSearch.value.trim();
    clearSearchBtn.classList.toggle('hidden', query.length === 0);
    if (searchDebounceTimer) {
      clearTimeout(searchDebounceTimer);
    }
    searchDebounceTimer = setTimeout(() => {
      loadDashboardData(query);
    }, 180);
  });

  clearSearchBtn.addEventListener('click', () => {
    tableSearch.value = '';
    clearSearchBtn.classList.add('hidden');
    loadDashboardData('');
    tableSearch.focus();
  });

  openImportBtn.addEventListener('click', openImportModal);
  emptyImportBtn.addEventListener('click', openImportModal);
  closeImportModalBtn.addEventListener('click', closeImportModal);
  cancelImportBtn.addEventListener('click', closeImportModal);
  submitImportBtn.addEventListener('click', submitImport);

  importModal.addEventListener('click', (event) => {
    if (event.target === importModal) {
      closeImportModal();
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !importModal.classList.contains('hidden')) {
      closeImportModal();
    }
  });

  importFileInput.addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      handleImportFileChoice(file);
    }
  });

  importDropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      importFileInput.click();
    }
  });

  importDropzone.addEventListener('dragover', (event) => {
    event.preventDefault();
    importDropzone.classList.add('drag-active');
  });

  importDropzone.addEventListener('dragleave', () => {
    importDropzone.classList.remove('drag-active');
  });

  importDropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    importDropzone.classList.remove('drag-active');
    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) {
      handleImportFileChoice(file);
    }
  });

  window.addEventListener('dragover', (event) => event.preventDefault());
  window.addEventListener('drop', (event) => event.preventDefault());

  logoutBtn.addEventListener('click', async () => {
    await fetch('/api/auth/logout', { method: 'POST' }).catch(() => {});
    window.location.assign('/login');
  });

  loadDashboardData();
})();

