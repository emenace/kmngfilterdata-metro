(() => {
  const PAGE_SIZE = 20;

  const logoutBtn = document.getElementById('logoutBtn');
  const uploadSection = document.getElementById('uploadSection');
  const dukToggle = document.getElementById('dukToggle');
  const dropzonesGrid = document.getElementById('dropzonesGrid');
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const excelDropHint = document.getElementById('excelDropHint');
  const excelSelectedBadge = document.getElementById('excelSelectedBadge');

  const dukUploadWrap = document.getElementById('dukUploadWrap');
  const dukDropzone = document.getElementById('dukDropzone');
  const dukFileInput = document.getElementById('dukFileInput');
  const dukSelectedBadge = document.getElementById('dukSelectedBadge');

  const combineActions = document.getElementById('combineActions');
  const combineStatusHint = document.getElementById('combineStatusHint');
  const processCombineBtn = document.getElementById('processCombineBtn');

  const processingPanel = document.getElementById('processingPanel');
  const processingFileMeta = document.getElementById('processingFileMeta');
  const errorBox = document.getElementById('errorBox');
  const errorMessage = document.getElementById('errorMessage');
  const resultsPanel = document.getElementById('resultsPanel');
  const resetBtn = document.getElementById('resetBtn');

  const sourceFilename = document.getElementById('sourceFilename');
  const sourceMeta = document.getElementById('sourceMeta');
  const statSourceRows = document.getElementById('statSourceRows');
  const statFilteredRows = document.getElementById('statFilteredRows');
  const statFilteredRowsSub = document.getElementById('statFilteredRowsSub');
  const kpiDukAddedCard = document.getElementById('kpiDukAddedCard');
  const statDukAdded = document.getElementById('statDukAdded');
  const statDukAddedSub = document.getElementById('statDukAddedSub');
  const statRemovedCols = document.getElementById('statRemovedCols');
  const statKeptCols = document.getElementById('statKeptCols');
  const processingDuration = document.getElementById('processingDuration');
  const integrityChecksum = document.getElementById('integrityChecksum');

  const downloadXlsx = document.getElementById('downloadXlsx');
  const sizeXlsx = document.getElementById('sizeXlsx');
  const nameXlsx = document.getElementById('nameXlsx');

  const downloadCsv = document.getElementById('downloadCsv');
  const sizeCsv = document.getElementById('sizeCsv');
  const nameCsv = document.getElementById('nameCsv');

  const downloadSqlite = document.getElementById('downloadSqlite');
  const sizeSqlite = document.getElementById('sizeSqlite');
  const nameSqlite = document.getElementById('nameSqlite');

  const dukSummaryCard = document.getElementById('dukSummaryCard');
  const dukSummaryCounts = document.getElementById('dukSummaryCounts');
  const dukSourceMeta = document.getElementById('dukSourceMeta');
  const dukAddedList = document.getElementById('dukAddedList');

  const colAuditCounts = document.getElementById('colAuditCounts');
  const removedColumnsList = document.getElementById('removedColumnsList');
  const keptColumnsList = document.getElementById('keptColumnsList');

  const previewSubtitle = document.getElementById('previewSubtitle');
  const previewSearch = document.getElementById('previewSearch');
  const previewThead = document.getElementById('previewThead');
  const previewTbody = document.getElementById('previewTbody');
  const pageInfo = document.getElementById('pageInfo');
  const prevPage = document.getElementById('prevPage');
  const nextPage = document.getElementById('nextPage');

  let selectedExcelFile = null;
  let selectedDukFile = null;
  let previewHeaders = [];
  let previewAllRows = [];
  let previewFilteredRows = [];
  let currentPage = 1;

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes < 1024) {
      return `${bytes || 0} B`;
    }
    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  }

  function showError(message) {
    errorMessage.textContent = message;
    errorBox.classList.remove('hidden');
  }

  function hideError() {
    errorBox.classList.add('hidden');
    errorMessage.textContent = '';
  }

  function updateStagedSelectionUI() {
    if (selectedExcelFile) {
      excelSelectedBadge.textContent = `Selected: ${selectedExcelFile.name} (${formatBytes(selectedExcelFile.size)})`;
      excelSelectedBadge.classList.remove('hidden');
      dropzone.classList.add('has-file');
    } else {
      excelSelectedBadge.textContent = '';
      excelSelectedBadge.classList.add('hidden');
      dropzone.classList.remove('has-file');
    }

    if (selectedDukFile) {
      dukSelectedBadge.textContent = `Selected: ${selectedDukFile.name} (${formatBytes(selectedDukFile.size)})`;
      dukSelectedBadge.classList.remove('hidden');
      dukDropzone.classList.add('has-file');
    } else {
      dukSelectedBadge.textContent = '';
      dukSelectedBadge.classList.add('hidden');
      dukDropzone.classList.remove('has-file');
    }

    const readyToCombine = Boolean(selectedExcelFile && selectedDukFile);
    processCombineBtn.disabled = !readyToCombine;
    if (readyToCombine) {
      combineStatusHint.textContent = 'Both files are ready. Click "Process & Combine Files" to filter and merge.';
    } else if (selectedExcelFile && !selectedDukFile) {
      combineStatusHint.textContent = 'Base Excel selected. Now upload the DUK HTML support file.';
    } else if (!selectedExcelFile && selectedDukFile) {
      combineStatusHint.textContent = 'DUK HTML selected. Now upload the base Excel file.';
    } else {
      combineStatusHint.textContent = 'Select both the base Excel file and the DUK HTML file to continue.';
    }
  }

  function syncDukToggleUI() {
    const enabled = dukToggle.checked;
    if (enabled) {
      dropzonesGrid.classList.add('dual');
      dukUploadWrap.classList.remove('hidden');
      combineActions.classList.remove('hidden');
      excelDropHint.textContent = 'Supports legacy .xls (BIFF8) and .xlsx up to 100 MB.';
    } else {
      dropzonesGrid.classList.remove('dual');
      dukUploadWrap.classList.add('hidden');
      combineActions.classList.add('hidden');
      excelDropHint.textContent = 'Supports legacy .xls (BIFF8) and .xlsx up to 100 MB. Processing starts automatically upon upload.';
      selectedDukFile = null;
      dukFileInput.value = '';
    }
    updateStagedSelectionUI();
  }

  function setProcessingState(isProcessing, excelFile, dukFile) {
    if (isProcessing && excelFile) {
      uploadSection.classList.add('hidden');
      resultsPanel.classList.add('hidden');
      hideError();
      const summary = dukFile
        ? `${excelFile.name} (${formatBytes(excelFile.size)}) + ${dukFile.name} (${formatBytes(dukFile.size)})`
        : `${excelFile.name} (${formatBytes(excelFile.size)})`;
      processingFileMeta.textContent = summary;
      processingPanel.classList.remove('hidden');
    } else {
      processingPanel.classList.add('hidden');
    }
  }

  function resetToUpload() {
    selectedExcelFile = null;
    selectedDukFile = null;
    fileInput.value = '';
    dukFileInput.value = '';
    previewSearch.value = '';
    resultsPanel.classList.add('hidden');
    processingPanel.classList.add('hidden');
    hideError();
    uploadSection.classList.remove('hidden');
    syncDukToggleUI();
  }

  function renderChipList(container, items, className) {
    container.replaceChildren();
    if (!items || items.length === 0) {
      const emptyNote = document.createElement('span');
      emptyNote.className = 'chip';
      emptyNote.textContent = 'None';
      container.appendChild(emptyNote);
      return;
    }

    const frag = document.createDocumentFragment();
    for (const item of items) {
      const span = document.createElement('span');
      span.className = `chip ${className}`;
      span.textContent = item;
      frag.appendChild(span);
    }
    container.appendChild(frag);
  }

  function renderDukSummary(duk) {
    if (!duk || !duk.enabled) {
      dukSummaryCard.classList.add('hidden');
      dukAddedList.replaceChildren();
      return;
    }

    dukSummaryCard.classList.remove('hidden');
    dukSummaryCounts.textContent = `${duk.addedFromDukCount} added, ${duk.alreadyInExcelCount} already in Excel`;
    const satkerNote = duk.satkerTitle ? ` • Satker: "${duk.satkerTitle}"` : '';
    dukSourceMeta.textContent = `Support file: ${duk.sourceFile} (${duk.totalDukRows} DUK rows parsed${satkerNote})`;

    const frag = document.createDocumentFragment();
    if (!duk.addedPeople || duk.addedPeople.length === 0) {
      const li = document.createElement('li');
      li.className = 'duk-added-item';
      li.textContent = 'All employees from the DUK HTML file were already present in the base Excel dataset.';
      frag.appendChild(li);
    } else {
      for (const person of duk.addedPeople) {
        const li = document.createElement('li');
        li.className = 'duk-added-item';

        const nameDiv = document.createElement('div');
        nameDiv.className = 'duk-added-name';
        nameDiv.textContent = person.nama || '(Unnamed)';

        const idDiv = document.createElement('div');
        idDiv.className = 'duk-added-ids';
        const statusPart = person.statusPegawai ? ` • ${person.statusPegawai}` : '';
        idDiv.textContent = `${person.nip || '-'} / ${person.nipBaru || '-'}${statusPart}`;

        li.appendChild(nameDiv);
        li.appendChild(idDiv);
        frag.appendChild(li);
      }
    }
    dukAddedList.replaceChildren(frag);
  }

  function renderPreviewTable() {
    const totalMatches = previewFilteredRows.length;
    const totalPages = Math.max(1, Math.ceil(totalMatches / PAGE_SIZE));
    if (currentPage > totalPages) {
      currentPage = totalPages;
    }

    const startIdx = (currentPage - 1) * PAGE_SIZE;
    const pageRows = previewFilteredRows.slice(startIdx, startIdx + PAGE_SIZE);

    const headTr = document.createElement('tr');
    const numTh = document.createElement('th');
    numTh.textContent = 'NO';
    headTr.appendChild(numTh);

    for (const headerName of previewHeaders) {
      const th = document.createElement('th');
      th.textContent = headerName;
      headTr.appendChild(th);
    }
    previewThead.replaceChildren(headTr);

    const bodyFrag = document.createDocumentFragment();
    if (pageRows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = Math.max(1, previewHeaders.length + 1);
      td.textContent = 'No matching rows in preview.';
      tr.appendChild(td);
      bodyFrag.appendChild(tr);
    } else {
      for (let i = 0; i < pageRows.length; i += 1) {
        const entry = pageRows[i];
        const tr = document.createElement('tr');
        const idxTd = document.createElement('td');
        idxTd.textContent = String(entry.no);
        tr.appendChild(idxTd);

        for (let c = 0; c < previewHeaders.length; c += 1) {
          const td = document.createElement('td');
          const val = entry.row[c];
          td.textContent = val === undefined || val === null ? '' : String(val);
          tr.appendChild(td);
        }
        bodyFrag.appendChild(tr);
      }
    }
    previewTbody.replaceChildren(bodyFrag);

    if (totalMatches === 0) {
      pageInfo.textContent = 'Showing 0 rows';
    } else {
      pageInfo.textContent = `Showing ${startIdx + 1}–${startIdx + pageRows.length} of ${totalMatches} preview rows (Page ${currentPage} of ${totalPages})`;
    }
    prevPage.disabled = currentPage <= 1;
    nextPage.disabled = currentPage >= totalPages;
  }

  function renderResults(data) {
    sourceFilename.textContent = data.duk && data.duk.enabled
      ? `${data.sourceFile} + ${data.duk.sourceFile}`
      : data.sourceFile;
    sourceMeta.textContent = `Worksheet: "${data.sheetName}" • Filter column ${data.filter.columnLetter} (${data.filter.column}) • SQLite table: "${data.tableName}"`;

    statSourceRows.textContent = data.stats.totalSourceRows.toLocaleString();
    statFilteredRows.textContent = data.stats.filteredRowCount.toLocaleString();
    if (data.duk && data.duk.enabled) {
      statFilteredRowsSub.textContent = `${data.stats.baseExcelFilteredRows.toLocaleString()} base + ${data.stats.addedFromDukCount.toLocaleString()} DUK`;
      kpiDukAddedCard.classList.remove('hidden');
      statDukAdded.textContent = `+${data.stats.addedFromDukCount.toLocaleString()}`;
      statDukAddedSub.textContent = `${data.duk.alreadyInExcelCount.toLocaleString()} already in Excel`;
    } else {
      statFilteredRowsSub.textContent = 'Matched SATKER_3';
      kpiDukAddedCard.classList.add('hidden');
    }
    statRemovedCols.textContent = data.stats.removedColumnCount.toLocaleString();
    statKeptCols.textContent = data.stats.outputColumnCount.toLocaleString();

    processingDuration.textContent = `${data.stats.processingTimeMs} ms`;
    integrityChecksum.textContent = data.integrity.checksum;

    downloadXlsx.href = data.downloads.xlsx.url;
    downloadXlsx.download = data.downloads.xlsx.filename;
    sizeXlsx.textContent = formatBytes(data.downloads.xlsx.bytes);
    nameXlsx.textContent = data.downloads.xlsx.filename;

    downloadCsv.href = data.downloads.csv.url;
    downloadCsv.download = data.downloads.csv.filename;
    sizeCsv.textContent = formatBytes(data.downloads.csv.bytes);
    nameCsv.textContent = data.downloads.csv.filename;

    downloadSqlite.href = data.downloads.sqlite.url;
    downloadSqlite.download = data.downloads.sqlite.filename;
    sizeSqlite.textContent = formatBytes(data.downloads.sqlite.bytes);
    nameSqlite.textContent = `${data.downloads.sqlite.filename} (table: ${data.tableName})`;

    renderDukSummary(data.duk);

    colAuditCounts.textContent = `${data.stats.removedColumnCount} removed, ${data.stats.outputColumnCount} kept`;
    renderChipList(removedColumnsList, data.columns.removed, 'chip-removed');
    renderChipList(keptColumnsList, data.columns.kept, 'chip-kept');

    previewHeaders = data.preview.headers || [];
    previewAllRows = (data.preview.rows || []).map((row, idx) => ({ no: idx + 1, row }));
    previewFilteredRows = previewAllRows;
    currentPage = 1;
    previewSearch.value = '';

    if (data.preview.totalRows > previewAllRows.length) {
      previewSubtitle.textContent = `Showing first ${previewAllRows.length} of ${data.preview.totalRows.toLocaleString()} output rows. Download any file above for the full dataset.`;
    } else {
      previewSubtitle.textContent = `Showing all ${data.preview.totalRows.toLocaleString()} output rows.`;
    }

    renderPreviewTable();
    resultsPanel.classList.remove('hidden');
  }

  async function submitFiles(excelFile, dukFile) {
    if (!excelFile) {
      showError('Please choose a base Excel file (.xls or .xlsx).');
      return;
    }

    setProcessingState(true, excelFile, dukFile);

    try {
      const formData = new FormData();
      formData.append('file', excelFile);
      if (dukToggle.checked && dukFile) {
        formData.append('includeDuk', 'true');
        formData.append('dukFile', dukFile);
      }

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      if (response.status === 401) {
        window.location.assign('/login');
        return;
      }

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error || `Server returned status ${response.status}.`);
      }

      setProcessingState(false);
      renderResults(payload);
    } catch (err) {
      setProcessingState(false);
      uploadSection.classList.remove('hidden');
      showError(err.message || 'Failed to upload and process the uploaded files.');
    } finally {
      fileInput.value = '';
      dukFileInput.value = '';
    }
  }

  function handleExcelSelection(file) {
    if (!file) {
      return;
    }
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.xls') && !lowerName.endsWith('.xlsx')) {
      showError('Invalid base Excel file type. Please select a .xls or .xlsx file.');
      fileInput.value = '';
      return;
    }

    hideError();
    selectedExcelFile = file;
    updateStagedSelectionUI();

    if (!dukToggle.checked) {
      submitFiles(selectedExcelFile, null);
    }
  }

  function handleDukSelection(file) {
    if (!file) {
      return;
    }
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.html') && !lowerName.endsWith('.htm')) {
      showError('Invalid DUK support file type. Please select an .html or .htm file.');
      dukFileInput.value = '';
      return;
    }

    hideError();
    selectedDukFile = file;
    updateStagedSelectionUI();
  }

  logoutBtn.addEventListener('click', async () => {
    await fetch('/api/auth/logout', { method: 'POST' }).catch(() => {});
    window.location.assign('/login');
  });

  dukToggle.addEventListener('change', () => {
    hideError();
    syncDukToggleUI();
  });

  processCombineBtn.addEventListener('click', () => {
    if (!selectedExcelFile || !selectedDukFile) {
      showError('Please upload both the base Excel file and the DUK HTML support file.');
      return;
    }
    submitFiles(selectedExcelFile, selectedDukFile);
  });

  fileInput.addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      handleExcelSelection(file);
    }
  });

  dukFileInput.addEventListener('change', (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      handleDukSelection(file);
    }
  });

  dropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      fileInput.click();
    }
  });

  dukDropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      dukFileInput.click();
    }
  });

  dropzone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropzone.classList.add('drag-active');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-active');
  });

  dropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropzone.classList.remove('drag-active');
    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) {
      handleExcelSelection(file);
    }
  });

  dukDropzone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dukDropzone.classList.add('drag-active');
  });

  dukDropzone.addEventListener('dragleave', () => {
    dukDropzone.classList.remove('drag-active');
  });

  dukDropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    dukDropzone.classList.remove('drag-active');
    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) {
      handleDukSelection(file);
    }
  });

  window.addEventListener('dragover', (event) => event.preventDefault());
  window.addEventListener('drop', (event) => event.preventDefault());

  resetBtn.addEventListener('click', resetToUpload);

  previewSearch.addEventListener('input', () => {
    const query = previewSearch.value.trim().toLowerCase();
    if (!query) {
      previewFilteredRows = previewAllRows;
    } else {
      previewFilteredRows = previewAllRows.filter((entry) =>
        String(entry.no).includes(query) ||
        entry.row.some((cell) => String(cell).toLowerCase().includes(query))
      );
    }
    currentPage = 1;
    renderPreviewTable();
  });

  prevPage.addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage -= 1;
      renderPreviewTable();
    }
  });

  nextPage.addEventListener('click', () => {
    const totalPages = Math.max(1, Math.ceil(previewFilteredRows.length / PAGE_SIZE));
    if (currentPage < totalPages) {
      currentPage += 1;
      renderPreviewTable();
    }
  });

  syncDukToggleUI();
})();

