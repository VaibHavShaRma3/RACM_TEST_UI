// ─── Constants ──────────────────────────────────────────────────────────────
const RACM_FIELDS = [
  'Process Area', 'Sub-Process', 'Risk ID', 'Risk Description', 'Risk Category',
  'Risk Type', 'Control ID', 'Control Activity', 'Control Objective', 'Control Type',
  'Control Nature', 'Control Frequency', 'Control Owner',
  'Control description as per SOP', 'Testing Attributes', 'Evidence/Source',
  'Assertion Mapped', 'Compliance Reference', 'Risk Likelihood', 'Risk Impact',
  'Risk Rating', 'Mitigation Effectiveness', 'Gaps/Weaknesses Identified',
  'Source Quote', 'Extraction Confidence'
];

const FIELD_KEYS = RACM_FIELDS.map(f =>
  f.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/(^_|_$)/g, '')
);

const ACCEPTED_TYPES = ['.pdf', '.xlsx', '.xls', '.csv'];

const POLL_INTERVAL = 2000;

// ─── DOM Refs ───────────────────────────────────────────────────────────────
const $apiUrl = document.getElementById('api-url');
const $apiToken = document.getElementById('api-token');
const $btnHealth = document.getElementById('btn-health');
const $healthDot = document.getElementById('health-dot');
const $errorBanner = document.getElementById('error-banner');
const $dropZone = document.getElementById('drop-zone');
const $fileInput = document.getElementById('file-input');
const $fileNameDisplay = document.getElementById('file-name-display');
const $fileChipText = document.getElementById('file-chip-text');
const $promptInput = document.getElementById('prompt-input');
const $btnUpload = document.getElementById('btn-upload');
const $progressSection = document.getElementById('progress-section');
const $jobIdDisplay = document.getElementById('job-id-display');
const $progressFileName = document.getElementById('progress-file-name');
const $phaseBadge = document.getElementById('phase-badge');
const $progressMsg = document.getElementById('progress-msg');
const $progressBar = document.getElementById('progress-bar');
const $detailMsg = document.getElementById('detail-msg');
const $etaDisplay = document.getElementById('eta-display');
const $activityLog = document.getElementById('activity-log');
const $btnClearLog = document.getElementById('btn-clear-log');
const $resultsSection = document.getElementById('results-section');
const $filterRow = document.getElementById('filter-row');
const $headerRow = document.getElementById('header-row');
const $resultsBody = document.getElementById('results-body');
const $entryCount = document.getElementById('entry-count');
const $btnExportCsv = document.getElementById('btn-export-csv');
const $btnExportJson = document.getElementById('btn-export-json');
const $summarySection = document.getElementById('summary-section');
const $summaryContent = document.getElementById('summary-content');
const $summaryToggle = document.getElementById('summary-toggle');
const $btnCollapseSummary = document.getElementById('btn-collapse-summary');
const $detailModal = document.getElementById('detail-modal');
const $modalBody = document.getElementById('modal-body');
const $modalClose = document.getElementById('modal-close');
const $btnSettings = document.getElementById('btn-settings');
const $settingsPanel = document.getElementById('settings-panel');

// ─── State ──────────────────────────────────────────────────────────────────
let selectedFile = null;
let pollTimer = null;
let etaCountdownTimer = null;
let currentEtaSeconds = 0;
let rawResult = null;
let currentTab = 'detailed';
let currentEntries = [];
let columnFilters = {};
let lastDetailMsg = '';
let lastProgressMsg = '';
let lastPhase = '';

const PHASE_ORDER = ['extracting', 'chunking', 'analyzing', 'consolidating', 'deduplicating', 'summarizing'];

// ─── Helpers ────────────────────────────────────────────────────────────────
function getBaseUrl() {
  return $apiUrl.value.replace(/\/+$/, '');
}

function getHeaders() {
  return { 'Authorization': 'Bearer ' + $apiToken.value };
}

function showError(msg) {
  $errorBanner.textContent = msg;
  $errorBanner.classList.remove('hidden');
  setTimeout(() => $errorBanner.classList.add('hidden'), 8000);
}

function hideError() {
  $errorBanner.classList.add('hidden');
}

function getEntryValue(entry, fieldIndex) {
  const field = RACM_FIELDS[fieldIndex];
  const key = FIELD_KEYS[fieldIndex];
  return entry[field] ?? entry[key] ?? '';
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatEta(seconds) {
  if (seconds <= 0) return '';
  if (seconds < 60) return '< 1m';
  var m = Math.floor(seconds / 60);
  var s = Math.floor(seconds % 60);
  if (s === 0) return '~' + m + 'm';
  return '~' + m + 'm ' + s + 's';
}

function updateEtaDisplay() {
  if (currentEtaSeconds > 0) {
    $etaDisplay.textContent = formatEta(currentEtaSeconds) + ' remaining';
    $etaDisplay.classList.remove('hidden');
  } else {
    $etaDisplay.classList.add('hidden');
  }
}

function startEtaCountdown() {
  if (etaCountdownTimer) clearInterval(etaCountdownTimer);
  etaCountdownTimer = setInterval(function() {
    if (currentEtaSeconds > 0) {
      currentEtaSeconds = Math.max(0, currentEtaSeconds - 1);
      updateEtaDisplay();
    }
  }, 1000);
}

function stopEtaCountdown() {
  if (etaCountdownTimer) {
    clearInterval(etaCountdownTimer);
    etaCountdownTimer = null;
  }
  currentEtaSeconds = 0;
  $etaDisplay.classList.add('hidden');
}

// ─── Phase Steps ──────────────────────────────────────────────────────────
function updatePhaseSteps(currentPhase) {
  const currentIdx = PHASE_ORDER.indexOf(currentPhase);
  document.querySelectorAll('#phase-steps .step').forEach(function(stepEl) {
    var stepPhase = stepEl.dataset.phase;
    var stepIdx = PHASE_ORDER.indexOf(stepPhase);
    stepEl.classList.remove('active', 'done');
    if (stepIdx < 0) return;
    if (currentPhase === 'completed') {
      stepEl.classList.add('done');
    } else if (currentIdx >= 0) {
      if (stepIdx < currentIdx) stepEl.classList.add('done');
      else if (stepIdx === currentIdx) stepEl.classList.add('active');
    }
  });
}

function resetPhaseSteps() {
  document.querySelectorAll('#phase-steps .step').forEach(function(stepEl) {
    stepEl.classList.remove('active', 'done');
  });
}

// ─── Settings Toggle ──────────────────────────────────────────────────────
$btnSettings.addEventListener('click', function() {
  $settingsPanel.classList.toggle('hidden');
  $btnSettings.classList.toggle('active');
});

// ─── Summary Toggle ────────────────────────────────────────────────────────
$summaryToggle.addEventListener('click', function() {
  $summaryContent.classList.toggle('collapsed');
  $btnCollapseSummary.classList.toggle('collapsed');
});

function simpleMarkdown(text) {
  var html = escapeHtml(text);
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');
  html = html.replace(/\n(?!<)/g, '<br>');
  return html;
}

// ─── Activity Log ──────────────────────────────────────────────────────────
function logTime() {
  const now = new Date();
  return now.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

function addLogEntry(phase, msg, type) {
  const entry = document.createElement('div');
  entry.className = 'log-entry';

  const timeSpan = '<span class="log-time">' + logTime() + '</span>';
  const phaseSpan = phase ? '<span class="log-phase">[' + escapeHtml(phase.toUpperCase()) + ']</span>' : '';
  const msgClass = type === 'error' ? 'log-error' : type === 'complete' ? 'log-complete' : type === 'detail' ? 'log-detail' : 'log-msg';
  const msgSpan = '<span class="' + msgClass + '">' + escapeHtml(msg) + '</span>';

  entry.innerHTML = timeSpan + phaseSpan + msgSpan;
  $activityLog.appendChild(entry);
  $activityLog.scrollTop = $activityLog.scrollHeight;
}

$btnClearLog.addEventListener('click', function(e) {
  e.stopPropagation();
  $activityLog.innerHTML = '';
});

// ─── Health Check ───────────────────────────────────────────────────────────
$btnHealth.addEventListener('click', async () => {
  $healthDot.className = 'health-dot';
  $healthDot.title = 'Checking...';
  try {
    const res = await fetch(getBaseUrl() + '/health');
    if (res.ok) {
      $healthDot.classList.add('ok');
      $healthDot.title = 'Healthy';
    } else {
      $healthDot.classList.add('fail');
      $healthDot.title = 'Unhealthy: ' + res.status;
    }
  } catch (e) {
    $healthDot.classList.add('fail');
    $healthDot.title = 'Connection failed';
    showError('Health check failed: ' + e.message);
  }
});

// ─── File Selection ─────────────────────────────────────────────────────────
$dropZone.addEventListener('click', () => $fileInput.click());

$dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  $dropZone.classList.add('dragover');
});

$dropZone.addEventListener('dragleave', () => {
  $dropZone.classList.remove('dragover');
});

$dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  $dropZone.classList.remove('dragover');
  if (e.dataTransfer.files.length) {
    handleFileSelect(e.dataTransfer.files[0]);
  }
});

$fileInput.addEventListener('change', () => {
  if ($fileInput.files.length) {
    handleFileSelect($fileInput.files[0]);
  }
});

function handleFileSelect(file) {
  const ext = '.' + file.name.split('.').pop().toLowerCase();
  if (!ACCEPTED_TYPES.includes(ext)) {
    showError('Unsupported file type: ' + ext);
    return;
  }
  selectedFile = file;
  $fileChipText.textContent = file.name + ' (' + formatBytes(file.size) + ')';
  $fileNameDisplay.classList.remove('hidden');
  $btnUpload.disabled = false;
}

// ─── Upload ─────────────────────────────────────────────────────────────────
$btnUpload.addEventListener('click', async () => {
  if (!selectedFile) return;
  hideError();

  const formData = new FormData();
  formData.append('file', selectedFile);
  const prompt = $promptInput.value.trim();
  if (prompt) formData.append('prompt', prompt);

  var uploadSvg = $btnUpload.querySelector('svg');
  var uploadSvgHtml = uploadSvg ? uploadSvg.outerHTML : '';
  $btnUpload.classList.add('loading');
  $btnUpload.innerHTML = uploadSvgHtml + ' Uploading...';

  try {
    const res = await fetch(getBaseUrl() + '/api/jobs', {
      method: 'POST',
      headers: getHeaders(),
      body: formData
    });

    if (!res.ok) {
      const errText = await res.text().catch(() => '');
      if (res.status === 401) showError('Authentication failed (401). Check your API token.');
      else if (res.status === 413) showError('File too large (413).');
      else if (res.status === 400) showError('Bad request (400): ' + errText);
      else showError('Upload failed (' + res.status + '): ' + errText);
      return;
    }

    const data = await res.json();
    const jobId = data.job_id;

    // Show progress section
    $progressSection.classList.remove('hidden');
    $jobIdDisplay.textContent = jobId;
    $progressFileName.textContent = selectedFile.name;
    $phaseBadge.textContent = 'queued';
    $phaseBadge.className = 'phase-badge queued';
    $progressMsg.textContent = 'Waiting...';
    $progressBar.style.width = '0%';
    updateProgressGlow(0);
    $detailMsg.textContent = '';
    stopEtaCountdown();
    resetPhaseSteps();

    // Reset log state
    lastDetailMsg = '';
    lastProgressMsg = '';
    lastPhase = '';
    $activityLog.innerHTML = '';
    addLogEntry('system', 'Job submitted: ' + jobId, 'msg');
    addLogEntry('system', 'File: ' + selectedFile.name + ' (' + formatBytes(selectedFile.size) + ')', 'msg');

    // Hide results and summary from previous run
    $resultsSection.classList.add('hidden');
    $summarySection.classList.add('hidden');

    startPolling(jobId);
  } catch (e) {
    showError('Upload error: ' + e.message);
  } finally {
    $btnUpload.classList.remove('loading');
    $btnUpload.innerHTML = uploadSvgHtml + ' Upload & Analyze';
  }
});

// ─── Polling ────────────────────────────────────────────────────────────────
function startPolling(jobId) {
  if (pollTimer) clearInterval(pollTimer);

  pollTimer = setInterval(async () => {
    try {
      const res = await fetch(getBaseUrl() + '/api/jobs/' + jobId + '/status', {
        headers: getHeaders()
      });

      if (!res.ok) {
        showError('Status poll failed: ' + res.status);
        return;
      }

      const data = await res.json();
      const phase = (data.phase || 'unknown').toLowerCase();
      const pct = data.progress_pct ?? 0;
      const msg = data.progress_msg || '';
      const detail = data.detail_msg || '';
      const etaSec = data.eta_seconds ?? 0;

      // Update UI elements
      $phaseBadge.textContent = phase;
      $phaseBadge.className = 'phase-badge ' + phase;
      $progressMsg.textContent = msg;
      $progressBar.style.width = pct + '%';
  
      $detailMsg.textContent = detail;
      updatePhaseSteps(phase);

      // Update ETA — sync from server and start/continue countdown
      if (etaSec > 0 && phase !== 'completed' && phase !== 'failed') {
        currentEtaSeconds = etaSec;
        updateEtaDisplay();
        if (!etaCountdownTimer) startEtaCountdown();
      } else if (phase === 'completed' || phase === 'failed') {
        stopEtaCountdown();
      }

      // Log phase transitions
      if (phase !== lastPhase && phase !== 'unknown') {
        addLogEntry(phase, 'Phase started: ' + phase, 'msg');
        lastPhase = phase;
      }

      // Log new progress messages
      if (msg && msg !== lastProgressMsg) {
        addLogEntry(phase, msg, 'msg');
        lastProgressMsg = msg;
      }

      // Log new detail messages
      if (detail && detail !== lastDetailMsg) {
        addLogEntry(phase, detail, 'detail');
        lastDetailMsg = detail;
      }

      if (phase === 'completed') {
        clearInterval(pollTimer);
        pollTimer = null;
        stopEtaCountdown();
        addLogEntry('done', 'Job completed successfully!', 'complete');
        fetchResults(jobId);
      } else if (phase === 'failed') {
        clearInterval(pollTimer);
        pollTimer = null;
        stopEtaCountdown();
        addLogEntry('error', 'Job failed: ' + msg, 'error');
        showError('Job failed: ' + msg);
      }
    } catch (e) {
      showError('Polling error: ' + e.message);
    }
  }, POLL_INTERVAL);
}

// ─── Fetch Results ──────────────────────────────────────────────────────────
async function fetchResults(jobId) {
  try {
    const res = await fetch(getBaseUrl() + '/api/jobs/' + jobId + '/result', {
      headers: getHeaders()
    });

    if (!res.ok) {
      showError('Failed to fetch results: ' + res.status);
      return;
    }

    rawResult = await res.json();
    currentTab = 'detailed';
    document.querySelectorAll('.tab').forEach(t => {
      t.classList.toggle('active', t.dataset.tab === 'detailed');
    });

    const result = rawResult.result || rawResult;
    const nDetailed = (result.detailed_entries || []).length;
    const nSummary = (result.summary_entries || []).length;
    addLogEntry('result', 'Loaded ' + nDetailed + ' detailed + ' + nSummary + ' summary entries', 'complete');

    // Display summary narrative if available
    var narrative = result.summary_narrative || '';
    if (narrative) {
      $summaryContent.innerHTML = simpleMarkdown(narrative);
      $summaryContent.classList.remove('collapsed');
      $btnCollapseSummary.classList.remove('collapsed');
      $summarySection.classList.remove('hidden');
      addLogEntry('result', 'Executive summary loaded', 'complete');
    } else {
      $summarySection.classList.add('hidden');
    }

    renderResults();
    $resultsSection.classList.remove('hidden');
    $resultsSection.scrollIntoView({ behavior: 'smooth' });
  } catch (e) {
    showError('Results fetch error: ' + e.message);
  }
}

// ─── Render Results Table ───────────────────────────────────────────────────
function renderResults() {
  if (!rawResult) return;

  const result = rawResult.result || rawResult;

  const entries = currentTab === 'detailed'
    ? (result.detailed_entries || [])
    : (result.summary_entries || []);

  currentEntries = entries;

  $filterRow.innerHTML = '';
  $headerRow.innerHTML = '';
  columnFilters = {};

  RACM_FIELDS.forEach((field, i) => {
    const fth = document.createElement('th');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = 'Filter...';
    inp.dataset.colIndex = i;
    inp.addEventListener('input', onFilterChange);
    fth.appendChild(inp);
    $filterRow.appendChild(fth);

    const th = document.createElement('th');
    th.textContent = field;
    $headerRow.appendChild(th);
  });

  renderRows(entries);
}

function renderRows(entries) {
  $resultsBody.innerHTML = '';

  const filtered = entries.filter(entry => {
    return Object.entries(columnFilters).every(([colIdx, filterVal]) => {
      if (!filterVal) return true;
      const val = String(getEntryValue(entry, Number(colIdx))).toLowerCase();
      return val.includes(filterVal.toLowerCase());
    });
  });

  $entryCount.textContent = filtered.length + ' of ' + entries.length + ' entries';

  filtered.forEach((entry) => {
    const tr = document.createElement('tr');
    tr.addEventListener('click', () => openDetailModal(entry));

    RACM_FIELDS.forEach((field, colIdx) => {
      const td = document.createElement('td');
      const val = String(getEntryValue(entry, colIdx));
      td.textContent = val;
      td.title = val;

      if (field === 'Source Quote') {
        td.classList.add('col-source-quote');
      } else if (field === 'Extraction Confidence') {
        td.classList.add('col-extraction-confidence');
        const lower = val.toLowerCase();
        if (lower === 'extracted') td.classList.add('extracted');
        else if (lower === 'inferred') td.classList.add('inferred');
        else if (lower === 'partial') td.classList.add('partial');
      }

      tr.appendChild(td);
    });

    $resultsBody.appendChild(tr);
  });
}

function onFilterChange(e) {
  const colIdx = e.target.dataset.colIndex;
  const val = e.target.value;
  columnFilters[colIdx] = val;
  renderRows(currentEntries);
}

// ─── Tab Toggle ─────────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    currentTab = tab.dataset.tab;
    $filterRow.querySelectorAll('input').forEach(inp => inp.value = '');
    columnFilters = {};
    renderResults();
  });
});

// ─── Detail Modal ───────────────────────────────────────────────────────────
function openDetailModal(entry) {
  $modalBody.innerHTML = '';

  RACM_FIELDS.forEach((field, i) => {
    const val = String(getEntryValue(entry, i));
    const div = document.createElement('div');
    div.className = 'detail-field';
    if (field === 'Source Quote' || field === 'Extraction Confidence') {
      div.classList.add('highlight');
    }

    div.innerHTML =
      '<div class="detail-label">' + escapeHtml(field) + '</div>' +
      '<div class="detail-value">' + escapeHtml(val) + '</div>';

    $modalBody.appendChild(div);
  });

  $detailModal.classList.remove('hidden');
}

$modalClose.addEventListener('click', () => {
  $detailModal.classList.add('hidden');
});

$detailModal.addEventListener('click', (e) => {
  if (e.target === $detailModal) $detailModal.classList.add('hidden');
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') $detailModal.classList.add('hidden');
});

// ─── Export ─────────────────────────────────────────────────────────────────
function getExportBaseName() {
  var sopName = (selectedFile ? selectedFile.name : 'document')
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-zA-Z0-9_\-]+/g, '_')
    .replace(/(^_|_$)/g, '');
  var tabLabel = currentTab === 'summary' ? 'summary' : 'detailed';
  return 'RACM_' + sopName + '_' + tabLabel;
}

$btnExportCsv.addEventListener('click', () => {
  if (!currentEntries.length) return;

  const filtered = currentEntries.filter(entry => {
    return Object.entries(columnFilters).every(([colIdx, filterVal]) => {
      if (!filterVal) return true;
      const val = String(getEntryValue(entry, Number(colIdx))).toLowerCase();
      return val.includes(filterVal.toLowerCase());
    });
  });

  const rows = [RACM_FIELDS.map(csvEscape).join(',')];
  filtered.forEach(entry => {
    const row = RACM_FIELDS.map((_, i) => csvEscape(String(getEntryValue(entry, i))));
    rows.push(row.join(','));
  });

  downloadFile(getExportBaseName() + '.csv', rows.join('\n'), 'text/csv');
});

$btnExportJson.addEventListener('click', () => {
  if (!rawResult) return;
  downloadFile(getExportBaseName() + '.json', JSON.stringify(rawResult, null, 2), 'application/json');
});

function csvEscape(val) {
  if (val.includes(',') || val.includes('"') || val.includes('\n')) {
    return '"' + val.replace(/"/g, '""') + '"';
  }
  return val;
}

function downloadFile(name, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
