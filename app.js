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

const RISK_RATING_ORDER = { 'critical': 4, 'high': 3, 'medium': 2, 'low': 1 };

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
const $activityLog = document.getElementById('activity-log');
const $btnClearLog = document.getElementById('btn-clear-log');
const $resultsSection = document.getElementById('results-section');
const $filterRow = document.getElementById('filter-row');
const $headerRow = document.getElementById('header-row');
const $resultsBody = document.getElementById('results-body');
const $entryCount = document.getElementById('entry-count');
const $btnExportCsv = document.getElementById('btn-export-csv');
const $btnExportJson = document.getElementById('btn-export-json');
const $btnExportXlsx = document.getElementById('btn-export-xlsx');
const $summarySection = document.getElementById('summary-section');
const $summaryContent = document.getElementById('summary-content');
const $summaryToggle = document.getElementById('summary-toggle');
const $btnCollapseSummary = document.getElementById('btn-collapse-summary');
const $detailModal = document.getElementById('detail-modal');
const $modalBody = document.getElementById('modal-body');
const $modalClose = document.getElementById('modal-close');
const $settingsPanel = document.getElementById('settings-panel');
const $btnClearFile = document.getElementById('btn-clear-file');
const $navSettings = document.getElementById('nav-settings');
const $sidebar = document.getElementById('sidebar');
const $btnSidebarToggle = document.getElementById('btn-sidebar-toggle');
const $btnCancelJob = document.getElementById('btn-cancel-job');
const $btnDeleteReport = document.getElementById('btn-delete-report');
const $editActions = document.getElementById('edit-actions');
const $btnSaveChanges = document.getElementById('btn-save-changes');
const $btnDiscardChanges = document.getElementById('btn-discard-changes');
const $paginationControls = document.getElementById('pagination-controls');
const $pageInfo = document.getElementById('page-info');
const $btnPrevPage = document.getElementById('btn-prev-page');
const $btnNextPage = document.getElementById('btn-next-page');
const $uploadSection = document.getElementById('upload-section');

// ─── State ──────────────────────────────────────────────────────────────────
let selectedFile = null;
let pollTimer = null;
let rawResult = null;
let currentTab = 'detailed';
let currentEntries = [];
let columnFilters = {};
let lastDetailMsg = '';
let lastProgressMsg = '';
let lastPhase = '';
let isProcessing = false;
let currentJobId = null;

// Sorting state
let sortCol = -1;
let sortAsc = true;

// Pagination state
let pageSize = 25;
let currentPageNum = 0;

// Inline editing state
let pendingEdits = {};

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

function setEntryValue(entry, fieldIndex, value) {
  const field = RACM_FIELDS[fieldIndex];
  const key = FIELD_KEYS[fieldIndex];
  if (field in entry) {
    entry[field] = value;
  } else {
    entry[key] = value;
  }
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
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
$navSettings.addEventListener('click', function(e) {
  e.preventDefault();
  $settingsPanel.classList.toggle('hidden');
  $navSettings.classList.toggle('active');
});

// ─── Sidebar Mobile Toggle ────────────────────────────────────────────────
$btnSidebarToggle.addEventListener('click', function() {
  $sidebar.classList.toggle('open');
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
  $btnUpload.disabled = isProcessing;
}

function clearFileSelection() {
  selectedFile = null;
  $fileInput.value = '';
  $fileNameDisplay.classList.add('hidden');
  $fileChipText.textContent = '';
  $btnUpload.disabled = true;
}

$btnClearFile.addEventListener('click', function(e) {
  e.stopPropagation();
  clearFileSelection();
});

// ─── Upload ─────────────────────────────────────────────────────────────────
$btnUpload.addEventListener('click', async () => {
  if (!selectedFile || isProcessing) return;
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
    currentJobId = jobId;

    // Show progress section
    $progressSection.classList.remove('hidden');
    $jobIdDisplay.textContent = jobId;
    $progressFileName.textContent = selectedFile.name;
    $phaseBadge.textContent = 'queued';
    $phaseBadge.className = 'phase-badge queued';
    $progressMsg.textContent = 'Waiting...';
    $progressBar.style.width = '0%';
    $detailMsg.textContent = '';
    resetPhaseSteps();

    // Reset log state
    lastDetailMsg = '';
    lastProgressMsg = '';
    lastPhase = '';
    $activityLog.innerHTML = '';
    addLogEntry('system', 'Job submitted: ' + jobId, 'msg');
    addLogEntry('system', 'File: ' + selectedFile.name + ' (' + formatBytes(selectedFile.size) + ')', 'msg');

    // Lock upload during processing
    isProcessing = true;
    $btnUpload.disabled = true;

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

// ─── Cancel Job ─────────────────────────────────────────────────────────────
$btnCancelJob.addEventListener('click', async () => {
  if (!currentJobId) return;
  if (!confirm('Are you sure? This will stop the analysis and delete the job.')) return;

  try {
    const res = await fetch(getBaseUrl() + '/api/jobs/' + currentJobId, {
      method: 'DELETE',
      headers: getHeaders(),
    });

    if (pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }
    isProcessing = false;
    $btnUpload.disabled = !selectedFile;

    // Update UI to show cancelled state
    $phaseBadge.textContent = 'cancelled';
    $phaseBadge.className = 'phase-badge failed';
    $progressMsg.textContent = 'Cancelled by user';
    addLogEntry('system', 'Job cancelled by user', 'error');

    // After 2s, hide progress and show upload
    setTimeout(() => {
      $progressSection.classList.add('hidden');
    }, 2000);
  } catch (e) {
    showError('Cancel failed: ' + e.message);
  }
});

// ─── Delete Report ──────────────────────────────────────────────────────────
$btnDeleteReport.addEventListener('click', async () => {
  if (!currentJobId) return;
  if (!confirm('This will permanently delete this report. Continue?')) return;

  try {
    const res = await fetch(getBaseUrl() + '/api/jobs/' + currentJobId, {
      method: 'DELETE',
      headers: getHeaders(),
    });

    if (res.ok) {
      rawResult = null;
      currentEntries = [];
      currentJobId = null;
      pendingEdits = {};
      $editActions.classList.add('hidden');
      $resultsSection.classList.add('hidden');
      $summarySection.classList.add('hidden');
      $progressSection.classList.add('hidden');
      showError('Report deleted successfully');
    } else {
      showError('Delete failed: ' + res.status);
    }
  } catch (e) {
    showError('Delete failed: ' + e.message);
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

      // Update UI elements
      $phaseBadge.textContent = phase;
      $phaseBadge.className = 'phase-badge ' + phase;
      $progressMsg.textContent = msg;
      $progressBar.style.width = pct + '%';
      $detailMsg.textContent = detail;
      updatePhaseSteps(phase);

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
        isProcessing = false;
        $btnUpload.disabled = !selectedFile;
        addLogEntry('done', 'Job completed successfully!', 'complete');
        fetchResults(jobId);
      } else if (phase === 'failed') {
        clearInterval(pollTimer);
        pollTimer = null;
        isProcessing = false;
        $btnUpload.disabled = !selectedFile;
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
    currentJobId = jobId;
    currentTab = 'detailed';
    pendingEdits = {};
    $editActions.classList.add('hidden');
    document.querySelectorAll('.tab').forEach(t => {
      t.classList.toggle('active', t.dataset.tab === 'detailed');
    });

    const result = rawResult.result || rawResult;
    const nDetailed = (result.detailed_entries || []).length;
    const nSummary = (result.summary_entries || []).length;
    addLogEntry('result', 'Loaded ' + nDetailed + ' detailed + ' + nSummary + ' summary entries', 'complete');

    // Render visual summary dashboard
    renderSummaryDashboard(result);

    renderResults();
    $resultsSection.classList.remove('hidden');
    $resultsSection.scrollIntoView({ behavior: 'smooth' });
  } catch (e) {
    showError('Results fetch error: ' + e.message);
  }
}

// ─── Visual Summary Dashboard (F3) ─────────────────────────────────────────
function renderSummaryDashboard(result) {
  const entries = result.detailed_entries || [];
  if (entries.length === 0) {
    $summarySection.classList.add('hidden');
    return;
  }

  // Compute stats from detailed_entries
  const totalRisks = entries.length;
  const controlIds = new Set();
  const processAreas = new Set();
  const riskRatings = {};
  const controlTypes = {};
  const processAreaData = {};

  entries.forEach(entry => {
    const controlId = entry['Control ID'] || entry['control_id'] || '';
    const processArea = entry['Process Area'] || entry['process_area'] || '';
    const riskRating = entry['Risk Rating'] || entry['risk_rating'] || '';
    const controlType = entry['Control Type'] || entry['control_type'] || '';

    if (controlId) controlIds.add(controlId);
    if (processArea) processAreas.add(processArea);

    if (riskRating) {
      riskRatings[riskRating] = (riskRatings[riskRating] || 0) + 1;
    }
    if (controlType) {
      controlTypes[controlType] = (controlTypes[controlType] || 0) + 1;
    }

    if (processArea) {
      if (!processAreaData[processArea]) {
        processAreaData[processArea] = { risks: 0, controls: new Set(), topRisk: '' };
      }
      processAreaData[processArea].risks++;
      if (controlId) processAreaData[processArea].controls.add(controlId);
      const current = processAreaData[processArea].topRisk;
      if ((RISK_RATING_ORDER[riskRating.toLowerCase()] || 0) > (RISK_RATING_ORDER[current.toLowerCase()] || 0)) {
        processAreaData[processArea].topRisk = riskRating;
      }
    }
  });

  // Build dashboard HTML
  let html = '';

  // Stat cards
  html += '<div class="dashboard-stats">';
  html += '<div class="stat-card"><div class="stat-number">' + totalRisks + '</div><div class="stat-label">Total Risks</div></div>';
  html += '<div class="stat-card"><div class="stat-number">' + controlIds.size + '</div><div class="stat-label">Total Controls</div></div>';
  html += '<div class="stat-card"><div class="stat-number">' + processAreas.size + '</div><div class="stat-label">Process Areas</div></div>';
  html += '</div>';

  // Risk Rating distribution bar chart
  const ratingOrder = ['Critical', 'High', 'Medium', 'Low'];
  const ratingColors = { 'Critical': '#ef4444', 'High': '#f59e0b', 'Medium': '#3b82f6', 'Low': '#10b981' };
  const maxRating = Math.max(...Object.values(riskRatings), 1);

  html += '<div class="dashboard-section"><h4 class="dashboard-section-title">Risk Rating Distribution</h4>';
  html += '<div class="bar-chart">';
  ratingOrder.forEach(rating => {
    const count = 0;
    // Case-insensitive match
    let matchedCount = 0;
    Object.entries(riskRatings).forEach(([key, val]) => {
      if (key.toLowerCase() === rating.toLowerCase()) matchedCount += val;
    });
    if (matchedCount > 0 || true) {
      const pct = maxRating > 0 ? (matchedCount / maxRating) * 100 : 0;
      const color = ratingColors[rating] || '#9ca3af';
      html += '<div class="bar-row">';
      html += '<span class="bar-label">' + escapeHtml(rating) + '</span>';
      html += '<div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;background:' + color + '"></div></div>';
      html += '<span class="bar-count">' + matchedCount + '</span>';
      html += '</div>';
    }
  });
  html += '</div></div>';

  // Control type breakdown
  if (Object.keys(controlTypes).length > 0) {
    html += '<div class="dashboard-section"><h4 class="dashboard-section-title">Control Type Breakdown</h4>';
    html += '<div class="dashboard-stats control-type-stats">';
    Object.entries(controlTypes).forEach(([type, count]) => {
      html += '<div class="stat-card stat-card-sm"><div class="stat-number">' + count + '</div><div class="stat-label">' + escapeHtml(type) + '</div></div>';
    });
    html += '</div></div>';
  }

  // Process area breakdown table
  if (Object.keys(processAreaData).length > 0) {
    html += '<div class="dashboard-section"><h4 class="dashboard-section-title">By Process Area</h4>';
    html += '<table class="breakdown-table"><thead><tr><th>Process Area</th><th>Risks</th><th>Controls</th><th>Top Risk</th></tr></thead><tbody>';
    Object.entries(processAreaData)
      .sort((a, b) => b[1].risks - a[1].risks)
      .forEach(([area, data]) => {
        html += '<tr><td>' + escapeHtml(area) + '</td><td>' + data.risks + '</td><td>' + data.controls.size + '</td><td>' + escapeHtml(data.topRisk) + '</td></tr>';
      });
    html += '</tbody></table></div>';
  }

  // Collapsible AI summary narrative
  const narrative = result.summary_narrative || '';
  if (narrative) {
    html += '<details class="ai-summary-details"><summary class="ai-summary-toggle">AI Summary</summary>';
    html += '<div class="ai-summary-content">' + simpleMarkdown(narrative) + '</div>';
    html += '</details>';
  }

  $summaryContent.innerHTML = html;
  $summaryContent.classList.remove('collapsed');
  $btnCollapseSummary.classList.remove('collapsed');
  $summarySection.classList.remove('hidden');
  addLogEntry('result', 'Executive summary loaded', 'complete');
}

// ─── Render Results Table ───────────────────────────────────────────────────
function renderResults() {
  if (!rawResult) return;

  const result = rawResult.result || rawResult;

  const entries = currentTab === 'detailed'
    ? (result.detailed_entries || [])
    : (result.summary_entries || []);

  currentEntries = entries;

  // Reset sorting and pagination on tab/data change
  sortCol = -1;
  sortAsc = true;
  currentPageNum = 0;

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
    th.className = 'sortable';
    th.dataset.colIndex = i;
    th.addEventListener('click', onSortClick);
    $headerRow.appendChild(th);
  });

  renderRows(entries);
}

function getFilteredEntries(entries) {
  return entries.filter(entry => {
    return Object.entries(columnFilters).every(([colIdx, filterVal]) => {
      if (!filterVal) return true;
      const val = String(getEntryValue(entry, Number(colIdx))).toLowerCase();
      return val.includes(filterVal.toLowerCase());
    });
  });
}

function getSortedEntries(entries) {
  if (sortCol < 0) return entries;
  const sorted = [...entries];
  const field = RACM_FIELDS[sortCol];
  const isRiskRating = field === 'Risk Rating';

  sorted.sort((a, b) => {
    const aVal = String(getEntryValue(a, sortCol));
    const bVal = String(getEntryValue(b, sortCol));

    if (isRiskRating) {
      const aNum = RISK_RATING_ORDER[aVal.toLowerCase()] || 0;
      const bNum = RISK_RATING_ORDER[bVal.toLowerCase()] || 0;
      return sortAsc ? aNum - bNum : bNum - aNum;
    }
    return sortAsc ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });
  return sorted;
}

function renderRows(entries) {
  $resultsBody.innerHTML = '';

  let filtered = getFilteredEntries(entries);
  let sorted = getSortedEntries(filtered);

  // Pagination
  const totalItems = sorted.length;
  const effectivePageSize = pageSize === 'all' ? totalItems : pageSize;
  const totalPages = effectivePageSize > 0 ? Math.max(1, Math.ceil(totalItems / effectivePageSize)) : 1;
  if (currentPageNum >= totalPages) currentPageNum = totalPages - 1;
  if (currentPageNum < 0) currentPageNum = 0;

  const start = currentPageNum * effectivePageSize;
  const end = Math.min(start + effectivePageSize, totalItems);
  const pageEntries = sorted.slice(start, end);

  // Update counts and pagination info
  if (totalItems !== entries.length) {
    $entryCount.textContent = 'Showing ' + (start + 1) + '-' + end + ' of ' + totalItems + ' filtered (' + entries.length + ' total)';
  } else {
    $entryCount.textContent = 'Showing ' + (totalItems > 0 ? start + 1 : 0) + '-' + end + ' of ' + totalItems + ' entries';
  }

  $pageInfo.textContent = 'Page ' + (currentPageNum + 1) + ' of ' + totalPages;
  $btnPrevPage.disabled = currentPageNum <= 0;
  $btnNextPage.disabled = currentPageNum >= totalPages - 1;

  // Map from sorted entry back to original index for inline editing
  const result = rawResult.result || rawResult;
  const sourceEntries = currentTab === 'detailed'
    ? (result.detailed_entries || [])
    : (result.summary_entries || []);

  pageEntries.forEach((entry) => {
    const entryIdx = sourceEntries.indexOf(entry);
    const tr = document.createElement('tr');
    tr.addEventListener('click', (e) => {
      // Don't open modal if we're editing
      if (e.target.isContentEditable) return;
      openDetailModal(entry);
    });

    RACM_FIELDS.forEach((field, colIdx) => {
      const td = document.createElement('td');
      const val = String(getEntryValue(entry, colIdx));
      td.textContent = val;
      td.title = val;

      // Check if this cell has a pending edit
      const editKey = entryIdx + ':' + colIdx;
      if (pendingEdits[editKey] !== undefined) {
        td.classList.add('edited');
        td.textContent = pendingEdits[editKey];
      }

      if (field === 'Source Quote') {
        td.classList.add('col-source-quote');
      } else if (field === 'Extraction Confidence') {
        td.classList.add('col-extraction-confidence');
        const lower = (pendingEdits[editKey] !== undefined ? pendingEdits[editKey] : val).toLowerCase();
        if (lower === 'extracted') td.classList.add('extracted');
        else if (lower === 'inferred') td.classList.add('inferred');
        else if (lower === 'partial') td.classList.add('partial');
      }

      // Double-click to edit
      td.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        startCellEdit(td, entryIdx, colIdx);
      });

      tr.appendChild(td);
    });

    $resultsBody.appendChild(tr);
  });
}

// ─── Column Sorting (F6) ───────────────────────────────────────────────────
function onSortClick(e) {
  const colIdx = Number(e.currentTarget.dataset.colIndex);
  if (sortCol === colIdx) {
    sortAsc = !sortAsc;
  } else {
    sortCol = colIdx;
    sortAsc = true;
  }

  // Update header classes
  $headerRow.querySelectorAll('th').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
  });
  if (sortCol >= 0) {
    e.currentTarget.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');
  }

  currentPageNum = 0;
  renderRows(currentEntries);
}

function onFilterChange(e) {
  const colIdx = e.target.dataset.colIndex;
  const val = e.target.value;
  columnFilters[colIdx] = val;
  currentPageNum = 0;
  renderRows(currentEntries);
}

// ─── Pagination (F6) ───────────────────────────────────────────────────────
$btnPrevPage.addEventListener('click', () => {
  if (currentPageNum > 0) {
    currentPageNum--;
    renderRows(currentEntries);
  }
});

$btnNextPage.addEventListener('click', () => {
  currentPageNum++;
  renderRows(currentEntries);
});

document.querySelectorAll('.page-size-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.page-size-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const size = btn.dataset.size;
    pageSize = size === 'all' ? 'all' : Number(size);
    currentPageNum = 0;
    renderRows(currentEntries);
  });
});

// ─── Inline Cell Editing (F2) ──────────────────────────────────────────────
function startCellEdit(td, entryIdx, colIdx) {
  if (td.isContentEditable) return;
  td.contentEditable = true;
  td.classList.add('cell-editing');
  td.focus();

  // Select all text
  const range = document.createRange();
  range.selectNodeContents(td);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);

  function finishEdit() {
    td.contentEditable = false;
    td.classList.remove('cell-editing');
    td.removeEventListener('blur', onBlur);
    td.removeEventListener('keydown', onKeydown);

    const newVal = td.textContent.trim();
    const editKey = entryIdx + ':' + colIdx;
    const originalVal = String(getEntryValue(
      (currentTab === 'detailed'
        ? (rawResult.result || rawResult).detailed_entries
        : (rawResult.result || rawResult).summary_entries
      )[entryIdx], colIdx));

    if (newVal !== originalVal) {
      pendingEdits[editKey] = newVal;
      td.classList.add('edited');
    } else {
      delete pendingEdits[editKey];
      td.classList.remove('edited');
    }

    updateEditActionsVisibility();
  }

  function onBlur() { finishEdit(); }
  function onKeydown(e) {
    if (e.key === 'Enter') { e.preventDefault(); td.blur(); }
    if (e.key === 'Escape') {
      // Revert to original
      const editKey = entryIdx + ':' + colIdx;
      const originalVal = String(getEntryValue(
        (currentTab === 'detailed'
          ? (rawResult.result || rawResult).detailed_entries
          : (rawResult.result || rawResult).summary_entries
        )[entryIdx], colIdx));
      td.textContent = pendingEdits[editKey] !== undefined ? pendingEdits[editKey] : originalVal;
      td.contentEditable = false;
      td.classList.remove('cell-editing');
      td.removeEventListener('blur', onBlur);
      td.removeEventListener('keydown', onKeydown);
    }
  }

  td.addEventListener('blur', onBlur);
  td.addEventListener('keydown', onKeydown);
}

function updateEditActionsVisibility() {
  if (Object.keys(pendingEdits).length > 0) {
    $editActions.classList.remove('hidden');
  } else {
    $editActions.classList.add('hidden');
  }
}

// Save Changes
$btnSaveChanges.addEventListener('click', async () => {
  if (!currentJobId || Object.keys(pendingEdits).length === 0) return;

  const result = rawResult.result || rawResult;
  const detailedEntries = [...(result.detailed_entries || [])];
  const summaryEntries = [...(result.summary_entries || [])];

  // Apply pending edits to the correct array
  Object.entries(pendingEdits).forEach(([key, value]) => {
    const [entryIdx, colIdx] = key.split(':').map(Number);
    const targetEntries = currentTab === 'detailed' ? detailedEntries : summaryEntries;
    if (entryIdx >= 0 && entryIdx < targetEntries.length) {
      setEntryValue(targetEntries[entryIdx], colIdx, value);
    }
  });

  // Convert entries to plain objects for the API
  const detailedPlain = detailedEntries.map(e => {
    const obj = {};
    RACM_FIELDS.forEach((field, i) => { obj[field] = String(getEntryValue(e, i)); });
    return obj;
  });
  const summaryPlain = summaryEntries.map(e => {
    const obj = {};
    RACM_FIELDS.forEach((field, i) => { obj[field] = String(getEntryValue(e, i)); });
    return obj;
  });

  try {
    $btnSaveChanges.disabled = true;
    $btnSaveChanges.textContent = 'Saving...';

    const res = await fetch(getBaseUrl() + '/api/jobs/' + currentJobId + '/result', {
      method: 'PUT',
      headers: { ...getHeaders(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ detailed_entries: detailedPlain, summary_entries: summaryPlain }),
    });

    if (res.ok) {
      // Update local state
      if (rawResult.result) {
        rawResult.result.detailed_entries = detailedEntries;
        rawResult.result.summary_entries = summaryEntries;
      } else {
        rawResult.detailed_entries = detailedEntries;
        rawResult.summary_entries = summaryEntries;
      }
      pendingEdits = {};
      $editActions.classList.add('hidden');
      renderRows(currentEntries);
    } else {
      showError('Save failed: ' + res.status);
    }
  } catch (e) {
    showError('Save failed: ' + e.message);
  } finally {
    $btnSaveChanges.disabled = false;
    $btnSaveChanges.textContent = 'Save Changes';
  }
});

// Discard Changes
$btnDiscardChanges.addEventListener('click', () => {
  pendingEdits = {};
  $editActions.classList.add('hidden');
  renderRows(currentEntries);
});

// Navigate-away warning
window.addEventListener('beforeunload', (e) => {
  if (Object.keys(pendingEdits).length > 0) {
    e.preventDefault();
    e.returnValue = '';
  }
});

// ─── Tab Toggle ─────────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    if (Object.keys(pendingEdits).length > 0) {
      if (!confirm('You have unsaved changes. Switch tabs and discard them?')) return;
      pendingEdits = {};
      $editActions.classList.add('hidden');
    }
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
  var date = new Date().toISOString().slice(0, 10);
  return 'RACM_' + sopName + '_' + date;
}

$btnExportCsv.addEventListener('click', () => {
  if (!currentEntries.length) return;

  const filtered = getFilteredEntries(currentEntries);

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

// ─── XLSX Export (F1) ───────────────────────────────────────────────────────
$btnExportXlsx.addEventListener('click', () => {
  if (!rawResult) return;

  const result = rawResult.result || rawResult;
  const wb = XLSX.utils.book_new();

  // Sheet 1: Detailed RACM
  const detailedData = (result.detailed_entries || []).map(entry => {
    const row = {};
    RACM_FIELDS.forEach((field, i) => { row[field] = String(getEntryValue(entry, i)); });
    return row;
  });
  const ws1 = XLSX.utils.json_to_sheet(detailedData, { header: RACM_FIELDS });
  XLSX.utils.book_append_sheet(wb, ws1, 'Detailed RACM');

  // Sheet 2: Summary
  const summaryData = (result.summary_entries || []).map(entry => {
    const row = {};
    RACM_FIELDS.forEach((field, i) => { row[field] = String(getEntryValue(entry, i)); });
    return row;
  });
  const ws2 = XLSX.utils.json_to_sheet(summaryData, { header: RACM_FIELDS });
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');

  XLSX.writeFile(wb, getExportBaseName() + '.xlsx');
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
