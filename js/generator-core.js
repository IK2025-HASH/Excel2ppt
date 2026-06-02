/* ===== generator-core.js =====
   Global state, UI wiring, file I/O, audience filters, init.
   Depends on: sparrow-utils.js, excel-parser.js.
   ===== */

/* Fail loud */
window.onerror = function (msg, src, line) { alert('JS error line ' + line + ': ' + msg); return false; };

/* Nav */
function goHub() { window.location.href = 'index.html'; }
function goDashboard() { window.location.href = 'dashboard.html'; }

/* ✅ FIX 2: Epic normaliser (replaces normDomain usage for filtering) */
function normEpic(val) {
  const s = String(val || '').trim();
  return s || 'Activities';
}

/* State — ✅ FIX 2: domains → epics */
const state = {
  files: {}, fileOrder: [], fileConfigs: [],
  activeFileIdx: null, folderName: null, selectionProfile: null,
  inclusion: { taskLevels: new Set(), epics: new Set() },
  _levelsFound: new Set([1, 2]),
  _epicsFound: new Set(['Activities'])
};

/* ✅ FIX 2: rowMatchesInclusion uses epics instead of domains */
function rowMatchesInclusion(row) {
  const lvl = normTaskLevel(row.taskLevel);
  const epic = normEpic(row.epic);
  if (!state.inclusion.taskLevels.size || !state.inclusion.epics.size) return false;
  return state.inclusion.taskLevels.has(lvl) && state.inclusion.epics.has(epic);
}

/* Status UI */
function showFilesStatus(msg, type) {
  const b = $('filesStatus'); if (!b) return;
  b.textContent = msg; b.className = 'status-box ' + type; b.style.display = 'block';
}
/* ✅ FIX 3: separate profile status box */
function showProfileStatus(msg, type) {
  const b = $('profileStatus'); if (!b) return;
  b.textContent = msg; b.className = 'status-box ' + type; b.style.display = 'block';
}
function showGenStatus(msg, type) {
  const b = $('genStatus'); if (!b) return;
  b.textContent = msg; b.className = 'status-box ' + type; b.style.display = 'block';
}
function setProgress(pct, label) {
  $('progWrap').style.display = 'block';
  $('progFill').style.width = pct + '%';
  $('progLabel').textContent = label;
}
/* ✅ FIX 2: status strip shows epics instead of domains */
function updateStatusStrip() {
  const allRows = Object.values(state.files).flatMap(x => x.planRows || []);
  $('stFolder').textContent = 'FOLDER: ' + (state.folderName || 'n/a');
  $('stProfile').textContent = 'PROFILE: ' + (state.selectionProfile ? 'found' : 'none');
  $('stData').textContent = `DATA: ${state.fileOrder.length} files / ${allRows.length} rows`;

  const lvlTxt = [...state.inclusion.taskLevels].sort((a, b) => a - b).join(',') || 'none';
  const epicTxt = [...state.inclusion.epics].slice(0, 4).join(',') + (state.inclusion.epics.size > 4 ? '…' : '') || 'none';
  $('stAudience').textContent = `AUDIENCE: L${lvlTxt} · ${epicTxt}`;
}

/* Pickers */
function openFolderPicker() {
  const inp = $('folderInput');
  inp.value = '';
  inp.onchange = async (e) => {
    const files = Array.from(e.target.files || []);
    await handleStandaloneSelection(files, { source: 'folder' });
  };
  inp.click();
}
function openFilesPicker() {
  const inp = $('fileInput');
  inp.value = '';
  inp.onchange = async (e) => {
    const files = Array.from(e.target.files || []);
    await handleStandaloneSelection(files, { source: 'files' });
  };
  inp.click();
}
function detectFolderName(files) {
  const f = files.find(x => x && x.webkitRelativePath);
  if (!f) return null;
  return (f.webkitRelativePath.split('/')[0]) || null;
}

/* Month selects */
function fillMonthSelect(sel, selectedIdx) {
  sel.innerHTML = MONTH_NAMES.map((n, i) => `<option value="${i}" ${i === selectedIdx ? 'selected' : ''}>${n}</option>`).join('');
}

/* Milestones */
function addMilestone(preset) {
  const list = $('milestoneList');
  const div = document.createElement('div');
  div.className = 'arch-card';
  const mo = MONTH_NAMES.map((n, i) => `<option value="${i}">${n}</option>`).join('');
  div.innerHTML = `
    <div class="arch-grid">
      <div><label>Label</label><input type="text" class="ms-label" value="${safeHtml(preset?.label || 'MILESTONE')}"></div>
      <div><label>Start M</label><select class="ms-sM">${mo}</select></div>
      <div><label>Start Y</label><input type="number" class="ms-sY" value="${safeHtml(preset?.sY ?? 2026)}" min="2020" max="2035"></div>
      <div><label>End M</label><select class="ms-eM">${mo}</select></div>
      <div><label>End Y</label><input type="number" class="ms-eY" value="${safeHtml(preset?.eY ?? 2026)}" min="2020" max="2035"></div>
      <div><label>Colour</label><input type="color" class="ms-col" value="${safeHtml(preset?.col || '#0D47A1')}"></div>
      <div style="display:flex;justify-content:flex-end"><button class="xbtn" title="Remove" onclick="this.closest('.arch-card').remove()">×</button></div>
    </div>`;
  list.appendChild(div);
  div.querySelector('.ms-sM').value = String(preset?.sM ?? 0);
  div.querySelector('.ms-eM').value = String(preset?.eM ?? 9);
}
function getMilestones() {
  const out = [];
  document.querySelectorAll('#milestoneList .arch-card').forEach(card => {
    const label = (card.querySelector('.ms-label')?.value || '').trim();
    if (!label) return;
    const sM = parseInt(card.querySelector('.ms-sM')?.value || '0', 10);
    const sY = parseInt(card.querySelector('.ms-sY')?.value || '2026', 10);
    const eM = parseInt(card.querySelector('.ms-eM')?.value || '0', 10);
    const eY = parseInt(card.querySelector('.ms-eY')?.value || '2026', 10);
    const col = (card.querySelector('.ms-col')?.value || '#0D47A1').replace('#', '');
    out.push({ label, start: new Date(sY, sM, 1), end: new Date(eY, eM + 1, 0), color: col });
  });
  return out;
}

/* Bands */
function addBand(preset) {
  const tbody = $('bandsBody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${safeHtml(preset?.label || 'NEW BAND')}"></td>
    <td><select class="b-sM"></select></td>
    <td><input type="number" value="${safeHtml(preset?.sY ?? 2026)}" min="2020" max="2035" style="width:70px"></td>
    <td><select class="b-eM"></select></td>
    <td><input type="number" value="${safeHtml(preset?.eY ?? 2026)}" min="2020" max="2035" style="width:70px"></td>
    <td><input type="color" value="${safeHtml(preset?.col || '#607d8b')}"></td>
    <td><button class="ghost danger" onclick="this.closest('tr').remove()">✕</button></td>`;
  tbody.appendChild(tr);
  fillMonthSelect(tr.querySelector('.b-sM'), preset?.sM ?? 9);
  fillMonthSelect(tr.querySelector('.b-eM'), preset?.eM ?? 11);
}
function getBands() {
  const out = [];
  document.querySelectorAll('#bandsBody tr').forEach(tr => {
    const label = (tr.querySelector('input[type=text]')?.value || '').trim();
    if (!label) return;
    const sM = parseInt(tr.querySelector('select.b-sM')?.value || '0', 10);
    const sY = parseInt(tr.querySelectorAll('input[type=number]')[0]?.value || '2026', 10);
    const eM = parseInt(tr.querySelector('select.b-eM')?.value || '11', 10);
    const eY = parseInt(tr.querySelectorAll('input[type=number]')[1]?.value || '2026', 10);
    const col = (tr.querySelector('input[type=color]')?.value || '#607d8b').replace('#', '');
    out.push({ label, start: new Date(sY, sM, 1), end: new Date(eY, eM + 1, 0), color: col });
  });
  return out;
}

/* Programme window */
function computeProgrammeWindow() {
  const sMonth = parseInt($('startMonth').value || '0', 10);
  const sYear = parseInt($('startYear').value || '2026', 10);
  const eMonth = parseInt($('endMonth').value || '11', 10);
  const eYear = parseInt($('endYear').value || '2027', 10);
  return { sDate: new Date(sYear, sMonth, 1), eDate: new Date(eYear, eMonth + 1, 0), sYear, eYear };
}

/* ========= Multi-select dropdown UI ========= */
function toggleMsMenu(which) {
  const menu = $('msMenu_' + which);
  const isOpen = menu.classList.contains('open');
  closeAllMsMenus();
  if (!isOpen) menu.classList.add('open');
}
function closeAllMsMenus() { document.querySelectorAll('.ms-menu.open').forEach(m => m.classList.remove('open')); }
document.addEventListener('click', (e) => { if (!e.target.closest('.ms-wrap')) closeAllMsMenus(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeAllMsMenus(); });

/* ✅ FIX 2: msSelectAll/msSelectNone use epics instead of domains */
function msSelectAll(which) {
  if (which === 'taskLevels') state.inclusion.taskLevels = new Set([...state._levelsFound]);
  if (which === 'epics') state.inclusion.epics = new Set([...state._epicsFound]);
  syncMsChecksFromState();
  onAudienceChanged();
}
function msSelectNone(which) {
  if (which === 'taskLevels') state.inclusion.taskLevels = new Set();
  if (which === 'epics') state.inclusion.epics = new Set();
  syncMsChecksFromState();
  onAudienceChanged();
}
function buildTaskLevelMenu() {
  const host = $('msList_taskLevels'); host.innerHTML = '';
  const levels = [...state._levelsFound].sort((a, b) => a - b);
  if (!levels.length) { host.innerHTML = '<div class="ms-empty">No levels found yet.</div>'; return; }
  levels.forEach(lvl => {
    const row = document.createElement('label');
    row.className = 'ms-item';
    row.innerHTML = `<input type="checkbox" data-lvl="${lvl}"><div><b>Level ${lvl}</b> <span class="tag">(Task Level)</span></div>`;
    row.querySelector('input').addEventListener('change', () => { readMsStateFromChecks(); onAudienceChanged(); });
    host.appendChild(row);
  });
}
/* ✅ FIX 2: buildEpicMenu replaces buildDomainMenu */
function buildEpicMenu() {
  const host = $('msList_epics'); host.innerHTML = '';
  const epics = [...state._epicsFound].sort((a, b) => String(a).localeCompare(String(b)));
  if (!epics.length) { host.innerHTML = '<div class="ms-empty">No epics found yet.</div>'; return; }
  epics.forEach(ep => {
    const row = document.createElement('label');
    row.className = 'ms-item';
    row.innerHTML = `<input type="checkbox"><div><b>${safeHtml(ep)}</b> <span class="tag">(Epic)</span></div>`;
    const cb = row.querySelector('input');
    cb.dataset.epic = ep;
    cb.addEventListener('change', () => { readMsStateFromChecks(); onAudienceChanged(); });
    host.appendChild(row);
  });
}
function buildMultiSelectMenus() {
  buildTaskLevelMenu();
  buildEpicMenu();
  syncMsChecksFromState();
}
/* ✅ FIX 2: sync checks uses epics */
function syncMsChecksFromState() {
  document.querySelectorAll('#msList_taskLevels input[type=checkbox]').forEach(cb => {
    const lvl = parseInt(cb.getAttribute('data-lvl') || '0', 10);
    cb.checked = state.inclusion.taskLevels.has(lvl);
  });
  document.querySelectorAll('#msList_epics input[type=checkbox]').forEach(cb => {
    const ep = cb.dataset.epic || '';
    cb.checked = state.inclusion.epics.has(ep);
  });
  updateMsLabels();
}
/* ✅ FIX 2: read checks uses epics */
function readMsStateFromChecks() {
  const levels = new Set();
  document.querySelectorAll('#msList_taskLevels input[type=checkbox]').forEach(cb => {
    if (cb.checked) levels.add(parseInt(cb.getAttribute('data-lvl') || '0', 10));
  });
  const epics = new Set();
  document.querySelectorAll('#msList_epics input[type=checkbox]').forEach(cb => {
    if (cb.checked) epics.add(cb.dataset.epic || '');
  });
  state.inclusion.taskLevels = levels;
  state.inclusion.epics = epics;
  updateMsLabels();
}
/* ✅ FIX 2: labels show epics */
function updateMsLabels() {
  const lvl = [...state.inclusion.taskLevels].sort((a, b) => a - b);
  const ep = [...state.inclusion.epics].sort((a, b) => String(a).localeCompare(String(b)));
  $('msTaskLevelLabel').textContent = lvl.length ? ('Selected: ' + lvl.join(', ')) : 'Selected: (none)';
  $('msEpicLabel').textContent = ep.length ? ('Selected: ' + (ep.slice(0, 3).join(', ') + (ep.length > 3 ? '…' : ''))) : 'Selected: (none)';
  updateStatusStrip();
}

/* ✅ FIX 2: Presets use epics instead of domains */
function presetExec() {
  state.inclusion.taskLevels = new Set([1]);
  state.inclusion.epics = new Set([...state._epicsFound]);
  syncMsChecksFromState(); onAudienceChanged();
}
function presetDelivery() {
  state.inclusion.taskLevels = new Set([2]);
  state.inclusion.epics = new Set([...state._epicsFound]);
  syncMsChecksFromState(); onAudienceChanged();
}
function presetCert() {
  state.inclusion.taskLevels = new Set([2]);
  state.inclusion.epics = new Set([...state._epicsFound]);
  syncMsChecksFromState(); onAudienceChanged();
}
function presetAll() {
  state.inclusion.taskLevels = new Set([...state._levelsFound]);
  state.inclusion.epics = new Set([...state._epicsFound]);
  syncMsChecksFromState(); onAudienceChanged();
}

/* Audience re-render */
function onAudienceChanged() {
  updateStatusStrip();
  buildFileTiles();
  if (state.activeFileIdx !== null) renderRowTableFile(state.activeFileIdx);
  renderHtmlPreviewTiles();
}

/* ========= Profile ========= */
/* ✅ FIX 2: profile saves/loads epics instead of domains */
function buildSelectionProfileObject() {
  const progName = ($('progName')?.value || '').trim() || 'Programme';
  return {
    version: 3,
    programme: progName,
    savedAt: new Date().toISOString(),
    includedTaskLevels: [...state.inclusion.taskLevels].sort((a, b) => a - b),
    includedEpics: [...state.inclusion.epics].sort((a, b) => String(a).localeCompare(String(b))),
    files: (state.fileConfigs || []).map(fc => ({
      filename: fc.filename,
      displayName: fc.displayName,
      col1Color: fc.col1Color,
      col2Color: fc.col2Color,
      selectedIds: (fc.rowConfigs || []).filter(rc => rc.selected).map(rc => rc.row.id),
      rowColorById: Object.fromEntries((fc.rowConfigs || []).filter(rc => rc.color).map(rc => [rc.row.id, rc.color]))
    }))
  };
}
/* ✅ FIX 3: downloadSelectionProfile uses showProfileStatus */
function downloadSelectionProfile() {
  if (!state.fileConfigs.length) { showProfileStatus('Load Excel files first.', 'info'); return; }
  downloadJson(buildSelectionProfileObject(), 'selection-profile.json');
  showProfileStatus('Downloaded selection-profile.json. Copy into the Excel folder.', 'ok');
}
function importSelectionProfileClick() {
  const inp = $('importProfileInput');
  inp.value = '';
  inp.onchange = async (e) => {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    try {
      const prof = JSON.parse(await f.text());
      state.selectionProfile = prof;
      applySelectionProfileObject(prof);
      buildFileTiles();
      if (state.activeFileIdx !== null) renderRowTableFile(state.activeFileIdx);
      renderHtmlPreviewTiles();
      updateStatusStrip();
      showProfileStatus('Imported selection-profile.json', 'ok');
    } catch (err) {
      console.error(err);
      showProfileStatus('Failed to import selection-profile.json', 'err');
    }
  };
  inp.click();
}
/* ✅ FIX 2: applySelectionProfileObject loads epics (with backward compat for old domain profiles) */
function applySelectionProfileObject(profile) {
  if (profile) {
    if (Array.isArray(profile.includedTaskLevels)) state.inclusion.taskLevels = new Set(profile.includedTaskLevels.map(normTaskLevel));
    if (Array.isArray(profile.includedEpics)) {
      state.inclusion.epics = new Set(profile.includedEpics.map(normEpic));
    } else if (Array.isArray(profile.includedDomains)) {
      // backward compat: old v2 profiles had domains — select all epics instead
      state.inclusion.epics = new Set([...state._epicsFound]);
    }
  }
  if (profile && profile.files) {
    const byFile = new Map(profile.files.map(x => [x.filename, x]));
    state.fileConfigs.forEach(fc => {
      const s = byFile.get(fc.filename);
      if (!s) return;
      if (s.displayName) fc.displayName = s.displayName;
      if (s.col1Color) fc.col1Color = s.col1Color;
      if (s.col2Color) fc.col2Color = s.col2Color;
      const sel = new Set(s.selectedIds || []);
      const cmap = s.rowColorById || {};
      fc.rowConfigs.forEach(rc => {
        rc.selected = sel.has(rc.row.id);
        if (cmap[rc.row.id]) rc.color = cmap[rc.row.id];
      });
    });
  }
  buildMultiSelectMenus();
}

/* ========= Load + UI ========= */
async function handleStandaloneSelection(files, opts) {
  state.files = {}; state.fileOrder = []; state.fileConfigs = []; state.activeFileIdx = null;
  state.folderName = (opts?.source === 'folder') ? detectFolderName(files) : null;
  state.selectionProfile = null;

  $('fileGrid').innerHTML = '';
  $('wsTileList').innerHTML = '';
  $('rowTableContainer').innerHTML = '<div style="font-size:11px;color:var(--muted);padding:4px 2px">No workstream selected yet.</div>';
  $('raidTableContainer').innerHTML = '<div style="font-size:11px;color:var(--muted);padding:4px 2px">No workstream selected yet.</div>';
  $('htmlTileGrid').innerHTML = '<div style="font-size:11px;color:var(--muted);padding:4px 2px">No selected workstreams yet.</div>';
  $('filesStatus').style.display = 'none';
  $('genStatus').style.display = 'none';
  updateStatusStrip();

  /* ✅ FIX 1: only count .xlsx files */
  const excelFiles = files.filter(f => /\.xlsx?$/i.test(f.name));
  const profileFile = files.find(f => f.name === 'selection-profile.json') || null;

  if (!excelFiles.length) {
    showFilesStatus('No Excel files found in selection.', 'err');
    return;
  }

  showFilesStatus('Loading ' + excelFiles.length + ' Excel file(s)…', 'info');

  for (const f of excelFiles) {
    const id = 'fc_' + f.name.replace(/[^a-z0-9]/gi, '_');
    addFileCardShell(id, f.name);
    try {
      const parsed = await parseExcel(f);
      state.files[f.name] = parsed;
      state.fileOrder.push(f.name);
      updateFileCard(id, f.name, parsed.planRows.length, parsed.raidRows.length > 0);
    } catch (e) {
      console.error(e);
      updateFileCardError(id, f.name, e?.message || String(e));
    }
  }

  buildFileConfigs();

  /* ✅ FIX 2: discover levels + epics (not domains) */
  const levelsFound = new Set();
  const epicsFound = new Set();
  state.fileConfigs.forEach(fc => {
    fc.rowConfigs.forEach(rc => {
      levelsFound.add(normTaskLevel(rc.row.taskLevel));
      epicsFound.add(normEpic(rc.row.epic));
    });
  });
  state._levelsFound = levelsFound.size ? levelsFound : new Set([1, 2]);
  state._epicsFound = epicsFound.size ? epicsFound : new Set(['Activities']);

  /* default select all */
  state.inclusion.taskLevels = new Set([...state._levelsFound]);
  state.inclusion.epics = new Set([...state._epicsFound]);

  buildMultiSelectMenus();
  buildFileTiles();
  renderHtmlPreviewTiles();
  updateStatusStrip();

  /* auto-apply folder profile */
  if (profileFile) {
    try {
      const prof = JSON.parse(await profileFile.text());
      state.selectionProfile = prof;
      applySelectionProfileObject(prof);
      buildFileTiles();
      if (state.activeFileIdx !== null) renderRowTableFile(state.activeFileIdx);
      renderHtmlPreviewTiles();
      updateStatusStrip();
      /* ✅ FIX 1: message shows excelFiles.length not total files */
      showFilesStatus('Loaded ' + excelFiles.length + ' Excel file(s) + auto-applied selection-profile.json', 'ok');
    } catch (e) {
      showFilesStatus('Loaded ' + excelFiles.length + ' Excel file(s). (selection-profile.json present but not parsed)', 'info');
    }
  } else {
    /* ✅ FIX 1: message shows excelFiles.length */
    showFilesStatus('Loaded ' + excelFiles.length + ' Excel file(s).', 'ok');
  }

  if (state.fileConfigs.length > 0) {
    openFileConfig(0);
  }
}

/* File cards */
function addFileCardShell(id, name) {
  const grid = $('fileGrid');
  const div = document.createElement('div');
  div.className = 'file-card';
  div.id = id;
  div.innerHTML = `<div style="font-size:18px">📊</div>
    <div style="flex:1;min-width:0">
      <div class="file-name"></div>
      <div class="file-meta"></div>
    </div>
    <div class="badge info">Loading…</div>`;
  grid.appendChild(div);
  div.querySelector('.file-name').textContent = name;
  div.querySelector('.file-meta').textContent = 'Reading…';
}
function updateFileCard(id, name, rowCount, hasRaid) {
  const div = $(id); if (!div) return;
  div.querySelector('.file-meta').textContent = `${rowCount} row(s)${hasRaid ? ' · RAID ✓' : ''}`;
  const b = div.querySelector('.badge'); b.textContent = 'OK'; b.className = 'badge ok';
}
function updateFileCardError(id, name, msg) {
  const div = $(id); if (!div) return;
  div.querySelector('.file-meta').textContent = 'Error: ' + msg;
  const b = div.querySelector('.badge'); b.textContent = 'Error'; b.className = 'badge err';
}

/* Build file configs */
function buildFileConfigs() {
  state.fileConfigs = [];
  state.fileOrder.forEach((fname, idx) => {
    const data = state.files[fname] || { planRows: [], raidRows: [] };
    state.fileConfigs.push({
      filename: fname,
      displayName: deriveDisplayNameFromFile(fname) || fname,
      col1Color: '#' + PALETTE[idx % PALETTE.length],
      col2Color: '#' + PALETTE[(idx + 3) % PALETTE.length],
      rowConfigs: (data.planRows || []).map(r => ({ row: r, selected: true, color: null })),
      raidRows: (data.raidRows || [])
    });
  });
}

/* Workstream tiles */
function includedSelectedCountForFc(fc) {
  return (fc.rowConfigs || []).reduce((n, rc) => n + ((rc.selected && rowMatchesInclusion(rc.row)) ? 1 : 0), 0);
}
function visibleRowsForFc(fc) {
  return (fc.rowConfigs || []).filter(rc => rowMatchesInclusion(rc.row));
}
function buildFileTiles() {
  const host = $('wsTileList');
  host.innerHTML = '';
  state.fileConfigs.forEach((fc, idx) => {
    const div = document.createElement('div');
    div.className = 'ws-tile' + (state.activeFileIdx === idx ? ' active' : '');
    div.onclick = () => openFileConfig(idx);
    div.innerHTML = `
      <div class="ws-swatch" style="background:${fc.col1Color}"></div>
      <div class="ws-info">
        <div class="ws-name">${safeHtml(fc.displayName)}</div>
        <div class="ws-count">${includedSelectedCountForFc(fc)} included / ${fc.rowConfigs.length} total</div>
      </div>`;
    host.appendChild(div);
  });
}
function openFileConfig(i) {
  state.activeFileIdx = i;
  document.querySelectorAll('.ws-tile').forEach((t, idx) => t.classList.toggle('active', idx === i));
  renderRowTableFile(i);
  renderRaidTableFile(i);
}

/* Row table */
function renderRowTableFile(i) {
  const container = $('rowTableContainer');
  const fc = state.fileConfigs[i];
  if (!fc) { container.innerHTML = ''; return; }

  const visible = visibleRowsForFc(fc);
  if (!visible.length) {
    container.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:6px 2px">No rows match filters.</div>';
    return;
  }

  let html = `<div class="row-table-wrap"><table><thead><tr>
    <th></th><th>ID</th><th>Lvl</th><th>Domain</th><th>Epic</th><th>Activity</th><th>Status</th><th>Start</th><th>End</th><th>Colour</th>
  </tr></thead><tbody>`;

  visible.forEach(rc => {
    const r = rc.row;
    const stCls = r.status === 'In Progress' ? 'in-progress' : (r.status === 'Completed' ? 'completed' : 'not-started');
    const color = rc.color || getBarColorHex(r.activity);
    html += `<tr class="${rc.selected ? 'selected' : ''}">
      <td><input type="checkbox" ${rc.selected ? 'checked' : ''} onchange="toggleRow('${safeHtml(fc.filename)}','${safeHtml(r.id)}',this.checked)"></td>
      <td style="font-family:var(--mono);color:var(--accent)">${safeHtml(r.id)}</td>
      <td style="font-family:var(--mono)">${safeHtml(normTaskLevel(r.taskLevel))}</td>
      <td style="color:var(--muted)">${safeHtml(normDomain(r.domain))}</td>
      <td style="font-weight:700">${safeHtml(r.epic || '')}</td>
      <td style="white-space:normal;word-break:break-word">${safeHtml(r.activity || '')}</td>
      <td><span class="status-pill-small ${stCls}">${safeHtml(r.status)}</span></td>
      <td style="font-family:var(--mono);font-size:10px">${fmtDate(r.start)}</td>
      <td style="font-family:var(--mono);font-size:10px">${fmtDate(r.end)}</td>
      <td><input type="color" value="${safeHtml(color)}" onchange="setRowColor('${safeHtml(fc.filename)}','${safeHtml(r.id)}',this.value)"></td>
    </tr>`;
  });

  html += '</tbody></table></div>';
  container.innerHTML = html;
}
function toggleRow(filename, rowId, checked) {
  const fc = state.fileConfigs.find(x => x.filename === filename); if (!fc) return;
  const rc = fc.rowConfigs.find(x => String(x.row.id) === String(rowId)); if (!rc) return;
  rc.selected = checked;
  buildFileTiles();
  renderHtmlPreviewTiles();
}
function setRowColor(filename, rowId, color) {
  const fc = state.fileConfigs.find(x => x.filename === filename); if (!fc) return;
  const rc = fc.rowConfigs.find(x => String(x.row.id) === String(rowId)); if (!rc) return;
  rc.color = color;
}
function selectAllRowsActive(val) {
  const i = state.activeFileIdx;
  if (i === null || i === undefined) return;
  const fc = state.fileConfigs[i];
  visibleRowsForFc(fc).forEach(rc => rc.selected = val);
  renderRowTableFile(i);
  buildFileTiles();
  renderHtmlPreviewTiles();
}

/* RAID */
function renderRaidTableFile(i) {
  const container = $('raidTableContainer');
  const fc = state.fileConfigs[i];
  if (!fc || !fc.raidRows || !fc.raidRows.length) {
    container.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:4px 2px">No RAID items.</div>';
    return;
  }
  let html = `<div class="raid-table"><table><thead><tr>
    <th>ID</th><th>Title</th><th>Description</th><th>Mitigation</th><th>Impact</th>
  </tr></thead><tbody>`;
  fc.raidRows.forEach(r => {
    html += `<tr>
      <td style="font-family:var(--mono);color:var(--accent)">${safeHtml(r.id)}</td>
      <td>${safeHtml(r.title || '')}</td>
      <td>${safeHtml(r.description || '')}</td>
      <td>${safeHtml(r.mitigation || '')}</td>
      <td>${safeHtml(r.impact || '')}</td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  container.innerHTML = html;
}

/* HTML preview tiles */
function renderHtmlPreviewTiles() {
  const grid = $('htmlTileGrid');
  const eligible = (state.fileConfigs || []).filter(fc => includedSelectedCountForFc(fc) > 0);

  if (!eligible.length) {
    grid.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:4px 2px">No selected workstreams (after filters).</div>';
    return;
  }

  grid.innerHTML = '';
  eligible.forEach(fc => {
    const cnt = includedSelectedCountForFc(fc);
    const tile = document.createElement('div');
    tile.className = 'html-tile';
    tile.onclick = () => openWorkstreamHtml(fc.filename);
    tile.innerHTML = `
      <div class="sw" style="background:${fc.col1Color}"></div>
      <div class="meta">
        <div class="nm">${safeHtml(fc.displayName)}</div>
        <div class="sub">${cnt} included selected row(s)</div>
      </div>
      <div class="go">OPEN</div>`;
    grid.appendChild(tile);
  });
}

/* Helper getters used by ppt-engine.js */
function getIncludedSelectedRowsForFc(fc) {
  return fc.rowConfigs
    .filter(rc => rc.selected && rowMatchesInclusion(rc.row))
    .map(rc => ({ ...rc.row, color: rc.color || null }));
}

function getAllRowsFlatIncluded() {
  const rows = [];
  state.fileConfigs.forEach(fc => {
    fc.rowConfigs.forEach(rc => {
      if (!(rc.selected && rowMatchesInclusion(rc.row))) return;
      rows.push({ ...rc.row, color: rc.color || null, sourceDisplay: fc.displayName });
    });
  });
  return rows;
}

function isQaRow(r) {
  const haystack = (String(r.activity || '') + ' ' + String(r.epic || '')).toLowerCase();
  return /\b(qa|test|testing|sit|uat|oat|e2e|regression|nfr|non[-\s]?functional|performance|security test)\b/i.test(haystack);
}

/* Init */
(function init() {
  fillMonthSelect($('startMonth'), 0);
  fillMonthSelect($('endMonth'), 11);

  $('fileGrid').innerHTML = '<div style="font-size:11px;color:var(--muted)">No files loaded yet.</div>';

  // default bands + milestones
  addBand({ label: 'CHANGE FREEZE', sM: 9, sY: 2026, eM: 11, eY: 2026, col: '#546e7a' });
  addMilestone({ label: 'CIC', sM: 0, sY: 2026, eM: 9, eY: 2026, col: '#0D47A1' });
  addMilestone({ label: 'Part VII', sM: 3, sY: 2026, eM: 11, eY: 2027, col: '#880E4F' });

  updateStatusStrip();
})();