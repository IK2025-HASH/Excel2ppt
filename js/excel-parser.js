/* ===== excel-parser.js =====
   SheetJS-based Excel parsing: plan sheet + RAID sheet.
   Depends on: sparrow-utils.js, SheetJS (XLSX global).
   ===== */

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        const planSN = wb.SheetNames.find(n => String(n).toLowerCase().includes('plan')) || wb.SheetNames[0];
        const planRows = parsePlanSheet(wb.Sheets[planSN], file.name);
        const raidSN = wb.SheetNames.find(n => String(n).toLowerCase().includes('raid'));
        const raidRows = raidSN ? parseRaidSheet(wb.Sheets[raidSN]) : [];
        resolve({ planRows, raidRows });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Read failed'));
    reader.readAsBinaryString(file);
  });
}

function parsePlanSheet(ws, fname) {
  if (!ws) return [];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const gc = (r, c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell ? cell.v : null; };

  let hRow = -1;
  for (let r = 0; r <= Math.min(25, range.e.r); r++) {
    for (let c = 0; c <= 12; c++) {
      const v = String(gc(r, c) || '').trim().toUpperCase();
      if (v === 'ID') { hRow = r; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) hRow = 5;

  const colMap = {};
  for (let c = 0; c <= range.e.c; c++) {
    const vRaw = String(gc(hRow, c) || '').trim();
    const v = vRaw.toUpperCase().replace(/[\s\/\-]+/g, '_');
    if (!v) continue;
    const looksLikeCombined = (v.includes('DEPEND') && (v.includes('TASK') || v.includes('MILESTONE') || v.includes('ACTIVITY') || v.includes('DESCRIPTION')));

    if (v === 'ID') colMap.id = c;
    if (v.includes('WORKSTREAM')) colMap.workstream = c;
    if (v.includes('TASK') && v.includes('LEVEL')) colMap.taskLevel = c;
    if (v.includes('DOMAIN')) colMap.domain = c;
    if (v.includes('EPIC') || v.includes('FEATURE') || v.includes('DELIVERABLE') || v.includes('FEAT')) colMap.epic = c;

    if (looksLikeCombined || v.includes('ACTIVITY') || v.includes('TASK') || v.includes('MILESTONE') || v.includes('DESCRIPTION')) {
      if (colMap.activity === undefined) colMap.activity = c;
      continue;
    }
    if (v.includes('PROGRESS')) colMap.progress = c;
    if (v.includes('START')) colMap.start = c;
    if (v.includes('END')) colMap.end = c;
  }

  if (colMap.id === undefined) colMap.id = 1;
  if (colMap.workstream === undefined) colMap.workstream = 2;
  if (colMap.epic === undefined) colMap.epic = 3;
  if (colMap.activity === undefined) colMap.activity = 4;
  if (colMap.progress === undefined) colMap.progress = 15;
  if (colMap.start === undefined) colMap.start = 16;
  if (colMap.end === undefined) colMap.end = 17;

  const pd = v => {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number' && v > 30000) return new Date((v - 25569) * 86400 * 1000);
    if (typeof v === 'string') {
      const dmy = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmy) {
        const d = new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
        return isNaN(d.getTime()) ? null : d;
      }
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  };

  const rows = [];
  for (let r = hRow + 1; r <= range.e.r; r++) {
    const id = gc(r, colMap.id);
    if (!id) continue;
    const ids = String(id).trim();
    if (!ids || ids.toUpperCase() === 'ID') continue;

    const start = pd(gc(r, colMap.start));
    const end = pd(gc(r, colMap.end));
    const prog = gc(r, colMap.progress);

    rows.push({
      id: ids,
      workstream: String(gc(r, colMap.workstream) || '').trim(),
      epic: String(gc(r, colMap.epic) || '').trim(),
      activity: String(gc(r, colMap.activity) || '').trim(),
      taskLevel: normTaskLevel(gc(r, colMap.taskLevel)),
      domain: normDomain(gc(r, colMap.domain)),
      progress: prog,
      start, end,
      status: deriveStatus(prog, start),
      sourceFile: fname,
      color: null
    });
  }
  return rows;
}

function parseRaidSheet(ws) {
  if (!ws) return [];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const gc = (r, c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell ? String(cell.v || '').trim() : ''; };
  const out = [];
  for (let r = 1; r <= range.e.r; r++) {
    const id = gc(r, 0);
    if (!id) continue;
    out.push({ id, title: gc(r, 1), description: gc(r, 2), mitigation: gc(r, 3), impact: gc(r, 4) });
  }
  return out;
}