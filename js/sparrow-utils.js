/* ===== sparrow-utils.js =====
   Constants, DOM shorthand, and pure helpers shared across all modules.
   No dependencies — load first.
   ===== */

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const PALETTE = ['CC0000','1565C0','2E7D32','6A1B9A','006064','E65100','4E342E','37474F','880E4F','1A237E','BF360C','00695C'];

const $ = (id) => document.getElementById(id);

function safeHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function clampPct(p) { return Math.max(0, Math.min(100, p)); }

function normPct(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'string') {
    const s = v.trim();
    if (s.endsWith('%')) { const n = parseFloat(s.slice(0, -1)); return isFinite(n) ? clampPct(n) : 0; }
    const n = parseFloat(s);
    if (!isFinite(n)) return 0;
    if (n > 0 && n <= 1) return clampPct(n * 100);
    return clampPct(n);
  }
  if (typeof v === 'number') {
    if (!isFinite(v)) return 0;
    if (v > 0 && v <= 1) return clampPct(v * 100);
    return clampPct(v);
  }
  return 0;
}

function pctStr(v) {
  const p = normPct(v);
  const r = Math.round(p * 10) / 10;
  return String(r).endsWith('.0') ? String(Math.round(r)) : String(r);
}

function deriveStatus(progRaw, start) {
  const p = normPct(progRaw);
  if (p >= 100) return 'Completed';
  if (p > 0) return 'In Progress';
  if (start && start < new Date()) return 'In Progress';
  return 'Not Started';
}

function fmtDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function deriveDisplayNameFromFile(filename) {
  return String(filename || '')
    .replace(/\.xlsx?$/i, '')
    .replace(/Detail(ed)?_Plan_Plan_/i, '')
    .replace(/Detail(ed)?_Plan_/i, '')
    .replace(/_v\d+(\s*\(\d+\))?$/i, '')
    .replace(/_IK$/i, '')
    .replace(/_/g, ' ')
    .trim();
}

function getBarColorHex(a) {
  a = (a || '').toLowerCase();
  if (a.includes('uat')) return '#1565C0';
  if (a.includes('sit')) return '#2E7D32';
  if (a.includes('qa') || a.includes('test')) return '#2E7D32';
  if (a.includes('design')) return '#E65100';
  return '#546E7A';
}

function pptHex(c) { return String(c || '').replace('#', '').toUpperCase(); }

function contrast(hex) {
  hex = String(hex || '').replace('#', '');
  const r = parseInt(hex.substr(0, 2), 16) || 0;
  const g = parseInt(hex.substr(2, 2), 16) || 0;
  const b = parseInt(hex.substr(4, 2), 16) || 0;
  return ((0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.55) ? '1A1A1A' : 'FFFFFF';
}

function genTimestamp() {
  const n = new Date();
  const dd = String(n.getDate()).padStart(2, '0');
  const mm = String(n.getMonth() + 1).padStart(2, '0');
  const yyyy = n.getFullYear();
  const HH = String(n.getHours()).padStart(2, '0');
  const MIN = String(n.getMinutes()).padStart(2, '0');
  return { display: `${dd}/${mm}/${yyyy} ${HH}:${MIN}`, file: `${dd}-${mm}-${yyyy}_${HH}-${MIN}` };
}

/* Level / Domain normalisation */
function normTaskLevel(v) {
  if (v === null || v === undefined || v === '') return 2;
  if (typeof v === 'number') return (v === 1 || v === 2 || v === 3) ? v : 2;
  const m = String(v).match(/(\d+)/);
  const n = m ? parseInt(m[1], 10) : 2;
  return (n === 1 || n === 2 || n === 3) ? n : 2;
}

function normDomain(v) {
  const s = String(v ?? '').trim();
  return s ? s : 'Unspecified';
}

function downloadJson(obj, fileName) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}