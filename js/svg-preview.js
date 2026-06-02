/* ===== svg-preview.js =====
   HTML/SVG timeline preview builder.
   Depends on: sparrow-utils.js, generator-core.js (computeProgrammeWindow, getMilestones, getBands, etc.).
   ===== */

/* BAR LABELS — ALWAYS VISIBLE, no truncation */
function fitSvgTextToWidth(text, pxWidth) {
  return String(text ?? '');
}

/* Workstream HTML preview (PPT-like SVG) — returns { headerSvg, bodySvg } */
function buildWorkstreamPptLikeSVG(rows, fc) {
  const { sDate, eDate, sYear, eYear } = computeProgrammeWindow();
  const totMs = Math.max(1, eDate - sDate);

  const W = 1400;
  const LW = 220;
  const GX = LW;
  const GW = W - LW - 30;

  const HDR1 = 22, HDR2 = 18, HDR3 = 18;
  const milestones = getMilestones();
  const MS_H = milestones.length ? 14 : 0;
  const headerH = HDR1 + HDR2 + HDR3 + MS_H;

  const groups = new Map();
  rows.forEach(r => {
    const epic = (r.epic || 'Activities').trim() || 'Activities';
    if (!groups.has(epic)) groups.set(epic, []);
    groups.get(epic).push(r);
  });

  let rowCount = 0;
  groups.forEach(g => rowCount += g.length);
  const ROW_H = Math.max(20, Math.min(28, Math.floor(520 / Math.max(rowCount, 1))));
  const BODY_H = Math.max(220, rowCount * ROW_H + 18);

  const d2x = d => {
    if (!(d instanceof Date) || isNaN(d)) return GX;
    const ms = Math.max(0, Math.min(d - sDate, totMs));
    return GX + (ms / totMs) * GW;
  };

  const cEpic = pptHex(fc.col1Color || '#CC0000');
  const bands = getBands();

  /* ========= HEADER SVG (sticky) ========= */
  let hSvg = `<svg viewBox="0 0 ${W} ${headerH}" xmlns="http://www.w3.org/2000/svg">`;
  hSvg += `<rect x="0" y="0" width="${W}" height="${headerH}" fill="#ffffff"/>`;

  /* Title bar */
  hSvg += `<rect x="0" y="0" width="${W}" height="${HDR1}" fill="#CC0000"/>`;
  hSvg += `<text x="${W / 2}" y="15" text-anchor="middle" font-family="Calibri" font-size="12" font-weight="bold" fill="#ffffff">${safeHtml(fc.displayName)}</text>`;

  /* Year bar */
  const yY = HDR1;
  hSvg += `<rect x="0" y="${yY}" width="${W}" height="${HDR2}" fill="#0D4B6E"/>`;
  for (let yr = sYear; yr <= eYear; yr++) {
    const x1 = Math.max(d2x(new Date(yr, 0, 1)), GX);
    const x2 = Math.min(d2x(new Date(yr + 1, 0, 1)), GX + GW);
    const w = x2 - x1;
    if (w > 0) {
      hSvg += `<text x="${x1 + w / 2}" y="${yY + 13}" text-anchor="middle" font-family="Calibri" font-size="10" font-weight="bold" fill="#ffffff">${yr}</text>`;
    }
  }

  /* Month bar */
  const mY = HDR1 + HDR2;
  hSvg += `<rect x="0" y="${mY}" width="${W}" height="${HDR3}" fill="#1565C0"/>`;
  hSvg += `<text x="${LW / 2}" y="${mY + 13}" text-anchor="middle" font-family="Calibri" font-size="8" font-weight="bold" fill="#ffffff">Epic</text>`;
  let hdrCur = new Date(sDate.getFullYear(), sDate.getMonth(), 1);
  while (hdrCur <= eDate) {
    const x1 = Math.max(d2x(hdrCur), GX);
    const x2 = Math.min(d2x(new Date(hdrCur.getFullYear(), hdrCur.getMonth() + 1, 1)), GX + GW);
    const w = x2 - x1;
    if (w > 3) {
      hSvg += `<rect x="${x1}" y="${mY}" width="1" height="${HDR3}" fill="#ffffff" opacity="0.6"/>`;
      hSvg += `<text x="${x1 + 3}" y="${mY + 13}" font-family="Calibri" font-size="7" fill="#ffffff">${MONTHS[hdrCur.getMonth()]}</text>`;
    }
    hdrCur.setMonth(hdrCur.getMonth() + 1);
  }

  /* Milestones */
  if (MS_H) {
    const msY = HDR1 + HDR2 + HDR3;
    hSvg += `<rect x="${GX}" y="${msY}" width="${GW}" height="${MS_H}" fill="#E8EDF4"/>`;
    milestones.forEach(m => {
      const bx = Math.max(d2x(m.start), GX);
      const ex = Math.min(d2x(m.end), GX + GW);
      const bw = ex - bx; if (bw <= 0) return;
      hSvg += `<rect x="${bx}" y="${msY + 1}" width="${bw}" height="${MS_H - 2}" rx="2" fill="#${m.color}" opacity="0.9"/>`;
      hSvg += `<text x="${bx + bw / 2}" y="${msY + MS_H / 2 + 1}" text-anchor="middle" dominant-baseline="middle" font-family="Calibri" font-size="7" font-weight="bold" fill="#${contrast(m.color)}">${safeHtml(m.label)}</text>`;
    });
  }

  hSvg += '</svg>';

  /* ========= BODY SVG (scrollable) ========= */
  let bSvg = `<svg viewBox="0 0 ${W} ${BODY_H}" xmlns="http://www.w3.org/2000/svg">`;
  bSvg += `<rect x="0" y="0" width="${W}" height="${BODY_H}" fill="#ffffff"/>`;

  let curY = 0;
  let alt = 0;

  for (const [epic, gRows] of groups) {
    const gH = gRows.length * ROW_H;

    /* Epic label band */
    bSvg += `<rect x="0" y="${curY}" width="${LW}" height="${gH - 1}" fill="#${cEpic}"/>`;
    bSvg += `<text x="${LW / 2}" y="${curY + gH / 2}" text-anchor="middle" dominant-baseline="middle" font-family="Calibri" font-size="8" font-weight="bold" fill="#${contrast(cEpic)}">${safeHtml(String(epic).slice(0, 30))}</text>`;

    /* Row background */
    bSvg += `<rect x="${GX}" y="${curY}" width="${GW}" height="${gH - 1}" fill="${alt % 2 === 0 ? '#F7F7F7' : '#EFEFEF'}"/>`;

    /* Vertical bands */
    bands.forEach(b => {
      const bx = Math.max(d2x(b.start), GX);
      const ex = Math.min(d2x(b.end), GX + GW);
      const bw = ex - bx; if (bw <= 0) return;
      bSvg += `<rect x="${bx}" y="${curY}" width="${bw}" height="${gH - 1}" fill="#${b.color}" opacity="0.12"/>`;
    });

    /* ✅ NEW: Month grid lines (vertical, matching PPT) */
    let mCur = new Date(sDate.getFullYear(), sDate.getMonth(), 1);
    while (mCur <= eDate) {
      const x = d2x(mCur);
      if (x >= GX && x <= GX + GW) {
        bSvg += `<line x1="${x}" y1="${curY}" x2="${x}" y2="${curY + gH - 1}" stroke="#D8D8D8" stroke-width="0.5"/>`;
      }
      mCur.setMonth(mCur.getMonth() + 1);
    }

    /* Bars + labels (geometry UNCHANGED from working version) */
    gRows.forEach(r => {
      const pct = normPct(r.progress);
      if (r.start && r.end && r.end >= r.start) {
        const bx = Math.max(d2x(r.start), GX);
        const ex = Math.min(d2x(r.end), GX + GW);
        let bw = ex - bx; if (bw < 3) bw = 3;

        const barH = Math.max(7, Math.floor(ROW_H * 0.36));
        const barY = curY + Math.max(7, Math.floor(ROW_H * 0.32));
        const bc = pptHex(r.color || getBarColorHex(r.activity));

        /* bar */
        bSvg += `<rect x="${bx}" y="${barY}" width="${bw}" height="${barH}" rx="2" fill="#${bc}" opacity="${(pct > 0 && pct < 100) ? 0.35 : 1}"/>`;
        if (pct > 0 && pct < 100) {
          bSvg += `<rect x="${bx}" y="${barY}" width="${bw * (pct / 100)}" height="${barH}" rx="2" fill="#${bc}"/>`;
        }

        /* label ON the bar — UNCHANGED */
        const rawLabel = `${r.id || ''} - ${r.activity || ''}`.trim();
        const label = fitSvgTextToWidth(rawLabel, bw);
        const labelY = barY + barH / 2;
        bSvg += `<text x="${bx + 3}" y="${labelY}" dominant-baseline="middle" font-family="Calibri" font-size="7" font-weight="bold" fill="#222">${safeHtml(label)}</text>`;
      }

      bSvg += `<line x1="${GX}" y1="${curY + ROW_H - 1}" x2="${GX + GW}" y2="${curY + ROW_H - 1}" stroke="#CCCCCC" stroke-width="0.5"/>`;
      curY += ROW_H;
    });

    bSvg += `<line x1="0" y1="${curY}" x2="${GX + GW}" y2="${curY}" stroke="#9E9E9E" stroke-width="1"/>`;
    alt++;
  }

  bSvg += '</svg>';

  return { headerSvg: hSvg, bodySvg: bSvg };
}

function buildWorkstreamHtml(fc) {
  const progName = ($('progName')?.value || '').trim() || 'Programme';
  const { sDate, eDate } = computeProgrammeWindow();

  const rows = fc.rowConfigs
    .filter(rc => rc.selected && rowMatchesInclusion(rc.row))
    .map(rc => ({ ...rc.row, color: rc.color }));

  const result = buildWorkstreamPptLikeSVG(rows, fc);

  return `<!doctype html>
<html lang="en"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>${safeHtml(progName)} · ${safeHtml(fc.displayName)} · HTML Preview</title>
<style>
body{font-family:Arial,Helvetica,sans-serif;background:#0b1020;color:#e9eefc;margin:0;padding:18px}
h1{margin:0 0 6px;font-size:18px}
.meta{font-family:Consolas,monospace;font-size:11px;color:#8fa0d8;line-height:1.6;margin-bottom:10px}
.card{background:#11172b;border:1px solid #243044;border-radius:12px;padding:0;overflow:hidden;max-height:85vh;display:flex;flex-direction:column}
.sticky-hdr{flex-shrink:0;background:#fff;border-bottom:1px solid #ccc}
.sticky-hdr svg{width:100%;height:auto;display:block}
.scroll-body{flex:1;overflow-y:auto;background:#fff}
.scroll-body svg{width:100%;height:auto;display:block}
</style></head><body>
<h1>${safeHtml(progName)} — ${safeHtml(fc.displayName)} (HTML Preview)</h1>
<div class="meta">
File: ${safeHtml(fc.filename)} · Included rows: ${rows.length}<br/>
Programme window: ${safeHtml(fmtDate(sDate))} → ${safeHtml(fmtDate(eDate))}
</div>
<div class="card">
<div class="sticky-hdr">${result.headerSvg}</div>
<div class="scroll-body">${result.bodySvg}</div>
</div>
</body></html>`;
}

function openWorkstreamHtml(filename) {
  const fc = state.fileConfigs.find(x => x.filename === filename);
  if (!fc) { alert('Workstream not found'); return; }
  const html = buildWorkstreamHtml(fc);
  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  window.open(url, '_blank');
  setTimeout(() => URL.revokeObjectURL(url), 15000);
}