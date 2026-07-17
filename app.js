'use strict';

/* ================================================================
   RadPoint — Interactive radiology hotspots for PowerPoint slides
   Content Add-in · Vanilla JS · HTML5 Canvas · Office.js
   ----------------------------------------------------------------
   Zero-server architecture: the image (Base64), polygon hotspots
   (normalized 0–100% coordinates), windowing presets and labels are
   serialized to JSON and stored in the slide via
   Office.context.document.settings, so the .pptx is self-contained.
   ================================================================ */

/* ------------------------------ Constants ------------------------------ */
const SETTINGS_KEY = 'radpoint.state.v1';
const MAX_UPLOAD_BYTES = 2 * 1024 * 1024; // compress uploads bigger than this
const TARGET_DATAURL_CHARS = 2000000;     // ~1.5 MB binary once Base64-decoded
const MAX_COMPRESS_DIM = 1600;            // px, longest edge after compression
const CLOSE_SNAP_PCT = 2.5;               // % distance that snaps a click onto the first vertex
const MIN_POINTS = 3;
const HINT_AFTER_MISSES = 3;              // consecutive misses before the quadrant hint
const REVEAL_MS = 550;                    // fade-in duration for revealed outlines

const COLORS = {
  pathology: '#f59e0b',
  anatomy: '#22d3ee',
  success: '#4ade80',
  draft: '#38bdf8',
};

/* ------------------------------ Geometry ------------------------------ */
/* All coordinates are normalized percentages (0–100) of the image, never
   raw pixels, so hotspots survive any canvas / slide / window resize.    */

/** Standard ray-casting point-in-polygon test. */
function pointInPolygon(x, y, pts) {
  let inside = false;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const xi = pts[i].x, yi = pts[i].y;
    const xj = pts[j].x, yj = pts[j].y;
    const crosses = ((yi > y) !== (yj > y)) &&
      (x < ((xj - xi) * (y - yi)) / (yj - yi) + xi);
    if (crosses) inside = !inside;
  }
  return inside;
}

/** Area-weighted (shoelace) centroid, with a vertex-mean fallback for
    degenerate, near-zero-area polygons. */
function polygonCentroid(pts) {
  let area = 0, cx = 0, cy = 0;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const cross = pts[j].x * pts[i].y - pts[i].x * pts[j].y;
    area += cross;
    cx += (pts[j].x + pts[i].x) * cross;
    cy += (pts[j].y + pts[i].y) * cross;
  }
  area /= 2;
  if (Math.abs(area) < 1e-7) {
    const n = pts.length || 1;
    let sx = 0, sy = 0;
    pts.forEach(p => { sx += p.x; sy += p.y; });
    return { x: sx / n, y: sy / n };
  }
  return { x: cx / (6 * area), y: cy / (6 * area) };
}

/** Difficulty tolerance: inflate the polygon about its centroid by a
    percentage margin (+5% … +25%). */
function expandPolygon(pts, marginPct) {
  if (!marginPct) return pts;
  const c = polygonCentroid(pts);
  const f = 1 + marginPct / 100;
  return pts.map(p => ({ x: c.x + (p.x - c.x) * f, y: c.y + (p.y - c.y) * f }));
}

function hitTest(hotspot, x, y, marginPct) {
  return pointInPolygon(x, y, expandPolygon(hotspot.points, marginPct || 0));
}

function dist(a, b) { return Math.hypot(a.x - b.x, a.y - b.y); }
function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }

/* -------------------------------- State -------------------------------- */
function defaultState() {
  return {
    version: 1,
    mode: 'pathology',                 // 'pathology' | 'anatomy'
    image: null,                       // { dataUrl, w, h }
    windowing: { ww: 100, wl: 100 },   // window width → contrast %, level → brightness %
    hotspots: [],                      // { id, label, pearl, tolerance, points: [{x,y}] }
  };
}

/** Defensive rehydration of persisted JSON (tolerates junk/older schemas). */
function normalizeState(raw) {
  const s = defaultState();
  if (!raw || typeof raw !== 'object') return s;
  if (raw.mode === 'anatomy') s.mode = 'anatomy';
  if (raw.image && typeof raw.image.dataUrl === 'string' && raw.image.dataUrl) {
    s.image = {
      dataUrl: raw.image.dataUrl,
      w: Math.max(1, +raw.image.w || 1),
      h: Math.max(1, +raw.image.h || 1),
    };
  }
  if (raw.windowing) {
    s.windowing.ww = clamp(+raw.windowing.ww || 100, 20, 300);
    s.windowing.wl = clamp(+raw.windowing.wl || 100, 20, 300);
  }
  if (Array.isArray(raw.hotspots)) {
    s.hotspots = raw.hotspots
      .filter(h => h && Array.isArray(h.points) && h.points.length >= MIN_POINTS)
      .map(h => ({
        id: String(h.id || uid()),
        label: String(h.label || ''),
        pearl: String(h.pearl || ''),
        tolerance: clamp(+h.tolerance || 10, 0, 25),
        points: h.points.map(p => ({
          x: clamp(+p.x || 0, 0, 100),
          y: clamp(+p.y || 0, 0, 100),
        })),
      }));
  }
  return s;
}

function freshSession() {
  return {
    found: new Set(),     // hotspot ids identified this run
    armedId: null,        // anatomy: structure currently selected in the sidebar
    missStreak: 0,        // consecutive misses → smart hint at HINT_AFTER_MISSES
    reveals: new Map(),   // hotspot id → timestamp (drives the fade-in animation)
  };
}

let state = defaultState();
let session = freshSession();

/* --------------------------- Runtime flags ---------------------------- */
let inOffice = false;   // actually hosted inside PowerPoint
let isPlay = false;     // presentation / read-only view → learner interface
let isPreview = false;  // creator-initiated "Test Run" of the learner interface
let draft = null;       // { points: [] } while outlining a new polygon
let editingId = null;   // hotspot id currently open in the config dialog
let hoveredId = null;   // hotspot row hovered in the edit list
let mousePct = null;    // cursor position while drafting (rubber-band segment)
let els = {};
let rafId = null;
let saveTimer = null;
let toastTimer = null;

/* ------------------------------ Utilities ------------------------------ */
function uid() { return 'h' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function byId(id) { return document.getElementById(id); }
function findHotspot(id) { return state.hotspots.find(h => h.id === id) || null; }
function modeColor() { return state.mode === 'anatomy' ? COLORS.anatomy : COLORS.pathology; }
function firstUnfound() { return state.hotspots.find(h => !session.found.has(h.id)) || null; }
function timeNow() {
  const d = new Date();
  return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
}
function hexA(hex, a) {
  const n = parseInt(hex.slice(1), 16);
  return 'rgba(' + ((n >> 16) & 255) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + a + ')';
}

/* ------------------------------ DOM cache ------------------------------ */
function cacheDom() {
  ['app', 'toolbar', 'btnModePathology', 'btnModeAnatomy', 'btnUpload', 'fileInput',
   'btnTest', 'btnSave', 'statusText', 'saveStatus', 'devBanner',
   'canvasFrame', 'canvasWrap', 'baseImage', 'overlay', 'fxLayer', 'emptyState',
   'progressChip', 'btnExitTest',
   'drawBar', 'drawInfo', 'btnUndoPt', 'btnClosePoly', 'btnCancelPoly',
   'wlBar', 'wwSlider', 'wlSlider', 'wwVal', 'wlVal', 'btnResetWL',
   'sidebar', 'editPanel', 'hotspotEmpty', 'hotspotList',
   'anatomyPanel', 'targetProgress', 'targetList',
   'pearlPanel', 'pearlTitle', 'pearlBody', 'btnClosePearl',
   'configModal', 'configTitle', 'labelCaption', 'labelField', 'pearlCaption',
   'pearlField', 'tolRow', 'tolVal', 'tolInput',
   'btnDeleteHotspot', 'btnCancelHotspot', 'btnSaveHotspot', 'toast']
    .forEach(id => { els[id] = byId(id); });
}

/* ------------------------------- Boot --------------------------------- */
function boot() {
  cacheDom();
  bindUi();
  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(info => {
      inOffice = !!(info && info.host);
      if (inOffice) {
        loadPersistedState();
        wireViewDetection();
      } else {
        els.devBanner.hidden = false; // opened in a plain browser (development)
      }
      start();
    });
  } else {
    els.devBanner.hidden = false;
    start();
  }
}

function start() {
  session = freshSession();
  if (state.image) els.baseImage.src = state.image.dataUrl;
  applyMode();
  syncWlInputs();
  setPlay(isPlay, isPreview);
}

function loadPersistedState() {
  try {
    const raw = Office.context.document.settings.get(SETTINGS_KEY);
    if (typeof raw === 'string' && raw.length) state = normalizeState(JSON.parse(raw));
    else if (raw && typeof raw === 'object') state = normalizeState(raw);
  } catch (err) {
    console.error('RadPoint: could not parse saved state', err);
  }
}

/** Presentation-mode detection: the active view is "read" during a
    slideshow (and in reading view), "edit" in the normal editor. */
function wireViewDetection() {
  const doc = Office.context.document;
  if (typeof doc.getActiveViewAsync === 'function') {
    doc.getActiveViewAsync(res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        setPlay(res.value === 'read');
      }
    });
  }
  if (typeof doc.addHandlerAsync === 'function') {
    doc.addHandlerAsync(Office.EventType.ActiveViewChanged, e => {
      setPlay(e.activeView === 'read');
    });
  }
}

/* --------------------------- Mode switching ---------------------------- */
function setPlay(play, preview) {
  isPlay = !!play;
  isPreview = !!(play && preview);
  draft = null;
  mousePct = null;
  hoveredId = null;
  session = freshSession();
  els.app.classList.remove('drafting');
  els.app.classList.toggle('mode-play', isPlay);
  els.app.classList.toggle('mode-edit', !isPlay);
  els.btnExitTest.hidden = !isPreview;
  closeModalSilent();
  closePearl();
  clearFx();
  buildSidebar();
  updateProgressChip();
  updateStatus();
  layoutCanvas();
}

function setMode(mode) {
  if (state.mode === mode) return;
  cancelDraft(true);
  state.mode = mode;
  applyMode();
  draw();
  scheduleSave();
}

function applyMode() {
  const anat = state.mode === 'anatomy';
  els.app.classList.toggle('game-anatomy', anat);
  els.app.classList.toggle('game-pathology', !anat);
  els.btnModePathology.classList.toggle('active', !anat);
  els.btnModeAnatomy.classList.toggle('active', anat);
  els.app.classList.toggle('has-image', !!state.image);
  buildSidebar();
  updateStatus();
}

/* ------------------------------ UI wiring ------------------------------ */
function bindUi() {
  els.btnModePathology.addEventListener('click', () => setMode('pathology'));
  els.btnModeAnatomy.addEventListener('click', () => setMode('anatomy'));

  els.btnUpload.addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', e => {
    if (e.target.files && e.target.files[0]) onFileChosen(e.target.files[0]);
    e.target.value = '';
  });

  els.btnSave.addEventListener('click', () => persist(true));
  els.btnTest.addEventListener('click', () => {
    if (!state.image) { toast('Upload an image first.'); return; }
    setPlay(true, true);
  });
  els.btnExitTest.addEventListener('click', () => setPlay(false));

  // Empty state doubles as an upload target (edit mode only)
  els.emptyState.addEventListener('click', () => { if (!isPlay) els.fileInput.click(); });
  ['dragover', 'dragenter'].forEach(t => els.canvasFrame.addEventListener(t, e => {
    if (isPlay) return;
    e.preventDefault();
    els.canvasFrame.classList.add('dragging');
  }));
  ['dragleave', 'drop'].forEach(t => els.canvasFrame.addEventListener(t, e => {
    e.preventDefault();
    els.canvasFrame.classList.remove('dragging');
  }));
  els.canvasFrame.addEventListener('drop', e => {
    if (isPlay) return;
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) onFileChosen(f);
  });

  // Canvas interactions
  els.overlay.addEventListener('click', onCanvasClick);
  els.overlay.addEventListener('dblclick', e => {
    e.preventDefault();
    if (!isPlay && draft) finishDraft(true);
  });
  els.overlay.addEventListener('contextmenu', e => {
    e.preventDefault();
    if (!isPlay && draft) undoPoint();
  });
  els.overlay.addEventListener('mousemove', e => {
    if (!isPlay && draft) { mousePct = eventPct(e); draw(); }
  });
  els.overlay.addEventListener('mouseleave', () => {
    if (mousePct) { mousePct = null; draw(); }
  });

  // Drawing action bar
  els.btnUndoPt.addEventListener('click', undoPoint);
  els.btnClosePoly.addEventListener('click', () => finishDraft(false));
  els.btnCancelPoly.addEventListener('click', () => cancelDraft(false));

  // Windowing controls
  els.wwSlider.addEventListener('input', () => setWindowing(+els.wwSlider.value, state.windowing.wl));
  els.wlSlider.addEventListener('input', () => setWindowing(state.windowing.ww, +els.wlSlider.value));
  els.btnResetWL.addEventListener('click', () => { setWindowing(100, 100); syncWlInputs(); });

  // Pearl panel + config dialog
  els.btnClosePearl.addEventListener('click', closePearl);
  els.btnSaveHotspot.addEventListener('click', saveConfig);
  els.btnCancelHotspot.addEventListener('click', cancelConfig);
  els.btnDeleteHotspot.addEventListener('click', deleteFromConfig);
  els.tolInput.addEventListener('input', () => {
    els.tolVal.textContent = '+' + els.tolInput.value + '%';
  });
  els.configModal.addEventListener('mousedown', e => {
    if (e.target === els.configModal) cancelConfig();
  });

  document.addEventListener('keydown', onKeyDown);

  // Responsive canvas
  if (typeof ResizeObserver !== 'undefined') {
    new ResizeObserver(() => layoutCanvas()).observe(els.canvasFrame);
  }
  window.addEventListener('resize', layoutCanvas);
}

function onKeyDown(e) {
  const tag = e.target && e.target.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA') {
    if (e.key === 'Escape' && !els.configModal.hidden) cancelConfig();
    return;
  }
  if (e.key === 'Escape') {
    if (!els.configModal.hidden) cancelConfig();
    else if (draft) cancelDraft(false);
    else closePearl();
  } else if (!isPlay && draft) {
    if (e.key === 'Enter') { e.preventDefault(); finishDraft(false); }
    else if (e.key === 'Backspace' || e.key.toLowerCase() === 'z') { e.preventDefault(); undoPoint(); }
  }
}

/* --------------------------- Canvas geometry --------------------------- */
function cssW() { return els.overlay.clientWidth || 1; }
function cssH() { return els.overlay.clientHeight || 1; }
function pctToPx(p) { return { x: (p.x / 100) * cssW(), y: (p.y / 100) * cssH() }; }

function eventPct(e) {
  const r = els.overlay.getBoundingClientRect();
  return {
    x: clamp(((e.clientX - r.left) / r.width) * 100, 0, 100),
    y: clamp(((e.clientY - r.top) / r.height) * 100, 0, 100),
  };
}

/** Fit the image (and its overlay canvas) inside the stage, preserving
    aspect ratio, at device-pixel-ratio resolution. */
function layoutCanvas() {
  const has = !!state.image;
  els.emptyState.hidden = has;
  els.canvasWrap.style.display = has ? 'block' : 'none';
  els.app.classList.toggle('has-image', has);
  if (!has) return;

  const availW = els.canvasFrame.clientWidth - 20;
  const availH = els.canvasFrame.clientHeight - 20;
  if (availW < 40 || availH < 40) return;

  const ratio = state.image.w / state.image.h;
  let w = availW, h = w / ratio;
  if (h > availH) { h = availH; w = h * ratio; }

  els.canvasWrap.style.width = Math.round(w) + 'px';
  els.canvasWrap.style.height = Math.round(h) + 'px';

  const dpr = window.devicePixelRatio || 1;
  els.overlay.width = Math.round(w * dpr);
  els.overlay.height = Math.round(h * dpr);
  els.overlay.getContext('2d').setTransform(dpr, 0, 0, dpr, 0, 0);
  draw();
}

/* ------------------------------- Drawing ------------------------------- */
function draw() {
  const ctx = els.overlay.getContext('2d');
  ctx.clearRect(0, 0, cssW(), cssH());
  if (!state.image) return;
  if (isPlay) drawPlay(ctx);
  else drawEdit(ctx);
}

function tracePath(ctx, pts) {
  ctx.beginPath();
  pts.forEach((p, i) => {
    const q = pctToPx(p);
    if (i === 0) ctx.moveTo(q.x, q.y);
    else ctx.lineTo(q.x, q.y);
  });
  ctx.closePath();
}

function drawEdit(ctx) {
  const color = modeColor();
  state.hotspots.forEach((h, idx) => {
    const hot = h.id === hoveredId || h.id === editingId;
    tracePath(ctx, h.points);
    ctx.fillStyle = hexA(color, hot ? 0.28 : 0.14);
    ctx.fill();
    ctx.setLineDash([]);
    ctx.strokeStyle = color;
    ctx.lineWidth = hot ? 3 : 2;
    ctx.stroke();
    const c = pctToPx(polygonCentroid(h.points));
    drawChip(ctx, c.x, c.y, h.label || defaultName(idx), color);
  });
  if (draft && draft.points.length) drawDraft(ctx);
}

function defaultName(idx) {
  return (state.mode === 'anatomy' ? 'Structure ' : 'Finding ') + (idx + 1);
}

function drawDraft(ctx) {
  const pts = draft.points;

  // Committed segments
  ctx.beginPath();
  pts.forEach((p, i) => {
    const q = pctToPx(p);
    if (i === 0) ctx.moveTo(q.x, q.y);
    else ctx.lineTo(q.x, q.y);
  });
  ctx.setLineDash([]);
  ctx.strokeStyle = COLORS.draft;
  ctx.lineWidth = 2;
  ctx.stroke();

  // Rubber-band segment to the cursor
  if (mousePct) {
    const last = pctToPx(pts[pts.length - 1]);
    const m = pctToPx(mousePct);
    ctx.beginPath();
    ctx.moveTo(last.x, last.y);
    ctx.lineTo(m.x, m.y);
    ctx.setLineDash([6, 5]);
    ctx.strokeStyle = hexA(COLORS.draft, 0.7);
    ctx.lineWidth = 1.5;
    ctx.stroke();
    ctx.setLineDash([]);
  }

  // Anchor points (first anchor doubles as the "close" target)
  pts.forEach((p, i) => {
    const q = pctToPx(p);
    ctx.beginPath();
    ctx.arc(q.x, q.y, i === 0 ? 6 : 4, 0, Math.PI * 2);
    ctx.fillStyle = (i === 0 && pts.length >= MIN_POINTS) ? COLORS.success : COLORS.draft;
    ctx.fill();
    ctx.strokeStyle = '#0a0e12';
    ctx.lineWidth = 1.5;
    ctx.stroke();
  });
}

function drawPlay(ctx) {
  const now = performance.now();
  let animating = false;
  state.hotspots.forEach(h => {
    if (!session.found.has(h.id)) return; // hidden until identified
    const t0 = session.reveals.get(h.id);
    let a = 1;
    if (t0 !== undefined) {
      a = clamp((now - t0) / REVEAL_MS, 0, 1);
      if (a < 1) animating = true;
    }
    const color = state.mode === 'anatomy' ? COLORS.anatomy : COLORS.success;
    tracePath(ctx, h.points);
    ctx.fillStyle = hexA(color, 0.16 * a);
    ctx.fill();
    ctx.setLineDash([]);
    ctx.strokeStyle = hexA(color, a);
    ctx.lineWidth = 2.5;
    ctx.stroke();

    // Anatomy: permanent label pin at the polygon centroid
    if (state.mode === 'anatomy') {
      const c = pctToPx(polygonCentroid(h.points));
      drawPin(ctx, c.x, c.y, h.label || 'Structure', a);
    }
  });
  if (animating) ensureRaf();
}

function drawChip(ctx, x, y, text, color) {
  ctx.font = '600 11px -apple-system, "Segoe UI", Roboto, sans-serif';
  const w = ctx.measureText(text).width + 12;
  const h = 18;
  const bx = clamp(x - w / 2, 2, cssW() - w - 2);
  const by = clamp(y - h / 2, 2, cssH() - h - 2);
  roundRect(ctx, bx, by, w, h, 4);
  ctx.fillStyle = 'rgba(10, 14, 18, 0.82)';
  ctx.fill();
  ctx.strokeStyle = hexA(color, 0.6);
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.fillStyle = '#e8eef5';
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, bx + 6, by + h / 2 + 0.5);
}

function drawPin(ctx, x, y, label, a) {
  // Anchor dot at the centroid
  ctx.beginPath();
  ctx.arc(x, y, 4, 0, Math.PI * 2);
  ctx.fillStyle = hexA(COLORS.anatomy, a);
  ctx.fill();
  ctx.strokeStyle = 'rgba(10, 14, 18, ' + (0.9 * a) + ')';
  ctx.lineWidth = 1.5;
  ctx.stroke();

  ctx.font = '600 11px -apple-system, "Segoe UI", Roboto, sans-serif';
  const w = ctx.measureText(label).width + 14;
  const h = 20;
  let bx = x + 10, by = y - h - 10;
  if (bx + w > cssW() - 2) bx = x - w - 10;
  if (by < 2) by = y + 10;
  bx = clamp(bx, 2, cssW() - w - 2);
  by = clamp(by, 2, cssH() - h - 2);

  // Leader line from the dot to the label chip
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(bx + w / 2, by + h);
  ctx.strokeStyle = hexA(COLORS.anatomy, 0.55 * a);
  ctx.lineWidth = 1;
  ctx.stroke();

  roundRect(ctx, bx, by, w, h, 5);
  ctx.fillStyle = 'rgba(11, 60, 74, ' + (0.92 * a) + ')';
  ctx.fill();
  ctx.strokeStyle = hexA(COLORS.anatomy, 0.8 * a);
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.fillStyle = 'rgba(224, 247, 255, ' + a + ')';
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, bx + 7, by + h / 2 + 0.5);
}

/** Path-based rounded rect (ctx.roundRect is missing in older WebViews). */
function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

/** One-shot rAF loop that keeps repainting while reveal fades are active. */
function ensureRaf() {
  if (rafId) return;
  rafId = requestAnimationFrame(() => { rafId = null; draw(); });
}

/* --------------------------- Click dispatch ---------------------------- */
function onCanvasClick(e) {
  if (!state.image) return;
  const pt = eventPct(e);
  if (isPlay) playClick(pt);
  else editClick(pt);
}

/* ====================== EDIT MODE: polygon authoring ====================== */
function editClick(pt) {
  if (!draft) {
    draft = { points: [] };
    els.app.classList.add('drafting');
  }
  const pts = draft.points;
  if (pts.length >= MIN_POINTS && dist(pt, pts[0]) <= CLOSE_SNAP_PCT) {
    finishDraft(false);
    return;
  }
  pts.push(pt);
  updateDrawBar();
  draw();
}

function undoPoint() {
  if (!draft || !draft.points.length) return;
  draft.points.pop();
  if (!draft.points.length) cancelDraft(true);
  else { updateDrawBar(); draw(); }
}

function cancelDraft(silent) {
  const had = draft && draft.points.length;
  draft = null;
  mousePct = null;
  els.app.classList.remove('drafting');
  updateDrawBar();
  draw();
  if (had && !silent) toast('Outline discarded.');
}

function finishDraft(fromDblClick) {
  if (!draft) return;
  const pts = draft.points;
  // A double-click registers a click first, leaving a duplicate anchor.
  if (fromDblClick && pts.length > MIN_POINTS &&
      dist(pts[pts.length - 1], pts[pts.length - 2]) < 0.8) {
    pts.pop();
  }
  if (pts.length < MIN_POINTS) { toast('A hotspot needs at least 3 points.'); return; }
  openConfig(null);
}

function updateDrawBar() {
  if (!draft) { els.drawInfo.textContent = ''; return; }
  const n = draft.points.length;
  els.drawInfo.textContent = 'Points: ' + n + (n >= MIN_POINTS
    ? ' — click the first point or press Enter to close'
    : ' — keep clicking to outline the region');
}

/* ------------------------ Hotspot config dialog ------------------------ */
function openConfig(hotspot) {
  editingId = hotspot ? hotspot.id : null;
  const anat = state.mode === 'anatomy';
  els.configTitle.textContent = hotspot ? 'Edit hotspot' : 'New hotspot';
  els.labelCaption.textContent = anat ? 'Structure name / label (required)' : 'Finding title (optional)';
  els.pearlCaption.textContent = anat ? 'Anatomical pearl (optional)' : 'Clinical explanation / pearl (required)';
  els.labelField.placeholder = anat ? 'e.g., Left renal vein' : 'e.g., Acute appendicitis';
  els.labelField.value = hotspot ? hotspot.label : '';
  els.pearlField.value = hotspot ? hotspot.pearl : '';
  els.tolInput.value = hotspot ? hotspot.tolerance : 10;
  els.tolVal.textContent = '+' + els.tolInput.value + '%';
  els.tolRow.hidden = anat;
  els.btnDeleteHotspot.hidden = !hotspot;
  els.configModal.hidden = false;
  els.labelField.focus();
  draw();
}

function saveConfig() {
  const anat = state.mode === 'anatomy';
  const label = els.labelField.value.trim();
  const pearl = els.pearlField.value.trim();
  if (anat && !label) { toast('Enter a structure name.'); return; }
  if (!anat && !pearl) { toast('Enter the clinical explanation.'); return; }

  if (editingId) {
    const h = findHotspot(editingId);
    if (h) {
      h.label = label;
      h.pearl = pearl;
      h.tolerance = clamp(+els.tolInput.value, 5, 25);
    }
  } else if (draft) {
    state.hotspots.push({
      id: uid(),
      label: label,
      pearl: pearl,
      tolerance: clamp(+els.tolInput.value, 5, 25),
      points: draft.points.slice(),
    });
    draft = null;
    els.app.classList.remove('drafting');
    updateDrawBar();
  }
  closeModalSilent();
  buildSidebar();
  updateStatus();
  draw();
  scheduleSave();
}

function cancelConfig() {
  const wasNew = !editingId;
  closeModalSilent();
  if (wasNew && draft) toast('Outline kept — keep adjusting, or press Esc to discard.');
  draw();
}

function deleteFromConfig() {
  if (editingId) {
    const id = editingId;
    state.hotspots = state.hotspots.filter(h => h.id !== id);
    toast('Hotspot deleted.');
  }
  closeModalSilent();
  buildSidebar();
  updateStatus();
  draw();
  scheduleSave();
}

function closeModalSilent() {
  editingId = null;
  els.configModal.hidden = true;
}

/* ==================== PRESENTATION MODE: game loops ==================== */
function playClick(pt) {
  if (!state.hotspots.length) { toast('No hotspots have been set up on this slide.'); return; }
  if (state.mode === 'anatomy') anatomyClick(pt);
  else pathologyClick(pt);
}

/** Pathology: click anywhere; any hotspot hit reveals its clinical pearl. */
function pathologyClick(pt) {
  // Clicking an already-identified finding re-opens its pearl (no penalty).
  for (const h of state.hotspots) {
    if (session.found.has(h.id) && hitTest(h, pt.x, pt.y, h.tolerance)) {
      showPearl(h);
      return;
    }
  }
  for (const h of state.hotspots) {
    if (!session.found.has(h.id) && hitTest(h, pt.x, pt.y, h.tolerance)) {
      identify(h);
      showPearl(h);
      return;
    }
  }
  registerMiss(pt, firstUnfound());
}

/** Anatomy: clicks are evaluated only against the armed structure. */
function anatomyClick(pt) {
  if (!session.armedId) {
    toast('Pick a structure from the list, then click it on the image.');
    return;
  }
  const h = findHotspot(session.armedId);
  if (!h) return;
  if (hitTest(h, pt.x, pt.y, 0)) {
    session.armedId = null;
    identify(h);
    if (h.pearl) showPearl(h); // optional anatomical pearl
  } else {
    registerMiss(pt, h);
  }
}

function identify(h) {
  session.found.add(h.id);
  session.missStreak = 0;
  session.reveals.set(h.id, performance.now());
  ensureRaf();
  buildSidebar();
  updateProgressChip();
  if (session.found.size === state.hotspots.length) {
    toast(state.mode === 'anatomy'
      ? 'All structures identified — well done.'
      : 'All findings identified — well done.');
  }
}

function registerMiss(pt, target) {
  spawnPing(pt);
  session.missStreak += 1;
  if (session.missStreak >= HINT_AFTER_MISSES) {
    session.missStreak = 0;
    if (target) showQuadrantHint(target);
  }
}

/* ------------------------------ Feedback FX ---------------------------- */
/** Red ripple at the exact click coordinates (CSS keyframe animation). */
function spawnPing(pt) {
  const el = document.createElement('div');
  el.className = 'ping';
  el.style.left = pt.x + '%';
  el.style.top = pt.y + '%';
  els.fxLayer.appendChild(el);
  el.addEventListener('animationend', () => el.remove());
}

/** Smart hint: pulse the image quadrant containing the target's centroid. */
function showQuadrantHint(h) {
  const c = polygonCentroid(h.points);
  const el = document.createElement('div');
  el.className = 'quadrant-hint';
  el.style.left = (c.x >= 50 ? 50 : 0) + '%';
  el.style.top = (c.y >= 50 ? 50 : 0) + '%';
  els.fxLayer.appendChild(el);
  el.addEventListener('animationend', () => el.remove());
}

function clearFx() { els.fxLayer.innerHTML = ''; }

/* ------------------------------ Pearl panel ---------------------------- */
function showPearl(h) {
  els.pearlTitle.textContent = h.label || (state.mode === 'anatomy' ? 'Structure' : 'Teaching point');
  els.pearlBody.textContent = h.pearl || '';
  els.pearlPanel.classList.add('open');
}

function closePearl() { els.pearlPanel.classList.remove('open'); }

/* ------------------------------- Sidebar ------------------------------- */
function buildSidebar() {
  buildHotspotList();
  buildTargetList();
}

function buildHotspotList() {
  const list = els.hotspotList;
  list.innerHTML = '';
  els.hotspotEmpty.hidden = state.hotspots.length > 0 || !state.image;
  state.hotspots.forEach((h, idx) => {
    const li = document.createElement('li');
    li.className = 'hotspot-row';

    const swatch = document.createElement('span');
    swatch.className = 'swatch';

    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = h.label || defaultName(idx);

    const meta = document.createElement('span');
    meta.className = 'meta';
    meta.textContent = h.points.length + ' pts' +
      (state.mode === 'pathology' ? ' · +' + h.tolerance + '%' : '');

    const edit = document.createElement('button');
    edit.type = 'button';
    edit.className = 'row-btn';
    edit.textContent = 'Edit';
    edit.addEventListener('click', e => { e.stopPropagation(); openConfig(h); });

    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'row-btn danger';
    del.textContent = '✕';
    del.title = 'Delete hotspot';
    del.addEventListener('click', e => {
      e.stopPropagation();
      state.hotspots = state.hotspots.filter(x => x.id !== h.id);
      buildSidebar();
      updateStatus();
      draw();
      scheduleSave();
    });

    li.append(swatch, name, meta, edit, del);
    li.addEventListener('mouseenter', () => { hoveredId = h.id; draw(); });
    li.addEventListener('mouseleave', () => { hoveredId = null; draw(); });
    li.addEventListener('click', () => openConfig(h));
    list.appendChild(li);
  });
}

function buildTargetList() {
  const list = els.targetList;
  list.innerHTML = '';
  const total = state.hotspots.length;
  els.targetProgress.textContent = total ? session.found.size + ' / ' + total : '';
  state.hotspots.forEach((h, idx) => {
    const done = session.found.has(h.id);
    const li = document.createElement('li');
    li.className = 'target' + (done ? ' done' : '') + (session.armedId === h.id ? ' armed' : '');

    const box = document.createElement('span');
    box.className = 'box';
    box.textContent = done ? '✓' : '';

    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = h.label || 'Structure ' + (idx + 1);

    li.append(box, name);
    li.addEventListener('click', () => {
      if (!isPlay || done) return;
      session.armedId = h.id;   // arm this structure
      session.missStreak = 0;
      buildTargetList();
    });
    list.appendChild(li);
  });
}

function updateProgressChip() {
  const show = isPlay && state.mode === 'pathology' && state.hotspots.length > 0;
  els.progressChip.hidden = !show;
  if (show) {
    els.progressChip.textContent = 'Findings ' + session.found.size + ' / ' + state.hotspots.length;
  }
}

/* ------------------------------ Windowing ------------------------------ */
function setWindowing(ww, wl) {
  state.windowing.ww = clamp(ww, 20, 300);
  state.windowing.wl = clamp(wl, 20, 300);
  applyWindowing();
  // Creator adjustments persist as the slide's preset; learner tweaks
  // during a presentation are per-session only (document is read-only).
  if (!isPlay) scheduleSave();
}

function applyWindowing() {
  els.baseImage.style.filter =
    'contrast(' + state.windowing.ww + '%) brightness(' + state.windowing.wl + '%)';
  els.wwVal.textContent = state.windowing.ww + '%';
  els.wlVal.textContent = state.windowing.wl + '%';
}

function syncWlInputs() {
  els.wwSlider.value = state.windowing.ww;
  els.wlSlider.value = state.windowing.wl;
  applyWindowing();
}

/* --------------------------- Image ingestion --------------------------- */
function onFileChosen(file) {
  if (!file.type || file.type.indexOf('image/') !== 0) {
    toast('Please choose an image file.');
    return;
  }
  setStatus('Loading image…');
  readAsDataURL(file)
    .then(url => loadImage(url).then(img => {
      const needsCompression = file.size > MAX_UPLOAD_BYTES || url.length > TARGET_DATAURL_CHARS;
      const finalUrl = needsCompression ? compressImage(img) : url;
      return loadImage(finalUrl).then(finalImg => {
        const hadHotspots = state.hotspots.length > 0;
        state.image = {
          dataUrl: finalUrl,
          w: finalImg.naturalWidth,
          h: finalImg.naturalHeight,
        };
        state.hotspots = [];
        session = freshSession();
        draft = null;
        els.app.classList.remove('drafting');
        els.baseImage.src = finalUrl;
        applyMode();
        layoutCanvas();
        scheduleSave();
        setStatus('');
        toast(hadHotspots
          ? 'Image replaced — previous hotspots were cleared.'
          : (needsCompression ? 'Image compressed for embedding and loaded.' : 'Image loaded.'));
      });
    }))
    .catch(err => {
      console.error('RadPoint: image load failed', err);
      setStatus('');
      toast('Could not load that image.');
    });
}

function readAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

/** Re-encode as JPEG via canvas, stepping down quality then dimensions
    until the Base64 payload fits comfortably in the settings store —
    prevents bloated .pptx files. */
function compressImage(img) {
  let scale = Math.min(1, MAX_COMPRESS_DIM / Math.max(img.naturalWidth, img.naturalHeight));
  const qualities = [0.92, 0.85, 0.75, 0.62, 0.5];
  let best = null;
  for (let pass = 0; pass < 4; pass++) {
    const w = Math.max(1, Math.round(img.naturalWidth * scale));
    const h = Math.max(1, Math.round(img.naturalHeight * scale));
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#000';                 // flatten transparency to black
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    for (const q of qualities) {
      const url = canvas.toDataURL('image/jpeg', q);
      if (!best || url.length < best.length) best = url;
      if (url.length <= TARGET_DATAURL_CHARS) return url;
    }
    scale *= 0.75;
  }
  return best;
}

/* ----------------------------- Persistence ----------------------------- */
function scheduleSave() {
  if (isPlay || !inOffice) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => persist(false), 900);
}

/** Serialize the full state to JSON and write it into the slide's XML
    via the document settings store. */
function persist(manual) {
  if (!inOffice) {
    if (manual) toast('Running outside PowerPoint — nothing was saved.');
    return;
  }
  clearTimeout(saveTimer);
  let json;
  try {
    json = JSON.stringify(state);
  } catch (err) {
    toast('Could not serialize slide state.');
    return;
  }
  if (json.length > 4500000) {
    toast('Warning: the embedded image is very large — consider re-uploading a smaller one.');
  }
  const settings = Office.context.document.settings;
  settings.set(SETTINGS_KEY, json);
  setStatus('Saving…');
  settings.saveAsync(res => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      setStatus('Saved ' + timeNow());
      if (manual) toast('Saved into the slide.');
    } else {
      setStatus('Save failed');
      toast('Save failed: ' + (res.error ? res.error.message : 'unknown error'));
    }
  });
}

/* ------------------------------ Status bar ----------------------------- */
function updateStatus() {
  const n = state.hotspots.length;
  els.statusText.textContent =
    (state.image ? '' : 'No image · ') + n + (n === 1 ? ' hotspot' : ' hotspots');
}

function setStatus(msg) { els.saveStatus.textContent = msg || ''; }

function toast(msg) {
  els.toast.textContent = msg;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove('show'), 2800);
}

/* --------------------------- Boot / test hooks ------------------------- */
// Pure helpers are exported so the geometry can be unit-tested under Node.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    pointInPolygon: pointInPolygon,
    polygonCentroid: polygonCentroid,
    expandPolygon: expandPolygon,
    hitTest: hitTest,
    normalizeState: normalizeState,
    defaultState: defaultState,
    clamp: clamp,
  };
}
if (typeof window !== 'undefined') {
  boot();
}
