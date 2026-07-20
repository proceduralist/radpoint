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

/** Distinct label colors for anatomy mode (auto-assigned, user-changeable). */
const PALETTE = ['#22d3ee', '#f59e0b', '#4ade80', '#f472b6', '#a78bfa',
                 '#fb923c', '#60a5fa', '#facc15', '#2dd4bf', '#f87171'];

/** First palette color not already in `used` (cycles once all are taken). */
function pickColor(used) {
  for (const c of PALETTE) if (used.indexOf(c) === -1) return c;
  return PALETTE[used.length % PALETTE.length];
}

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

/** A hotspot is hit when the click lands in ANY of its pieces. */
function hitTest(hotspot, x, y, marginPct) {
  const parts = hotspotParts(hotspot);
  for (const p of parts) {
    if (pointInPolygon(x, y, expandPolygon(p, marginPct || 0))) return true;
  }
  return false;
}

function distToSegment(p, a, b) {
  const dx = b.x - a.x, dy = b.y - a.y;
  const l2 = dx * dx + dy * dy;
  const t = l2 ? clamp(((p.x - a.x) * dx + (p.y - a.y) * dy) / l2, 0, 1) : 0;
  return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy));
}

function distToEdges(p, pts) {
  let d = Infinity;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    d = Math.min(d, distToSegment(p, pts[j], pts[i]));
  }
  return d;
}

/** Visual anchor guaranteed to lie inside the polygon — the interior point
    with the greatest clearance from every edge (pole of inaccessibility).
    The centroid only seeds the search; it wins only when nothing is deeper,
    so thin or winding shapes never get a boundary-hugging dot. Memoized. */
const anchorCache = typeof WeakMap !== 'undefined' ? new WeakMap() : null;
function polygonAnchor(pts) {
  if (anchorCache) {
    const hit = anchorCache.get(pts);
    if (hit) return hit;
  }
  let best = null, bestD = -1;
  const c = polygonCentroid(pts);
  if (pointInPolygon(c.x, c.y, pts)) { best = c; bestD = distToEdges(c, pts); }

  let minX = 100, minY = 100, maxX = 0, maxY = 0;
  pts.forEach(p => {
    minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x);
    minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y);
  });

  // Exhaustive grid pass — always runs, keeps the deepest interior point.
  const N = 64;
  for (let i = 0; i <= N; i++) {
    for (let j = 0; j <= N; j++) {
      const x = minX + ((maxX - minX) * i) / N;
      const y = minY + ((maxY - minY) * j) / N;
      if (!pointInPolygon(x, y, pts)) continue;
      const d = distToEdges({ x: x, y: y }, pts);
      if (d > bestD) { bestD = d; best = { x: x, y: y }; }
    }
  }

  // Local refinement around the best cell (halving steps, polylabel-style).
  if (best) {
    let step = Math.max(maxX - minX, maxY - minY) / N;
    for (let pass = 0; pass < 3; pass++) {
      const cur = best;
      for (let di = -2; di <= 2; di++) {
        for (let dj = -2; dj <= 2; dj++) {
          const x = cur.x + (di * step) / 2;
          const y = cur.y + (dj * step) / 2;
          if (!pointInPolygon(x, y, pts)) continue;
          const d = distToEdges({ x: x, y: y }, pts);
          if (d > bestD) { bestD = d; best = { x: x, y: y }; }
        }
      }
      step /= 2;
    }
  }

  // Pathological fallback: edge midpoints nudged inward, then the centroid.
  if (!best) {
    for (let i = 0, j = pts.length - 1; i < pts.length && !best; j = i++) {
      const mx = (pts[i].x + pts[j].x) / 2, my = (pts[i].y + pts[j].y) / 2;
      const nx = -(pts[i].y - pts[j].y), ny = pts[i].x - pts[j].x;
      const len = Math.hypot(nx, ny) || 1;
      for (const s of [0.5, -0.5]) {
        const x = mx + (nx / len) * s, y = my + (ny / len) * s;
        if (pointInPolygon(x, y, pts)) { best = { x: x, y: y }; break; }
      }
    }
  }
  if (!best) best = c;
  if (anchorCache) anchorCache.set(pts, best);
  return best;
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
    hotspots: [],                      // { id, label, pearl, tolerance, color, labelPos,
                                       //   parts: [ [{x,y}...], ... ] }  ← one structure,
                                       //   one or more discontiguous polygon pieces
  };
}

/** A hotspot's polygon pieces. Tolerates the legacy single-polygon `points`
    shape so older slides keep working. */
function hotspotParts(h) {
  if (h && Array.isArray(h.parts) && h.parts.length) return h.parts;
  if (h && Array.isArray(h.points) && h.points.length) return [h.points];
  return [];
}

/** |Shoelace| area of one polygon piece (normalized units). */
function partArea(pts) {
  let a = 0;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    a += pts[j].x * pts[i].y - pts[i].x * pts[j].y;
  }
  return Math.abs(a / 2);
}

/** The biggest piece — where the single shared label is anchored. */
function largestPart(parts) {
  let best = parts[0], bestA = -1;
  for (const p of parts) { const a = partArea(p); if (a > bestA) { bestA = a; best = p; } }
  return best || [];
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
      .map(h => {
        if (!h) return null;
        // New shape = parts[]; legacy shape = a single points[]. Accept both.
        let rawParts = Array.isArray(h.parts) ? h.parts
                     : Array.isArray(h.points) ? [h.points] : [];
        const parts = rawParts
          .filter(Array.isArray)
          .map(part => part.map(p => ({
            x: clamp(+p.x || 0, 0, 100),
            y: clamp(+p.y || 0, 0, 100),
          })))
          .filter(part => part.length >= MIN_POINTS);
        if (!parts.length) return null;
        return {
          id: String(h.id || uid()),
          label: String(h.label || ''),
          pearl: String(h.pearl || ''),
          tolerance: clamp(+h.tolerance || 10, 0, 25),
          color: (typeof h.color === 'string' && /^#[0-9a-f]{6}$/i.test(h.color))
            ? h.color.toLowerCase() : null,
          labelPos: (h.labelPos && isFinite(+h.labelPos.x) && isFinite(+h.labelPos.y))
            ? { x: clamp(+h.labelPos.x, 0, 100), y: clamp(+h.labelPos.y, 0, 100) }
            : null,
          parts: parts,
        };
      })
      .filter(Boolean);
    // Older saves have no colors — assign distinct ones.
    s.hotspots.forEach(h => {
      if (!h.color) h.color = pickColor(s.hotspots.map(x => x.color).filter(Boolean));
    });
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
let draft = null;       // { parts: [[...]], points: [] } while outlining a new structure
let addingRegion = false; // true while drawing an extra piece for the structure being named
let regionStash = null; // { label, pearl, tol, color } preserved between pieces
let editingId = null;   // hotspot id currently open in the config dialog
let hoveredId = null;   // hotspot row hovered in the edit list
let mousePct = null;    // cursor position while drafting (rubber-band segment)
let labelDrag = null;   // { id, moved } while dragging a label chip (edit mode)
let labelRects = [];    // px rects of label chips from the last edit-mode draw
let suppressClick = false; // swallow the click that follows a chip mousedown
let editColor = PALETTE[0]; // color selected in the config dialog
let els = {};
let rafId = null;
let saveTimer = null;
let toastTimer = null;

/* ------------------------------ Utilities ------------------------------ */
function uid() { return 'h' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function byId(id) { return document.getElementById(id); }
function findHotspot(id) { return state.hotspots.find(h => h.id === id) || null; }
function modeColor() { return state.mode === 'anatomy' ? COLORS.anatomy : COLORS.pathology; }
/** Per-hotspot color in anatomy mode; uniform mode color in pathology. */
function hotspotColor(h) {
  return state.mode === 'anatomy' ? (h.color || COLORS.anatomy) : modeColor();
}
/** Label chip under a normalized point, if any (topmost chip wins). */
function chipAt(pt) {
  const x = (pt.x / 100) * cssW(), y = (pt.y / 100) * cssH();
  for (let i = labelRects.length - 1; i >= 0; i--) {
    const r = labelRects[i];
    if (x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h) return r;
  }
  return null;
}
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
   'pearlField', 'tolRow', 'tolVal', 'tolInput', 'colorRow', 'colorSwatches', 'colorPicker',
   'regionRow', 'regionInfo', 'btnAddRegion', 'btnRemoveRegion',
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
  labelDrag = null;
  session = freshSession();
  els.app.classList.remove('drafting');
  els.overlay.classList.remove('over-label', 'drag-label');
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
    if (labelDrag) return; // handled by the document-level drag handler
    if (!isPlay && draft) { mousePct = eventPct(e); draw(); }
    else if (!isPlay && state.image) {
      els.overlay.classList.toggle('over-label', !!chipAt(eventPct(e)));
    }
  });
  els.overlay.addEventListener('mouseleave', () => {
    els.overlay.classList.remove('over-label');
    if (mousePct) { mousePct = null; draw(); }
  });

  // Label chips: mousedown starts a drag; a still click opens the editor.
  els.overlay.addEventListener('mousedown', e => {
    if (isPlay || draft || e.button !== 0) return;
    const r = chipAt(eventPct(e));
    if (r) {
      labelDrag = { id: r.id, moved: false };
      els.overlay.classList.add('drag-label');
      e.preventDefault();
    }
  });
  document.addEventListener('mousemove', e => {
    if (!labelDrag) return;
    const h = findHotspot(labelDrag.id);
    if (h) { labelDrag.moved = true; h.labelPos = eventPct(e); draw(); }
  });
  document.addEventListener('mouseup', () => {
    if (!labelDrag) return;
    const id = labelDrag.id, moved = labelDrag.moved;
    labelDrag = null;
    els.overlay.classList.remove('drag-label');
    suppressClick = true;                       // swallow the trailing click
    setTimeout(() => { suppressClick = false; }, 0);
    if (moved) scheduleSave();
    else { const h = findHotspot(id); if (h) openConfig(h); }
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
  PALETTE.forEach(c => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'color-swatch';
    b.style.background = c;
    b.dataset.color = c;
    b.addEventListener('click', () => selectColor(c));
    els.colorSwatches.appendChild(b);
  });
  els.colorPicker.addEventListener('input', () => selectColor(els.colorPicker.value));
  els.btnAddRegion.addEventListener('click', beginAddRegion);
  els.btnRemoveRegion.addEventListener('click', removeLastRegion);
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
  labelRects = [];
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

/** Fill with the even-odd rule so the painted region matches the ray-cast
    hit test exactly, even for self-crossing outlines. Feature-detected. */
let FILL_RULE = 'evenodd';
try { document.createElement('canvas').getContext('2d').fill(FILL_RULE); }
catch (err) { FILL_RULE = null; }
function fillShape(ctx) { if (FILL_RULE) ctx.fill(FILL_RULE); else ctx.fill(); }

function drawEdit(ctx) {
  state.hotspots.forEach((h, idx) => {
    const color = hotspotColor(h);
    const hot = h.id === hoveredId || h.id === editingId;
    const parts = hotspotParts(h);

    // Every discontiguous piece is painted in the structure's own color.
    parts.forEach(pp => {
      tracePath(ctx, pp);
      ctx.fillStyle = hexA(color, hot ? 0.28 : 0.14);
      fillShape(ctx);
      ctx.setLineDash([]);
      ctx.strokeStyle = color;
      ctx.lineWidth = hot ? 3 : 2;
      ctx.stroke();
    });

    // One shared label — on the biggest piece, or at its dragged position.
    const c = pctToPx(polygonAnchor(largestPart(parts)));
    let lx = c.x, ly = c.y;
    if (h.labelPos) {
      const q = pctToPx(h.labelPos);
      lx = q.x; ly = q.y;
      // Leader line + anchor dot tie the moved label back to its region.
      ctx.beginPath();
      ctx.moveTo(c.x, c.y);
      ctx.lineTo(lx, ly);
      ctx.strokeStyle = hexA(color, 0.5);
      ctx.lineWidth = 1;
      ctx.stroke();
      ctx.beginPath();
      ctx.arc(c.x, c.y, 3, 0, Math.PI * 2);
      ctx.fillStyle = color;
      ctx.fill();
    }
    const rect = drawChip(ctx, lx, ly, h.label || defaultName(idx), color);
    labelRects.push({ id: h.id, x: rect.x, y: rect.y, w: rect.w, h: rect.h });
  });
  if (draft && (draft.points.length || draft.parts.length)) drawDraft(ctx);
}

function defaultName(idx) {
  return (state.mode === 'anatomy' ? 'Structure ' : 'Finding ') + (idx + 1);
}

function drawDraft(ctx) {
  // Pieces of this structure already closed during this drawing session.
  draft.parts.forEach(pp => {
    tracePath(ctx, pp);
    ctx.setLineDash([]);
    ctx.fillStyle = hexA(COLORS.draft, 0.12);
    fillShape(ctx);
    ctx.strokeStyle = COLORS.draft;
    ctx.lineWidth = 2;
    ctx.stroke();
  });

  const pts = draft.points;
  if (!pts.length) return; // between pieces: nothing in progress yet

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
    const color = state.mode === 'anatomy' ? (h.color || COLORS.anatomy) : COLORS.success;
    const parts = hotspotParts(h);
    // Reveal every piece of the structure in the same color.
    parts.forEach(pp => {
      tracePath(ctx, pp);
      ctx.fillStyle = hexA(color, 0.16 * a);
      fillShape(ctx);
      ctx.setLineDash([]);
      ctx.strokeStyle = hexA(color, a);
      ctx.lineWidth = 2.5;
      ctx.stroke();
    });

    // Anatomy: one permanent label pin, anchored on the biggest piece.
    if (state.mode === 'anatomy') {
      const c = pctToPx(polygonAnchor(largestPart(parts)));
      drawPin(ctx, c.x, c.y, h.label || 'Structure', a, color, h.labelPos);
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
  return { x: bx, y: by, w: w, h: h };
}

function drawPin(ctx, x, y, label, a, color, labelPos) {
  color = color || COLORS.anatomy;
  // Anchor dot at the centroid
  ctx.beginPath();
  ctx.arc(x, y, 4, 0, Math.PI * 2);
  ctx.fillStyle = hexA(color, a);
  ctx.fill();
  ctx.strokeStyle = 'rgba(10, 14, 18, ' + (0.9 * a) + ')';
  ctx.lineWidth = 1.5;
  ctx.stroke();

  ctx.font = '600 11px -apple-system, "Segoe UI", Roboto, sans-serif';
  const w = ctx.measureText(label).width + 14;
  const h = 20;
  let bx, by;
  if (labelPos) {
    // Creator-chosen position (chip centered on the stored point)
    const q = pctToPx(labelPos);
    bx = q.x - w / 2;
    by = q.y - h / 2;
  } else {
    // Automatic placement: up-right of the dot, flipped when near an edge
    bx = x + 10;
    by = y - h - 10;
    if (bx + w > cssW() - 2) bx = x - w - 10;
    if (by < 2) by = y + 10;
  }
  bx = clamp(bx, 2, cssW() - w - 2);
  by = clamp(by, 2, cssH() - h - 2);

  // Leader line from the dot to the label chip (chip is painted on top)
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(bx + w / 2, labelPos ? by + h / 2 : by + h);
  ctx.strokeStyle = hexA(color, 0.55 * a);
  ctx.lineWidth = 1;
  ctx.stroke();

  roundRect(ctx, bx, by, w, h, 5);
  ctx.fillStyle = 'rgba(10, 14, 18, ' + (0.9 * a) + ')';
  ctx.fill();
  ctx.strokeStyle = hexA(color, 0.85 * a);
  ctx.lineWidth = 1.2;
  ctx.stroke();
  ctx.fillStyle = 'rgba(232, 238, 245, ' + a + ')';
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
  if (suppressClick) { suppressClick = false; return; }
  if (!state.image) return;
  const pt = eventPct(e);
  if (isPlay) playClick(pt);
  else editClick(pt);
}

/* ====================== EDIT MODE: polygon authoring ====================== */
function editClick(pt) {
  if (!draft) {
    draft = { parts: [], points: [] };
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
  // Nothing left anywhere → abandon the draft; otherwise stay in progress.
  if (!draft.points.length && !draft.parts.length) cancelDraft(true);
  else { updateDrawBar(); draw(); }
}

function cancelDraft(silent) {
  if (!draft) return;
  // While adding an extra piece, Esc/Discard just drops THIS piece and
  // returns to naming with the pieces gathered so far.
  if (addingRegion) {
    addingRegion = false;
    draft.points = [];
    mousePct = null;
    updateDrawBar();
    openConfig(null, { preserve: true });
    return;
  }
  const had = draft.points.length || draft.parts.length;
  draft = null;
  regionStash = null;
  mousePct = null;
  els.app.classList.remove('drafting');
  updateDrawBar();
  draw();
  if (had && !silent) toast('Outline discarded.');
}

/** Close the current piece (if any), then open the naming dialog. Between
    pieces (no in-progress points) this simply finishes the structure. */
function finishDraft(fromDblClick) {
  if (!draft) return;
  const pts = draft.points;
  // A double-click registers a click first, leaving a duplicate anchor.
  if (fromDblClick && pts.length > MIN_POINTS &&
      dist(pts[pts.length - 1], pts[pts.length - 2]) < 0.8) {
    pts.pop();
  }
  if (pts.length >= MIN_POINTS) {
    draft.parts.push(pts.slice());
    draft.points = [];
  } else if (pts.length > 0) {
    toast('A piece needs at least 3 points.');
    return;
  }
  if (!draft.parts.length) { toast('Outline a region first.'); return; }
  const preserve = addingRegion || !!regionStash;
  addingRegion = false;
  openConfig(null, { preserve: preserve });
}

/** Config dialog → "Add another region": stash the fields, hide the dialog,
    and go back to the canvas to outline the next piece of this structure. */
function beginAddRegion() {
  if (editingId) return; // pieces are added while creating a structure
  regionStash = {
    label: els.labelField.value,
    pearl: els.pearlField.value,
    tol: els.tolInput.value,
    color: editColor,
  };
  addingRegion = true;
  if (!draft) draft = { parts: [], points: [] };
  draft.points = [];
  els.configModal.hidden = true;
  els.app.classList.add('drafting');
  updateDrawBar();
  draw();
  toast('Outline the next piece, then close it.');
}

/** Config dialog → "Remove last": drop the most recent piece while creating. */
function removeLastRegion() {
  if (!draft || draft.parts.length < 2) return;
  draft.parts.pop();
  const n = draft.parts.length;
  els.regionInfo.textContent = n + (n === 1 ? ' region' : ' regions');
  els.btnRemoveRegion.hidden = n < 2;
  draw();
}

function updateDrawBar() {
  if (!draft) { els.drawInfo.textContent = ''; return; }
  const n = draft.points.length;
  const done = draft.parts.length;
  if (n === 0 && done) {
    els.drawInfo.textContent = done + (done === 1 ? ' piece' : ' pieces') +
      ' outlined — click to add another, or press Enter to name it';
    return;
  }
  const region = done ? ' (piece ' + (done + 1) + ')' : '';
  els.drawInfo.textContent = 'Points: ' + n + region + (n >= MIN_POINTS
    ? ' — click the first point or press Enter to close'
    : ' — keep clicking to outline the region');
}

/* ------------------------ Hotspot config dialog ------------------------ */
function selectColor(c) {
  if (typeof c === 'string') c = c.toLowerCase();
  editColor = c;
  Array.prototype.forEach.call(els.colorSwatches.children, b => {
    b.classList.toggle('selected', b.dataset.color === c);
  });
  if (els.colorPicker && /^#[0-9a-f]{6}$/i.test(c)) els.colorPicker.value = c;
}

function openConfig(hotspot, opts) {
  opts = opts || {};
  editingId = hotspot ? hotspot.id : null;
  const anat = state.mode === 'anatomy';
  els.configTitle.textContent = hotspot ? 'Edit hotspot' : 'New hotspot';
  els.labelCaption.textContent = anat ? 'Structure name / label (required)' : 'Finding title (optional)';
  els.pearlCaption.textContent = anat ? 'Anatomical pearl (optional)' : 'Clinical explanation / pearl (required)';
  els.labelField.placeholder = anat ? 'e.g., Left renal vein' : 'e.g., Acute appendicitis';

  if (opts.preserve && regionStash) {
    // Returning from "Add another region" — keep what was already typed.
    els.labelField.value = regionStash.label;
    els.pearlField.value = regionStash.pearl;
    els.tolInput.value = regionStash.tol;
    selectColor(regionStash.color);
  } else {
    els.labelField.value = hotspot ? hotspot.label : '';
    els.pearlField.value = hotspot ? hotspot.pearl : '';
    els.tolInput.value = hotspot ? hotspot.tolerance : 10;
    const used = state.hotspots
      .filter(x => !hotspot || x.id !== hotspot.id)
      .map(x => x.color).filter(Boolean);
    selectColor((hotspot && hotspot.color) || pickColor(used));
  }
  els.tolVal.textContent = '+' + els.tolInput.value + '%';
  els.tolRow.hidden = anat;
  els.colorRow.hidden = !anat;

  // Region row: piece count + add/remove (adding pieces only while creating).
  const parts = hotspot ? hotspotParts(hotspot) : (draft ? draft.parts : []);
  const n = parts.length;
  els.regionInfo.textContent = n + (n === 1 ? ' region' : ' regions');
  els.btnAddRegion.hidden = !!hotspot;
  els.btnRemoveRegion.hidden = !!hotspot || n < 2;

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
      h.color = editColor;
    }
  } else if (draft) {
    // Fold in any still-open piece, then commit every piece as one structure.
    if (draft.points.length >= MIN_POINTS) {
      draft.parts.push(draft.points.slice());
      draft.points = [];
    }
    if (!draft.parts.length) { toast('Outline at least one region.'); return; }
    state.hotspots.push({
      id: uid(),
      label: label,
      pearl: pearl,
      tolerance: clamp(+els.tolInput.value, 5, 25),
      color: editColor,
      labelPos: null,
      parts: draft.parts.map(p => p.slice()),
    });
    draft = null;
    addingRegion = false;
    regionStash = null;
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
  if (wasNew && draft && draft.parts.length) {
    // Keep every piece gathered so far; return to the canvas so the user can
    // add more, press Enter to name it, or Esc to discard.
    regionStash = {
      label: els.labelField.value,
      pearl: els.pearlField.value,
      tol: els.tolInput.value,
      color: editColor,
    };
    addingRegion = false;
    els.app.classList.add('drafting');
    updateDrawBar();
    toast('Pieces kept — add another, press Enter to name, or Esc to discard.');
  } else if (wasNew && draft) {
    draft = null;
    regionStash = null;
    els.app.classList.remove('drafting');
    updateDrawBar();
  }
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

/** Smart hint: pulse the image quadrant containing the target. */
function showQuadrantHint(h) {
  const c = polygonAnchor(largestPart(hotspotParts(h)));
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
    if (state.mode === 'anatomy') {
      swatch.style.background = h.color || COLORS.anatomy;
      swatch.title = 'Click to change color';
      swatch.addEventListener('click', e => {
        e.stopPropagation();
        h.color = PALETTE[(PALETTE.indexOf(h.color) + 1) % PALETTE.length];
        buildSidebar();
        draw();
        scheduleSave();
      });
    }

    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = h.label || defaultName(idx);

    const parts = hotspotParts(h);
    const totalPts = parts.reduce((s, p) => s + p.length, 0);
    const meta = document.createElement('span');
    meta.className = 'meta';
    meta.textContent = (parts.length > 1 ? parts.length + ' pcs · ' : '') + totalPts + ' pts' +
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
    if (done && h.color) {   // tie the checkmark to the label color on the image
      box.style.borderColor = h.color;
      box.style.color = h.color;
    }

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
    polygonAnchor: polygonAnchor,
    distToEdges: distToEdges,
    expandPolygon: expandPolygon,
    hitTest: hitTest,
    hotspotParts: hotspotParts,
    largestPart: largestPart,
    partArea: partArea,
    normalizeState: normalizeState,
    defaultState: defaultState,
    clamp: clamp,
  };
}
if (typeof window !== 'undefined') {
  boot();
}
