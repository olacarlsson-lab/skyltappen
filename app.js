'use strict';

// ── Konstanter ────────────────────────────────────────────────────────────────
const API_URL = 'https://vgregion.entryscape.net/rowstore/dataset/1802e57a-b25d-4716-8437-9ed648bbae59/json';

const MM = 2.8346;   // 1 mm i PDF-punkter (1pt = 1/72 tum)

const AVERY = {
  cardW:  85.0 * MM,
  cardH:  54.0 * MM,
  left:   15.0 * MM,
  top:    13.5 * MM,
  colGap: 10.0 * MM,
  rowGap:  0.0 * MM,
  cols: 2,
  rows: 5,
  perPage: 10,
};

const DEFAULT_MARGINS = { top: 3.5, right: 5.0, bot: 5.0, left: 7.5 };
let margins = { ...DEFAULT_MARGINS };

function loadMargins() {
  try {
    const d = localStorage.getItem('vgr_margins');
    if (d) margins = { ...DEFAULT_MARGINS, ...JSON.parse(d) };
  } catch (e) { console.warn('loadMargins:', e); }
}
function saveMargins() {
  try { localStorage.setItem('vgr_margins', JSON.stringify(margins)); } catch (e) { handleStorageError(e); }
}
const pad = () => ({
  left:  margins.left  * MM,
  right: margins.right * MM,
  top:   margins.top   * MM,
  bot:   margins.bot   * MM,
});

const A4_W = 595.28;
const A4_H = 841.89;

const FIELD_META = [
  { key: 'creator', label: 'Konstnär'       },
  { key: 'title',   label: 'Konstverksnamn' },
  { key: 'artform', label: 'Konsttyp'       },
  { key: 'created', label: 'År'             },
  { key: 'id',      label: 'VG-nummer'      },
];

// Fält där användaren kan lägga in manuella radbryt med Enter.
const MULTILINE_FIELDS = new Set(['creator', 'title', 'artform']);

// ── State ─────────────────────────────────────────────────────────────────────
let artworks    = [];          // Array<{ id, creator, title, artform, created }>
let selected    = new Set();   // Indexes i artworks som är markerade (legacy, kept for ctx menu compat)
let history     = [];          // Undo-stack (JSON-snapshots)
let searchTerm  = '';
let projectName = '';
let logoBytes  = null;        // Uint8Array med PNG-data
let editingIdx = -1;
let addingNew  = false;       // true när modal öppnades via "Lägg till manuellt"

// WYSIWYG state
let selectedSlot    = null;  // global slot index of selected label
let lastAnchorSlot  = null;  // anchor för Shift-klick range-select
let dragSrc         = null;  // drag source slot index
let lastDragOverEl  = null;  // element currently showing drag-over highlight

// ── Hjälpare ──────────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Delar text på \n och wrapp:ar varje rad i en <div class="cls">. Tomma strängar → "".
function renderMultiline(text, cls) {
  if (!text) return '';
  return String(text)
    .split('\n')
    .map(line => `<div class="${cls}">${esc(line) || '&nbsp;'}</div>`)
    .join('');
}

// ── Toast notifications ───────────────────────────────────────────────────────
function showToast(msg, type = 'info') {
  const container = $('toast-container');
  const icons = { info: 'ℹ️', success: '✓', error: '✕' };

  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${icons[type] || 'ℹ️'}</span>
    <span class="toast-msg">${msg}</span>
  `;
  container.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('toast-out');
    toast.addEventListener('animationend', () => toast.remove(), { once: true });
  }, 3500);
}

// ── Logo (inbäddad base64 — fungerar oavsett protokoll) ───────────────────────
const LOGO_DATA_URL = 'data:image/png;base64,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';

// Konverterar data-URL till Uint8Array för pdf-lib
function dataUrlToBytes(dataUrl) {
  const b64 = dataUrl.split(',')[1];
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

async function loadLogo() {
  logoBytes = dataUrlToBytes(LOGO_DATA_URL);
}

// ── Session (localStorage) ────────────────────────────────────────────────────
let storageWarned = false;
function handleStorageError(e) {
  const quota = e && (e.name === 'QuotaExceededError' || e.code === 22);
  if (quota && !storageWarned) {
    storageWarned = true;
    if (typeof showToast === 'function') {
      showToast('Lagringsutrymmet är fullt – ändringar sparas inte. Exportera listan och rensa.', 'error');
    }
  } else if (!quota) {
    console.warn('localStorage-fel:', e);
  }
}
function saveSession() {
  try { localStorage.setItem('vgr_konstskylt', JSON.stringify(artworks)); } catch (e) { handleStorageError(e); }
}
function loadSession() {
  try {
    const d = localStorage.getItem('vgr_konstskylt');
    if (d) artworks = JSON.parse(d);
  } catch (e) { handleStorageError(e); }
}
function saveProjectName() {
  try { localStorage.setItem('vgr_konstskylt_project', projectName); } catch (e) { handleStorageError(e); }
}
function loadProjectName() {
  try { projectName = localStorage.getItem('vgr_konstskylt_project') || ''; } catch (e) { handleStorageError(e); }
}

// ── Inställningar (fasta värden) ──────────────────────────────────────────────
const settings = {
  confirmClear:     true,
  autoFetchOnPaste: false,
  maxUndo:          20,
  minFontScale:     0.60,
};

// ── Undo ──────────────────────────────────────────────────────────────────────
function pushHistory() {
  history.push(JSON.stringify(artworks));
  while (history.length > settings.maxUndo) history.shift();
}
function undo() {
  if (!history.length) return;
  artworks = JSON.parse(history.pop());
  selected.clear();
  selectedSlot = null;
  const sel = $('sort-select');
  if (sel) sel.value = '';
  renderPages();
  closePopover();
  saveSession();
  showToast('Ångrade senaste åtgärden.', 'info');
}

// ── API-hämtning ──────────────────────────────────────────────────────────────
async function fetchOne(vgId) {
  const cacheKey = 'vgr_api_' + vgId.toUpperCase();
  const cached   = sessionStorage.getItem(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch { /* fall through */ }
  }

  try {
    const r = await fetch(`${API_URL}?id=${encodeURIComponent(vgId)}`, {
      signal: AbortSignal.timeout(8000),
    });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const data = await r.json();
    const row = (data.results || []).find(
      r => (r.id || '').trim().toUpperCase() === vgId.toUpperCase()
    );
    if (!row) return null;
    const result = {
      id:      row.id      || vgId,
      creator: row.creator || '',
      title:   row.title   || '',
      artform: row.artform || '',
      created: row.created || '',
    };
    try { sessionStorage.setItem(cacheKey, JSON.stringify(result)); } catch { /* quota */ }
    return result;
  } catch (e) {
    return e.name === 'TimeoutError' ? 'timeout' : 'error';
  }
}

/** Parsar VG-nummer ur fritext.
 *  Accepterar: VG1234, VG-1234, vg1234 (var som helst i texten) och
 *  isolerade 3–6-siffriga tal som står ensamma på en rad eller mellan
 *  avgränsare (,; tab), så att löpande text inte ger falska träffar. */
function parseVgNumbers(raw) {
  const seen = new Set();
  const result = [];
  const push = digits => {
    const norm = 'VG' + (digits.replace(/^0+/, '') || '0');
    if (!seen.has(norm)) { seen.add(norm); result.push(norm); }
  };
  for (const m of raw.matchAll(/\bVG[-]?(\d+)\b/gi)) push(m[1]);
  for (const token of raw.split(/[\s,;]+/)) {
    if (/^\d{3,6}$/.test(token)) push(token);
  }
  return result;
}

async function fetchAll() {
  const ids = parseVgNumbers($('vg-input').value);
  if (!ids.length) {
    showToast('Inga giltiga VG-nummer hittades.', 'error');
    return;
  }

  pushHistory();
  const btn = $('btn-fetch');
  btn.disabled = true;
  btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="animation:spin 0.8s linear infinite"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg> Hämtar…`;

  const pw = $('progress-wrap');
  const pf = $('progress-fill');
  pw.style.display = 'block';
  pf.style.width = '0%';

  const toFetch = ids.filter(id => !artworks.some(a => a.id.toUpperCase() === id.toUpperCase()));
  const notFound = [], errors = [], timeouts = [];
  let done = 0;

  const CONCURRENCY = 5;
  const queue = [...toFetch];
  const results = new Map();

  async function worker() {
    while (queue.length) {
      const vgId = queue.shift();
      const result = await fetchOne(vgId);
      results.set(vgId, result);
      done++;
      pf.style.width = `${(done / toFetch.length) * 100}%`;
    }
  }

  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, toFetch.length) }, worker));

  for (const vgId of toFetch) {
    const result = results.get(vgId);
    if (result && typeof result === 'object') {
      artworks.push(result);
    } else {
      if (result === null)            notFound.push(vgId);
      else if (result === 'timeout')  timeouts.push(vgId);
      else                            errors.push(vgId);
      artworks.push({ id: vgId, creator: '', title: '', artform: '', created: '' });
    }
  }

  renderPages();

  pf.style.width = '100%';
  setTimeout(() => { pw.style.display = 'none'; }, 600);

  btn.disabled = false;
  btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 17 12 21 16 17"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.88 18.09A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.29"/></svg> Hämta data`;
  saveSession();

  const added = toFetch.length - notFound.length - errors.length - timeouts.length;
  let type = 'success';
  let msg = `${ids.length} ID:n behandlade`;
  if (added > 0) msg += `, ${added} tillagda`;
  if (notFound.length) { msg += `. Ej funna: ${notFound.join(', ')}`; type = 'info'; }
  if (timeouts.length) { msg += `. Tidsgräns: ${timeouts.join(', ')}`; type = 'error'; }
  if (errors.length)   { msg += `. Fel: ${errors.join(', ')}`; type = 'error'; }
  msg += '.';
  showToast(msg, type);
}

// ── WYSIWYG Page Rendering ────────────────────────────────────────────────────

const MM_TO_PX = 96 / 25.4; // CSS px per mm at 96dpi
const A4_W_MM = 210, A4_H_MM = 297;

let currentScale = 1;

function fitScale() {
  const scaler = $('page-scaler');
  if (!scaler) return 1;
  const avail = scaler.clientWidth - 48;
  return Math.min(1, avail / (A4_W_MM * MM_TO_PX));
}

function artworkMatchesSearch(aw) {
  if (!searchTerm) return true;
  if (!aw) return false;
  return [aw.creator, aw.title, aw.artform, aw.created, aw.id]
    .some(v => (v || '').toLowerCase().includes(searchTerm));
}

function renderPages() {
  const scaler = $('page-scaler');
  const label  = $('list-label');
  if (!scaler) return;

  currentScale = fitScale();

  const wPx = A4_W_MM * MM_TO_PX * currentScale;
  const hPx = A4_H_MM * MM_TO_PX * currentScale;

  scaler.innerHTML = '';

  if (!artworks.length) {
    label.textContent = '0 konstverk';
    scaler.innerHTML = `<div class="wysiwyg-empty">
      <svg class="wysiwyg-empty-icon" width="56" height="56" viewBox="0 0 56 56" fill="none">
        <rect x="4" y="10" width="48" height="36" rx="4" stroke="currentColor" stroke-width="1.5" fill="none" opacity="0.15"/>
        <rect x="10" y="16" width="36" height="16" rx="2" stroke="currentColor" stroke-width="1.5" fill="none" opacity="0.25"/>
        <line x1="10" y1="38" x2="30" y2="38" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" opacity="0.3"/>
        <line x1="10" y1="43" x2="20" y2="43" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" opacity="0.2"/>
      </svg>
      <h2>Kom igång</h2>
      <ol class="tutorial-steps">
        <li>
          <span class="tutorial-step-num">1</span>
          <div>
            <strong>Ange VG-nummer</strong>
            <p>Klistra in VG-nummer i textfältet – blandat med text, kommatecken eller mellanslag fungerar. Du kan också ladda från fil eller urklipp med knapparna bredvid.</p>
          </div>
        </li>
        <li>
          <span class="tutorial-step-num">2</span>
          <div>
            <strong>Hämta data</strong>
            <p>Klicka "Hämta data" för att ladda konstverksinformation automatiskt.</p>
          </div>
        </li>
        <li>
          <span class="tutorial-step-num">3</span>
          <div>
            <strong>Redigera skyltar</strong>
            <p>Dubbelklicka på en skylt för att redigera konstnär, titel, konsttyp och år. Högerklicka för fler alternativ.</p>
          </div>
        </li>
        <li>
          <span class="tutorial-step-num">4</span>
          <div>
            <strong>Slå ihop skyltar</strong>
            <p>Markera flera skyltar (klicka + Shift/Ctrl) och välj "Slå ihop" för att samla ett konstverk på en skylt.</p>
          </div>
        </li>
        <li>
          <span class="tutorial-step-num">5</span>
          <div>
            <strong>Generera PDF</strong>
            <p>Klicka "Generera PDF" i menyn ovan – färdig för utskrift på Avery C32010.</p>
          </div>
        </li>
      </ol>
    </div>`;
    return;
  }

  const totalPages = Math.max(1, Math.ceil(artworks.length / 10));

  let matchCount = 0;
  let firstMatchEl = null;

  for (let p = 0; p < totalPages; p++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'page-wrapper';
    wrapper.style.width  = `${wPx}px`;
    wrapper.style.height = `${hPx}px`;

    const page = document.createElement('div');
    page.className = 'a4-page';
    page.style.transform = `scale(${currentScale})`;

    // 10 label slots per page
    for (let row = 0; row < 5; row++) {
      for (let col = 0; col < 2; col++) {
        const slot = p * 10 + row * 2 + col;
        const aw = artworks[slot] ?? null;

        const isMatch = aw && artworkMatchesSearch(aw);
        const labelEl = createLabelEl(slot, aw);

        if (aw && searchTerm) {
          if (isMatch) {
            labelEl.classList.add('search-match');
            matchCount++;
            if (!firstMatchEl) firstMatchEl = labelEl;
          } else {
            labelEl.classList.add('search-dim');
          }
        }

        // Exact Avery coordinates
        const x = 15 + col * (85 + 10);   // mm
        const y = 13.5 + row * 54;         // mm
        labelEl.style.left = `${x}mm`;
        labelEl.style.top  = `${y}mm`;

        page.appendChild(labelEl);
      }
    }

    const pnum = document.createElement('div');
    pnum.className = 'page-num';
    pnum.textContent = totalPages > 1 ? `Sida ${p + 1} av ${totalPages}` : '';
    wrapper.appendChild(pnum);

    wrapper.appendChild(page);
    scaler.appendChild(wrapper);
  }

  if (searchTerm) {
    label.textContent = `${matchCount} av ${artworks.length} konstverk`;
    if (firstMatchEl) {
      requestAnimationFrame(() => {
        firstMatchEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
    }
  } else {
    const pageLabel = totalPages > 1 ? ` · ${totalPages} sidor` : '';
    label.textContent = `${artworks.length} konstverk${pageLabel}`;
  }

  const projPart = projectName ? `${projectName} · ` : '';
  document.title = `${projPart}${artworks.length} konstverk · VGR Konstskylt`;
}

function createLabelEl(slot, aw) {
  const el = document.createElement('div');
  el.className = `label-slot ${aw ? 'filled' : 'empty'}`;
  el.dataset.slot = slot;
  if (aw) el.dataset.step = aw.sizeStep || 0;
  if (aw) el.style.setProperty('--ks', (aw.kerning || 0) / 100);

  if (aw) {
    if (slot === selectedSlot) el.classList.add('selected');

    el.innerHTML = `
      <div class="slot-badge">${slot + 1}</div>
      <div class="label-content" style="padding:${margins.top}mm ${margins.right}mm ${margins.bot}mm ${margins.left}mm">
        <div class="label-text-area">
          ${renderMultiline(aw.creator, 'sign-creator')}
          ${renderMultiline(aw.title,   'sign-title')}
          ${renderMultiline(aw.artform, 'sign-artform')}
          ${aw.created ? `<div class="sign-year">${esc(aw.created)}</div>`    : ''}
        </div>
        <div class="sign-bottom">
          <span class="sign-id">${esc(aw.id)}</span>
          <img class="sign-logo" src="${LOGO_DATA_URL}" alt="">
        </div>
      </div>`;

    el.draggable = true;
  } else {
    el.textContent = '+';
    el.title = 'Lägg till konstverk';
  }

  return el;
}

function selectSlot(slot, el, event = null) {
  const multi = event?.ctrlKey || event?.metaKey;
  const range = event?.shiftKey;

  if (range && lastAnchorSlot !== null) {
    // Shift+klick — välj intervall från ankaret till slot
    document.querySelectorAll('.label-slot.selected').forEach(e => e.classList.remove('selected'));
    selected.clear();
    const [a, b] = [lastAnchorSlot, slot].sort((x, y) => x - y);
    for (let i = a; i <= b; i++) {
      if (!artworks[i]) continue;
      selected.add(i);
      document.querySelector(`.label-slot[data-slot="${i}"]`)?.classList.add('selected');
    }
    selectedSlot = slot;
    closePopover(/*keepSelection*/ true);
  } else if (multi) {
    // Ctrl/Cmd+klick — lägg till eller ta bort ur urvalet
    if (selected.has(slot)) {
      selected.delete(slot);
      el.classList.remove('selected');
    } else {
      selected.add(slot);
      el.classList.add('selected');
    }
    lastAnchorSlot = slot;
    closePopover(/*keepSelection*/ true);
    selectedSlot = null;
  } else {
    // Vanligt klick — välj bara, öppna inte popover
    selected.forEach(s => {
      document.querySelector(`.label-slot[data-slot="${s}"]`)?.classList.remove('selected');
    });
    selected.clear();
    selected.add(slot);
    selectedSlot = slot;
    lastAnchorSlot = slot;
    el.classList.add('selected');
    closePopover(/*keepSelection*/ true);
  }

  updateMergeBar();
}

function updateMergeBar() {
  const bar = $('merge-bar');
  if (!bar) return;
  if (selected.size >= 2) {
    $('merge-bar-count').textContent = `${selected.size} konstverk valda`;
    bar.classList.add('visible');
  } else {
    bar.classList.remove('visible');
  }
}

function updateLabelEl(slot) {
  // Update just one label's content without full re-render
  const el = document.querySelector(`.label-slot[data-slot="${slot}"]`);
  if (!el) return;
  const aw = artworks[slot];
  if (!aw) return;
  const ta = el.querySelector('.label-text-area');
  const id = el.querySelector('.sign-id');
  if (ta) {
    ta.innerHTML = `
      ${renderMultiline(aw.creator, 'sign-creator')}
      ${renderMultiline(aw.title,   'sign-title')}
      ${renderMultiline(aw.artform, 'sign-artform')}
      ${aw.created ? `<div class="sign-year">${esc(aw.created)}</div>`    : ''}`;
  }
  if (id) id.textContent = aw.id;
  el.dataset.step = aw.sizeStep || 0;
  el.style.setProperty('--ks', (aw.kerning || 0) / 100);
  const pop = $('edit-popover');
  if (pop && pop.style.display !== 'none') {
    $('pop-vg-id').textContent = aw.id;
  }
}

// ── Popover ───────────────────────────────────────────────────────────────────

function openPopover(slot, labelEl) {
  const aw = artworks[slot];
  if (!aw) return;

  const pop = $('edit-popover');
  $('pop-vg-id').textContent  = aw.id      || '';
  $('pop-creator').value       = aw.creator || '';
  $('pop-title').value         = aw.title   || '';
  $('pop-artform').value       = aw.artform || '';
  $('pop-year').value          = aw.created || '';
  $('pop-id').value            = aw.id      || '';

  const step = aw.sizeStep || 0;
  pop.querySelectorAll('.pop-step-btn').forEach(btn =>
    btn.classList.toggle('active', parseInt(btn.dataset.step) === step)
  );

  const kerning = aw.kerning || 0;
  $('pop-kerning').value = kerning;
  $('pop-kerning-val').textContent = formatKerning(kerning);

  pop.style.display = 'block';
  positionPopover(labelEl);
}

function positionPopover(labelEl) {
  const pop  = $('edit-popover');
  const rect = labelEl.getBoundingClientRect();
  const pw   = 264;
  const ph   = pop.offsetHeight || 280;

  // Prefer right side; fall back to left
  let left = rect.right + 12;
  if (left + pw > window.innerWidth - 8) left = rect.left - pw - 12;
  left = Math.max(8, left);

  let top = rect.top + (rect.height - ph) / 2;
  top = Math.max(8, Math.min(top, window.innerHeight - ph - 8));

  pop.style.left = `${left}px`;
  pop.style.top  = `${top}px`;
}

function closePopover(keepSelection = false) {
  $('edit-popover').style.display = 'none';
  if (!keepSelection) {
    selected.forEach(s => document.querySelector(`.label-slot[data-slot="${s}"]`)?.classList.remove('selected'));
    selected.clear();
    selectedSlot = null;
    $('merge-bar')?.classList.remove('visible');
  }
}

function formatKerning(v) {
  if (v === 0) return '0';
  const em = v / 100;
  return (em > 0 ? '+' : '') + em.toFixed(2);
}

function bindPopoverInputs() {
  const fields = [
    ['pop-creator', 'creator'],
    ['pop-title',   'title'],
    ['pop-artform', 'artform'],
    ['pop-year',    'created'],
    ['pop-id',      'id'],
  ];
  fields.forEach(([inputId, key]) => {
    $(inputId).addEventListener('input', e => {
      if (selectedSlot === null) return;
      artworks[selectedSlot][key] = e.target.value;
      updateLabelEl(selectedSlot);
      saveSession();
    });
  });
  document.querySelectorAll('.pop-step-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (selectedSlot === null) return;
      const step = parseInt(btn.dataset.step);
      artworks[selectedSlot].sizeStep = step;
      document.querySelectorAll('.pop-step-btn').forEach(b =>
        b.classList.toggle('active', parseInt(b.dataset.step) === step)
      );
      const el = document.querySelector(`.label-slot[data-slot="${selectedSlot}"]`);
      if (el) el.dataset.step = step;
      saveSession();
    });
  });

  $('pop-kerning').addEventListener('input', e => {
    if (selectedSlot === null) return;
    const k = parseInt(e.target.value, 10);
    artworks[selectedSlot].kerning = k;
    $('pop-kerning-val').textContent = formatKerning(k);
    const el = document.querySelector(`.label-slot[data-slot="${selectedSlot}"]`);
    if (el) el.style.setProperty('--ks', k / 100);
    saveSession();
  });

  $('pop-kerning-reset').addEventListener('click', () => {
    if (selectedSlot === null) return;
    artworks[selectedSlot].kerning = 0;
    $('pop-kerning').value = 0;
    $('pop-kerning-val').textContent = formatKerning(0);
    const el = document.querySelector(`.label-slot[data-slot="${selectedSlot}"]`);
    if (el) el.style.setProperty('--ks', 0);
    saveSession();
  });

  $('pop-close').addEventListener('click', closePopover);

  // Close popover when clicking on the A4 page background (not on a label)
  $('page-scaler')?.addEventListener('click', e => {
    if (!e.target.closest('.label-slot')) closePopover();
  });
}

// ── Context menu ──────────────────────────────────────────────────────────────
function showCtxMenu(e, slot) {
  e.preventDefault();

  // Om högerklick på ett redan flervalt objekt — behåll urvalet.
  // Annars välj bara det klickade objektet.
  if (!selected.has(slot) || selected.size < 2) {
    selected.forEach(s => document.querySelector(`.label-slot[data-slot="${s}"]`)?.classList.remove('selected'));
    selected.clear();
    selected.add(slot);
    const el = document.querySelector(`.label-slot[data-slot="${slot}"]`);
    if (el) el.classList.add('selected');
  }
  selectedSlot = slot;

  const n = selected.size;
  $('ctx-merge').classList.toggle('disabled', n < 2);
  const canSplit = n === 1 && artworks[slot]?.id.includes(',');
  $('ctx-split').classList.toggle('disabled', !canSplit);

  const menu = $('ctx-menu');
  menu.style.display = 'block';
  menu.style.left = `${Math.min(e.clientX, window.innerWidth  - 200)}px`;
  menu.style.top  = `${Math.min(e.clientY, window.innerHeight - 240)}px`;
}

function hideCtxMenu() { $('ctx-menu').style.display = 'none'; }

function handleCtxAction(e) {
  const item = e.target.closest('.ctx-item');
  if (!item || item.classList.contains('disabled')) return;
  hideCtxMenu();

  const action  = item.dataset.action;
  const indices = [...selected].sort((a, b) => a - b);

  switch (action) {
    case 'undo':      undo(); break;
    case 'edit':      if (indices.length) openEdit(indices[0]); break;
    case 'duplicate': if (indices.length) duplicateSlot(indices[0]); break;
    case 'move-up':   moveItems(indices, -1); break;
    case 'move-down': moveItems(indices,  1); break;
    case 'merge':     mergeSelected(indices); break;
    case 'split':     if (indices.length) splitSelected(indices[0]); break;
    case 'delete':    deleteSelected(indices); break;
  }
}

// ── Edit modal (for manual add) ───────────────────────────────────────────────
function openEdit(idx) {
  if (idx < 0 || idx >= artworks.length) return;
  editingIdx = idx;
  const aw   = artworks[idx];
  const body = $('modal-body');
  body.innerHTML = '';

  FIELD_META.forEach(({ key, label }) => {
    const div = document.createElement('div');
    div.className = 'modal-field';
    if (MULTILINE_FIELDS.has(key)) {
      div.innerHTML = `<label>${label}</label>
        <textarea data-key="${key}" rows="2" spellcheck="false">${esc(aw[key] || '')}</textarea>`;
    } else {
      div.innerHTML = `<label>${label}</label>
        <input type="text" data-key="${key}" value="${esc(aw[key] || '')}">`;
    }
    body.appendChild(div);
  });

  $('modal-overlay').classList.add('open');
  body.querySelector('[data-key]')?.focus();
}

function closeModal() {
  if (addingNew && editingIdx >= 0 && editingIdx < artworks.length) {
    artworks.splice(editingIdx, 1);
    selected.clear();
    renderPages();
    saveSession();
  }
  addingNew  = false;
  editingIdx = -1;
  $('modal-overlay').classList.remove('open');
}

function saveModal() {
  if (editingIdx < 0) return;

  const fields = {};
  document.querySelectorAll('#modal-body [data-key]').forEach(inp => {
    const val = MULTILINE_FIELDS.has(inp.dataset.key)
      ? inp.value.replace(/[ \t]+(\n|$)/g, '$1').replace(/\n{3,}/g, '\n\n').trim()
      : inp.value.trim();
    fields[inp.dataset.key] = val;
  });

  if (addingNew) {
    const hasContent = Object.values(fields).some(v => v !== '');
    if (!hasContent) {
      showToast('Fyll i minst ett fält för att lägga till ett konstverk.', 'error');
      return;
    }
    artworks.splice(editingIdx, 1);   // ta bort sentineln
    pushHistory();                     // snapshot utan den tomma posten
    artworks.push(fields);
    addingNew = false;
  } else {
    pushHistory();
    Object.entries(fields).forEach(([key, val]) => {
      artworks[editingIdx][key] = val;
    });
  }

  closeModal();
  renderPages();
  saveSession();
  showToast('Ändringarna sparade.', 'success');
}

// ── Åtgärder ──────────────────────────────────────────────────────────────────
function addManual() {
  addingNew = true;
  artworks.push({ id: '', creator: '', title: '', artform: '', created: '' });
  selected.clear();
  selected.add(artworks.length - 1);
  openEdit(artworks.length - 1);
}

function deleteSelected(indices) {
  if (!indices.length) return;
  pushHistory();
  [...indices].sort((a, b) => b - a).forEach(i => artworks.splice(i, 1));
  selected.clear();
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
  showToast(`${indices.length} konstverk borttaget.`, 'info');
}

function moveItems(indices, dir) {
  if (!indices.length) return;
  const sorted = [...indices].sort((a, b) => a - b);
  if (dir === -1 && sorted[0] === 0) return;
  if (dir ===  1 && sorted[sorted.length - 1] === artworks.length - 1) return;

  pushHistory();
  const items = sorted.map(i => artworks[i]);
  sorted.reverse().forEach(i => artworks.splice(i, 1));
  sorted.reverse();
  sorted.forEach((oldIdx, j) => {
    artworks.splice(oldIdx + dir, 0, items[j]);
  });

  selected.clear();
  sorted.forEach(i => selected.add(i + dir));
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
}

function mergeSelected(indices) {
  if (indices.length < 2) return;
  pushHistory();

  const items = indices.map(i => artworks[i]);
  const unique = arr => [...new Set(arr.filter(Boolean))];

  const years = unique(items.map(a => a.created));
  const anyMissingYear = items.some(a => !a.created);
  if (anyMissingYear) years.push('årtal ej registrerat');

  const merged = {
    id:      items.map(a => a.id).join(', '),
    creator: unique(items.map(a => a.creator)).join(', '),
    title:   unique(items.map(a => a.title)).join(', '),
    artform: unique(items.map(a => a.artform)).join(', '),
    created: years.join(', '),
  };

  const insertAt = indices[0];
  [...indices].sort((a, b) => b - a).forEach(i => artworks.splice(i, 1));
  artworks.splice(insertAt, 0, merged);

  selected.clear();
  selected.add(insertAt);
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
  showToast(`${indices.length} konstverk hopslagna.`, 'success');
}

async function splitSelected(idx) {
  const aw    = artworks[idx];
  const parts = aw.id.split(',').map(s => s.trim()).filter(Boolean);
  if (parts.length < 2) return;

  if (!confirm(
    `Dela upp i ${parts.length} separata skyltar och hämta om från databasen?\n\n` +
    parts.map(p => '• ' + p).join('\n')
  )) return;

  pushHistory();
  artworks.splice(idx, 1);
  selectedSlot = null;
  closePopover();
  renderPages();

  const btn = $('btn-fetch');
  btn.disabled = true;

  for (let i = 0; i < parts.length; i++) {
    const result = await fetchOne(parts[i]);
    const entry = (result && typeof result === 'object')
      ? result
      : { id: parts[i], creator: '', title: '', artform: '', created: '' };
    artworks.splice(idx + i, 0, entry);
    renderPages();
  }

  btn.disabled = false;
  selected.clear();
  renderPages();
  saveSession();
  showToast('Uppdelning klar.', 'success');
}

// ── Export / Import ───────────────────────────────────────────────────────────
function exportList() {
  if (!artworks.length) { showToast('Listan är tom — inget att exportera.', 'error'); return; }
  const name = prompt('Namn på exporten (valfritt):', projectName || '');
  if (name === null) return;
  projectName = name.trim();
  saveProjectName();
  const ts      = new Date().toISOString().slice(0, 16).replace('T', '_').replace(':', '-');
  const proj    = projectName ? projectName.replace(/[^a-zA-Z0-9åäöÅÄÖ]+/g, '_').replace(/^_|_$/g, '') + '_' : '';
  const payload = { projectName, artworks };
  const json    = JSON.stringify(payload, null, 2);
  const blob    = new Blob([json], { type: 'application/json' });
  const url     = URL.createObjectURL(blob);
  const a       = Object.assign(document.createElement('a'), {
    href: url, download: `vgr_${proj}${ts}.json`,
  });
  a.click();
  URL.revokeObjectURL(url);
  showToast('Lista exporterad.', 'success');
}

function importList(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const raw    = JSON.parse(ev.target.result);
      // Stöd både gammalt format (array) och nytt format ({ projectName, artworks })
      const parsed = Array.isArray(raw) ? raw : (Array.isArray(raw.artworks) ? raw.artworks : null);
      if (!parsed) throw new Error('Ogiltigt format');
      if (artworks.length && !confirm(
        `Importera ${parsed.length} konstverk?\n\nDen nuvarande listan (${artworks.length} konstverk) ersätts. Du kan ångra med Ctrl/Cmd+Z.`
      )) return;
      pushHistory();
      if (!Array.isArray(raw) && raw.projectName) {
        projectName = raw.projectName;
        const inp = $('project-name');
        if (inp) inp.value = projectName;
        saveProjectName();
      }
      artworks = parsed
        .filter(a => a && typeof a === 'object')
        .map(a => ({
          id:       String(a.id      || ''),
          creator:  String(a.creator || ''),
          title:    String(a.title   || ''),
          artform:  String(a.artform || ''),
          created:  String(a.created || ''),
          sizeStep: Number.isInteger(a.sizeStep) ? a.sizeStep : 0,
          kerning:  Number.isInteger(a.kerning)  ? a.kerning  : 0,
        }));
      selected.clear();
      selectedSlot = null;
      closePopover();
      renderPages();
      saveSession();
      showToast(`${artworks.length} konstverk importerade.`, 'success');
    } catch {
      showToast('Kunde inte läsa filen – kontrollera att det är en giltig JSON-export.', 'error');
    }
  };
  reader.readAsText(file, 'UTF-8');
  e.target.value = '';
}

// ── Sortering ─────────────────────────────────────────────────────────────────
function sortArtworks(key, dir) {
  if (!artworks.length) return;
  pushHistory();
  artworks.sort((a, b) => {
    const av = (a[key] || '').toLowerCase();
    const bv = (b[key] || '').toLowerCase();
    if (!av && bv) return 1;
    if (av && !bv) return -1;
    return dir * av.localeCompare(bv, 'sv');
  });
  selected.clear();
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
  const dirStr   = dir === 1 ? 'asc' : 'desc';
  const sel      = $('sort-select');
  if (sel) sel.value = `${key}-${dirStr}`;
  const dirLabel = dir === 1 ? 'A–Ö' : 'Ö–A';
  const keyLabel = { creator: 'Konstnär', title: 'Titel', artform: 'Konsttyp' }[key] || key;
  showToast(`Sorterat: ${keyLabel} ${dirLabel}.`, 'info');
}

// ── Duplicera ─────────────────────────────────────────────────────────────────
function duplicateSlot(idx) {
  if (idx < 0 || idx >= artworks.length) return;
  pushHistory();
  const copy = { ...artworks[idx] };
  artworks.splice(idx + 1, 0, copy);
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
  showToast('Skylt duplicerad.', 'success');
}

// ── Direktutskrift ────────────────────────────────────────────────────────────
function printLabels() {
  if (!artworks.length) { showToast('Inget att skriva ut.', 'error'); return; }
  window.print();
}

function clearAll() {
  if (!artworks.length) return;
  if (settings.confirmClear && !confirm('Rensa hela listan?')) return;
  pushHistory();
  artworks = [];
  selected.clear();
  selectedSlot = null;
  closePopover();
  renderPages();
  saveSession();
  showToast('Listan rensad.', 'info');
}

async function pasteClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    if (!text.trim()) return;
    const area = $('vg-input');
    const cur  = area.value.trim();
    area.value = cur ? cur + '\n' + text.trim() : text.trim();
    showToast('Klistrat in från urklipp.', 'info');
    if (settings.autoFetchOnPaste) fetchAll();
  } catch {
    showToast('Kunde inte läsa urklipp — klistra in manuellt.', 'error');
  }
}

function loadFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const isExcel = /\.(xlsx|xls|ods)$/i.test(file.name);

  if (isExcel) {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb  = XLSX.read(ev.target.result, { type: 'array' });
        const ws  = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Plocka ut alla celler som matchar VG-nummermönstret
        const vgNums = [];
        const seen   = new Set();
        const push = digits => {
          const norm = 'VG' + (digits.replace(/^0+/, '') || '0');
          if (!seen.has(norm)) { seen.add(norm); vgNums.push(norm); }
        };
        for (const row of rows) {
          for (const cell of row) {
            const str = String(cell).trim();
            if (/^\d{3,6}$/.test(str)) { push(str); continue; }
            for (const m of str.matchAll(/\bVG[-]?(\d+)\b/gi)) push(m[1]);
          }
        }

        if (!vgNums.length) {
          showToast('Inga VG-nummer hittades i filen.', 'error');
          return;
        }

        const area = $('vg-input');
        const cur  = area.value.trim();
        area.value = cur ? cur + '\n' + vgNums.join('\n') : vgNums.join('\n');
        showToast(`${vgNums.length} VG-nummer inlästa från ${file.name}.`, 'success');
        if (settings.autoFetchOnPaste) fetchAll();
      } catch {
        showToast('Kunde inte läsa Excel-filen.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = ev => {
      const area = $('vg-input');
      const cur  = area.value.trim();
      const text = ev.target.result.trim();
      area.value = cur ? cur + '\n' + text : text;
      showToast(`Fil inläst: ${file.name}`, 'info');
      if (settings.autoFetchOnPaste) fetchAll();
    };
    reader.readAsText(file, 'UTF-8');
  }
  e.target.value = '';
}

function handleKeydown(e) {
  const active  = document.activeElement;
  const inField = !!active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA');

  if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
    e.preventDefault();
    undo();
  }

  // Ctrl/Cmd+A — välj alla
  if ((e.ctrlKey || e.metaKey) && e.key === 'a' && !inField) {
    e.preventDefault();
    selected.clear();
    artworks.forEach((_, i) => {
      selected.add(i);
      document.querySelector(`.label-slot[data-slot="${i}"]`)?.classList.add('selected');
    });
    updateMergeBar();
  }

  if (!inField && (e.key === 'Delete' || e.key === 'Backspace')) {
    e.preventDefault();
    if (selectedSlot !== null) {
      deleteSelected([selectedSlot]);
    }
  }

  // Pilar utan modifier — navigera mellan skyltar
  if (!inField && !e.altKey && !e.ctrlKey && !e.metaKey && ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key) && selectedSlot !== null) {
    e.preventDefault();
    const step = (e.key === 'ArrowUp' || e.key === 'ArrowDown') ? 2 : 1;
    const dir  = (e.key === 'ArrowUp' || e.key === 'ArrowLeft') ? -step : step;
    const newSlot = selectedSlot + dir;
    if (newSlot >= 0 && newSlot < artworks.length) {
      document.querySelector(`.label-slot[data-slot="${selectedSlot}"]`)?.classList.remove('selected');
      selected.clear();
      selected.add(newSlot);
      selectedSlot = newSlot;
      const newEl = document.querySelector(`.label-slot[data-slot="${newSlot}"]`);
      if (newEl) newEl.classList.add('selected');
      newEl?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  // Alt+pilar — flytta vald skylt (↑/↓ = en rad = 2 steg, ←/→ = ett steg)
  if (!inField && e.altKey && ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key) && selectedSlot !== null) {
    e.preventDefault();
    const step = (e.key === 'ArrowUp' || e.key === 'ArrowDown') ? 2 : 1;
    const dir  = (e.key === 'ArrowUp' || e.key === 'ArrowLeft') ? -step : step;
    const newSlot = selectedSlot + dir;
    if (newSlot >= 0 && newSlot < artworks.length) {
      pushHistory();
      [artworks[selectedSlot], artworks[newSlot]] = [artworks[newSlot], artworks[selectedSlot]];
      selectedSlot = newSlot;
      selected.clear();
      selected.add(newSlot);
      renderPages();
      saveSession();
      requestAnimationFrame(() => {
        document.querySelector(`.label-slot[data-slot="${newSlot}"]`)
          ?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      });
    }
  }

  if (e.key === 'Enter' && active?.closest('#modal-overlay') && active.tagName !== 'TEXTAREA') {
    saveModal();
  }
  if (e.key === 'Escape') {
    closeModal();
    closeArtworkSearch();
    hideCtxMenu();
    closePopover();
    // Rensa sökfältet om det har innehåll
    const si = $('search-input');
    if (si && si.value) {
      si.value = '';
      searchTerm = '';
      renderPages();
    }
  }
}

// ── PDF-generering ────────────────────────────────────────────────────────────

/** Bryter text i ord och returnerar rader som ryms inom maxWidth. */
function wrapText(text, font, size, maxWidth, maxLines) {
  const words = text.split(/\s+/);
  const lines = [];
  let cur = '';

  for (const word of words) {
    const test = cur ? cur + ' ' + word : word;
    if (font.widthOfTextAtSize(test, size) > maxWidth && cur) {
      lines.push(cur);
      if (lines.length >= maxLines) { cur = ''; break; }
      cur = word;
    } else {
      cur = test;
    }
  }
  if (cur && lines.length < maxLines) lines.push(cur);
  return lines;
}

/** Respekterar manuella \n och wrapp:ar varje del med wrapText. */
function wrapWithBreaks(text, font, size, maxWidth, maxLines) {
  if (!text) return [];
  const segments = String(text).split('\n');
  const out = [];
  for (const seg of segments) {
    if (out.length >= maxLines) break;
    const trimmed = seg.trim();
    if (!trimmed) { out.push(''); continue; }
    const wrapped = wrapText(trimmed, font, size, maxWidth, maxLines - out.length);
    out.push(...wrapped);
  }
  return out.slice(0, maxLines);
}

async function buildPDF() {
  const showGrid = false;
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  const doc      = await PDFDocument.create();
  const fontReg  = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);

  let logoImg = null;
  if (logoBytes) {
    try { logoImg = await doc.embedPng(logoBytes); } catch { }
  }

  const black = rgb(0, 0, 0);

  const STEP_SCALE = [1.0, 0.85, 0.70];

  /** Ritar en enskild skylt. cardX/cardY = nedre vänstra hörnet i pt. */
  const drawSign = (page, aw, cardX, cardY) => {
    const p    = pad();
    const maxW = AVERY.cardW - p.left - p.right;
    const maxH = AVERY.cardH - p.top  - p.bot  - 8 * MM;

    const creator  = aw.creator || '';
    const title    = aw.title   || '';
    const artform  = aw.artform || '';
    const created  = aw.created || '';
    const stepMul  = STEP_SCALE[aw.sizeStep || 0] ?? 1.0;

    const layout = scale => {
      const s = scale * stepMul;
      const sc = Math.max(8, Math.round(14 * s));
      const st = Math.max(8, Math.round(14 * s));
      const sa = Math.max(7, Math.round(12 * s));
      const sy = Math.max(7, Math.round(10 * s));
      return {
        cl: creator ? wrapWithBreaks(creator, fontReg,  sc, maxW, 3) : [],
        tl: title   ? wrapWithBreaks(title,   fontBold, st, maxW, 3) : [],
        al: artform ? wrapWithBreaks(artform, fontReg,  sa, maxW, 3) : [],
        sc, st, sa, sy,
      };
    };

    const neededH = ({ cl, tl, al, sc, st, sa, sy }) => {
      let h = 0;
      if (cl.length) h += sc + sc * 1.2 * (cl.length - 1) + 1.5 * MM;
      if (tl.length) h += st * 1.05 + st * 1.3 * (tl.length - 1) + 2.0 * MM;
      if (al.length) h += sa + sa * 1.2 * (al.length - 1) + 1.5 * MM;
      if (created)   h += sy;
      return h;
    };

    const worstW = ({ cl, tl, al, sc, st, sa }) => {
      let w = 1.0;
      for (const l of cl) w = Math.max(w, fontReg .widthOfTextAtSize(l, sc) / maxW);
      for (const l of tl) w = Math.max(w, fontBold.widthOfTextAtSize(l, st) / maxW);
      for (const l of al) w = Math.max(w, fontReg .widthOfTextAtSize(l, sa) / maxW);
      return w;
    };

    let lay = layout(1.0);
    const overflow = Math.max(neededH(lay) / maxH, worstW(lay));
    if (overflow > 1.0 && settings.minFontScale < 1.0) {
      lay = layout(Math.max(1.0 / overflow, settings.minFontScale));
    }

    const { cl, tl, al, sc, st, sa, sy } = lay;
    const ks = (aw.kerning || 0) / 100;

    let curY = cardY + AVERY.cardH - p.top;

    // Konstnär
    if (cl.length) {
      curY -= sc;
      page.setCharacterSpacing(ks * sc);
      cl.forEach((line, j) => {
        page.drawText(line, { x: cardX + p.left, y: curY, font: fontReg, size: sc, color: black });
        if (j < cl.length - 1) curY -= sc * 1.2;
      });
      curY -= 1.5 * MM;
    }

    // Titel
    if (tl.length) {
      curY -= st * 1.05;
      page.setCharacterSpacing(ks * st);
      tl.forEach((line, j) => {
        page.drawText(line, { x: cardX + p.left, y: curY, font: fontBold, size: st, color: black });
        if (j < tl.length - 1) curY -= st * 1.3;
      });
      curY -= 2.0 * MM;
    }

    // Konsttyp
    if (al.length) {
      curY -= sa;
      page.setCharacterSpacing(ks * sa);
      al.forEach((line, j) => {
        page.drawText(line, { x: cardX + p.left, y: curY, font: fontReg, size: sa, color: black });
        if (j < al.length - 1) curY -= sa * 1.2;
      });
      curY -= 1.5 * MM;
    }

    // År
    if (created) {
      curY -= sy;
      page.setCharacterSpacing(ks * sy);
      page.drawText(created, { x: cardX + p.left, y: curY, font: fontReg, size: sy, color: black });
    }

    // Återställ teckensärning inför nedre rad
    page.setCharacterSpacing(0);

    // Nedre rad: VG-nummer + logotyp
    const bottomY = cardY + p.bot;
    if (aw.id) {
      page.drawText(aw.id, { x: cardX + p.left, y: bottomY, font: fontReg, size: 7, color: black });
    }
    if (logoImg) {
      const logoH  = 11 * MM;
      const dims   = logoImg.scale(1);
      const logoW  = logoH * (dims.width / dims.height);
      // Bilden har ~30% tomt utrymme under textens baslinje (135/454 px).
      // Vi placerar bilden så att texten linjerar med bottomY.
      // bottomWhitespace = (135/454) * logoH
      const bottomWS = (135 / dims.height) * logoH;
      page.drawImage(logoImg, {
        x: cardX + AVERY.cardW - p.right - logoW,
        y: bottomY - bottomWS,
        width: logoW, height: logoH,
      });
    }

    // Grid (bara i förhandsgranskningsversion)
    if (showGrid) {
      page.drawRectangle({
        x: cardX, y: cardY,
        width: AVERY.cardW, height: AVERY.cardH,
        borderColor: rgb(0.65, 0.65, 0.65),
        borderWidth: 0.4,
      });
    }
  };

  // Bygg sidor
  const totalPages = Math.ceil(artworks.length / AVERY.perPage);
  let page = null;
  for (let i = 0; i < artworks.length; i++) {
    const slot = i % AVERY.perPage;
    if (slot === 0) {
      page = doc.addPage([A4_W, A4_H]);
      // Sidfot med projektnamn och sidnummer
      if (projectName || totalPages > 1) {
        const pageNum  = Math.floor(i / AVERY.perPage) + 1;
        const footParts = [];
        if (projectName) footParts.push(projectName);
        if (totalPages > 1) footParts.push(`Sida ${pageNum} av ${totalPages}`);
        page.drawText(footParts.join('  ·  '), {
          x: AVERY.left, y: 6 * MM,
          font: fontReg, size: 7, color: rgb(0.55, 0.55, 0.55),
        });
      }
    }

    const col   = slot % AVERY.cols;
    const row   = Math.floor(slot / AVERY.cols);
    const cardX = AVERY.left + col * (AVERY.cardW + AVERY.colGap);
    const cardY = A4_H - AVERY.top - (row + 1) * AVERY.cardH - row * AVERY.rowGap;

    drawSign(page, artworks[i], cardX, cardY);
  }

  return doc.save();
}

async function generatePDF() {
  if (!artworks.length) {
    showToast('Hämta minst ett konstverk först.', 'error');
    return;
  }
  showToast('Genererar PDF…', 'info');
  try {
    const bytes = await buildPDF();
    const blob  = new Blob([bytes], { type: 'application/pdf' });
    const url   = URL.createObjectURL(blob);
    const a = Object.assign(document.createElement('a'), {
      href: url, download: 'vgr_skyltar.pdf',
    });
    a.click();
    URL.revokeObjectURL(url);
    showToast('PDF nedladdad.', 'success');
  } catch (e) {
    showToast('PDF-generering misslyckades: ' + e.message, 'error');
    console.error(e);
  }
}

// ── Konstverk-sök ────────────────────────────────────────────────────────────
let artworkIndex  = null;
let fuseInstance  = null;
let indexLoading  = false;

function openArtworkSearch() {
  $('artwork-search-overlay').classList.add('open');
  $('artwork-search-input').focus();
  if (!artworkIndex && !indexLoading) {
    $('search-results-list').innerHTML = '<p class="search-hint-text">Laddar register…</p>';
    setSearchLoadState('Förbereder…', 0);
    loadArtworkIndex();
  } else if (!artworkIndex && indexLoading) {
    $('search-results-list').innerHTML = '<p class="search-hint-text">Laddar register…</p>';
  }
}

function closeArtworkSearch() {
  $('artwork-search-overlay').classList.remove('open');
}

async function loadArtworkIndex() {
  indexLoading = true;
  try {
    const raw = localStorage.getItem('vgr_search_index');
    if (raw) {
      const { ts, data } = JSON.parse(raw);
      if (Date.now() - ts < 24 * 60 * 60 * 1000) {
        artworkIndex = data;
        buildFuse();
        setSearchLoadState(null);
        showCacheAge(ts);
        refreshSearchResults();
        return;
      }
    }
  } catch { }

  const LIMIT       = 500;
  const CONCURRENCY = 5;

  setSearchLoadState('Laddar konstverk…', 0);

  const fetchJson = url => fetch(url, { signal: AbortSignal.timeout(15000) }).then(r => {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
  });

  try {
    const first = await fetchJson(`${API_URL}?_limit=${LIMIT}&_offset=0`);
    let all     = first.results || [];
    const total = first.resultCount || all.length;

    const queue = [];
    for (let off = LIMIT; off < total; off += LIMIT) {
      queue.push(`${API_URL}?_limit=${LIMIT}&_offset=${off}`);
    }

    const totalBatches = Math.ceil(total / LIMIT) || 1;
    let done = 1;

    async function worker() {
      while (queue.length) {
        const url  = queue.shift();
        const data = await fetchJson(url);
        if (data.results) all.push(...data.results);
        done++;
        setSearchLoadState(`Laddar… ${Math.round(done / totalBatches * 100)} %`, done / totalBatches);
      }
    }

    await Promise.all(Array.from({ length: CONCURRENCY }, worker));

    artworkIndex = all;
    buildFuse();
    setSearchLoadState(null);
    refreshSearchResults();

    const ts = Date.now();
    try { localStorage.setItem('vgr_search_index', JSON.stringify({ ts, data: all })); } catch (e) { handleStorageError(e); }
    showCacheAge(ts);
  } catch (e) {
    setSearchLoadState('Kunde inte ladda register: ' + e.message);
  } finally {
    indexLoading = false;
  }
}

function buildFuse() {
  fuseInstance = new Fuse(artworkIndex, {
    includeScore:   true,
    threshold:      0.35,
    ignoreLocation: true,
    keys: ['creator', 'title'],
  });
}

function showCacheAge(ts) {
  const row = $('search-cache-row');
  if (!row) return;
  const d = new Date(ts);
  const dateStr = d.toLocaleDateString('sv-SE', { year: 'numeric', month: 'long', day: 'numeric' });
  $('search-cache-age').textContent = `Databasen inläst ${dateStr}`;
  row.hidden = false;
}

async function refreshArtworkIndex() {
  localStorage.removeItem('vgr_search_index');
  artworkIndex  = null;
  fuseInstance  = null;
  indexLoading  = false;
  $('search-cache-row').hidden = true;
  await loadArtworkIndex();
}

function setSearchLoadState(msg, progress) {
  const wrap = $('search-load-wrap');
  if (!msg) { wrap.hidden = true; return; }
  wrap.hidden = false;
  $('search-load-text').textContent = msg;
  const bar = $('search-load-bar');
  if (progress != null) bar.style.width = `${Math.round(progress * 100)}%`;
}

let _searchDebounce = null;
function onArtworkSearchInput() {
  clearTimeout(_searchDebounce);
  _searchDebounce = setTimeout(refreshSearchResults, 180);
}

function refreshSearchResults() {
  const list  = $('search-results-list');
  const query = ($('artwork-search-input')?.value || '').trim();
  if (!list) return;

  if (!query) {
    list.innerHTML = '<p class="search-hint-text">Börja skriva för att söka på konstnär eller titel.</p>';
    return;
  }
  if (!fuseInstance) {
    list.innerHTML = '<p class="search-hint-text">Laddar register, försök igen om ett ögonblick…</p>';
    return;
  }

  const hits = fuseInstance.search(query, { limit: 50 });
  if (!hits.length) {
    list.innerHTML = '<p class="search-hint-text">Inga träffar.</p>';
    return;
  }

  const existingIds = new Set(artworks.map(a => (a.id || '').toUpperCase()));

  list.innerHTML = hits.map(({ item }) => {
    const id      = String(item.id || '').trim();
    const already = existingIds.has(id.toUpperCase());
    const meta    = [item.artform, item.created, id].filter(Boolean).join(' · ');
    return `<div class="search-result-item${already ? ' search-result-item--added' : ''}">
      <div class="search-result-info">
        <span class="search-result-creator">${esc(item.creator || '–')}</span>
        <span class="search-result-title">${esc(item.title || '–')}</span>
        <span class="search-result-meta">${esc(meta)}</span>
      </div>
      <button class="btn btn-sm ${already ? 'btn-ghost' : 'btn-primary'} search-add-btn"
        ${already ? 'disabled' : ''}
        data-id="${esc(id)}">
        ${already ? 'Tillagd' : 'Lägg till'}
      </button>
    </div>`;
  }).join('');
}

function addFromSearch(rawId) {
  const item = artworkIndex?.find(r => String(r.id || '').trim().toUpperCase() === rawId.toUpperCase());
  if (!item) return;
  if (artworks.some(a => (a.id || '').toUpperCase() === rawId.toUpperCase())) return;

  pushHistory();
  artworks.push({
    id:       String(item.id      || '').trim(),
    creator:  String(item.creator || '').trim(),
    title:    String(item.title   || '').trim(),
    artform:  String(item.artform || '').trim(),
    created:  String(item.created || '').trim(),
    sizeStep: 0,
    kerning:  0,
  });
  renderPages();
  saveSession();
  showToast(`${rawId} – ${item.title || item.creator} tillagd.`, 'success');
  refreshSearchResults();
}

// ── Delegerade scaler-events (en uppsättning lyssnare för alla skyltar) ──────
function bindScalerEvents() {
  const scaler = $('page-scaler');

  scaler.addEventListener('click', e => {
    const slotEl = e.target.closest('.label-slot');
    if (!slotEl) {
      // Klick på bakgrunden – avmarkera och rensa sökning
      closePopover();
      if (searchTerm) {
        searchTerm = '';
        const si = $('search-input');
        if (si) si.value = '';
        renderPages();
      }
      return;
    }
    if (slotEl.classList.contains('empty')) {
      if ($('edit-popover').style.display !== 'none') { closePopover(false); return; }
      addManual();
    } else {
      if ($('edit-popover').style.display !== 'none') { closePopover(false); return; }
      selectSlot(parseInt(slotEl.dataset.slot), slotEl, e);
    }
  });

  scaler.addEventListener('dblclick', e => {
    const slotEl = e.target.closest('.label-slot.filled');
    if (!slotEl) return;
    openPopover(parseInt(slotEl.dataset.slot), slotEl);
  });

  scaler.addEventListener('contextmenu', e => {
    const slotEl = e.target.closest('.label-slot.filled');
    if (!slotEl) return;
    showCtxMenu(e, parseInt(slotEl.dataset.slot));
  });

  scaler.addEventListener('dragstart', e => {
    const slotEl = e.target.closest('.label-slot.filled');
    if (!slotEl) return;
    dragSrc = parseInt(slotEl.dataset.slot);
    e.dataTransfer.effectAllowed = 'move';
    setTimeout(() => slotEl.classList.add('dragging'), 0);
  });

  scaler.addEventListener('dragover', e => {
    e.preventDefault();
    if (dragSrc === null) return;
    const slotEl = e.target.closest('.label-slot.filled');
    if (!slotEl || parseInt(slotEl.dataset.slot) === dragSrc) return;
    if (slotEl !== lastDragOverEl) {
      if (lastDragOverEl) lastDragOverEl.classList.remove('drag-over');
      slotEl.classList.add('drag-over');
      lastDragOverEl = slotEl;
    }
  });

  scaler.addEventListener('dragleave', e => {
    if (lastDragOverEl && !scaler.contains(e.relatedTarget)) {
      lastDragOverEl.classList.remove('drag-over');
      lastDragOverEl = null;
    }
  });

  scaler.addEventListener('dragend', e => {
    e.target.closest('.label-slot')?.classList.remove('dragging');
    if (lastDragOverEl) { lastDragOverEl.classList.remove('drag-over'); lastDragOverEl = null; }
    dragSrc = null;
  });

  scaler.addEventListener('drop', e => {
    e.preventDefault();
    const slotEl = e.target.closest('.label-slot.filled');
    if (!slotEl || dragSrc === null) return;
    const slot = parseInt(slotEl.dataset.slot);
    if (slot === dragSrc) return;
    if (lastDragOverEl) { lastDragOverEl.classList.remove('drag-over'); lastDragOverEl = null; }
    pushHistory();
    const insertAt = slot > dragSrc ? slot - 1 : slot;
    const [item] = artworks.splice(dragSrc, 1);
    artworks.splice(insertAt, 0, item);
    selectedSlot = null;
    dragSrc = null;
    renderPages();
    closePopover();
    saveSession();
    const landEl = document.querySelector(`.label-slot[data-slot="${insertAt}"]`);
    if (landEl) {
      landEl.classList.add('just-dropped');
      landEl.addEventListener('animationend', () => landEl.classList.remove('just-dropped'), { once: true });
    }
  });
}

// ── Spinner keyframe ──────────────────────────────────────────────────────────
(function injectSpinnerCSS() {
  const style = document.createElement('style');
  style.textContent = `@keyframes spin { to { transform: rotate(360deg); } }`;
  document.head.appendChild(style);
})();

// ── Bootstrap ─────────────────────────────────────────────────────────────────
async function init() {
  loadMargins();
  loadSession();
  loadProjectName();

  renderPages();
  await loadLogo();

  // Marginaler
  const marginInputs = { top: $('margin-top'), right: $('margin-right'), bot: $('margin-bot'), left: $('margin-left') };
  Object.entries(marginInputs).forEach(([key, inp]) => {
    inp.value = margins[key];
    inp.addEventListener('change', () => {
      const v = parseFloat(inp.value);
      if (!isNaN(v) && v >= 0) { margins[key] = v; saveMargins(); renderPages(); }
    });
  });
  $('btn-margins-reset').addEventListener('click', () => {
    margins = { ...DEFAULT_MARGINS };
    saveMargins();
    Object.entries(marginInputs).forEach(([key, inp]) => { inp.value = margins[key]; });
    renderPages();
  });

  // Sidebar buttons
  $('btn-fetch')     .addEventListener('click', fetchAll);
  $('btn-add-manual').addEventListener('click', addManual);
  $('btn-file')      .addEventListener('click', () => $('file-input').click());
  $('btn-clipboard') .addEventListener('click', pasteClipboard);
  $('file-input')    .addEventListener('change', loadFile);

  // Header buttons
  $('btn-clear')   .addEventListener('click', clearAll);
  $('btn-generate').addEventListener('click', generatePDF);
  $('btn-export')  .addEventListener('click', exportList);
  $('btn-import')  .addEventListener('click', () => $('import-input').click());
  $('import-input').addEventListener('change', importList);
  $('btn-print')   .addEventListener('click', printLabels);
  $('btn-search')  .addEventListener('click', openArtworkSearch);

  // Artwork search modal
  $('artwork-search-close').addEventListener('click', closeArtworkSearch);
  $('artwork-search-overlay').addEventListener('click', e => {
    if (e.target === $('artwork-search-overlay')) closeArtworkSearch();
  });
  $('artwork-search-input').addEventListener('input', onArtworkSearchInput);
  $('search-refresh-btn') .addEventListener('click', refreshArtworkIndex);
  $('search-results-list').addEventListener('click', e => {
    const btn = e.target.closest('.search-add-btn');
    if (btn && !btn.disabled) addFromSearch(btn.dataset.id);
  });

  // Sort
  $('sort-select').addEventListener('change', e => {
    const val = e.target.value;
    if (!val) return;
    const [key, dir] = val.split('-');
    sortArtworks(key, dir === 'asc' ? 1 : -1);
    // Behåll valt alternativ som indikation på aktiv sortering
  });

  // Search (debounced så renderPages inte triggas per tangenttryckning)
  let _searchDebounceMain = null;
  $('search-input').addEventListener('input', e => {
    searchTerm = e.target.value.toLowerCase();
    clearTimeout(_searchDebounceMain);
    _searchDebounceMain = setTimeout(renderPages, 150);
  });

  // Modal
  $('modal-close').addEventListener('click', closeModal);
  $('modal-save') .addEventListener('click', saveModal);
  $('modal-reset').addEventListener('click', () => openEdit(editingIdx));
  $('modal-overlay').addEventListener('click', e => {
    if (e.target === $('modal-overlay')) closeModal();
  });

  // Context menu
  $('ctx-menu').addEventListener('click', handleCtxAction);
  document.addEventListener('click', e => {
    if (!e.target.closest('#ctx-menu')) hideCtxMenu();
  });
  document.addEventListener('contextmenu', e => {
    if (!e.target.closest('.label-slot')) hideCtxMenu();
  });

  // Keyboard
  document.addEventListener('keydown', handleKeydown);

  // Bind popover inputs
  bindPopoverInputs();

  // Close popover on header/sidebar click
  document.querySelector('.app-header')?.addEventListener('click', e => {
    if (!e.target.closest('#edit-popover')) closePopover();
  });

  // Merge-fält
  $('merge-bar-btn').addEventListener('click', () => {
    const indices = [...selected].sort((a, b) => a - b);
    mergeSelected(indices);
    $('merge-bar').classList.remove('visible');
  });
  $('merge-bar-cancel').addEventListener('click', () => {
    closePopover();
  });

  // Bulk textstorlek
  document.querySelectorAll('.merge-step-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const step = parseInt(btn.dataset.step);
      pushHistory();
      selected.forEach(idx => { if (artworks[idx]) artworks[idx].sizeStep = step; });
      renderPages();
      saveSession();
      showToast(`Textstorlek ändrad för ${selected.size} skyltar.`, 'info');
    });
  });

  // Tab-fälla i popovern
  $('edit-popover').addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    const focusable = [...$('edit-popover').querySelectorAll(
      'textarea, input, button:not([disabled])'
    )];
    if (!focusable.length) return;
    const first = focusable[0], last = focusable[focusable.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault(); last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault(); first.focus();
    }
  });

  // Delegerade events för alla label-slots (en gång, inte per skylt)
  bindScalerEvents();

  // Refit pages on resize (debounced)
  let _resizeDebounce = null;
  const ro = new ResizeObserver(() => {
    clearTimeout(_resizeDebounce);
    _resizeDebounce = setTimeout(renderPages, 100);
  });
  ro.observe($('page-viewport'));
}

init();
