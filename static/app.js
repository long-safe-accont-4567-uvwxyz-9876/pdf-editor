/* PDF 编辑器前端 —— 无框架，直接操作 DOM */

const $ = (id) => document.getElementById(id);

const FONT_CSS = {
  '微软雅黑': '"Microsoft YaHei", "PingFang SC", sans-serif',
  '黑体': 'SimHei, "Heiti SC", sans-serif',
  '宋体': 'SimSun, "Songti SC", serif',
  '楷体': 'KaiTi, "Kaiti SC", serif',
  '仿宋': 'FangSong, "Fangsong SC", serif',
  '等线': 'DengXian, "Microsoft YaHei", sans-serif',
  'Arial': 'Arial, Helvetica, sans-serif',
  'Times New Roman': '"Times New Roman", Times, serif',
  'Calibri': 'Calibri, sans-serif',
  'Courier New': '"Courier New", monospace',
  'Georgia': 'Georgia, serif',
  'Verdana': 'Verdana, sans-serif',
  'Cambria': 'Cambria, Georgia, serif',
};

const FONT_OPTIONS = Object.keys(FONT_CSS);

const state = {
  docId: null,
  name: '',
  pageCount: 0,
  pages: [],
  cur: 0,
  zoom: 1,
  dpi: 150,
  items: [],
  opMap: new Map(),
  images: [],
  selected: null,
  selectedImage: null,
  editing: null,
  busy: false,
  hits: [],
  hitIndex: -1,
  thumbObserver: null,
  suppressClick: false,
};

/* ------------------------------------------------------------ 基础工具 */

let toastTimer = null;
function toast(msg, isErr = false) {
  const el = $('toast');
  el.textContent = msg;
  el.className = 'toast show' + (isErr ? ' err' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 2400);
}

function setBusy(on) {
  state.busy = on;
  $('loading').hidden = !on;
}

const rgbCss = (a) => `rgb(${a[0]}, ${a[1]}, ${a[2]})`;
const rgbHex = (a) => '#' + a.map((v) => Math.max(0, Math.min(255, v | 0)).toString(16).padStart(2, '0')).join('');
const hexToRgb = (h) => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
const fontCss = (label) => FONT_CSS[label] || FONT_CSS['微软雅黑'];
const pxPerPt = () => (state.dpi / 72) * state.zoom;
const ptToPx = (pt) => pt * pxPerPt();

function debounce(fn, ms) {
  let t = null;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

async function apiGet(path) {
  const r = await fetch(path);
  if (!r.ok) {
    let msg = '请求失败';
    try { msg = (await r.json()).error || msg; } catch (e) { /* ignore */ }
    throw new Error(msg);
  }
  return r.json();
}

async function apiPost(path, body) {
  const r = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body || {}),
  });
  let data = null;
  try { data = await r.json(); } catch (e) { data = { ok: false, error: '返回解析失败' }; }
  if (!r.ok || data.ok === false) {
    const err = new Error(data.error || '操作失败');
    err.cancelled = !!data.cancelled;
    throw err;
  }
  return data;
}

/* ------------------------------------------------------------ 动效 */

/** anime.js 只用于需要 JS 编排的动画；失败即静默降级到 CSS */
function floatArt() {
  const el = $('emptyArt');
  if (!el) return;
  try {
    if (window.anime && anime.animate) {
      anime.animate(el, {
        opacity: [0, 1],
        translateY: [10, 0],
        scale: [0.97, 1],
        duration: 520,
        ease: 'outQuad',
      });
      anime.animate(el, {
        translateY: [-4, 4],
        duration: 2800,
        loop: true,
        alternate: true,
        ease: 'inOutSine',
        delay: 600,
      });
    }
  } catch (e) {
    el.style.opacity = '1';
  }
}

/** 面板字段依次浮现 */
function animatePanel(el) {
  if (!el) return;
  el.classList.remove('is-anim');
  void el.offsetWidth;
  el.classList.add('is-anim');
}

/** 页面切换时的轻微入场 */
function animatePage() {
  const page = $('page');
  page.classList.remove('is-fresh');
  void page.offsetWidth;
  page.classList.add('is-fresh');
}

/* ------------------------------------------------------------ 打开文档 */

async function openViaDialog() {
  if (state.busy) return;
  setBusy(true);
  try {
    const d = await apiPost('/api/dialog/open');
    await adoptDocument(d);
  } catch (e) {
    if (!e.cancelled) $('filePdf').click();
  } finally {
    setBusy(false);
  }
}

async function openViaFile(file) {
  if (!file) return;
  if (!/\.pdf$/i.test(file.name)) { toast('请选择 PDF 文件', true); return; }
  setBusy(true);
  try {
    const r = await fetch('/api/upload?name=' + encodeURIComponent(file.name), { method: 'POST', body: file });
    const d = await r.json();
    if (!r.ok || d.ok === false) throw new Error(d.error || '打开失败');
    await adoptDocument(d);
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
}

async function adoptDocument(d) {
  state.docId = d.docId;
  state.name = d.name;
  state.pageCount = d.pageCount;
  state.pages = d.pages || [];
  state.cur = 0;
  state.selected = null;
  state.selectedImage = null;
  state.editing = null;
  state.hits = [];

  document.body.classList.add('has-doc');
  $('empty').hidden = true;
  $('stage').hidden = false;
  $('btnExport').disabled = false;
  $('btnAddText').disabled = false;
  $('btnMosaic').disabled = false;
  $('btnFind').disabled = false;
  $('btnPage').disabled = false;
  $('pageCount').textContent = d.pageCount;

  buildThumbs();
  await fitToWidth();
  showInspector(null);
  refreshHistory();
}

async function reloadDoc(keepPage = true, refit = false) {
  const d = await apiGet(`/api/info?doc=${state.docId}`);
  const old = state.cur;
  state.pageCount = d.pageCount;
  state.pages = d.pages;
  $('pageCount').textContent = d.pageCount;
  state.cur = keepPage ? Math.max(0, Math.min(old, d.pageCount - 1)) : 0;
  state.selected = null;
  state.selectedImage = null;
  buildThumbs();
  markThumbActive();
  if (refit) await fitToWidth();
  else await renderPage();
  showInspector(null);
  refreshHistory();
}

function buildThumbs() {
  const box = $('thumbs');
  box.innerHTML = '';
  if (state.thumbObserver) state.thumbObserver.disconnect();
  state.thumbObserver = new IntersectionObserver((entries) => {
    entries.forEach((en) => {
      if (!en.isIntersecting) return;
      const el = en.target;
      state.thumbObserver.unobserve(el);
      const img = el.querySelector('img');
      if (img && !img.src) img.src = `/api/page?doc=${state.docId}&n=${el.dataset.n}&dpi=36&t=${Date.now()}`;
    });
  }, { root: box, rootMargin: '220px' });

  for (let i = 0; i < state.pageCount; i++) {
    const el = document.createElement('div');
    el.className = 'thumb' + (i === state.cur ? ' active' : '');
    el.dataset.n = i;
    const sz = state.pages[i] || state.pages[0];
    if (sz && sz[1]) el.style.setProperty('--ar', String(sz[0] / sz[1]));
    el.innerHTML = `<img alt="第 ${i + 1} 页"><span class="num">${i + 1}</span>`;
    el.addEventListener('click', () => gotoPage(i));
    box.appendChild(el);
    state.thumbObserver.observe(el);
  }
}

function markThumbActive() {
  document.querySelectorAll('.thumb').forEach((el) => {
    el.classList.toggle('active', +el.dataset.n === state.cur);
  });
}

/* ------------------------------------------------------------ 页面渲染 */

function pickDpi() {
  const dpr = window.devicePixelRatio || 1;
  return Math.max(110, Math.min(300, Math.round(96 * state.zoom * dpr)));
}

async function gotoPage(n) {
  if (!state.docId || n < 0 || n >= state.pageCount) return;
  if (state.editing) cancelEdit();
  state.cur = n;
  state.selected = null;
  state.selectedImage = null;
  markThumbActive();
  await renderPage();
  showInspector(null);
  const t = document.querySelector(`.thumb[data-n="${n}"]`);
  if (t) t.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

async function renderPage() {
  if (!state.docId) return;
  const n = state.cur;
  state.dpi = pickDpi();

  const info = await apiGet(`/api/spans?doc=${state.docId}&n=${n}`);

  state.items = (info.spans || []).map((s) => ({ ...s, kind: 'span' }));
  (info.freeTexts || []).forEach((f) => {
    state.items.push({
      id: f.id, kind: 'free', text: f.text, bbox: f.rect, origin: f.origin,
      size: f.size, color: f.color, font: f.font, fontLabel: guessLabel(f.font),
      bold: f.bold, italic: f.italic, align: f.align || 'left',
      offset: [0, 0], lineHeight: f.lineHeight || 1.2, page: f.page,
    });
  });

  state.opMap = new Map();
  (info.ops || []).forEach((o) => {
    if (o.type === 'text' && o.spanId) state.opMap.set(o.spanId, o);
    if (o.type === 'freeText' && o.id) state.opMap.set(o.id, o);
  });
  state.images = (info.images || []).filter((i) => i.rect[2] - i.rect[0] > 1 && i.rect[3] - i.rect[1] > 1);

  const page = $('page');
  page.style.width = ptToPx(info.width) + 'px';
  page.style.height = ptToPx(info.height) + 'px';

  const img = $('pageImg');
  await swapPageImage(img, `/api/page?doc=${state.docId}&n=${n}&dpi=${state.dpi}&t=${Date.now()}`);

  buildSpanLayer();
  buildObjectLayer();
  buildBgSampler(img);
  refreshThumb(n);
  animatePage();
  if (mosaic.on) sizeMosaicCanvas();
}

/** 编辑后同步当前页缩略图，避免左侧还是旧画面 */
function refreshThumb(n) {
  const im = document.querySelector(`.thumb[data-n="${n}"] img`);
  if (im) im.src = `/api/page?doc=${state.docId}&n=${n}&dpi=36&t=${Date.now()}`;
}

/** 先在后台把新页面图画好，再替换上去：旧画面保持到新画面就绪，不会闪白 */
function swapPageImage(img, url) {
  return new Promise((resolve) => {
    const pre = new Image();
    const done = () => {
      img.src = url;
      // 解码完成后再继续，避免尺寸抖动
      if (img.decode) img.decode().then(resolve).catch(resolve);
      else resolve();
    };
    pre.onload = done;
    pre.onerror = done;
    pre.src = url;
  });
}

function guessLabel(font) {
  const s = String(font || '').toLowerCase();
  if (/yahei|msyh/.test(s)) return '微软雅黑';
  if (/simsun|song/.test(s)) return '宋体';
  if (/simhei|heiti|hei/.test(s)) return '黑体';
  if (/kai/.test(s)) return '楷体';
  if (/fang/.test(s)) return '仿宋';
  if (/times/.test(s)) return 'Times New Roman';
  if (/arial|helv/.test(s)) return 'Arial';
  return '微软雅黑';
}

function findItem(id) {
  return state.items.find((it) => it.id === id) || null;
}

function itemText(it) {
  const op = state.opMap.get(it.id);
  if (!op) return it.text;
  return op.text !== undefined ? op.text : it.text;
}

function buildSpanLayer() {
  const layer = $('spanLayer');
  layer.innerHTML = '';

  state.items.forEach((it) => {
    const op = state.opMap.get(it.id);
    const div = document.createElement('div');
    div.className = 'span';
    div.dataset.id = it.id;

    if (op) div.classList.add('edited');
    if (state.selected === it.id) div.classList.add('sel');
    if (state.hits.some((h) => h.spanId === it.id)) div.classList.add('hit');

    const off = (op && op.offset) || [0, 0];
    const text = op && op.text !== undefined ? op.text : it.text;
    const size = (op && op.size) || it.size;

    let x0 = it.bbox[0] + off[0], y0 = it.bbox[1] + off[1];
    let x1 = it.bbox[2] + off[0], y1 = it.bbox[3] + off[1];

    if (it.kind === 'free') {
      x1 = Math.max(x1, x0 + 8);
      y1 = Math.max(y1, y0 + size * 1.2);
    } else if (op && op.text !== undefined) {
      x1 = Math.max(x1, x0 + estimateWidth(op.text, size, it.fontLabel));
    }

    const pad = 1.4;
    div.style.left = ptToPx(x0 - pad) + 'px';
    div.style.top = ptToPx(y0 - pad) + 'px';
    div.style.width = ptToPx(x1 - x0 + pad * 2) + 'px';
    div.style.height = ptToPx(y1 - y0 + pad * 2) + 'px';
    div.title = String(text || '').slice(0, 80);

    attachSpanEvents(div, it);
    layer.appendChild(div);
  });
}

function estimateWidth(text, size, fontLabel) {
  const cv = estimateWidth._cv || (estimateWidth._cv = document.createElement('canvas'));
  const ctx = cv.getContext('2d');
  ctx.font = `${size}px ${fontCss(fontLabel)}`;
  const lines = String(text || '').split('\n');
  return Math.max(...lines.map((l) => ctx.measureText(l).width), 0);
}

let bgSampler = null;
function buildBgSampler(img) {
  try {
    const cv = document.createElement('canvas');
    cv.width = img.naturalWidth;
    cv.height = img.naturalHeight;
    const ctx = cv.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(img, 0, 0);
    bgSampler = { cv, ctx };
  } catch (e) {
    bgSampler = null;
  }
}

function sampleBg(bbox) {
  if (!bgSampler) return '#ffffff';
  const s = pxPerPt();
  const x0 = Math.max(0, Math.floor(bbox[0] * s));
  const y0 = Math.max(0, Math.floor(bbox[1] * s));
  const x1 = Math.min(bgSampler.cv.width, Math.ceil(bbox[2] * s));
  const y1 = Math.min(bgSampler.cv.height, Math.ceil(bbox[3] * s));
  if (x1 <= x0 || y1 <= y0) return '#ffffff';

  try {
    const pad = Math.max(2, Math.round(1.5 * s));
    const counts = new Map();
    const put = (px, py) => {
      if (px < 0 || py < 0 || px >= bgSampler.cv.width || py >= bgSampler.cv.height) return;
      const d = bgSampler.ctx.getImageData(px, py, 1, 1).data;
      // 保留完整色值，量化会让遮挡色和页面差出一档，露出方块边界
      const k = (d[0] << 16) | (d[1] << 8) | d[2];
      counts.set(k, (counts.get(k) || 0) + 1);
    };
    for (let x = x0; x < x1; x += 2) { put(x, y0 - pad); put(x, y1 + pad); }
    for (let y = y0; y < y1; y += 2) { put(x0 - pad, y); put(x1 + pad, y); }
    if (!counts.size) return '#ffffff';
    let best = 0, bestN = -1;
    for (const [k, v] of counts) if (v > bestN) { bestN = v; best = k; }
    return `rgb(${(best >> 16) & 255},${(best >> 8) & 255},${best & 255})`;
  } catch (e) {
    return '#ffffff';
  }
}

/* ------------------------------------------------------------ 图片对象层 */

function buildObjectLayer() {
  const layer = $('objectLayer');
  layer.innerHTML = '';

  state.images.forEach((img) => {
    const box = document.createElement('div');
    box.className = 'imgbox' + (state.selectedImage === img.id ? ' sel' : '');
    box.dataset.id = img.id;
    const [x0, y0, x1, y1] = img.rect;
    box.style.left = ptToPx(x0) + 'px';
    box.style.top = ptToPx(y0) + 'px';
    box.style.width = ptToPx(x1 - x0) + 'px';
    box.style.height = ptToPx(y1 - y0) + 'px';

    const handle = document.createElement('div');
    handle.className = 'handle';
    box.appendChild(handle);

    box.addEventListener('mousedown', (e) => {
      if (e.target === handle) return;
      e.preventDefault();
      e.stopPropagation();
      startImageDrag(e, img, box, 'move');
    });
    handle.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      startImageDrag(e, img, box, 'resize');
    });
    box.addEventListener('click', (e) => {
      e.stopPropagation();
      selectImage(img.id);
    });

    layer.appendChild(box);
  });
}

function startImageDrag(e, img, box, mode) {
  const startX = e.clientX, startY = e.clientY;
  const rect = img.rect.slice();
  const scale = pxPerPt();
  let moved = false;

  const onMove = (ev) => {
    const dxPx = ev.clientX - startX;
    const dyPx = ev.clientY - startY;
    if (!moved && Math.hypot(dxPx, dyPx) < 2) return;
    moved = true;

    const dx = dxPx / scale;
    const dy = dyPx / scale;

    if (mode === 'move') {
      img.rect = [rect[0] + dx, rect[1] + dy, rect[2] + dx, rect[3] + dy];
    } else {
      const ratio = (rect[3] - rect[1]) / Math.max(1, rect[2] - rect[0]);
      const w = Math.max(8, (rect[2] - rect[0]) + dx);
      img.rect = [rect[0], rect[1], rect[0] + w, rect[1] + w * ratio];
    }

    box.style.left = ptToPx(img.rect[0]) + 'px';
    box.style.top = ptToPx(img.rect[1]) + 'px';
    box.style.width = ptToPx(img.rect[2] - img.rect[0]) + 'px';
    box.style.height = ptToPx(img.rect[3] - img.rect[1]) + 'px';
  };

  const onUp = async () => {
    window.removeEventListener('mousemove', onMove);
    if (!moved) return;
    try {
      await apiPost('/api/edit', {
        doc: state.docId,
        op: { type: 'imageUpdate', id: img.id, page: state.cur, rect: img.rect },
      });
      await renderPage();
      selectImage(img.id);
      refreshHistory();
    } catch (err) {
      toast(err.message, true);
    }
  };

  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp, { once: true });
}

/* ------------------------------------------------------------ 选中状态 */

function selectSpan(id) {
  state.selected = id;
  state.selectedImage = null;
  document.querySelectorAll('.span').forEach((el) => el.classList.toggle('sel', el.dataset.id === id));
  document.querySelectorAll('.imgbox').forEach((el) => el.classList.remove('sel'));
  showInspector(id, null);
}

function selectImage(id) {
  state.selectedImage = id;
  state.selected = null;
  document.querySelectorAll('.span').forEach((el) => el.classList.remove('sel'));
  document.querySelectorAll('.imgbox').forEach((el) => el.classList.toggle('sel', el.dataset.id === id));
  showInspector(null, id);
}

function showInspector(itemId, imageId) {
  $('insEmpty').hidden = true;
  $('insText').hidden = true;
  $('insImage').hidden = true;

  if (imageId) {
    const img = state.images.find((i) => i.id === imageId);
    if (!img) { $('insEmpty').hidden = false; return; }
    $('insImage').hidden = false;
    $('imgW').value = Math.round(img.rect[2] - img.rect[0]);
    $('imgH').value = Math.round(img.rect[3] - img.rect[1]);
    $('imgMeta').textContent = `x ${img.rect[0].toFixed(1)} · y ${img.rect[1].toFixed(1)}`;
    return;
  }

  const it = itemId ? findItem(itemId) : null;
  if (!it) { $('insEmpty').hidden = false; return; }

  $('insText').hidden = false;
  animatePanel($('insText'));

  const op = state.opMap.get(it.id);
  const size = (op && op.size) || it.size;
  const color = (op && op.color) || it.color;
  const bold = op ? !!op.bold : it.bold;
  const italic = op ? !!op.italic : it.italic;
  const align = (op && op.align) || it.align || 'left';
  const lh = (op && op.lineHeight) || it.lineHeight || 1.2;
  const off = (op && op.offset) || [0, 0];

  const sel = $('selFont');
  if (!sel.options.length) {
    FONT_OPTIONS.forEach((f) => {
      const o = document.createElement('option');
      o.value = f; o.textContent = f;
      sel.appendChild(o);
    });
  }
  sel.value = FONT_OPTIONS.includes(it.fontLabel) ? it.fontLabel : '微软雅黑';

  $('textContent').value = itemText(it);
  $('inpSize').value = size;
  $('inpColor').value = rgbHex(color);
  $('inpLineH').value = lh;
  $('inpX').value = Math.round((it.bbox[0] + off[0]) * 10) / 10;
  $('inpY').value = Math.round((it.bbox[1] + off[1]) * 10) / 10;

  $('btnBold').classList.toggle('on', bold);
  $('btnItalic').classList.toggle('on', italic);
  document.querySelectorAll('#alignSeg button').forEach((b) => b.classList.toggle('on', b.dataset.align === align));
  $('editFlag').hidden = !op;
}

/* ------------------------------------------------------------ 拖动 / 编辑 */

function attachSpanEvents(div, it) {
  div.addEventListener('mousedown', (e) => beginSpanDrag(e, it, div));
  div.addEventListener('dblclick', (e) => {
    e.preventDefault();
    e.stopPropagation();
    startEdit(it);
  });
  div.addEventListener('click', (e) => {
    e.stopPropagation();
    if (state.suppressClick) return;
    selectSpan(it.id);
  });
}

function beginSpanDrag(e, it, div) {
  if (e.button !== 0 || state.editing) return;

  const startX = e.clientX, startY = e.clientY;
  const scale = pxPerPt();
  const op = state.opMap.get(it.id);
  const baseOff = (op && op.offset) || [0, 0];
  let running = false;

  const onMove = (ev) => {
    const dxPx = ev.clientX - startX;
    const dyPx = ev.clientY - startY;
    if (!running && Math.hypot(dxPx, dyPx) < 4) return;
    if (!running) {
      running = true;
      state.suppressClick = true;
      enterDragVisual(div, it);
      document.body.style.cursor = 'grabbing';
    }
    div.style.transform = `translate(${dxPx}px, ${dyPx}px)`;
  };

  const onUp = async (ev) => {
    window.removeEventListener('mousemove', onMove);
    document.body.style.cursor = '';
    if (!running) {
      state.suppressClick = false;
      selectSpan(it.id);
      return;
    }

    const dxPt = (ev.clientX - startX) / scale;
    const dyPt = (ev.clientY - startY) / scale;
    div.style.transform = '';
    exitDragVisual(div, it);

    setTimeout(() => { state.suppressClick = false; }, 60);

    if (Math.abs(dxPt) < 0.3 && Math.abs(dyPt) < 0.3) return;

    // 先把元素就地落到新位置（视觉即时），再后台提交
    div.style.left = (parseFloat(div.style.left) + dxPt * scale) + 'px';
    div.style.top = (parseFloat(div.style.top) + dyPt * scale) + 'px';

    applyItemOp(it, { offset: [baseOff[0] + dxPt, baseOff[1] + dyPt] });
  };

  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp, { once: true });
}

function enterDragVisual(div, it) {
  const op = state.opMap.get(it.id);
  const text = op && op.text !== undefined ? op.text : it.text;
  const size = (op && op.size) || it.size;
  const color = (op && op.color) || it.color;
  div.classList.add('dragging');
  div.textContent = text;
  div.style.background = sampleBg(it.bbox);
  div.style.font = `${ptToPx(size)}px/1.2 ${fontCss(it.fontLabel)}`;
  div.style.color = rgbCss(color);
  div.style.width = 'auto';
  div.style.height = 'auto';
  div.style.padding = '1px 3px';
}

function exitDragVisual(div) {
  div.classList.remove('dragging');
  div.textContent = '';
  div.style.padding = '';
  div.style.width = '';
  div.style.height = '';
  div.style.font = '';
  div.style.color = '';
  div.style.background = '';
}

function startEdit(it) {
  cancelEdit();
  selectSpan(it.id);

  const op = state.opMap.get(it.id);
  const text = itemText(it);
  const size = (op && op.size) || it.size;
  const color = (op && op.color) || it.color;
  const bold = op ? !!op.bold : it.bold;
  const italic = op ? !!op.italic : it.italic;
  const align = (op && op.align) || it.align || 'left';

  const ta = document.createElement('textarea');
  ta.className = 'editbox';
  ta.value = text;

  const off = (op && op.offset) || [0, 0];
  // 无边框后，内容起点不再被 border 推移，补偿量从 2 降到 0.5，文字才不会跳
  const left = ptToPx(it.bbox[0] + off[0]) - 0.5;
  const top = ptToPx(it.bbox[1] + off[1]) - 0.5;
  const baseW = Math.max(ptToPx(it.bbox[2] - it.bbox[0]) + 3, 60);
  const baseH = Math.max(ptToPx(it.bbox[3] - it.bbox[1]) + 1, 14);

  Object.assign(ta.style, {
    left: Math.max(0, left) + 'px',
    top: Math.max(0, top) + 'px',
    width: baseW + 'px',
    minHeight: baseH + 'px',
    font: `${ptToPx(size)}px/1.2 ${fontCss(it.fontLabel)}`,
    color: rgbCss(color),
    background: sampleBg(it.bbox),
    textAlign: align,
    fontWeight: bold ? '700' : '400',
    fontStyle: italic ? 'italic' : 'normal',
  });

  $('page').appendChild(ta);
  state.editing = { item: it, ta, original: text, baseW, baseH };

  ta.addEventListener('input', autoGrow);
  ta.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitEdit(); }
  });

  requestAnimationFrame(() => {
    autoGrow();
    ta.focus();
    if (ta.dataset.caret) {
      const p = Math.min(Number(ta.dataset.caret) || 0, ta.value.length);
      ta.setSelectionRange(p, p);
    } else {
      ta.select();
    }
  });
}

let _measCtx = null;
function measureTextW(text, el) {
  if (!_measCtx) _measCtx = document.createElement('canvas').getContext('2d');
  const cs = getComputedStyle(el);
  _measCtx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  let w = 0;
  String(text || '').split('\n').forEach((line) => {
    w = Math.max(w, _measCtx.measureText(line).width);
  });
  return w;
}

function autoGrow() {
  if (!state.editing) return;
  const { ta, baseW, baseH } = state.editing;
  // 用实测文本宽度定宽，否则 textarea 先折行、scrollWidth 永远不比容器宽，字会被竖排挤成两行
  const pageW = ptToPx((state.pages[state.cur] || [595])[0]);
  const left = parseFloat(ta.style.left) || 0;
  const cap = Math.max(baseW, pageW - left - 6);
  const want = Math.max(baseW, measureTextW(ta.value, ta) + 6);
  ta.style.width = Math.min(want, cap) + 'px';
  ta.style.height = 'auto';
  ta.style.height = Math.max(ta.scrollHeight, baseH) + 'px';
}

function cancelEdit() {
  if (!state.editing) return;
  state.editing.ta.remove();
  state.editing = null;
}

async function commitEdit() {
  const ed = state.editing;
  if (!ed) return;
  const text = ed.ta.value;
  const it = ed.item;
  if (text === ed.original) { cancelEdit(); return; }
  state.editing = null;
  ed.ta.remove();
  await applyItemOp(it, { text });
}

/* ------------------------------------------------------------ 操作构建 */

function buildOp(it, patch) {
  const op = state.opMap.get(it.id) || {};
  const text = patch.text !== undefined ? patch.text : (op.text !== undefined ? op.text : it.text);
  const size = patch.size !== undefined ? patch.size : (op.size || it.size);
  const color = patch.color !== undefined ? patch.color : (op.color || it.color);
  const bold = patch.bold !== undefined ? patch.bold : (op.bold !== undefined ? op.bold : it.bold);
  const italic = patch.italic !== undefined ? patch.italic : (op.italic !== undefined ? op.italic : it.italic);
  const align = patch.align !== undefined ? patch.align : (op.align || it.align || 'left');
  const lh = patch.lineHeight !== undefined ? patch.lineHeight : (op.lineHeight || it.lineHeight || 1.2);
  const off = patch.offset !== undefined ? patch.offset : (op.offset || [0, 0]);

  if (it.kind === 'free') {
    const rect = patch.rect || it.bbox;
    return {
      type: 'freeText', id: it.id, page: state.cur,
      rect: [rect[0] + off[0], rect[1] + off[1], rect[2] + off[0], rect[3] + off[1]],
      text, size, color, bold, italic, align,
      font: patch.font !== undefined ? patch.font : it.font,
      lineHeight: lh,
    };
  }

  return {
    type: 'text', page: state.cur, spanId: it.id,
    text, size, color, bold, italic, align,
    font: patch.font !== undefined ? patch.font : it.font,
    lineHeight: lh,
    offset: off,
  };
}

async function applyItemOp(it, patch, opts = {}) {
  // 编辑是高频操作，不显示全屏遮罩——原地更新即可，否则每次都像卡住
  try {
    await apiPost('/api/edit', { doc: state.docId, op: buildOp(it, patch) });
    await renderPage();
    if (opts.keepSelection !== false) selectSpan(it.id);
    refreshHistory();
  } catch (e) {
    toast(e.message, true);
  }
}

/* ------------------------------------------------------------ 属性面板 */

const pushText = debounce(async () => {
  const it = findItem(state.selected);
  if (!it) return;
  await applyItemOp(it, { text: $('textContent').value });
}, 650);

$('textContent').addEventListener('input', pushText);
$('textContent').addEventListener('blur', () => {
  const it = findItem(state.selected);
  if (!it) return;
  const cur = state.opMap.get(it.id);
  const now = (cur && cur.text !== undefined) ? cur.text : it.text;
  if ($('textContent').value !== now) applyItemOp(it, { text: $('textContent').value });
});

$('selFont').addEventListener('change', (e) => {
  const it = findItem(state.selected);
  if (it) applyItemOp(it, { fontLabel: e.target.value, font: e.target.value });
});
$('inpSize').addEventListener('change', (e) => {
  const v = parseFloat(e.target.value);
  const it = findItem(state.selected);
  if (it && v >= 4 && v <= 300) applyItemOp(it, { size: v });
});
$('inpColor').addEventListener('input', (e) => {
  const it = findItem(state.selected);
  if (it) applyItemOp(it, { color: hexToRgb(e.target.value) });
});
$('inpLineH').addEventListener('change', (e) => {
  const v = parseFloat(e.target.value);
  const it = findItem(state.selected);
  if (it && v >= 0.6 && v <= 4) applyItemOp(it, { lineHeight: v });
});
$('inpX').addEventListener('change', (e) => {
  const it = findItem(state.selected);
  if (!it) return;
  const op = state.opMap.get(it.id);
  const off = (op && op.offset) || [0, 0];
  const target = parseFloat(e.target.value);
  if (Number.isFinite(target)) applyItemOp(it, { offset: [target - it.bbox[0], off[1]] });
});
$('inpY').addEventListener('change', (e) => {
  const it = findItem(state.selected);
  if (!it) return;
  const op = state.opMap.get(it.id);
  const off = (op && op.offset) || [0, 0];
  const target = parseFloat(e.target.value);
  if (Number.isFinite(target)) applyItemOp(it, { offset: [off[0], target - it.bbox[1]] });
});
$('btnBold').addEventListener('click', () => {
  const it = findItem(state.selected);
  if (!it) return;
  const op = state.opMap.get(it.id);
  applyItemOp(it, { bold: !(op ? op.bold : it.bold) });
});
$('btnItalic').addEventListener('click', () => {
  const it = findItem(state.selected);
  if (!it) return;
  const op = state.opMap.get(it.id);
  applyItemOp(it, { italic: !(op ? op.italic : it.italic) });
});
document.querySelectorAll('#alignSeg button').forEach((b) => {
  b.addEventListener('click', () => {
    const it = findItem(state.selected);
    if (it) applyItemOp(it, { align: b.dataset.align });
  });
});

$('btnRedact').addEventListener('click', async () => {
  const it = findItem(state.selected);
  if (!it) return;
  // 取样当前底色，盖上去才不会在白纸之外的地方露馅
  const bg = sampleBg(it.bbox);
  const m = bg.match(/(\d+)\D+(\d+)\D+(\d+)/);
  const fill = m ? [Number(m[1]) / 255, Number(m[2]) / 255, Number(m[3]) / 255] : [1, 1, 1];
  setBusy(true);
  try {
    const op = it.kind === 'free'
      ? { type: 'remove', id: it.id }
      : { type: 'redact', page: state.cur, spanId: it.id, fill };
    await apiPost('/api/edit', { doc: state.docId, op });
    state.selected = null;
    await renderPage();
    showInspector(null);
    refreshHistory();
    toast('已遮盖');
  } catch (e) {
    toast('遮盖失败：' + e.message, true);
  } finally {
    setBusy(false);
  }
});

$('btnDelete').addEventListener('click', async () => {
  const it = findItem(state.selected);
  if (!it) return;
  if (!confirm('删除这段文字？')) return;
  setBusy(true);
  try {
    const op = it.kind === 'free'
      ? { type: 'remove', id: it.id }
      : { type: 'delete', page: state.cur, spanId: it.id };
    await apiPost('/api/edit', { doc: state.docId, op });
    state.selected = null;
    await renderPage();
    showInspector(null);
    refreshHistory();
    toast('已删除');
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
});

$('btnDelImage').addEventListener('click', async () => {
  const img = state.images.find((i) => i.id === state.selectedImage);
  if (!img) return;
  if (!confirm('删除这张图片？')) return;
  try {
    setBusy(true);
    await apiPost('/api/edit', { doc: state.docId, op: { type: 'remove', id: img.id } });
    state.selectedImage = null;
    await renderPage();
    showInspector(null);
    refreshHistory();
    toast('已删除图片');
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
});

async function resizeImage(targetW, targetH) {
  const img = state.images.find((i) => i.id === state.selectedImage);
  if (!img) return;
  const rect = [img.rect[0], img.rect[1], img.rect[0] + targetW, img.rect[1] + targetH];
  await apiPost('/api/edit', { doc: state.docId, op: { type: 'imageUpdate', id: img.id, page: state.cur, rect } });
  await renderPage();
  selectImage(img.id);
}

$('imgW').addEventListener('change', (e) => {
  const img = state.images.find((i) => i.id === state.selectedImage);
  if (!img) return;
  const w = Math.max(8, parseFloat(e.target.value) || 8);
  const ratio = (img.rect[3] - img.rect[1]) / Math.max(1, img.rect[2] - img.rect[0]);
  resizeImage(w, w * ratio);
});
$('imgH').addEventListener('change', (e) => {
  const img = state.images.find((i) => i.id === state.selectedImage);
  if (!img) return;
  const h = Math.max(8, parseFloat(e.target.value) || 8);
  resizeImage(img.rect[2] - img.rect[0], h);
});

/* ------------------------------------------------------------ 视图 */

async function setZoom(z) {
  state.zoom = Math.max(0.2, Math.min(6, z));
  $('zoomLabel').textContent = Math.round(state.zoom * 100) + '%';
  if (state.docId) await renderPage();
}

async function fitToWidth() {
  if (!state.docId) return;
  const avail = $('viewport').clientWidth - 64;
  const pw = state.pages[state.cur][0];
  const dpi = pickDpi();
  state.zoom = Math.max(0.2, Math.min(3, avail / (pw * (dpi / 72))));
  $('zoomLabel').textContent = Math.round(state.zoom * 100) + '%';
  await renderPage();
}

$('zoomIn').addEventListener('click', () => setZoom(state.zoom * 1.15));
$('zoomOut').addEventListener('click', () => setZoom(state.zoom / 1.15));
$('zoomFit').addEventListener('click', () => fitToWidth());

$('viewport').addEventListener('wheel', (e) => {
  if (!e.ctrlKey || !state.docId) return;
  e.preventDefault();
  setZoom(state.zoom * (e.deltaY < 0 ? 1.1 : 0.9));
}, { passive: false });

$('page').addEventListener('click', (e) => {
  if (e.target === $('page') || e.target === $('pageImg') || e.target === $('spanLayer')) {
    if (state.editing) commitEdit();
    selectSpan(null);
  }
});

/* ------------------------------------------------------------ 新增内容 */

$('btnAddText').addEventListener('click', async () => {
  const w = state.pages[state.cur][0];
  const h = state.pages[state.cur][1];
  const size = 16;
  const rect = [w * 0.2, h * 0.42, w * 0.8, h * 0.42 + size * 1.6];
  try {
    const r = await apiPost('/api/edit', {
      doc: state.docId,
      op: { type: 'freeText', page: state.cur, rect, text: '新建文字，双击修改', size, color: [30, 30, 30], font: '微软雅黑' },
    });
    await renderPage();
    refreshHistory();
    const it = findItem((r.op && r.op.id) || r.id);
    if (it) { selectSpan(it.id); startEdit(it); }
    else toast('已新增文字，请点击选中后修改', true);
  } catch (e) {
    toast(e.message, true);
  }
});

$('fileImage').addEventListener('change', async (e) => {
  const f = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!f || !state.docId) return;
  const data = await new Promise((resolve) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.readAsDataURL(f);
  });
  const pw = state.pages[state.cur][0];
  const ph = state.pages[state.cur][1];
  const w = Math.min(pw * 0.5, 260);
  const rect = [(pw - w) / 2, ph * 0.3, (pw + w) / 2, ph * 0.3 + w * 0.72];
  try {
    setBusy(true);
    await apiPost('/api/edit', {
      doc: state.docId,
      op: { type: 'image', page: state.cur, rect, data, keepRatio: true },
    });
    await renderPage();
    refreshHistory();
    toast('图片已插入，可拖动或缩放');
  } catch (err) {
    toast(err.message, true);
  } finally {
    setBusy(false);
  }
});

/* ------------------------------------------------------------ 查找替换 */

function toggleFind(show) {
  const bar = $('findbar');
  const on = show === undefined ? bar.hidden : show;
  bar.hidden = !on;
  if (on) { $('findInput').focus(); $('findInput').select(); }
  else clearHits();
}

function clearHits() {
  state.hits = [];
  state.hitIndex = -1;
  $('findCount').textContent = '0';
  document.querySelectorAll('.span').forEach((el) => el.classList.remove('hit', 'hit-current'));
}

const runFind = debounce(async () => {
  const q = $('findInput').value.trim();
  if (!state.docId || !q) { clearHits(); return; }
  try {
    const r = await apiGet(`/api/find?doc=${state.docId}&q=${encodeURIComponent(q)}`);
    state.hits = r.hits || [];
    state.hitIndex = state.hits.length ? 0 : -1;
    $('findCount').textContent = state.hits.length ? `1/${state.hits.length}` : '0';
    if (state.hits.length) await renderPageAndHighlight(0);
    else await renderPage();
  } catch (e) {
    toast(e.message, true);
  }
}, 320);

async function renderPageAndHighlight(idx) {
  const h = state.hits[idx];
  if (!h) return;
  if (h.page !== state.cur) {
    state.cur = h.page;
    markThumbActive();
  }
  await renderPage();
  document.querySelectorAll('.span').forEach((el) => {
    el.classList.toggle('hit-current', el.dataset.id === h.spanId);
  });
  selectSpan(h.spanId);
  const el = document.querySelector(`.span[data-id="${h.spanId}"]`);
  if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
}

$('findInput').addEventListener('input', runFind);
$('findInput').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); stepHit(e.shiftKey ? -1 : 1); }
  if (e.key === 'Escape') toggleFind(false);
});

async function stepHit(dir) {
  if (!state.hits.length) return;
  state.hitIndex = (state.hitIndex + dir + state.hits.length) % state.hits.length;
  $('findCount').textContent = `${state.hitIndex + 1}/${state.hits.length}`;
  await renderPageAndHighlight(state.hitIndex);
}

$('findNext').addEventListener('click', () => stepHit(1));
$('findPrev').addEventListener('click', () => stepHit(-1));
$('findClose').addEventListener('click', () => toggleFind(false));
$('btnFind').addEventListener('click', () => toggleFind());

$('replOne').addEventListener('click', async () => {
  if (state.hitIndex < 0 || !state.hits.length) { toast('没有选中的匹配项', true); return; }
  const h = state.hits[state.hitIndex];
  const it = findItem(h.spanId);
  if (!it) return;
  const from = $('findInput').value;
  const cur = itemText(it);
  if (!cur.includes(from)) { toast('该处已不包含查找内容', true); return; }
  await applyItemOp(it, { text: cur.replace(from, $('replInput').value) });
  await runFind();
});

$('replAll').addEventListener('click', async () => {
  const from = $('findInput').value.trim();
  if (!from) return;
  if (!confirm(`把全文中的「${from}」全部替换？`)) return;
  try {
    setBusy(true);
    const r = await apiPost('/api/replace', { doc: state.docId, find: from, replace: $('replInput').value });
    await renderPage();
    refreshHistory();
    clearHits();
    toast(`已替换 ${r.count} 处`);
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
});

/* ------------------------------------------------------------ 页面操作 */

$('btnPage').addEventListener('click', (e) => {
  e.stopPropagation();
  $('pageMenu').classList.toggle('open');
});
document.addEventListener('click', () => $('pageMenu').classList.remove('open'));

$('pageMenu').addEventListener('click', async (e) => {
  const btn = e.target.closest('button');
  if (!btn) return;
  $('pageMenu').classList.remove('open');
  const act = btn.dataset.act;

  if (act === 'delete') {
    if (state.pageCount <= 1) { toast('至少保留一页', true); return; }
    if (!confirm(`删除第 ${state.cur + 1} 页？删除后此操作不可撤销。`)) return;
  }

  try {
    setBusy(true);
    const body = {
      doc: state.docId,
      action: act,
      page: state.cur,
      width: state.pages[state.cur] ? state.pages[state.cur][0] : undefined,
      height: state.pages[state.cur] ? state.pages[state.cur][1] : undefined,
    };
    await apiPost('/api/pageop', body);
    await reloadDoc(act !== 'delete', true);
    toast(act === 'rotate' ? '已旋转 90°' : act === 'insert' ? '已添加空白页' : '已删除页面');
  } catch (err) {
    toast(err.message, true);
  } finally {
    setBusy(false);
  }
});

/* ------------------------------------------------------------ 历史 */

async function refreshHistory() {
  if (!state.docId) return;
  try {
    const d = await apiGet(`/api/info?doc=${state.docId}`);
    $('btnUndo').disabled = !d.canUndo;
    $('btnRedo').disabled = !d.canRedo;
    $('docStat').textContent = `${d.name} · ${d.pageCount} 页` + (d.opCount ? ` · ${d.opCount} 处修改` : '');
    if (d.pageCount !== state.pageCount) {
      state.pageCount = d.pageCount;
      $('pageCount').textContent = d.pageCount;
    }
  } catch (e) { /* ignore */ }
}

$('btnUndo').addEventListener('click', async () => {
  try {
    setBusy(true);
    const r = await apiPost('/api/undo', { doc: state.docId });
    if (!r.ok) { toast('没有可撤销的操作'); return; }
    await reloadDoc();
    toast('已撤销');
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
});

$('btnRedo').addEventListener('click', async () => {
  try {
    setBusy(true);
    const r = await apiPost('/api/redo', { doc: state.docId });
    if (!r.ok) { toast('没有可重做的操作'); return; }
    await reloadDoc();
    toast('已重做');
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
});

/* ------------------------------------------------------------ 导出 */

async function doExport() {
  if (!state.docId) return;
  try {
    setBusy(true);
    const r = await apiPost('/api/export', { doc: state.docId });
    toast('已导出：' + r.path);
    if (r.path) apiPost('/api/reveal', { path: r.path }).catch(() => {});
  } catch (e) {
    if (e.cancelled) return;
    if (/dialog|tk|保存/i.test(e.message)) await downloadFallback();
    else toast(e.message, true);
  } finally {
    setBusy(false);
  }
}

async function downloadFallback() {
  const r = await fetch('/api/download', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ doc: state.docId }),
  });
  if (!r.ok) { toast('导出失败', true); return; }
  const blob = await r.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = state.name.replace(/\.pdf$/i, '') + '_已编辑.pdf';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  toast('已下载导出文件');
}

/** Ctrl+Shift+S：不弹保存框，直接覆盖到原文件同目录 */
async function quickExport() {
  if (!state.docId) return;
  try {
    setBusy(true);
    const r = await apiPost('/api/quickexport', { doc: state.docId });
    toast('已保存：' + r.path);
    apiPost('/api/reveal', { path: r.path }).catch(() => {});
  } catch (e) {
    toast(e.message, true);
  } finally {
    setBusy(false);
  }
}

/* ------------------------------------------------------------ 事件绑定 */

$('btnOpen').addEventListener('click', openViaDialog);
$('emptyOpen').addEventListener('click', openViaDialog);
$('filePdf').addEventListener('change', async (e) => {
  const f = e.target.files && e.target.files[0];
  e.target.value = '';
  if (f) await openViaFile(f);
});
$('btnExport').addEventListener('click', doExport);

document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  const k = e.key.toLowerCase();

  if (mod && k === 'o') { e.preventDefault(); openViaDialog(); return; }
  if (mod && k === 's') {
    e.preventDefault();
    if (!state.docId) return;
    if (e.shiftKey) quickExport();
    else doExport();
    return;
  }
  if (mod && k === 'f') { e.preventDefault(); if (state.docId) toggleFind(true); return; }
  if (mod && k === 'y') { e.preventDefault(); if (state.docId && !state.editing) $('btnRedo').click(); return; }
  if (mod && k === 'z') {
    e.preventDefault();
    if (!state.docId || state.editing) return;
    if (e.shiftKey) $('btnRedo').click();
    else $('btnUndo').click();
    return;
  }
  if (e.key === 'Escape') {
    if (mosaic.on) { setMosaicMode(false); return; }
    if (state.editing) { e.preventDefault(); cancelEdit(); return; }
    if (!$('findbar').hidden) { toggleFind(false); return; }
  }

  if (state.editing) return;
  const tag = (e.target.tagName || '').toLowerCase();
  if (tag === 'input' || tag === 'textarea' || tag === 'select') return;

  if ((e.key === 'Delete' || e.key === 'Backspace') && (state.selected || state.selectedImage)) {
    e.preventDefault();
    if (state.selectedImage) $('btnDelImage').click();
    else $('btnDelete').click();
    return;
  }

  // 选中文字后直接敲键就改字，和 Word / 微信一样，不用先双击
  if (state.selected && !mod && !e.altKey && e.key.length === 1) {
    const it = findItem(state.selected);
    if (!it) return;
    e.preventDefault();
    startEdit(it);
    const ta = state.editing && state.editing.ta;
    if (ta) {
      ta.value = e.key;
      ta.dataset.caret = String(e.key.length);
      autoGrow();
    }
    return;
  }

  if (state.selected && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    const it = findItem(state.selected);
    if (!it) return;
    const op = state.opMap.get(it.id);
    const off = (op && op.offset) || [0, 0];
    const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
    const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
    applyItemOp(it, { offset: [off[0] + dx, off[1] + dy] });
    return;
  }

  if (!state.selected && (e.key === 'PageDown' || e.key === 'PageUp')) {
    e.preventDefault();
    gotoPage(state.cur + (e.key === 'PageDown' ? 1 : -1));
  }
});

let dragDepth = 0;
window.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; $('dropzone').classList.add('show'); });
window.addEventListener('dragover', (e) => { e.preventDefault(); });
window.addEventListener('dragleave', () => { dragDepth--; if (dragDepth <= 0) { dragDepth = 0; $('dropzone').classList.remove('show'); } });
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  $('dropzone').classList.remove('show');
  const f = e.dataTransfer.files && e.dataTransfer.files[0];
  if (f) openViaFile(f);
});

let resizeTimer = null;
window.addEventListener('resize', () => {
  if (!state.docId) return;
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => fitToWidth(), 240);
});

/* ------------------------------------------------------------ 马赛克涂抹 */
const mosaic = {
  on: false,
  drawing: false,
  pts: [],         // 屏幕坐标轨迹
  sizePx: 46,      // 笔刷直径（屏幕像素）
  blockPt: 5.0,    // 马赛克块大小（PDF 点）
  tmp: document.createElement('canvas'),
  mask: document.createElement('canvas'),
};

/** 屏幕像素 → PDF 点 的比例，用页面图像的实际显示宽度换算 */
function mosaicScale() {
  const img = $('pageImg');
  const pw = (state.pages[state.cur] || [595])[0];
  return img.clientWidth ? img.clientWidth / pw : pxPerPt();
}

function sizeMosaicCanvas() {
  const cv = $('mosaicCanvas');
  const img = $('pageImg');
  const w = img.clientWidth;
  const h = img.clientHeight;
  if (!w || !h) return;
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  cv.width = Math.round(w * dpr);
  cv.height = Math.round(h * dpr);
  cv.style.width = w + 'px';
  cv.style.height = h + 'px';
  const ctx = cv.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.imageSmoothingEnabled = false;
}

/**
 * 整块重绘：把整条轨迹的包围盒一次性做成马赛克，再用轨迹圆做遮罩。
 * 逐点各自采样会导致色块互相覆盖、看起来斑驳。
 */
function redrawMosaic() {
  const cv = $('mosaicCanvas');
  const img = $('pageImg');
  const pts = mosaic.pts;
  const W = img.clientWidth;
  const H = img.clientHeight;
  if (!W || !H) return;

  const ctx = cv.getContext('2d');
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, cv.width, cv.height);
  ctx.restore();
  if (!pts.length || !img.naturalWidth) return;

  const r = mosaic.sizePx / 2;
  let x0 = W, y0 = H, x1 = 0, y1 = 0;
  for (const p of pts) {
    if (p.x < x0) x0 = p.x;
    if (p.x > x1) x1 = p.x;
    if (p.y < y0) y0 = p.y;
    if (p.y > y1) y1 = p.y;
  }
  x0 = Math.max(0, x0 - r);
  y0 = Math.max(0, y0 - r);
  x1 = Math.min(W, x1 + r);
  y1 = Math.min(H, y1 + r);
  const w = x1 - x0;
  const h = y1 - y0;
  if (w < 2 || h < 2) return;

  const k = img.naturalWidth / W;
  // 与最终效果保持同一块尺寸，做到所见即所得
  const blockPx = Math.max(2, mosaic.blockPt * mosaicScale());
  const bw = Math.max(2, Math.round(w / blockPx));
  const bh = Math.max(2, Math.round(h / blockPx));
  const t = mosaic.tmp;
  t.width = bw;
  t.height = bh;
  const tc = t.getContext('2d');
  tc.imageSmoothingEnabled = true;
  tc.drawImage(img, x0 * k, y0 * k, w * k, h * k, 0, 0, bw, bh);

  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(t, 0, 0, bw, bh, x0, y0, w, h);

  const m = mosaic.mask;
  if (m.width !== Math.round(W) || m.height !== Math.round(H)) {
    m.width = Math.round(W);
    m.height = Math.round(H);
  }
  const mc = m.getContext('2d');
  mc.clearRect(0, 0, m.width, m.height);
  mc.fillStyle = '#fff';
  for (const p of pts) {
    mc.beginPath();
    mc.arc(p.x, p.y, r, 0, Math.PI * 2);
    mc.fill();
  }
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(m, 0, 0, m.width, m.height, 0, 0, W, H);
  ctx.globalCompositeOperation = 'source-over';
}

function addMosaicPoint(clientX, clientY) {
  const cv = $('mosaicCanvas');
  const box = cv.getBoundingClientRect();
  const x = clientX - box.left;
  const y = clientY - box.top;
  if (x < 0 || y < 0 || x > box.width || y > box.height) return;
  const last = mosaic.pts[mosaic.pts.length - 1];
  if (last && Math.hypot(x - last.x, y - last.y) < 2) return;
  mosaic.pts.push({ x, y });
  redrawMosaic();
}

async function commitMosaic() {
  const pts = mosaic.pts;
  mosaic.pts = [];
  if (!pts.length) return;
  const s = mosaicScale();
  const r = mosaic.sizePx / 2 / s;
  const strokes = pts.map((p) => ({ x: p.x / s, y: p.y / s, r }));
  setBusy(true);
  try {
    await apiPost('/api/edit', {
      doc: state.docId,
      op: { type: 'mosaic', page: state.cur, strokes, block: mosaic.blockPt },
    });
    await renderPage();
    refreshHistory();
  } catch (e) {
    toast('马赛克失败：' + e.message, true);
  } finally {
    setBusy(false);
    if (mosaic.on) sizeMosaicCanvas();
  }
}

function setMosaicMode(on) {
  mosaic.on = !!on;
  $('btnMosaic').classList.toggle('active', mosaic.on);
  $('page').classList.toggle('mosaic-mode', mosaic.on);
  $('mosaicCanvas').hidden = !mosaic.on;
  if (mosaic.on) {
    if (state.editing) cancelEdit();
    selectSpan(null);
    sizeMosaicCanvas();
    toast('按住鼠标涂抹打马赛克，松手生效；滚轮调笔刷大小');
  } else {
    mosaic.pts = [];
  }
}

$('btnMosaic').addEventListener('click', () => setMosaicMode(!mosaic.on));

$('mosaicCanvas').addEventListener('mousedown', (e) => {
  if (!mosaic.on || e.button !== 0) return;
  e.preventDefault();
  mosaic.drawing = true;
  mosaic.pts = [];
  addMosaicPoint(e.clientX, e.clientY);
});

window.addEventListener('mousemove', (e) => {
  if (!mosaic.drawing) return;
  addMosaicPoint(e.clientX, e.clientY);
});

window.addEventListener('mouseup', () => {
  if (!mosaic.drawing) return;
  mosaic.drawing = false;
  commitMosaic();
});

$('mosaicCanvas').addEventListener('wheel', (e) => {
  if (!mosaic.on) return;
  e.preventDefault();
  mosaic.sizePx = Math.max(14, Math.min(220, mosaic.sizePx + (e.deltaY < 0 ? 8 : -8)));
  redrawMosaic();
}, { passive: false });



$('zoomLabel').textContent = '100%';
floatArt();

/* 双击 PDF / 拖到 exe 图标上启动时，URL 会带 file 参数，直接打开 */
(async function openFromQuery() {
  const q = new URLSearchParams(location.search);
  const file = q.get('file');
  if (!file) return;
  setBusy(true);
  try {
    const d = await apiPost('/api/openlocal', { path: file });
    await adoptDocument(d);
  } catch (e) {
    toast('无法打开该文件：' + e.message, true);
  } finally {
    setBusy(false);
  }
})();
