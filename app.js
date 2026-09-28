// Super PDF v1.1 — front-end (runs inside Edge app window served by SuperPDF.exe)
import * as pdfjsLib from './pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
const { PDFDocument, degrees } = window.PDFLib;

const $ = (id) => document.getElementById(id);
const TOKEN = new URLSearchParams(location.search).get('t') || '';
const IMG_EXT = ['jpg', 'jpeg', 'jfif', 'png', 'gif', 'bmp', 'webp'];
const TIFF_EXT = ['tif', 'tiff', 'mtiff'];
const WORD_EXT = ['doc', 'docx', 'docm', 'rtf', 'odt', 'txt', 'htm', 'html', 'xml', 'wpd', 'dot', 'dotx'];
const EXCEL_EXT = ['xls', 'xlsx', 'xlsm', 'xlsb', 'csv', 'ods'];
const PPT_EXT = ['ppt', 'pptx', 'pps', 'ppsx', 'odp'];
const A4 = [595.28, 841.89];
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// ------------------------------------------------------------------ state
const sources = new Map();   // id -> { id, name, bytes, pdf (pdfjs doc), password, blank }
const docs = [];             // tabs: { id, name, pages:[{uid,src,index,rot}], selected:Set, lastClicked, undo:[], redo:[], dirty, scroll:{} }
let D = null;                // active doc
let mode = 'organize';
let host = { host: false };
let clip = null;             // { items:[{src,index,rot}], marker, from }
let seq = 1;
const uid = () => 'p' + (seq++);

function newDoc(name = 'เอกสารใหม่', pages = []) {
  const d = { id: 'd' + (seq++), name, pages, selected: new Set(), lastClicked: null, undo: [], redo: [], dirty: false, scroll: {} };
  docs.push(d);
  return d;
}
const isBlankDoc = (d) => d && !d.pages.length && !d.dirty;

// ------------------------------------------------------------------ host API
async function api(path, opts = {}) {
  const res = await fetch(path, { ...opts, headers: { ...(opts.headers || {}), 'X-Token': TOKEN } });
  if (!res.ok) throw new Error(await res.text() || ('HTTP ' + res.status));
  return res;
}

async function initHost() {
  if (!PACKED) { try { host = await (await api('/api/info')).json(); } catch { host = { host: false }; } }
  const o = $('stOffice');
  if (host.host) {
    const tag = (ok, n) => `<span class="${ok ? 'ok' : 'no'}">${ok ? '✔' : '✖'} ${n}</span>`;
    o.innerHTML = tag(host.word, 'Word') + ' &nbsp;' + tag(host.excel, 'Excel') + ' &nbsp;' + tag(host.powerpoint, 'PowerPoint') +
      (host.libreoffice ? ' &nbsp;' + tag(true, 'LibreOffice') : '');
    $('officeHint').textContent = (host.word || host.excel || host.libreoffice)
      ? 'ไฟล์ Word / Excel / PowerPoint จะถูกแปลงด้วย Microsoft Office ในเครื่องโดยอัตโนมัติ'
      : 'ไม่พบ Microsoft Office ในเครื่อง — ยังใช้งาน PDF และรูปภาพได้ตามปกติ';
    loadFontList();
    setInterval(() => api('/api/ping').catch(() => {}), 20000);
    window.addEventListener('pagehide', () => navigator.sendBeacon('/api/bye?t=' + TOKEN));
    pollPending();
  } else {
    loadFontList();
    o.textContent = PACKED ? 'เวอร์ชันเว็บ · ทำงานในเครื่อง ไม่อัปโหลดไฟล์' : 'โหมดเบราว์เซอร์ (แปลง Office ได้เฉพาะเมื่อเปิดผ่าน SuperPDF.exe)';
    $('officeHint').textContent = PACKED
      ? 'ไฟล์ทั้งหมดประมวลผลในเครื่องนี้ ไม่ถูกส่งขึ้นอินเทอร์เน็ต · Word/Excel ให้ส่งออกเป็น PDF จากแอป Office ก่อน'
      : 'เปิดผ่าน SuperPDF.exe เพื่อแปลงไฟล์ Word / Excel ด้วย Microsoft Office';
  }
}

// files passed on the command line / "Open with" (also from a 2nd launch) -> open as tabs
let polling = false;
async function pollPending() {
  if (!polling) {
    polling = true;
    try {
      const list = await (await api('/api/pending')).json();
      if (list.length) {
        const files = [];
        for (const it of list) files.push({ name: it.name, path: it.path, bytes: new Uint8Array(await (await api('/api/file?i=' + it.i)).arrayBuffer()) });
        await openAsTabs(files);
        window.focus();
      }
    } catch {}
    polling = false;
  }
  setTimeout(pollPending, 1500);
}

// ------------------------------------------------------------------ UI helpers
function busy(text) { $('busyText').textContent = text; $('busy').classList.remove('hidden'); }
function unbusy() { $('busy').classList.add('hidden'); }

function toast(msg, kind = '', actions = [], ms = 5000) {
  const el = document.createElement('div');
  el.className = 'toast ' + kind;
  const m = document.createElement('div'); m.className = 'tmsg'; m.textContent = msg; el.appendChild(m);
  for (const [label, fn] of actions) {
    const b = document.createElement('button'); b.textContent = label;
    b.onclick = () => { fn(); el.remove(); }; el.appendChild(b);
  }
  const x = document.createElement('button'); x.textContent = '✕'; x.onclick = () => el.remove(); el.appendChild(x);
  $('toasts').appendChild(el);
  if (ms) setTimeout(() => el.remove(), ms + actions.length * 4000);
}

// generic dialog: returns { btn, values } ; btn = value of pressed button or null (cancel)
function dialog({ title, body = '', buttons = [['cancel', 'ยกเลิก'], ['ok', 'ตกลง', true]], focus }) {
  return new Promise((resolve) => {
    $('modalTitle').textContent = title;
    $('modalBody').innerHTML = body;
    const bb = $('modalBtns'); bb.innerHTML = '';
    const form = $('modalForm');
    const done = (btn) => {
      const values = {};
      for (const el of form.elements) {
        if (!el.name) continue;
        if (el.type === 'radio') { if (el.checked) values[el.name] = el.value; }
        else if (el.type === 'checkbox') values[el.name] = el.checked;
        else values[el.name] = el.value;
      }
      $('modal').classList.add('hidden');
      document.removeEventListener('keydown', onKey, true);
      resolve({ btn, values });
    };
    for (const [val, label, primary] of buttons) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'tbtn' + (primary ? ' primary' : '');
      b.textContent = label; b.onclick = () => done(val === 'cancel' ? null : val); bb.appendChild(b);
    }
    const primaryBtn = buttons.find((b) => b[2]);
    const onKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); done(null); }
      if (e.key === 'Enter' && primaryBtn) { e.preventDefault(); e.stopPropagation(); done(primaryBtn[0]); }
    };
    document.addEventListener('keydown', onKey, true);
    $('modal').classList.remove('hidden');
    setTimeout(() => { const f = focus ? form.querySelector(focus) : bb.querySelector('.primary'); f && f.focus(); }, 30);
  });
}

async function askPassword(fileName, retry) {
  const r = await dialog({
    title: 'ไฟล์นี้มีรหัสผ่าน',
    body: `<p>${retry ? 'รหัสผ่านไม่ถูกต้อง ลองอีกครั้ง — ' : 'กรุณาใส่รหัสผ่านเพื่อเปิด '}${esc(fileName)}</p><input name="pw" type="password">`,
    focus: 'input',
  });
  return r.btn ? r.values.pw : null;
}

const ext = (name) => (name.split('.').pop() || '').toLowerCase();
const baseName = (name) => name.replace(/\.[^.]+$/, '');
const readFile = (f) => f.arrayBuffer().then((b) => ({ name: f.name, bytes: new Uint8Array(b) }));
const pad = (n, w = 3) => String(n).padStart(w, '0');
const dirOf = (p) => p.replace(/[\\/][^\\/]*$/, '');

// pick files with the Windows dialog (so we know their real paths); falls back to <input type=file>
async function pickFiles(inputId, title) {
  if (!host.host) { $(inputId).click(); return null; }
  try {
    const list = await (await api('/api/openfiles?title=' + encodeURIComponent(title))).json();
    const files = [];
    for (const it of list) {
      busy('กำลังอ่าน ' + it.name + ' …');
      files.push({ name: it.name, path: it.path, bytes: new Uint8Array(await (await api('/api/file?i=' + it.i)).arrayBuffer()) });
    }
    unbusy();
    return files;
  } catch (e) { unbusy(); toast('เปิดไฟล์ไม่สำเร็จ: ' + e.message, 'err'); return []; }
}
async function doOpen() { const fs = await pickFiles('fileOpen', 'เปิดไฟล์'); if (fs) await openPicked(fs); }

// ------------------------------------------------------------------ importing
const IS_TOUCH = matchMedia('(hover: none) and (pointer: coarse)').matches;
const PACKED = !!window.SUPERPDF_PACKED; // web build: cmaps & standard fonts packed into 2 files
const packCache = {};
const loadPack = (file) => (packCache[file] ||= fetch('./' + file).then((r) => r.json()));
const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
class PackedCMapFactory {
  constructor() {}
  async fetch({ name }) { const m = await loadPack('cmaps.json'); if (!m[name]) throw new Error('CMap ' + name); return { cMapData: b64(m[name]), isCompressed: true }; }
}
class PackedFontFactory {
  constructor() {}
  async fetch({ filename }) { const m = await loadPack('fonts.json'); if (!m[filename]) throw new Error('font ' + filename); return b64(m[filename]); }
}
const CDN = window.SUPERPDF_CDN; // GitHub Pages build: big libraries come from cdn.jsdelivr.net
const pdfjsOpts = () => CDN
  ? { cMapUrl: CDN + 'cmaps/', cMapPacked: true, standardFontDataUrl: CDN + 'standard_fonts/' }
  : PACKED
  ? { CMapReaderFactory: PackedCMapFactory, StandardFontDataFactory: PackedFontFactory, useWorkerFetch: false, cMapUrl: './', cMapPacked: true, standardFontDataUrl: './' }
  : { cMapUrl: './pdfjs/cmaps/', cMapPacked: true, standardFontDataUrl: './pdfjs/standard_fonts/' };

async function openPdfSource(name, bytes) {
  let retry = false, password;
  while (true) {
    const task = pdfjsLib.getDocument({
      data: bytes.slice(), password, ...pdfjsOpts(),
      isEvalSupported: false,
    });
    try {
      const pdf = await task.promise;
      const id = 's' + (seq++);
      sources.set(id, { id, name, bytes, pdf, password });
      return sources.get(id);
    } catch (e) {
      if (e && e.name === 'PasswordException') {
        unbusy();
        password = await askPassword(name, retry);
        retry = true;
        if (password == null) throw new Error('ยกเลิกการเปิดไฟล์ ' + name);
        busy('กำลังเปิด ' + name + ' …');
        continue;
      }
      throw new Error('เปิดไฟล์ ' + name + ' ไม่ได้ (ไฟล์อาจเสียหาย)');
    }
  }
}

function loadImageEl(blob) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(blob);
    const im = new Image();
    im.onload = () => res(im);
    im.onerror = () => { URL.revokeObjectURL(url); rej(new Error('อ่านรูปภาพไม่ได้')); };
    im.src = url;
  });
}

function jpegOrientation(b) {
  if (b[0] !== 0xFF || b[1] !== 0xD8) return 1;
  let o = 2;
  while (o < b.length - 4) {
    if (b[o] !== 0xFF) return 1;
    const marker = b[o + 1], len = (b[o + 2] << 8) | b[o + 3];
    if (marker === 0xE1 && b[o + 4] === 0x45 && b[o + 5] === 0x78) {
      const t = o + 10, le = b[t] === 0x49;
      const r16 = (p) => le ? b[p] | (b[p + 1] << 8) : (b[p] << 8) | b[p + 1];
      const r32 = (p) => le ? (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0 : ((b[p] << 24) | (b[p + 1] << 16) | (b[p + 2] << 8) | b[p + 3]) >>> 0;
      const ifd = t + r32(t + 4), n = r16(ifd);
      for (let i = 0; i < n; i++) { const e = ifd + 2 + i * 12; if (r16(e) === 0x0112) return r16(e + 8); }
      return 1;
    }
    if (marker === 0xDA) return 1;
    o += 2 + len;
  }
  return 1;
}

function canvasToBytes(canvas, type, q) {
  return new Promise((res) => canvas.toBlob((b) => b.arrayBuffer().then((a) => res(new Uint8Array(a))), type, q));
}

async function imagesToPdf(images) {
  const doc = await PDFDocument.create();
  const fitA4 = $('imgMode').value === 'a4';
  for (const im of images) {
    const emb = im.kind === 'jpg' ? await doc.embedJpg(im.bytes) : await doc.embedPng(im.bytes);
    let pw, ph, dw, dh, x, y;
    if (fitA4) {
      const land = emb.width > emb.height;
      [pw, ph] = land ? [A4[1], A4[0]] : A4;
      const m = 18, s = Math.min((pw - 2 * m) / emb.width, (ph - 2 * m) / emb.height);
      dw = emb.width * s; dh = emb.height * s; x = (pw - dw) / 2; y = (ph - dh) / 2;
    } else {
      const dpi = im.dpi || 96;
      pw = dw = emb.width * 72 / dpi; ph = dh = emb.height * 72 / dpi; x = y = 0;
    }
    const pg = doc.addPage([pw, ph]);
    pg.drawImage(emb, { x, y, width: dw, height: dh });
  }
  return await doc.save();
}

async function imageFileToPdf(f) {
  const e = ext(f.name);
  let item;
  if ((e === 'jpg' || e === 'jpeg' || e === 'jfif') && jpegOrientation(f.bytes) === 1) item = { kind: 'jpg', bytes: f.bytes };
  else if (e === 'png') item = { kind: 'png', bytes: f.bytes };
  else {
    const im = await loadImageEl(new Blob([f.bytes]));
    const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight;
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0);
    const hasAlpha = e === 'gif' || e === 'webp';
    item = hasAlpha ? { kind: 'png', bytes: await canvasToBytes(c, 'image/png') } : { kind: 'jpg', bytes: await canvasToBytes(c, 'image/jpeg', 0.92) };
  }
  return imagesToPdf([item]);
}

async function tiffToPdf(f) {
  const buf = () => f.bytes.buffer.slice(f.bytes.byteOffset, f.bytes.byteOffset + f.bytes.byteLength);
  const ifds = UTIF.decode(buf());
  const frames = ifds.filter((d) => d.t256 && d.t257 && !(d.t254 && d.t254[0] & 1));
  if (!frames.length) throw new Error('อ่านไฟล์ TIFF ไม่ได้');
  const items = [];
  for (const ifd of frames) {
    UTIF.decodeImage(buf(), ifd, ifds);
    const rgba = UTIF.toRGBA8(ifd);
    const c = document.createElement('canvas'); c.width = ifd.width; c.height = ifd.height;
    c.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, rgba.length), ifd.width, ifd.height), 0, 0);
    const bw = ifd.t258 && ifd.t258[0] === 1;
    const dpi = ifd.t282 ? (Array.isArray(ifd.t282[0]) ? ifd.t282[0][0] / (ifd.t282[0][1] || 1) : ifd.t282[0]) : 0;
    items.push({ kind: bw ? 'png' : 'jpg', bytes: await canvasToBytes(c, bw ? 'image/png' : 'image/jpeg', 0.9), dpi: dpi > 20 ? dpi : 0 });
  }
  return imagesToPdf(items);
}

async function officeToPdf(f) {
  if (!host.host) throw new Error(IS_TOUCH || PACKED
    ? `แปลง ${f.name} ในมือถือไม่ได้ — เปิดไฟล์ในแอป Word/Excel แล้วเลือก "ส่งออก/Export เป็น PDF" จากนั้นนำ PDF มาเปิดที่นี่`
    : 'แปลง ' + f.name + ' ไม่ได้: ต้องเปิดโปรแกรมผ่าน SuperPDF.exe และมี Microsoft Office');
  const e = ext(f.name);
  const app = EXCEL_EXT.includes(e) ? 'Excel' : PPT_EXT.includes(e) ? 'PowerPoint' : 'Word';
  busy(`กำลังแปลง ${f.name} ด้วย Microsoft ${app} … (อาจใช้เวลาสักครู่)`);
  const res = await api('/api/convert?name=' + encodeURIComponent(f.name), { method: 'POST', body: f.bytes });
  return new Uint8Array(await res.arrayBuffer());
}

async function fileToPages(f) {
  const e = ext(f.name);
  let pdfBytes;
  busy('กำลังเปิด ' + f.name + ' …');
  if (e === 'pdf' || (f.bytes[0] === 0x25 && f.bytes[1] === 0x50 && f.bytes[2] === 0x44 && f.bytes[3] === 0x46)) pdfBytes = f.bytes;
  else if (IMG_EXT.includes(e)) pdfBytes = await imageFileToPdf(f);
  else if (TIFF_EXT.includes(e)) pdfBytes = await tiffToPdf(f);
  else if (WORD_EXT.includes(e) || EXCEL_EXT.includes(e) || PPT_EXT.includes(e)) pdfBytes = await officeToPdf(f);
  else throw new Error('ไม่รองรับไฟล์ชนิด .' + e + ' (' + f.name + ')');
  const src = await openPdfSource(f.name, pdfBytes);
  const out = [];
  for (let i = 0; i < src.pdf.numPages; i++) out.push({ uid: uid(), src: src.id, index: i, rot: 0 });
  return out;
}

async function filesToPages(files) {
  const added = [], errors = [];
  try {
    for (const f of files) {
      try { added.push({ f, pages: await fileToPages(f) }); }
      catch (e) { errors.push(e.message || String(e)); }
    }
  } finally { unbusy(); }
  errors.forEach((m) => toast(m, 'err', [], 9000));
  return added;
}

// open each file in its own tab (or merge into one tab)
async function openAsTabs(files, { merge = false } = {}) {
  if (!files.length) return;
  const res = await filesToPages(files);
  if (!res.length) return;
  const groups = merge ? [{ f: res[0].f, pages: res.flatMap((r) => r.pages), merged: res.length > 1 }] : res;
  let last;
  for (const g of groups) {
    const name = baseName(g.f.name) + (g.merged ? '_รวม' : '');
    let d = isBlankDoc(D) && !last ? D : newDoc();
    d.name = name; d.pages = g.pages; d.selected = new Set(); d.undo = []; d.redo = [];
    d.dirty = g.merged || ext(g.f.name) !== 'pdf';
    d.path = !g.merged && g.f.path && ext(g.f.name) === 'pdf' ? g.f.path : null;  // enables "save over original"
    d.dir = g.f.path ? dirOf(g.f.path) : null;
    last = d;
  }
  switchTo(last);
}

// insert files into the active doc
async function insertFiles(files, at = null) {
  if (!files.length) return;
  if (!D.pages.length) return openAsTabs(files, { merge: true });
  const res = await filesToPages(files);
  const added = res.flatMap((r) => r.pages);
  if (!added.length) return;
  pushUndo();
  const pos = at == null ? D.pages.length : at;
  D.pages.splice(pos, 0, ...added);
  D.selected = new Set(added.map((p) => p.uid));
  D.dirty = true;
  refresh();
  toast(`เพิ่ม ${added.length} หน้าแล้ว`, 'ok', [], 2500);
  cardEls.get(added[0].uid)?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

async function openPicked(files) {
  if (!files.length) return;
  let merge = false;
  if (files.length > 1) {
    const r = await dialog({
      title: `เปิด ${files.length} ไฟล์`,
      body: '<p>ต้องการเปิดไฟล์อย่างไร?</p>',
      buttons: [['cancel', 'ยกเลิก'], ['tabs', 'เปิดแยกแท็บ'], ['merge', 'รวมเป็น PDF เดียว', true]],
    });
    if (!r.btn) return;
    merge = r.btn === 'merge';
  }
  await openAsTabs(files, { merge });
}

function insertionIndex() {
  if (!D.selected.size) return D.pages.length;
  let last = -1;
  D.pages.forEach((p, i) => { if (D.selected.has(p.uid)) last = i; });
  return last + 1;
}
const selPages = () => D.pages.filter((p) => D.selected.has(p.uid));

// ------------------------------------------------------------------ undo
const cloneAnn = (a) => ({ ...a, id: a.id });
const clonePage = (p, fresh = true) => ({ ...p, uid: fresh ? uid() : p.uid, annots: (p.annots || []).map(cloneAnn) });
function snapshot() { return { pages: D.pages.map((p) => clonePage(p, false)), name: D.name }; }
function pushUndo() { D.undo.push(snapshot()); if (D.undo.length > 100) D.undo.shift(); D.redo = []; }
function restore(s) { D.pages = s.pages; D.name = s.name; D.selected = new Set([...D.selected].filter((u) => D.pages.some((p) => p.uid === u))); D.dirty = true; refresh(); }
function undo() { if (!D.undo.length) return; D.redo.push(snapshot()); restore(D.undo.pop()); }
function redo() { if (!D.redo.length) return; D.undo.push(snapshot()); restore(D.redo.pop()); }

// ------------------------------------------------------------------ tabs
function renderTabs() {
  const box = $('tabs');
  box.replaceChildren(...docs.map((d) => {
    const t = document.createElement('div');
    t.className = 'tab' + (d === D ? ' active' : '');
    t.dataset.id = d.id;
    t.title = (d.path || d.name) + (d.pages.length ? ` (${d.pages.length} หน้า)` : '');
    t.innerHTML = `<svg class="ticon"><use href="#i-blank"/></svg><span class="tname">${esc(d.name)}</span>${d.dirty && d.pages.length ? '<span class="tdirty" title="ยังไม่ได้บันทึก"></span>' : ''}<button class="tclose" title="ปิดแท็บ (Ctrl+W)"><svg><use href="#i-x"/></svg></button>`;
    return t;
  }));
  box.querySelector('.tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function switchTo(d) {
  if (D && D !== d) D.scroll = { organize: $('organize').scrollTop, reader: $('reader').scrollTop };
  D = d;
  readerKey = '';
  refresh();
  $('organize').scrollTop = d.scroll.organize || 0;
  if (mode === 'read') setTimeout(() => { $('reader').scrollTop = d.scroll.reader || 0; }, 80);
}

async function closeTab(d) {
  if (d.dirty && d.pages.length) {
    const r = await dialog({
      title: 'ปิดแท็บ', body: `<p>“${esc(d.name)}” ยังไม่ได้บันทึก ต้องการบันทึกก่อนปิดหรือไม่?</p>`,
      buttons: [['cancel', 'ยกเลิก'], ['discard', 'ไม่บันทึก'], ['save', 'บันทึก', true]],
    });
    if (!r.btn) return;
    if (r.btn === 'save') { if (D !== d) switchTo(d); if (!(await saveAll())) return; }
  }
  const i = docs.indexOf(d);
  docs.splice(i, 1);
  if (!docs.length) newDoc();
  if (D === d || !docs.includes(D)) { D = null; switchTo(docs[Math.min(i, docs.length - 1)]); }
  else refresh();
}

$('tabs').addEventListener('click', (e) => {
  const t = e.target.closest('.tab'); if (!t) return;
  const d = docs.find((x) => x.id === t.dataset.id);
  if (e.target.closest('.tclose')) closeTab(d); else if (d !== D) switchTo(d);
});
$('tabs').addEventListener('dblclick', (e) => { const t = e.target.closest('.tab'); if (t && !e.target.closest('.tclose')) renameDoc(docs.find((x) => x.id === t.dataset.id)); });
$('tabs').addEventListener('auxclick', (e) => { if (e.button === 1) { const t = e.target.closest('.tab'); if (t) closeTab(docs.find((x) => x.id === t.dataset.id)); } });
$('tabs').addEventListener('contextmenu', (e) => {
  const t = e.target.closest('.tab'); if (!t) return;
  e.preventDefault();
  const d = docs.find((x) => x.id === t.dataset.id);
  showMenu([
    { label: 'บันทึกแท็บนี้…', icon: 'save', action: () => { switchTo(d); saveAll(); } },
    { label: 'บันทึกเป็นไฟล์ใหม่…', icon: 'extract', action: () => { switchTo(d); saveAll(true); } },
    { label: 'เปลี่ยนชื่อ…', icon: 'blank', action: () => renameDoc(d) },
    { label: 'ทำสำเนาแท็บ', icon: 'dup', action: () => { const n = newDoc(d.name + ' (สำเนา)', d.pages.map((p) => clonePage(p))); n.dirty = true; switchTo(n); } },
    { sep: true },
    { label: 'ปิดแท็บ', icon: 'x', kbd: 'Ctrl+W', action: () => closeTab(d) },
    { label: 'ปิดแท็บอื่นทั้งหมด', disabled: docs.length < 2, action: async () => { for (const o of docs.filter((x) => x !== d)) await closeTab(o); } },
  ], e.clientX, e.clientY);
});
$('btnNewTab').onclick = () => switchTo(newDoc());

async function renameDoc(d) {
  const r = await dialog({ title: 'เปลี่ยนชื่อเอกสาร', body: `<input name="n" type="text" value="${esc(d.name)}">`, focus: 'input' });
  if (r.btn && r.values.n.trim()) { d.name = r.values.n.trim(); refresh(); }
}

// ------------------------------------------------------------------ thumbnails
const thumbCache = new Map();
const thumbQueue = [];
let thumbRunning = 0;
const tkey = (p) => `${p.src}:${p.index}:${p.rot}:${p.av || 0}`;
const rkey = (p) => `${p.src}:${p.index}:${p.rot}`;

function requestThumb(p) {
  const k = tkey(p);
  if (thumbCache.has(k)) { applyThumb(k); return; }
  thumbQueue.push({ p: { ...p }, k });
  pumpThumbs();
}
function pumpThumbs() {
  while (thumbRunning < 3 && thumbQueue.length) {
    const job = thumbQueue.pop();
    if (thumbCache.has(job.k)) { applyThumb(job.k); continue; }
    thumbRunning++;
    renderThumb(job.p).then((url) => { thumbCache.set(job.k, url); applyThumb(job.k); })
      .catch(() => {}).finally(() => { thumbRunning--; pumpThumbs(); });
  }
}
function applyThumb(k) {
  document.querySelectorAll(`.card[data-key="${k}"] .thumb`).forEach((t) => {
    t.innerHTML = ''; const im = new Image(); im.src = thumbCache.get(k); t.appendChild(im);
  });
}
async function renderThumb(p) {
  const page = await sources.get(p.src).pdf.getPage(p.index + 1);
  const rotation = (page.rotate + p.rot) % 360;
  const v1 = page.getViewport({ scale: 1, rotation });
  const vp = page.getViewport({ scale: 360 / Math.max(v1.width, v1.height), rotation });
  const c = document.createElement('canvas'); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: g, viewport: vp }).promise;
  drawAnnots(g, p, page, vp.scale);
  const blob = await new Promise((r) => c.toBlob(r, IS_TOUCH ? 'image/jpeg' : 'image/png', 0.85));
  c.width = c.height = 0;
  return URL.createObjectURL(blob);
}

const thumbObserver = new IntersectionObserver((entries) => {
  for (const en of entries) {
    if (!en.isIntersecting) continue;
    const card = en.target; thumbObserver.unobserve(card);
    const p = D.pages.find((x) => x.uid === card.dataset.uid);
    if (p) requestThumb(p);
  }
}, { root: $('organize'), rootMargin: '400px' });

// ------------------------------------------------------------------ organize view
const cardEls = new Map();

function makeCard(p) {
  const el = document.createElement('div');
  el.className = 'card';
  el.dataset.uid = p.uid;
  el.innerHTML = `<div class="thumb"><div class="ph"></div></div>
    <div class="meta"><span class="num"></span><span class="src"></span></div>
    <div class="actions">
      <button data-act="rl" title="หมุนซ้าย"><svg><use href="#i-rl"/></svg></button>
      <button data-act="rr" title="หมุนขวา"><svg><use href="#i-rr"/></svg></button>
      <button data-act="ins" title="แทรกไฟล์ต่อจากหน้านี้"><svg><use href="#i-plus"/></svg></button>
      <button data-act="del" class="del" title="ลบหน้านี้"><svg><use href="#i-trash"/></svg></button>
      <button data-act="more" class="more" title="เมนู"><svg><use href="#i-more"/></svg></button>
    </div>`;
  return el;
}

function updateCard(el, p, i) {
  const k = tkey(p);
  el.querySelector('.num').textContent = i + 1;
  const s = sources.get(p.src);
  el.querySelector('.src').textContent = s.blank ? 'หน้าว่าง' : `${s.name} · น.${p.index + 1}`;
  el.querySelector('.src').title = s.name;
  el.classList.toggle('selected', D.selected.has(p.uid));
  el.removeAttribute('data-count');
  if (el.dataset.key !== k) {
    el.dataset.key = k;
    const t = el.querySelector('.thumb');
    if (thumbCache.has(k)) { t.innerHTML = ''; const im = new Image(); im.src = thumbCache.get(k); t.appendChild(im); }
    else { t.innerHTML = '<div class="ph"></div>'; }
  }
  if (!thumbCache.has(k)) thumbObserver.observe(el);
}

function renderGrid() {
  const alive = new Set(docs.flatMap((d) => d.pages.map((p) => p.uid)));
  for (const [u, el] of cardEls) if (!alive.has(u)) { thumbObserver.unobserve(el); cardEls.delete(u); }
  const els = D.pages.map((p, i) => {
    let el = cardEls.get(p.uid);
    if (!el) { el = makeCard(p); cardEls.set(p.uid, el); }
    updateCard(el, p, i);
    return el;
  });
  $('grid').replaceChildren(...els);
}

function paintSelection() {
  for (const el of $('grid').children) el.classList.toggle('selected', D.selected.has(el.dataset.uid));
  updateStatus();
}

function updateStatus() {
  const has = D.pages.length > 0, nSel = D.selected.size;
  $('stDoc').textContent = has ? `${D.name}${D.dirty ? ' •' : ''}  —  ${D.pages.length} หน้า` : 'ยังไม่มีเอกสาร';
  $('stSel').textContent = nSel ? `เลือก ${nSel} หน้า` : '';
  for (const id of ['btnDelete', 'btnRotL', 'btnRotR', 'btnCopy', 'btnCut']) $(id).disabled = !nSel;
  $('btnPaste').disabled = false;
  $('btnSplit').disabled = !has; $('btnConvert').disabled = !has; $('btnSave').disabled = !has;
  $('btnUndo').disabled = !D.undo.length; $('btnRedo').disabled = !D.redo.length;
}

function refresh() {
  const has = D.pages.length > 0;
  $('empty').classList.toggle('hidden', has);
  $('organize').classList.toggle('hidden', !has || mode !== 'organize');
  $('reader').classList.toggle('hidden', !has || mode !== 'read');
  document.body.className = 'mode-' + mode + (has ? '' : ' noDoc');
  if (has && mode === 'organize') renderGrid();
  if (has && mode === 'read') { renderReader(); renderAllAnnLayers(); }
  updateAnnBar();
  updateStatus();
  renderTabs();
  $('pageTotal').textContent = D.pages.length;
  document.title = has ? `${D.name}${D.dirty ? ' *' : ''} — Super PDF` : 'Super PDF';
}

// click selection & card buttons
let lastPointer = 'mouse';
window.addEventListener('pointerdown', (e) => { lastPointer = e.pointerType; }, true);
$('grid').addEventListener('click', (e) => {
  const card = e.target.closest('.card');
  if (!card) return;
  const u = card.dataset.uid;
  const act = e.target.closest('button')?.dataset.act;
  if (act) {
    e.stopPropagation();
    if (act === 'del') deletePages([u]);
    if (act === 'rl') rotatePages([u], -90);
    if (act === 'rr') rotatePages([u], 90);
    if (act === 'ins') pickInsert(D.pages.findIndex((p) => p.uid === u) + 1);
    if (act === 'more') {
      if (!D.selected.has(u)) { D.selected = new Set([u]); D.lastClicked = u; paintSelection(); }
      const r = e.target.closest('button').getBoundingClientRect();
      showMenu(pageMenu(D.pages.findIndex((p) => p.uid === u)), r.left - 200, r.bottom + 4);
    }
    return;
  }
  if (IS_TOUCH || lastPointer === 'touch') { // tap = toggle selection
    D.selected.has(u) ? D.selected.delete(u) : D.selected.add(u);
    D.lastClicked = u; paintSelection(); return;
  }
  if (e.shiftKey && D.lastClicked) {
    const a = D.pages.findIndex((p) => p.uid === D.lastClicked), b = D.pages.findIndex((p) => p.uid === u);
    if (!(e.ctrlKey || e.metaKey)) D.selected.clear();
    for (let i = Math.min(a, b); i <= Math.max(a, b); i++) D.selected.add(D.pages[i].uid);
  } else if (e.ctrlKey || e.metaKey) {
    D.selected.has(u) ? D.selected.delete(u) : D.selected.add(u);
    D.lastClicked = u;
  } else {
    D.selected = new Set([u]); D.lastClicked = u;
  }
  paintSelection();
});
$('grid').addEventListener('dblclick', (e) => {
  const card = e.target.closest('.card'); if (!card) return;
  const i = D.pages.findIndex((p) => p.uid === card.dataset.uid);
  setMode('read'); setTimeout(() => goToPage(i + 1), 60);
});

// ---- rubber-band (marquee) selection on empty space
let marquee = null;
$('organize').appendChild($('marquee'));
$('organize').addEventListener('mousedown', (e) => {
  if (e.button !== 0 || e.target.closest('.card')) return;
  if (lastPointer === 'touch') { if (!e.target.closest('.card')) { D.selected.clear(); paintSelection(); } return; }
  const org = $('organize');
  if (e.clientX - org.getBoundingClientRect().left >= org.clientWidth) return; // on scrollbar
  e.preventDefault();
  const r = org.getBoundingClientRect();
  marquee = {
    x0: e.clientX - r.left + org.scrollLeft, y0: e.clientY - r.top + org.scrollTop,
    cx: e.clientX, cy: e.clientY, moved: false,
    base: (e.ctrlKey || e.shiftKey) ? new Set(D.selected) : new Set(),
  };
  const move = (ev) => { marquee.cx = ev.clientX; marquee.cy = ev.clientY; updateMarquee(); };
  const up = () => {
    window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up);
    cancelAnimationFrame(marquee.raf);
    $('marquee').classList.add('hidden');
    if (!marquee.moved) { D.selected = marquee.base; paintSelection(); }
    marquee = null;
  };
  window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  const tick = () => { // auto-scroll near edges
    if (!marquee) return;
    const rr = org.getBoundingClientRect(); let dy = 0;
    if (marquee.cy > rr.bottom - 40) dy = Math.min(30, (marquee.cy - (rr.bottom - 40)) / 2 + 4);
    else if (marquee.cy < rr.top + 30) dy = -Math.min(30, ((rr.top + 30) - marquee.cy) / 2 + 4);
    if (dy && marquee.moved) { org.scrollTop += dy; updateMarquee(); }
    marquee.raf = requestAnimationFrame(tick);
  };
  marquee.raf = requestAnimationFrame(tick);
});
function updateMarquee() {
  const org = $('organize'), r = org.getBoundingClientRect();
  const x1 = marquee.cx - r.left + org.scrollLeft, y1 = marquee.cy - r.top + org.scrollTop;
  if (!marquee.moved && Math.hypot(x1 - marquee.x0, y1 - marquee.y0) < 5) return;
  marquee.moved = true;
  const L = Math.min(marquee.x0, x1), T = Math.min(marquee.y0, y1), W = Math.abs(x1 - marquee.x0), H = Math.abs(y1 - marquee.y0);
  Object.assign($('marquee').style, { left: L + 'px', top: T + 'px', width: W + 'px', height: H + 'px' });
  $('marquee').classList.remove('hidden');
  const sel = new Set(marquee.base);
  for (const el of $('grid').children) {
    const cl = el.offsetLeft, ct = el.offsetTop, cr = cl + el.offsetWidth, cb = ct + el.offsetHeight;
    if (cl < L + W && cr > L && ct < T + H && cb > T) sel.add(el.dataset.uid);
  }
  D.selected = sel;
  paintSelection();
}

// ---- drag to reorder (moves all selected pages together); drop on a tab = copy there
let dragInfo = null;
new Sortable($('grid'), {
  animation: 180,
  ghostClass: 'ghost', chosenClass: 'chosen', dragClass: 'drag-el',
  forceFallback: true, fallbackTolerance: 4,
  delay: 260, delayOnTouchOnly: true, touchStartThreshold: 8,
  scroll: true, scrollSensitivity: 90, scrollSpeed: 18, bubbleScroll: true,
  filter: '.actions button', preventOnFilter: false,
  onStart: (evt) => {
    const u = evt.item.dataset.uid;
    if (!D.selected.has(u)) { D.selected = new Set([u]); D.lastClicked = u; paintSelection(); }
    const n = D.selected.size;
    dragInfo = { uid: u, tab: null };
    if (n > 1) {
      evt.item.dataset.count = n;
      document.querySelectorAll('.sortable-fallback').forEach((f) => { f.dataset.count = n; });
      for (const el of $('grid').children) if (el !== evt.item && D.selected.has(el.dataset.uid)) el.classList.add('ghost');
    }
    pushUndo();
    window.addEventListener('mousemove', dragTrack);
  },
  onEnd: (evt) => {
    window.removeEventListener('mousemove', dragTrack);
    document.querySelectorAll('.tab.droptarget').forEach((t) => t.classList.remove('droptarget'));
    for (const el of $('grid').children) el.classList.remove('ghost');
    evt.item.removeAttribute('data-count');
    const target = dragInfo && dragInfo.tab;
    dragInfo = null;
    if (target && target !== D) {           // dropped on another tab -> copy pages there
      D.pages = D.undo.pop().pages;
      const items = selPages();
      copyPagesToDoc(items, target);
      refresh();
      return;
    }
    const order = [...$('grid').children].map((el) => el.dataset.uid);
    const byId = new Map(D.pages.map((p) => [p.uid, p]));
    const dragged = evt.item.dataset.uid;
    let finalOrder = order;
    if (D.selected.size > 1) {
      const group = D.pages.filter((p) => D.selected.has(p.uid)).map((p) => p.uid);
      const rest = order.filter((u) => !D.selected.has(u) || u === dragged);
      const at = rest.indexOf(dragged);
      rest.splice(at, 1, ...group);
      finalOrder = rest;
    }
    if (finalOrder.join() === D.pages.map((p) => p.uid).join()) { D.undo.pop(); renderGrid(); return; }
    D.pages = finalOrder.map((u) => byId.get(u));
    D.dirty = true;
    refresh();
  },
});
function dragTrack(e) {
  if (!dragInfo) return;
  const el = document.elementFromPoint(e.clientX, e.clientY);
  const t = el && el.closest && el.closest('.tab');
  document.querySelectorAll('.tab.droptarget').forEach((x) => { if (x !== t) x.classList.remove('droptarget'); });
  const d = t ? docs.find((x) => x.id === t.dataset.id) : null;
  if (t && d !== D) t.classList.add('droptarget');
  dragInfo.tab = d && d !== D ? d : null;
}

function copyPagesToDoc(items, target, { at = null, silent = false } = {}) {
  const copies = items.map((p) => ({ uid: uid(), src: p.src, index: p.index, rot: p.rot, av: p.av || 0, annots: (p.annots || []).map((a) => ({ ...a, id: annId() })) }));
  const cur = D;
  D = target; pushUndo(); D = cur;
  target.pages.splice(at == null ? target.pages.length : at, 0, ...copies);
  target.selected = new Set(copies.map((p) => p.uid));
  target.dirty = true;
  if (target.pages.length === copies.length && target.name === 'เอกสารใหม่') target.name = cur.name + ' (แยก)';
  if (!silent) toast(`คัดลอก ${copies.length} หน้าไปยังแท็บ “${target.name}” แล้ว`, 'ok', [['ไปที่แท็บ', () => switchTo(target)]], 4000);
  return copies;
}

function deletePages(uids) {
  if (!uids.length) return;
  pushUndo();
  const s = new Set(uids);
  D.pages = D.pages.filter((p) => !s.has(p.uid));
  uids.forEach((u) => D.selected.delete(u));
  D.dirty = true;
  refresh();
  toast(`ลบ ${uids.length} หน้าแล้ว`, '', [['เลิกทำ', undo]], 4000);
}

function rotatePages(uids, d) {
  if (!uids.length) return;
  pushUndo();
  const s = new Set(uids);
  D.pages = D.pages.map((p) => s.has(p.uid) ? { ...p, rot: (p.rot + d + 360) % 360 } : p);
  D.dirty = true;
  refresh();
}

function duplicatePages() {
  const items = selPages(); if (!items.length) return;
  pushUndo();
  const copies = items.map((p) => clonePage(p));
  D.pages.splice(insertionIndex(), 0, ...copies);
  D.selected = new Set(copies.map((p) => p.uid));
  D.dirty = true; refresh();
}

async function addBlankPage(at = null) {
  let size = A4;
  const pos = at == null ? insertionIndex() : at;
  const refIdx = Math.min(pos, D.pages.length) - 1;
  if (refIdx >= 0) {
    try {
      const p = D.pages[refIdx];
      const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
      const v = pg.getViewport({ scale: 1, rotation: (pg.rotate + p.rot) % 360 });
      size = [v.width, v.height];
    } catch {}
  }
  const doc = await PDFDocument.create();
  doc.addPage(size);
  const src = await openPdfSource('หน้าว่าง', await doc.save());
  src.blank = true;
  const np = { uid: uid(), src: src.id, index: 0, rot: 0 };
  pushUndo();
  D.pages.splice(pos, 0, np);
  D.selected = new Set([np.uid]);
  D.dirty = true;
  if (mode !== 'organize') setMode('organize');
  refresh();
  cardEls.get(np.uid)?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

// ------------------------------------------------------------------ clipboard (copy / cut / paste)
function doCopy(cut = false) {
  const items = selPages();
  if (!items.length) return false;
  clip = { items: items.map((p) => ({ src: p.src, index: p.index, rot: p.rot, av: p.av || 0, annots: (p.annots || []).map(cloneAnn) })), marker: 'SuperPDF-clip-' + (seq++), from: D.name };
  if (cut) {
    pushUndo();
    const s = new Set(items.map((p) => p.uid));
    D.pages = D.pages.filter((p) => !s.has(p.uid));
    D.selected.clear(); D.dirty = true;
    refresh();
  }
  updateStatus();
  toast(`${cut ? 'ตัด' : 'คัดลอก'} ${items.length} หน้าแล้ว — เลือกตำแหน่ง/แท็บแล้วกด Ctrl+V เพื่อวาง`, '', cut ? [['เลิกทำ', undo]] : [], 3500);
  return true;
}

function doPaste(at = null) {
  if (!clip) { toast('ยังไม่มีหน้าที่คัดลอกไว้', ''); return; }
  const pos = at == null ? insertionIndex() : at;
  const copies = copyPagesToDoc(clip.items, D, { at: pos, silent: true });
  if (D.name === 'เอกสารใหม่' && D.pages.length === copies.length) D.name = clip.from + ' (คัดลอก)';
  if (mode !== 'organize') setMode('organize');
  refresh();
  cardEls.get(copies[0].uid)?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  toast(`วาง ${copies.length} หน้าแล้ว`, 'ok', [], 2000);
}

// system clipboard integration: keyboard Ctrl+C / Ctrl+X / Ctrl+V
const typingInField = () => { const a = document.activeElement; return a && (a.tagName === 'INPUT' && a.type !== 'range' && a.type !== 'color' && a.type !== 'checkbox' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT' || a.isContentEditable); };
const modalOpen = () => !$('modal').classList.contains('hidden');
let pendingCut = false, copyHandled = false, pasteHandled = false;
document.addEventListener('copy', (e) => {
  if (typingInField() || modalOpen() || !D.selected.size) return;
  copyHandled = true;
  if (doCopy(pendingCut)) { e.clipboardData.setData('text/plain', clip.marker); e.preventDefault(); }
  pendingCut = false;
});
document.addEventListener('paste', async (e) => {
  if (typingInField() || modalOpen()) return;
  pasteHandled = true;
  e.preventDefault();
  const pos = insertionIndex();
  const data = await readPasteEvent(e);
  if (!data.files.length && !data.text && host.host) return pasteAny(pos); // e.g. files copied in Explorer
  await applyPasted(data, pos);
});

// ------------------------------------------------------------------ paste from other apps
// Sources: our own copied pages, files (PDF / images / Office) and pictures copied in other apps,
// text (becomes a new page), or a link to a PDF / picture.
const TYPE_EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/tiff': 'tiff', 'application/pdf': 'pdf' };
async function readPasteEvent(e) {
  const cd = e.clipboardData, files = [];
  let n = 0;
  for (const f of cd.files) {
    n++;
    const r = await readFile(f);
    const e2 = TYPE_EXT[f.type];
    const generic = !f.name || /^image\.(png|jpe?g)$/i.test(f.name);
    files.push({ ...r, name: generic ? `ภาพที่วาง_${n}.${e2 || 'png'}` : f.name });
  }
  return { files, text: cd.getData('text/plain') || '' };
}
async function insertOrOpen(files, pos) {
  if (!D.pages.length) await openAsTabs(files, { merge: true });
  else await insertFiles(files, pos);
}
async function applyPasted({ files, text }, pos) {
  if (annClip && text === annClip.marker) { if (mode === 'read') return pasteAnn(); }
  if (clip && text === clip.marker) return doPaste(pos);
  if (mode === 'read' && !files.length && text && text.trim() && !/^https?:\/\/\S+$/i.test(text.trim())) return pasteAnn(text.replace(/\r\n?/g, '\n').trim());
  if (files.length) return insertOrOpen(files, pos);
  if (text && text.trim()) return pasteText(text.trim(), pos);
  if (clip) return doPaste(pos);
  toast('ไม่พบไฟล์ รูปภาพ หรือข้อความในคลิปบอร์ด', '', [], 4000);
}

async function pasteAny(at = null) {
  const pos = at == null ? insertionIndex() : at;
  // 1) Windows (SuperPDF.exe): read the real clipboard — files copied in Explorer, pictures, text
  if (host.host) {
    try {
      const r = await (await api('/api/clipboard')).json();
      const files = [];
      for (const it of r.files) {
        busy('กำลังอ่าน ' + it.name + ' …');
        files.push({ name: it.name, path: it.clip ? null : it.path, bytes: new Uint8Array(await (await api('/api/file?i=' + it.i)).arrayBuffer()) });
      }
      unbusy();
      return applyPasted({ files, text: r.text || '' }, pos);
    } catch (e) { unbusy(); toast('อ่านคลิปบอร์ดไม่สำเร็จ: ' + e.message, 'err'); return; }
  }
  // 2) browser / iPhone: async clipboard API (iPhone shows a small "วาง/Paste" bubble to confirm)
  if (navigator.clipboard && navigator.clipboard.read) {
    try {
      const items = await navigator.clipboard.read();
      const files = []; let text = '', n = 0;
      for (const it of items) {
        const t = it.types.find((x) => TYPE_EXT[x]);
        if (t) { n++; files.push({ name: `ภาพที่วาง_${n}.${TYPE_EXT[t]}`.replace('ภาพที่วาง_' + n + '.pdf', 'ไฟล์ที่วาง_' + n + '.pdf'), bytes: new Uint8Array(await (await it.getType(t)).arrayBuffer()) }); continue; }
        if (!text && it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
      }
      return applyPasted({ files, text }, pos);
    } catch (e) { /* not allowed / unsupported -> paste box */ }
  }
  // 3) fallback: a box where the user long-presses / right-clicks and chooses "Paste"
  const data = await pasteBox();
  if (data) return applyPasted(data, pos);
}

function pasteBox() {
  return new Promise((resolve) => {
    let got = null;
    dialog({
      title: 'วางจากแอปอื่น',
      body: `<p>${IS_TOUCH ? 'แตะค้างในกรอบด้านล่าง แล้วเลือก <b>“วาง”</b>' : 'คลิกในกรอบด้านล่าง แล้วกด <b>Ctrl+V</b>'}</p>
        <div id="pasteZone" class="pastezone" contenteditable="true" inputmode="none"></div>`,
      buttons: [['cancel', 'ยกเลิก']],
    }).then(() => resolve(got));
    const z = $('pasteZone');
    setTimeout(() => z.focus(), 60);
    z.addEventListener('paste', async (e) => {
      e.preventDefault();
      got = await readPasteEvent(e);
      $('modalBtns').querySelector('button').click();
    });
  });
}

async function pasteText(text, pos) {
  if (/^https?:\/\/\S+$/i.test(text)) return pasteUrl(text, pos);
  const r = await dialog({
    title: 'วางข้อความเป็นหน้าใหม่',
    body: `<p>ข้อความ ${text.length.toLocaleString()} ตัวอักษร จะถูกจัดเป็นหน้ากระดาษ A4 (บันทึกเป็นภาพ)</p><div class="pathbox" style="max-height:140px;overflow:auto;white-space:pre-wrap">${esc(text.slice(0, 600))}${text.length > 600 ? '…' : ''}</div>`,
    buttons: [['cancel', 'ยกเลิก'], ['ok', 'สร้างหน้า', true]],
  });
  if (!r.btn) return;
  busy('กำลังสร้างหน้าจากข้อความ …');
  try {
    const bytes = await textToPdf(text);
    unbusy();
    await insertOrOpen([{ name: 'ข้อความที่วาง.pdf', bytes }], pos);
  } catch (e) { unbusy(); toast('สร้างหน้าไม่สำเร็จ: ' + e.message, 'err'); }
}

async function pasteUrl(url, pos) {
  let name = decodeURIComponent((url.split(/[?#]/)[0].split('/').pop() || 'ไฟล์จากลิงก์'));
  busy('กำลังดาวน์โหลดจากลิงก์ …');
  try {
    let bytes, type = '';
    if (host.host) {
      const res = await api('/api/fetchurl?url=' + encodeURIComponent(url));
      type = res.headers.get('X-Type') || ''; bytes = new Uint8Array(await res.arrayBuffer());
    } else {
      const res = await fetch(url, { mode: 'cors' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      type = res.headers.get('Content-Type') || ''; bytes = new Uint8Array(await res.arrayBuffer());
    }
    unbusy();
    const t = TYPE_EXT[type.split(';')[0].trim()];
    if (!/\.[a-z0-9]{2,5}$/i.test(name)) name += '.' + (t || 'pdf');
    await insertOrOpen([{ name, bytes }], pos);
  } catch (e) {
    unbusy();
    toast(host.host ? 'ดาวน์โหลดจากลิงก์ไม่สำเร็จ: ' + e.message
      : 'เว็บไซต์ต้นทางไม่อนุญาตให้ดึงไฟล์โดยตรง — เปิดลิงก์ใน Safari แล้วบันทึกลง “ไฟล์” ก่อน จากนั้นเปิดในแอปนี้', 'err', [], 9000);
  }
}

// render plain text (Thai supported) onto A4 pages at 150 dpi
async function textToPdf(text) {
  const W = 1240, H = 1754, M = 110, FS = 29, LH = Math.round(FS * 1.55);
  const font = `${FS}px "Leelawadee UI","Thonburi","Noto Sans Thai","Segoe UI",Tahoma,sans-serif`;
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const g = c.getContext('2d'); g.font = font;
  const seg = window.Intl && Intl.Segmenter ? new Intl.Segmenter('th', { granularity: 'word' }) : null;
  const words = (s) => seg ? [...seg.segment(s)].map((x) => x.segment) : s.split(/(\s+)/);
  const lines = [];
  for (const para of text.replace(/\r\n?/g, '\n').split('\n')) {
    let line = '';
    for (const w of words(para.replace(/\t/g, '    '))) {
      if (g.measureText(line + w).width <= W - 2 * M) { line += w; continue; }
      if (line) lines.push(line);
      line = w.trimStart();
      while (g.measureText(line).width > W - 2 * M) { // very long token: hard break
        let k = line.length; while (k > 1 && g.measureText(line.slice(0, k)).width > W - 2 * M) k--;
        lines.push(line.slice(0, k)); line = line.slice(k);
      }
    }
    lines.push(line);
  }
  const perPage = Math.floor((H - 2 * M) / LH);
  const doc = await PDFDocument.create();
  for (let i = 0; i < lines.length; i += perPage) {
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    g.fillStyle = '#111'; g.font = font; g.textBaseline = 'top';
    lines.slice(i, i + perPage).forEach((l, k) => g.fillText(l, M, M + k * LH));
    const img = await doc.embedPng(await canvasToBytes(c, 'image/png'));
    doc.addPage(A4).drawImage(img, { x: 0, y: 0, width: A4[0], height: A4[1] });
  }
  c.width = c.height = 0;
  return doc.save();
}
function menuCopy(cut) {
  if (!D.selected.size) return;
  pendingCut = cut; copyHandled = false;
  try { document.execCommand('copy'); } catch {}
  if (!copyHandled) { // copy event unavailable -> do it directly
    doCopy(cut); pendingCut = false;
    try { navigator.clipboard.writeText(clip.marker); } catch {}
  }
}
$('btnCopy').onclick = () => menuCopy(false);
$('btnCut').onclick = () => menuCopy(true);
$('btnPaste').onclick = () => pasteAny();

// ------------------------------------------------------------------ context menus
let ctxEl = null;
function closeMenu() { if (ctxEl) { ctxEl.remove(); ctxEl = null; } }
function buildMenu(items) {
  const m = document.createElement('div'); m.className = 'ctx';
  for (const it of items) {
    if (it.sep) { const s = document.createElement('div'); s.className = 'ctxsep'; m.appendChild(s); continue; }
    if (it.head) { const h = document.createElement('div'); h.className = 'ctxhead'; h.textContent = it.head; m.appendChild(h); continue; }
    const el = document.createElement('div');
    el.className = 'ctxitem' + (it.disabled ? ' disabled' : '') + (it.danger ? ' danger' : '');
    el.innerHTML = `${it.icon ? `<svg><use href="#i-${it.icon}"/></svg>` : '<svg></svg>'}<span class="ctxlabel">${esc(it.label)}</span>${it.kbd ? `<span class="kbd">${it.kbd}</span>` : ''}${it.sub ? '<svg><use href="#i-chev"/></svg>' : ''}`;
    if (it.sub) {
      const sm = buildMenu(it.sub); sm.classList.add('ctxsub'); el.appendChild(sm);
      el.addEventListener('click', (e) => { if (e.target.closest('.ctxsub')) return; e.stopPropagation(); el.parentElement.querySelectorAll(':scope>.ctxitem.open').forEach((o) => { if (o !== el) o.classList.remove('open'); }); el.classList.toggle('open'); });
    }
    else el.addEventListener('click', (e) => { e.stopPropagation(); closeMenu(); it.action && it.action(); });
    m.appendChild(el);
  }
  return m;
}
function showMenu(items, x, y) {
  closeMenu();
  ctxEl = buildMenu(items);
  ctxEl.style.left = '0px'; ctxEl.style.top = '0px';
  document.body.appendChild(ctxEl);
  const w = ctxEl.offsetWidth, h = ctxEl.offsetHeight;
  ctxEl.style.left = Math.max(4, Math.min(x, innerWidth - w - 4)) + 'px';
  ctxEl.style.top = Math.max(4, Math.min(y, innerHeight - h - 4)) + 'px';
  if (x + w * 2 > innerWidth) ctxEl.querySelectorAll('.ctxsub').forEach((s) => s.classList.add('left'));
  ctxEl.querySelectorAll('.ctxsub').forEach((s) => { // keep submenus on screen vertically
    s.parentElement.addEventListener('mouseenter', () => {
      s.style.top = '-5px';
      const r = s.getBoundingClientRect();
      if (r.bottom > innerHeight - 4) s.style.top = (-5 - (r.bottom - innerHeight + 8)) + 'px';
    });
  });
}
window.addEventListener('mousedown', (e) => { if (ctxEl && !ctxEl.contains(e.target)) closeMenu(); }, true);
window.addEventListener('blur', closeMenu);
window.addEventListener('resize', closeMenu);
$('organize').addEventListener('scroll', closeMenu, { passive: true });

const n = () => D.selected.size;
const splitMenu = () => [
  { label: 'แยกไปแท็บใหม่ (คัดลอก)', icon: 'tab', disabled: !n(), action: () => splitToTab(false) },
  { label: 'ย้ายไปแท็บใหม่ (ตัดออก)', icon: 'cut', disabled: !n(), action: () => splitToTab(true) },
  { sep: true },
  { label: 'บันทึกเป็น PDF ไฟล์เดียว…', icon: 'extract', disabled: !n(), action: () => splitToFile(false) },
  { label: 'บันทึกแยกทีละหน้า (1 หน้า = 1 ไฟล์)…', icon: 'split', disabled: !n(), action: () => splitToFile(true) },
  { label: 'แยกทุก ๆ N หน้า…', icon: 'split', disabled: !D.pages.length, action: () => splitEvery() },
];
const convertMenu = () => [
  { head: n() ? `แปลง ${n()} หน้าที่เลือก` : 'แปลงทุกหน้า' },
  { label: 'รูปภาพ JPG…', icon: 'img', action: () => convertImages('jpg') },
  { label: 'รูปภาพ PNG…', icon: 'img', action: () => convertImages('png') },
  { label: 'TIFF (1 หน้า = 1 ไฟล์)…', icon: 'img', action: () => convertTiff(false) },
  { label: 'TIFF หลายหน้า (.tiff / .mtiff)…', icon: 'img', action: () => convertTiff(true) },
  { label: 'Word (.docx)…', icon: 'word', disabled: !host.host || !host.word, action: () => convertWord() },
  { label: 'ข้อความ (.txt)…', icon: 'text', action: () => convertText() },
];
const toTabMenu = () => [
  ...docs.filter((d) => d !== D).map((d) => ({ label: d.name, icon: 'tab', action: () => copyPagesToDoc(selPages(), d) })),
  ...(docs.length > 1 ? [{ sep: true }] : []),
  { label: 'แท็บใหม่', icon: 'plus', action: () => splitToTab(false) },
];

function pageMenu(idx) {
  const hasClip = !!clip;
  return [
    { label: 'เพิ่มข้อความในหน้านี้', icon: 'text', action: () => { setMode('read'); setTimeout(() => { goToPage(idx + 1); setAnnTool(true); }, 80); } },
    { label: n() > 1 ? `เปิดอ่านหน้า ${idx + 1}` : 'เปิดอ่านหน้านี้', icon: 'read', action: () => { setMode('read'); setTimeout(() => goToPage(idx + 1), 60); } },
    { sep: true },
    { label: `คัดลอก${n() > 1 ? ` (${n()} หน้า)` : ''}`, icon: 'copy', kbd: 'Ctrl+C', action: () => menuCopy(false) },
    { label: `ตัด${n() > 1 ? ` (${n()} หน้า)` : ''}`, icon: 'cut', kbd: 'Ctrl+X', action: () => menuCopy(true) },
    { label: 'วางก่อนหน้านี้', icon: 'paste', action: () => pasteAny(idx) },
    { label: 'วางหลังหน้านี้', icon: 'paste', kbd: 'Ctrl+V', action: () => pasteAny(idx + 1) },
    { label: 'ทำสำเนาหน้า', icon: 'dup', kbd: 'Ctrl+D', action: duplicatePages },
    { sep: true },
    { label: 'หมุนซ้าย 90°', icon: 'rl', action: () => rotatePages([...D.selected], -90) },
    { label: 'หมุนขวา 90°', icon: 'rr', action: () => rotatePages([...D.selected], 90) },
    { label: 'หมุน 180°', icon: 'rr', action: () => rotatePages([...D.selected], 180) },
    { sep: true },
    { label: 'แทรก', icon: 'plus', sub: [
      { label: 'แทรกไฟล์ก่อนหน้านี้…', icon: 'add', action: () => pickInsert(idx) },
      { label: 'แทรกไฟล์หลังหน้านี้…', icon: 'add', action: () => pickInsert(idx + 1) },
      { sep: true },
      { label: 'หน้าว่างก่อนหน้านี้', icon: 'blank', action: () => addBlankPage(idx) },
      { label: 'หน้าว่างหลังหน้านี้', icon: 'blank', action: () => addBlankPage(idx + 1) },
    ] },
    { label: 'แยกหน้าที่เลือก (Split)', icon: 'split', sub: splitMenu() },
    { label: 'แปลงหน้าที่เลือก', icon: 'convert', sub: convertMenu() },
    { label: 'คัดลอกไปยังแท็บ', icon: 'tab', sub: toTabMenu() },
    { sep: true },
    { label: 'เลือกทั้งหมด', icon: 'selall', kbd: 'Ctrl+A', action: selectAll },
    { label: `ลบ${n() > 1 ? ` (${n()} หน้า)` : 'หน้านี้'}`, icon: 'trash', kbd: 'Delete', danger: true, action: () => deletePages(selPages().map((p) => p.uid)) },
  ];
}
function emptyAreaMenu() {
  return [
    { label: 'วางท้ายเอกสาร', icon: 'paste', kbd: 'Ctrl+V', action: () => pasteAny(D.pages.length) },
    { label: 'เพิ่มไฟล์ท้ายเอกสาร…', icon: 'add', action: () => pickInsert(D.pages.length) },
    { label: 'หน้าว่างท้ายเอกสาร', icon: 'blank', action: () => addBlankPage(D.pages.length) },
    { sep: true },
    { label: 'เลือกทั้งหมด', icon: 'selall', kbd: 'Ctrl+A', action: selectAll },
    { label: 'แยกไฟล์', icon: 'split', sub: splitMenu() },
    { label: 'แปลงไฟล์', icon: 'convert', sub: convertMenu() },
    { sep: true },
    { label: 'บันทึก…', icon: 'save', kbd: 'Ctrl+S', action: () => saveAll() },
    { label: 'บันทึกเป็นไฟล์ใหม่…', icon: 'extract', kbd: 'Ctrl+Shift+S', action: () => saveAll(true) },
  ];
}
function selectAll() { D.selected = new Set(D.pages.map((p) => p.uid)); paintSelection(); }

// block the browser's own right-click menu everywhere; show ours where it makes sense
document.addEventListener('contextmenu', (e) => {
  if (e.target.closest('input[type=text], input[type=password], input[type=number], textarea, .annot.editing, .pastezone')) return; // keep copy/paste for text boxes
  e.preventDefault();
  if (!D.pages.length) { if (e.target.closest('#empty')) showMenu([{ label: 'วางจากคลิปบอร์ด', icon: 'paste', kbd: 'Ctrl+V', action: () => pasteAny(0) }, { label: 'เปิดไฟล์…', icon: 'open', kbd: 'Ctrl+O', action: doOpen }], e.clientX, e.clientY); return; }
  const annEl2 = e.target.closest('.annot');
  if (annEl2 && !annEl2.classList.contains('editing')) {
    const L = annEl2.closest('.annlayer'); const p = D.pages[+L.dataset.i];
    selectAnn({ uid: p.uid, id: annEl2.dataset.id });
    showMenu([
      { label: 'แก้ไขข้อความ', icon: 'text', kbd: 'Enter', action: startEdit },
      { sep: true },
      { label: 'คัดลอก', icon: 'copy', kbd: 'Ctrl+C', action: () => copyAnn(false) },
      { label: 'ตัด', icon: 'cut', kbd: 'Ctrl+X', action: () => copyAnn(true) },
      { label: 'วาง', icon: 'paste', kbd: 'Ctrl+V', disabled: !annClip, action: () => pasteAnn() },
      { label: 'ทำสำเนา', icon: 'dup', kbd: 'Ctrl+D', action: () => { copyAnn(false); pasteAnn(); } },
      { sep: true },
      { label: 'ตัวหนา', icon: 'text', action: () => applyAnnStyle({ bold: !findAnn(annSel).bold }) },
      { label: 'ตัวเอียง', icon: 'text', action: () => applyAnnStyle({ italic: !findAnn(annSel).italic }) },
      { sep: true },
      { label: 'ลบกล่องข้อความ', icon: 'trash', kbd: 'Delete', danger: true, action: deleteAnn },
    ], e.clientX, e.clientY);
    return;
  }
  if (e.target.closest('.annot.editing')) return;
  const card = e.target.closest('.card');
  const rpage = e.target.closest('.rpage');
  if (card || rpage) {
    const idx = card ? D.pages.findIndex((p) => p.uid === card.dataset.uid) : +rpage.dataset.i;
    const u = D.pages[idx].uid;
    if (!D.selected.has(u)) { D.selected = new Set([u]); D.lastClicked = u; paintSelection(); }
    showMenu(pageMenu(idx), e.clientX, e.clientY);
  } else if (e.target.closest('#organize, #reader')) {
    showMenu(emptyAreaMenu(), e.clientX, e.clientY);
  }
});
$('btnSplit').onclick = (e) => { const r = e.currentTarget.getBoundingClientRect(); showMenu(splitMenu(), r.left, r.bottom + 4); };
$('btnConvert').onclick = (e) => { const r = e.currentTarget.getBoundingClientRect(); showMenu(convertMenu(), r.left, r.bottom + 4); };

// ------------------------------------------------------------------ split
function splitToTab(move) {
  const items = selPages(); if (!items.length) return;
  const d = newDoc(D.name + (move ? ' (ย้าย)' : ' (แยก)'), items.map((p) => clonePage(p)));
  d.dirty = true;
  if (move) {
    pushUndo();
    D.pages = D.pages.filter((p) => !D.selected.has(p.uid)); D.selected.clear(); D.dirty = true;
  }
  switchTo(d);
  toast(`${move ? 'ย้าย' : 'แยก'} ${items.length} หน้าไปแท็บใหม่แล้ว`, 'ok', [], 3000);
}
async function splitToFile(each) {
  const items = selPages(); if (!items.length) return;
  if (!each) return saveList(items, D.name + '_หน้าที่เลือก');
  const frames = [];
  const idx = new Map(D.pages.map((p, i) => [p.uid, i + 1]));
  try {
    for (let k = 0; k < items.length; k++) {
      busy(`กำลังแยกไฟล์ … ${k + 1}/${items.length}`);
      const r = await buildPdfInner([items[k]], true);
      frames.push({ name: `หน้า${pad(idx.get(items[k].uid))}.pdf`, bytes: r.bytes });
    }
  } catch (e) { unbusy(); toast('แยกไฟล์ไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  await saveMany(frames, D.name, 'pdf');
}
async function splitEvery() {
  const r = await dialog({
    title: 'แยกเอกสารทุก ๆ N หน้า',
    body: `<div class="field"><label>จำนวนหน้าต่อไฟล์</label><input name="n" type="text" inputmode="numeric" value="1"></div><p>เอกสาร ${D.pages.length} หน้า จะถูกแบ่งเป็นหลายไฟล์ตามลำดับ</p>`,
    focus: 'input',
  });
  if (!r.btn) return;
  const size = Math.max(1, parseInt(r.values.n, 10) || 1);
  const frames = [];
  try {
    for (let s = 0, k = 1; s < D.pages.length; s += size, k++) {
      busy(`กำลังแยกไฟล์ … ส่วนที่ ${k}`);
      const part = D.pages.slice(s, s + size);
      const res = await buildPdfInner(part, true);
      const label = part.length > 1 ? `หน้า${pad(s + 1)}-${pad(s + part.length)}` : `หน้า${pad(s + 1)}`;
      frames.push({ name: label + '.pdf', bytes: res.bytes });
    }
  } catch (e) { unbusy(); toast('แยกไฟล์ไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  await saveMany(frames, D.name, 'pdf');
}

// ------------------------------------------------------------------ convert
const convList = () => (D.selected.size ? selPages() : D.pages);
async function convertImages(kind) {
  const list = convList(); if (!list.length) return;
  const r = await dialog({
    title: `แปลงเป็นรูปภาพ ${kind.toUpperCase()}`,
    body: `<div class="field"><label>ความละเอียด</label><select name="dpi">
      <option value="96">96 dpi — ไฟล์เล็ก (ดูบนจอ)</option><option value="150" selected>150 dpi — มาตรฐาน</option>
      <option value="200">200 dpi — ชัด</option><option value="300">300 dpi — คมมาก (สำหรับพิมพ์)</option></select></div>
      <p>จะได้ ${list.length} ไฟล์ (1 หน้า = 1 รูป)</p>`,
  });
  if (!r.btn) return;
  const dpi = +r.values.dpi;
  const frames = [];
  const idx = new Map(D.pages.map((p, i) => [p.uid, i + 1]));
  try {
    for (let k = 0; k < list.length; k++) {
      busy(`กำลังแปลงเป็นรูปภาพ … ${k + 1}/${list.length}`);
      const c = await renderPageCanvas(list[k], dpi / 72);
      frames.push({ name: `หน้า${pad(idx.get(list[k].uid))}.${kind}`, bytes: await canvasToBytes(c, kind === 'jpg' ? 'image/jpeg' : 'image/png', 0.9) });
      c.width = c.height = 0;
    }
  } catch (e) { unbusy(); toast('แปลงไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  if (frames.length === 1) await saveBytes(frames[0].bytes, `${D.name}_${frames[0].name}`, kind);
  else await saveMany(frames, D.name, kind);
}
async function convertTiff(multi) {
  const list = convList(); if (!list.length) return;
  const r = await dialog({
    title: multi ? 'แปลงเป็น TIFF หลายหน้า (ไฟล์เดียว)' : 'แปลงเป็น TIFF (1 หน้า = 1 ไฟล์)',
    body: `<div class="row2">
      <div class="field"><label>ความละเอียด</label><select name="dpi">
        <option value="150">150 dpi</option><option value="200" selected>200 dpi (แนะนำ)</option><option value="300">300 dpi</option></select></div>
      <div class="field"><label>สี</label><select name="mode">
        <option value="bw" selected>ขาวดำ (CCITT G4) — ไฟล์เล็กที่สุด</option><option value="gray">ระดับเทา (LZW)</option><option value="rgb">สี (LZW)</option></select></div>
      </div>
      ${multi ? `<div class="field"><label>นามสกุลไฟล์</label><select name="ext">
        <option value="tiff">.tiff</option><option value="tif">.tif</option><option value="mtiff">.mtiff</option></select></div>` : ''}
      <p>${list.length} หน้า${multi ? ' → 1 ไฟล์' : ` → ${list.length} ไฟล์`} &nbsp;·&nbsp; ขาวดำ G4 เหมาะกับเอกสารข้อความ/ระบบสแกนเอกสาร</p>`,
  });
  if (!r.btn) return;
  const dpi = +r.values.dpi, cmode = r.values.mode, extn = multi ? r.values.ext : 'tif';
  const idx = new Map(D.pages.map((p, i) => [p.uid, i + 1]));
  const pagesOut = [];
  try {
    for (let k = 0; k < list.length; k++) {
      busy(`กำลังแปลงเป็น TIFF … หน้า ${k + 1}/${list.length}`);
      await new Promise((res) => setTimeout(res, 0));
      const c = await renderPageCanvas(list[k], dpi / 72);
      pagesOut.push({ n: idx.get(list[k].uid), page: TiffEnc.fromCanvas(c, cmode, dpi, 160) });
      c.width = c.height = 0;
    }
    busy('กำลังเขียนไฟล์ TIFF …');
  } catch (e) { unbusy(); toast('แปลง TIFF ไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  if (multi) {
    await saveBytes(TiffEnc.encode(pagesOut.map((x) => x.page)), `${D.name}.${extn}`, extn);
  } else if (pagesOut.length === 1) {
    await saveBytes(TiffEnc.encode([pagesOut[0].page]), `${D.name}_หน้า${pad(pagesOut[0].n)}.tif`, 'tif');
  } else {
    await saveMany(pagesOut.map((x) => ({ name: `หน้า${pad(x.n)}.tif`, bytes: TiffEnc.encode([x.page]) })), D.name, 'tif');
  }
}

async function convertWord() {
  const list = convList(); if (!list.length) return;
  if (!host.host) { toast('การแปลงเป็น Word ต้องเปิดโปรแกรมผ่าน SuperPDF.exe และมี Microsoft Word', 'err'); return; }
  try {
    const { bytes } = await buildPdfInner(list, true);
    busy('กำลังแปลงเป็น Word ด้วย Microsoft Word … (อาจใช้เวลาสักครู่)');
    const res = await api('/api/pdf2docx', { method: 'POST', body: bytes });
    const docx = new Uint8Array(await res.arrayBuffer());
    unbusy();
    await saveBytes(docx, D.name + '.docx', 'docx');
  } catch (e) { unbusy(); toast('แปลงเป็น Word ไม่สำเร็จ: ' + e.message, 'err', [], 9000); }
}
async function convertText() {
  const list = convList(); if (!list.length) return;
  let out = '';
  try {
    for (let k = 0; k < list.length; k++) {
      busy(`กำลังดึงข้อความ … ${k + 1}/${list.length}`);
      const pg = await sources.get(list[k].src).pdf.getPage(list[k].index + 1);
      const tc = await pg.getTextContent();
      let t = '';
      for (const it of tc.items) { t += it.str; if (it.hasEOL) t += '\r\n'; }
      out += `===== หน้า ${D.pages.indexOf(list[k]) + 1} =====\r\n${t.trim()}\r\n\r\n`;
    }
  } finally { unbusy(); }
  if (!out.replace(/=====.*=====|\s/g, '')) toast('ไม่พบข้อความในหน้าเหล่านี้ (อาจเป็นเอกสารสแกน/รูปภาพ)', '', [], 6000);
  const bytes = new TextEncoder().encode('﻿' + out);
  await saveBytes(bytes, D.name + '.txt', 'txt');
}

// ------------------------------------------------------------------ reader view
let zoom = 0;
let readerKey = '';
const readerObserver = new IntersectionObserver((entries) => {
  for (const en of entries) if (en.isIntersecting) paintReaderPage(en.target);
}, { root: $('reader'), rootMargin: '600px 0px' });

async function pageSize(p) {
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const v = pg.getViewport({ scale: 1, rotation: (pg.rotate + p.rot) % 360 });
  const b = pg.getViewport({ scale: 1, rotation: pg.rotate });
  return [v.width, v.height, b.width, b.height];
}
function currentScale(maxW) {
  if (zoom) return zoom;
  const avail = Math.min($('reader').clientWidth - (innerWidth < 760 ? 16 : 90), 980);
  return Math.max(0.25, Math.min(3, avail / maxW));
}
async function renderReader(force = false) {
  const key = D.id + '#' + D.pages.map(rkey).join('|') + '#' + zoom + '#' + $('reader').clientWidth;
  if (!force && key === readerKey) return;
  readerKey = key;
  const wrap = $('readerPages');
  const pages = D.pages;
  const sizes = await Promise.all(pages.map(pageSize));
  const ws = sizes.map((s) => s[0]).sort((a, b) => a - b);
  const scale = currentScale(ws[Math.floor(ws.length / 2)]);
  $('zoomLabel').textContent = Math.round(scale * 100) + '%';
  readerObserver.disconnect();
  const els = pages.map((p, i) => {
    const d = document.createElement('div');
    d.className = 'rpage';
    d.dataset.i = i;
    d.style.width = Math.floor(sizes[i][0] * scale) + 'px';
    d.style.height = Math.floor(sizes[i][1] * scale) + 'px';
    d.dataset.scale = scale;
    d.innerHTML = `<span class="rnum">${i + 1}</span>`;
    const L = document.createElement('div');
    L.className = 'annlayer';
    const bw = sizes[i][2] * scale, bh = sizes[i][3] * scale;
    Object.assign(L.style, { width: bw + 'px', height: bh + 'px', left: (sizes[i][0] * scale - bw) / 2 + 'px', top: (sizes[i][1] * scale - bh) / 2 + 'px', transform: `rotate(${p.rot}deg)` });
    L.dataset.i = i; L.dataset.scale = scale;
    d.appendChild(L);
    readerObserver.observe(d);
    return d;
  });
  wrap.replaceChildren(...els);
  renderAllAnnLayers();
}
async function paintReaderPage(d) {
  if (d.dataset.painted) return;
  d.dataset.painted = '1';
  const p = D.pages[+d.dataset.i]; if (!p) return;
  const c = await renderPageCanvas(p, +d.dataset.scale * (window.devicePixelRatio || 1), false);
  d.insertBefore(c, d.firstChild);
}
async function renderPageCanvas(p, scale, withAnnots = true) {
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const vp = pg.getViewport({ scale, rotation: (pg.rotate + p.rot) % 360 });
  const c = document.createElement('canvas'); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await pg.render({ canvasContext: g, viewport: vp }).promise;
  if (withAnnots) drawAnnots(g, p, pg, scale);
  return c;
}
function goToPage(n) { const el = $('readerPages').children[n - 1]; if (el) el.scrollIntoView({ block: 'start' }); }

$('reader').addEventListener('scroll', () => {
  closeMenu();
  const r = $('reader'), mid = r.scrollTop + 40;
  const kids = $('readerPages').children;
  let cur = 1;
  for (let i = 0; i < kids.length; i++) { if (kids[i].offsetTop <= mid) cur = i + 1; else break; }
  if (document.activeElement !== $('pageInput')) $('pageInput').value = cur;
}, { passive: true });

function setZoom(z) {
  const r = $('reader');
  const ratio = r.scrollTop / Math.max(1, r.scrollHeight);
  zoom = z;
  renderReader(true).then(() => { r.scrollTop = ratio * r.scrollHeight; });
}
function zoomStep(dir) {
  const cur = parseFloat($('zoomLabel').textContent) / 100 || 1;
  const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
  const z = dir > 0 ? steps.find((s) => s > cur + 0.001) : [...steps].reverse().find((s) => s < cur - 0.001);
  setZoom(z || cur);
}
function setMode(m) {
  if (m !== 'read') { if (annEditing) document.activeElement.blur(); annTool = false; annSel = null; }
  mode = m;
  $('viewOrganize').classList.toggle('active', m === 'organize');
  $('viewRead').classList.toggle('active', m === 'read');
  refresh();
  if (m === 'read' && D.selected.size) {
    const i = D.pages.findIndex((p) => D.selected.has(p.uid));
    setTimeout(() => goToPage(i + 1), 60);
  }
}

// ------------------------------------------------------------------ text annotations (เพิ่มข้อความ)
// annotation: { id, x, y (pt, top-left, in the page's own orientation), text, size(pt), font, color, bold, italic }
const ANN_LH = 1.35, ANN_PAD = 0.2;
const annId = () => 'a' + (seq++);
let annTool = false;              // "add text" tool active
let annSel = null;                // { uid, id } selected annotation
let annEditing = false;
let annClip = null;               // copied annotations
let annDefaults = { font: 'Tahoma', size: 16, color: '#1f3a93', bold: false, italic: false };
const annFont = (a, px) => `${a.italic ? 'italic ' : ''}${a.bold ? '700' : '400'} ${px}px "${a.font}", "Leelawadee UI", Tahoma, sans-serif`;
const mctx = document.createElement('canvas').getContext('2d');
function annMetrics(a) { // in pt
  mctx.font = annFont(a, a.size);
  const lines = (a.text || '').split('\n');
  let w = 0; for (const l of lines) w = Math.max(w, mctx.measureText(l).width);
  const m = mctx.measureText('กÅgjปฐ');
  const asc = m.fontBoundingBoxAscent ?? a.size * 0.9, desc = m.fontBoundingBoxDescent ?? a.size * 0.25;
  const pad = a.size * ANN_PAD, lh = a.size * ANN_LH;
  return { lines, pad, lh, asc, desc, w: w + 2 * pad, h: lines.length * lh + 2 * pad, half: (lh - asc - desc) / 2 };
}
// draw annotation text onto a 2D context already scaled so that 1 unit = 1 pt of base space
function paintAnn(g, a) {
  const m = annMetrics(a);
  g.save(); g.translate(a.x, a.y); if (a.r) g.rotate(a.r * Math.PI / 180);
  g.font = annFont(a, a.size); g.fillStyle = a.color; g.textBaseline = 'alphabetic';
  m.lines.forEach((l, i) => g.fillText(l, m.pad, m.pad + i * m.lh + m.half + m.asc));
  g.restore();
  return m;
}
// box-local point (u,v) -> base-space point, honouring the box rotation a.r (clockwise, degrees)
function annLocal(a, u, v) {
  const t = (a.r || 0) * Math.PI / 180;
  return [a.x + u * Math.cos(t) - v * Math.sin(t), a.y + u * Math.sin(t) + v * Math.cos(t)];
}
// draw all annotations of page p onto a canvas rendered with rotation (pg.rotate + p.rot) at `scale`
function drawAnnots(g, p, pg, scale) {
  if (!p.annots || !p.annots.length) return;
  const b = pg.getViewport({ scale: 1, rotation: pg.rotate });
  const W = b.width * scale, H = b.height * scale;
  g.save();
  if (p.rot === 90) { g.translate(H, 0); g.rotate(Math.PI / 2); }
  else if (p.rot === 180) { g.translate(W, H); g.rotate(Math.PI); }
  else if (p.rot === 270) { g.translate(0, W); g.rotate(-Math.PI / 2); }
  g.scale(scale, scale);
  for (const a of p.annots) paintAnn(g, a);
  g.restore();
}
// put annotations into a copied pdf-lib page: crisp image (looks exactly like on screen) + invisible text (searchable)
async function embedAnnots(out, pdfPage, p) {
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const vb = pg.getViewport({ scale: 1, rotation: pg.rotate });
  const R = pg.rotate || 0;
  for (const a of p.annots) {
    if (!a.text || !a.text.trim()) continue;
    const m = annMetrics(a), K = 4;
    const c = document.createElement('canvas'); c.width = Math.ceil(m.w * K); c.height = Math.ceil(m.h * K);
    const g = c.getContext('2d'); g.scale(K, K); paintAnn(g, { ...a, x: 0, y: 0, r: 0 });
    const img = await out.embedPng(await canvasToBytes(c, 'image/png'));
    c.width = c.height = 0;
    const rot = ((R - (a.r || 0)) % 360 + 360) % 360;
    const [px, py] = vb.convertToPdfPoint(...annLocal(a, 0, m.h));
    pdfPage.drawImage(img, { x: px, y: py, width: m.w, height: m.h, rotate: degrees(rot) });
    const font = await annPdfFont(out, a.font);
    if (font) {
      m.lines.forEach((l, i) => {
        if (!l.trim()) return;
        const [tx, ty] = vb.convertToPdfPoint(...annLocal(a, m.pad, m.pad + i * m.lh + m.half + m.asc));
        try { pdfPage.drawText(l, { x: tx, y: ty, size: a.size, font, opacity: 0, rotate: degrees(((R - (a.r || 0)) % 360 + 360) % 360) }); } catch {}
      });
    }
  }
}
const fontBytesCache = new Map();
async function annPdfFont(out, family) { // real font file from Windows (for the invisible, searchable text layer)
  if (!host.host || !window.fontkit) return null;
  if (out.__fonts.has(family)) return out.__fonts.get(family);
  let f = null;
  try {
    if (!fontBytesCache.has(family)) {
      const r = await fetch('/api/fontfile?family=' + encodeURIComponent(family), { headers: { 'X-Token': TOKEN } });
      fontBytesCache.set(family, r.ok ? new Uint8Array(await r.arrayBuffer()) : null);
    }
    const bytes = fontBytesCache.get(family);
    if (bytes) { out.registerFontkit(window.fontkit); f = await out.embedFont(bytes, { subset: true }); }
  } catch { f = null; }
  out.__fonts.set(family, f);
  return f;
}

// ---- reader overlay (live, editable)
const pageByUid = (u) => D.pages.find((p) => p.uid === u);
const findAnn = (sel) => { const p = sel && pageByUid(sel.uid); return p && (p.annots || []).find((a) => a.id === sel.id); };
function touchPage(p) { p.av = (p.av || 0) + 1; D.dirty = true; }
function renderAllAnnLayers() { document.querySelectorAll('.annlayer').forEach(renderAnnLayer); }
function renderAnnLayer(L) {
  const p = D.pages[+L.dataset.i]; if (!p) return;
  const s = +L.dataset.scale;
  if (annEditing && annSel && annSel.uid === p.uid && L.querySelector('.annot.editing')) return; // don't disturb typing
  L.classList.toggle('tool', annTool);
  L.replaceChildren(...(p.annots || []).map((a) => {
    const el = document.createElement('div');
    el.className = 'annot' + (annSel && annSel.id === a.id ? ' selected' : '');
    el.dataset.id = a.id;
    Object.assign(el.style, {
      left: a.x * s + 'px', top: a.y * s + 'px', padding: a.size * ANN_PAD * s + 'px',
      font: annFont(a, a.size * s), lineHeight: ANN_LH, color: a.color,
      transform: a.r ? `rotate(${a.r}deg)` : '', transformOrigin: '0 0',
    });
    el.textContent = a.text;
    return el;
  }));
}
function selectAnn(sel, edit = false) {
  annSel = sel; annEditing = false;
  renderAllAnnLayers(); updateAnnBar();
  if (edit && sel) startEdit();
}
function annEl(sel) { return sel && document.querySelector(`.annlayer .annot[data-id="${sel.id}"]`); }
function startEdit() {
  const el = annEl(annSel); if (!el) return;
  const a = findAnn(annSel);
  annEditing = true;
  el.classList.add('editing');
  el.contentEditable = 'plaintext-only';
  if (el.contentEditable !== 'plaintext-only') el.contentEditable = 'true';
  el.focus();
  const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
  const sl = getSelection(); sl.removeAllRanges(); sl.addRange(r);
  const before = a.text;
  el.oninput = () => { /* live */ };
  el.onblur = () => finishEdit(el, before);
  el.onkeydown = (e) => { if (e.key === 'Escape') { e.preventDefault(); el.blur(); } e.stopPropagation(); };
}
function finishEdit(el, before) {
  if (!annEditing) return;
  annEditing = false;
  el.onblur = el.onkeydown = el.oninput = null;
  el.contentEditable = 'false';
  const p = pageByUid(annSel?.uid), a = findAnn(annSel);
  if (!p || !a) return;
  const text = el.innerText.replace(/\n$/, '');
  if (text !== before || !text.trim()) {
    if (before !== '') pushUndo(); // a brand-new box already has its undo step
    if (!text.trim()) { p.annots = p.annots.filter((x) => x.id !== a.id); annSel = null; }
    else a.text = text;
    touchPage(p);
  }
  refresh();
}
function newAnnAt(p, x, y, text = '') {
  pushUndo();
  const a = { id: annId(), x, y, text, ...annDefaults };
  p.annots = [...(p.annots || []), a];
  touchPage(p);
  return a;
}
function annPoint(L, e) { // pointer -> pt in base space (handles page rotation)
  const s = +L.dataset.scale, r = L.getBoundingClientRect();
  const p = D.pages[+L.dataset.i];
  const cx = e.clientX - (r.left + r.width / 2), cy = e.clientY - (r.top + r.height / 2);
  const t = -(p.rot || 0) * Math.PI / 180;
  const ux = cx * Math.cos(t) - cy * Math.sin(t), uy = cx * Math.sin(t) + cy * Math.cos(t);
  return [(ux + L.offsetWidth / 2) / s, (uy + L.offsetHeight / 2) / s];
}
$('readerPages').addEventListener('pointerdown', (e) => {
  const L = e.target.closest('.annlayer');
  if (!L || e.button !== 0) return;
  const p = D.pages[+L.dataset.i];
  const el = e.target.closest('.annot');
  if (el && el.classList.contains('editing')) return;            // let the caret move
  if (!el) {
    if (annEditing) { document.activeElement.blur(); }
    if (annTool) {
      e.preventDefault();
      const [x, y] = annPoint(L, e);
      // on a rotated page the new box is counter-rotated so it reads upright on screen
      const r = (360 - (p.rot || 0)) % 360;
      const probe = { x, y, r };
      const [ax, ay] = annLocal(probe, -annDefaults.size * ANN_PAD, -annDefaults.size * ANN_LH / 2);
      const a = newAnnAt(p, ax, ay);
      a.r = r;
      annTool = false; // one box per click of the tool (like Acrobat "Add text")
      selectAnn({ uid: p.uid, id: a.id }, true);
    } else if (annSel) selectAnn(null);
    return;
  }
  e.preventDefault();
  const sel = { uid: p.uid, id: el.dataset.id };
  const wasSel = annSel && annSel.id === sel.id;
  if (!wasSel) selectAnn(sel);
  const a = findAnn(sel), s = +L.dataset.scale;
  const [sx, sy] = annPoint(L, e), ox = a.x, oy = a.y;
  let moved = false;
  const box = annEl(sel);
  const move = (ev) => {
    const [mx, my] = annPoint(L, ev);
    if (!moved && Math.hypot(mx - sx, my - sy) * s < 3) return;
    if (!moved) { moved = true; pushUndo(); }
    a.x = ox + (mx - sx); a.y = oy + (my - sy);
    box.style.left = a.x * s + 'px'; box.style.top = a.y * s + 'px';
  };
  const up = () => {
    window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
    if (moved) { touchPage(p); refresh(); }
    else if (wasSel) startEdit();
  };
  window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
});
$('readerPages').addEventListener('dblclick', (e) => { if (e.target.closest('.annot') && !annEditing) startEdit(); });

function deleteAnn() {
  const p = pageByUid(annSel?.uid); if (!p) return;
  pushUndo(); p.annots = p.annots.filter((a) => a.id !== annSel.id); touchPage(p); annSel = null; refresh();
}
function copyAnn(cut) {
  const a = findAnn(annSel); if (!a) return;
  annClip = { ann: { ...a }, marker: 'SuperPDF-text-' + (seq++) };
  try { navigator.clipboard.writeText(annClip.marker); } catch {}
  if (cut) deleteAnn();
  toast(cut ? 'ตัดกล่องข้อความแล้ว' : 'คัดลอกกล่องข้อความแล้ว — กด Ctrl+V เพื่อวาง (วางหน้าอื่น/แท็บอื่นได้)', '', [], 2500);
}
function currentReaderPage() {
  if (annSel && pageByUid(annSel.uid)) return pageByUid(annSel.uid);
  const n = Math.max(1, Math.min(D.pages.length, +$('pageInput').value || 1));
  return D.pages[n - 1];
}
function pasteAnn(textOverride) {
  const p = currentReaderPage(); if (!p) return;
  let a;
  if (textOverride != null) {
    const mid = document.querySelector(`.annlayer[data-i="${D.pages.indexOf(p)}"]`);
    const y = mid ? Math.max(20, ($('reader').scrollTop - mid.parentElement.offsetTop + 120) / +mid.dataset.scale) : 60;
    a = newAnnAt(p, 60, y, textOverride);
    a.r = (360 - (p.rot || 0)) % 360;
  } else {
    pushUndo();
    const src = annClip.ann;
    const same = annSel && pageByUid(annSel.uid) === p;
    a = { ...src, id: annId(), x: src.x + (same ? 12 : 0), y: src.y + (same ? 12 : 0) };
    p.annots = [...(p.annots || []), a]; touchPage(p);
    annClip.ann = { ...a };
  }
  selectAnn({ uid: p.uid, id: a.id });
  refresh();
}
function applyAnnStyle(patch) {
  Object.assign(annDefaults, patch);
  const a = findAnn(annSel);
  if (a) { pushUndo(); Object.assign(a, patch); touchPage(pageByUid(annSel.uid)); annEditing = false; refresh(); }
  updateAnnBar();
}
function setAnnTool(on) { annTool = on; if (on && mode !== 'read') setMode('read'); renderAllAnnLayers(); updateAnnBar(); }

// ---- format bar
function updateAnnBar() {
  const bar = $('annbar'); if (!bar) return;
  bar.classList.toggle('hidden', mode !== 'read' || !D.pages.length);
  const a = findAnn(annSel) || annDefaults;
  $('annTool').classList.toggle('active', annTool);
  if (document.activeElement !== $('annFont')) {
    if (![...$('annFont').options].some((o) => o.value === a.font)) $('annFont').insertAdjacentHTML('afterbegin', `<option>${esc(a.font)}</option>`);
    $('annFont').value = a.font;
  }
  if (document.activeElement !== $('annSize')) $('annSize').value = a.size;
  $('annColor').value = a.color;
  $('annBold').classList.toggle('active', !!a.bold);
  $('annItalic').classList.toggle('active', !!a.italic);
  const has = !!findAnn(annSel);
  for (const id of ['annDel', 'annCopy', 'annCut', 'annDup']) $(id).disabled = !has;
  $('annPaste').disabled = !annClip;
  $('annHint').textContent = annTool ? 'คลิกบนหน้าเพื่อวางข้อความ' : has ? 'ลากเพื่อย้าย · คลิกอีกครั้ง/ดับเบิลคลิกเพื่อแก้ข้อความ' : '';
  $('annUndo').disabled = !D.undo.length; $('annRedo').disabled = !D.redo.length;
}
async function loadFontList() {
  let fams = [];
  if (host.host) { try { fams = await (await api('/api/fonts')).json(); } catch {} }
  const prefer = ['TH Sarabun New', 'TH SarabunPSK', 'Sarabun', 'Tahoma', 'Leelawadee UI', 'Leelawadee', 'Angsana New', 'AngsanaUPC', 'Cordia New', 'Browallia New', 'TH Niramit AS', 'TH Charmonman', 'Arial', 'Times New Roman', 'Calibri', 'Cambria', 'Segoe UI', 'Courier New', 'Thonburi'];
  const set = new Set(fams);
  const top = fams.length ? prefer.filter((f) => set.has(f)) : ['Tahoma', 'Leelawadee UI', 'Thonburi', 'Sarabun', 'Arial', 'Times New Roman', 'Courier New'];
  const rest = fams.filter((f) => !top.includes(f) && !f.startsWith('@'));
  $('annFont').innerHTML = `<optgroup label="แนะนำ">${top.map((f) => `<option style="font-family:'${esc(f)}'">${esc(f)}</option>`).join('')}</optgroup>` +
    (rest.length ? `<optgroup label="ฟอนต์ทั้งหมดในเครื่อง">${rest.map((f) => `<option>${esc(f)}</option>`).join('')}</optgroup>` : '');
  if (top.length && !top.includes(annDefaults.font)) annDefaults.font = top[0];
  if (!top.length && fams.length) annDefaults.font = fams[0];
  if (![...$('annFont').options].some((o) => o.value === annDefaults.font)) $('annFont').insertAdjacentHTML('afterbegin', `<option>${esc(annDefaults.font)}</option>`);
  $('annFont').value = annDefaults.font;
}
$('annTool').onclick = () => setAnnTool(!annTool);
$('annFont').onchange = (e) => applyAnnStyle({ font: e.target.value });
$('annSize').onchange = (e) => applyAnnStyle({ size: Math.max(4, Math.min(300, +e.target.value || 16)) });
$('annSize').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } };
$('annSizeDn').onclick = () => applyAnnStyle({ size: Math.max(4, (findAnn(annSel) || annDefaults).size - 1) });
$('annSizeUp').onclick = () => applyAnnStyle({ size: Math.min(300, (findAnn(annSel) || annDefaults).size + 1) });
$('annColor').oninput = (e) => applyAnnStyle({ color: e.target.value });
document.querySelectorAll('#annbar .swatch').forEach((b) => { b.onclick = () => applyAnnStyle({ color: b.dataset.c }); });
$('annBold').onclick = () => applyAnnStyle({ bold: !(findAnn(annSel) || annDefaults).bold });
$('annItalic').onclick = () => applyAnnStyle({ italic: !(findAnn(annSel) || annDefaults).italic });
$('annDel').onclick = deleteAnn;
$('annCopy').onclick = () => copyAnn(false);
$('annCut').onclick = () => copyAnn(true);
$('annPaste').onclick = () => annClip && pasteAnn();
$('annDup').onclick = () => { if (findAnn(annSel)) { copyAnn(false); pasteAnn(); } };
$('annUndo').onclick = () => { annSel = null; undo(); };
$('annRedo').onclick = () => { annSel = null; redo(); };
document.querySelectorAll('#annbar button').forEach((b) => b.addEventListener('mousedown', (e) => e.preventDefault())); // keep editing focus

// keyboard for selected text boxes (reader mode, not typing)
window.addEventListener('keydown', (e) => {
  if (mode !== 'read' || modalOpen() || typingInField()) return;
  const c = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
  const a = findAnn(annSel);
  if (!c && !e.altKey && k === 't') { e.preventDefault(); setAnnTool(!annTool); return; }
  if (!a) { if (e.key === 'Escape' && annTool) setAnnTool(false); return; }
  let handled = true;
  if (c && k === 'c') copyAnn(false);
  else if (c && k === 'x') copyAnn(true);
  else if (c && k === 'd') { copyAnn(false); pasteAnn(); }
  else if (e.key === 'Delete' || e.key === 'Backspace') deleteAnn();
  else if (e.key === 'Enter' || e.key === 'F2') startEdit();
  else if (e.key === 'Escape') selectAnn(null);
  else if (e.key.startsWith('Arrow')) {
    const d = e.shiftKey ? 10 : 1;
    pushUndo();
    if (e.key === 'ArrowLeft') a.x -= d; if (e.key === 'ArrowRight') a.x += d;
    if (e.key === 'ArrowUp') a.y -= d; if (e.key === 'ArrowDown') a.y += d;
    touchPage(pageByUid(annSel.uid)); refresh();
  } else handled = false;
  if (handled) { e.preventDefault(); e.stopImmediatePropagation(); }
}, true);

// ------------------------------------------------------------------ saving
async function rasterPage(out, p) {
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const v1 = pg.getViewport({ scale: 1, rotation: (pg.rotate + p.rot) % 360 });
  const c = await renderPageCanvas(p, 200 / 72);
  const img = await out.embedJpg(await canvasToBytes(c, 'image/jpeg', 0.88));
  c.width = c.height = 0;
  const np = out.addPage([v1.width, v1.height]);
  np.drawImage(img, { x: 0, y: 0, width: v1.width, height: v1.height });
}

const libCache = new Map(); // src -> PDFDocument|null (parsed once)
async function libDoc(srcId) {
  if (libCache.has(srcId)) return libCache.get(srcId);
  const s = sources.get(srcId);
  let d = null;
  if (!s.password) { try { d = await PDFDocument.load(s.bytes, { updateMetadata: false }); } catch { d = null; } }
  libCache.set(srcId, d);
  return d;
}

async function buildPdf(list) {
  try { return await buildPdfInner(list); } finally { unbusy(); }
}
async function buildPdfInner(list, quiet = false) {
  const out = await PDFDocument.create();
  out.__fonts = new Map();
  const libDocs = new Map();
  for (const p of list) if (!libDocs.has(p.src)) libDocs.set(p.src, await libDoc(p.src));
  const copied = new Map();
  for (const [src, d] of libDocs) {
    if (!d) continue;
    const idx = list.filter((p) => p.src === src).map((p) => p.index);
    const cps = await out.copyPages(d, idx);
    idx.forEach((ix, k) => { const key = src + ':' + ix; if (!copied.has(key)) copied.set(key, []); copied.get(key).push(cps[k]); });
  }
  let rastered = 0;
  for (let i = 0; i < list.length; i++) {
    const p = list[i];
    if (!quiet) busy(`กำลังสร้าง PDF … หน้า ${i + 1}/${list.length}`);
    if (!libDocs.get(p.src)) { await rasterPage(out, p); rastered++; continue; }
    const pg = copied.get(p.src + ':' + p.index).shift();
    out.addPage(pg);
    if (p.annots && p.annots.length) await embedAnnots(out, pg, p);
    if (p.rot) pg.setRotation(degrees((pg.getRotation().angle + p.rot) % 360));
  }
  out.setProducer('Super PDF'); out.setCreator('Super PDF');
  out.setTitle(D.name || 'Document');
  return { bytes: await out.save({ useObjectStreams: true }), rastered };
}

const cleanName = (s) => s.replace(/[\\/:*?"<>|]/g, '_');

// save one file (host dialog, or browser download)
async function saveBytes(bytes, name, extn = 'pdf', onSaved = null) {
  name = cleanName(name);
  if (host.host) {
    try {
      busy('รอเลือกตำแหน่งบันทึกไฟล์ …');
      const dir = D && D.dir ? '&dir=' + encodeURIComponent(D.dir) : '';
      const r = await (await api(`/api/save?ext=${extn}&name=${encodeURIComponent(name)}${dir}`, { method: 'POST', body: bytes })).json();
      unbusy();
      if (r.ok) {
        if (onSaved) onSaved(r.path);
        toast('บันทึกแล้ว: ' + r.path, 'ok', [
          ['เปิดไฟล์', () => api('/api/openpdf', { method: 'POST', body: r.path })],
          ['เปิดโฟลเดอร์', () => api('/api/reveal', { method: 'POST', body: r.path })],
        ], 8000);
        return true;
      }
      if (r.error) toast('บันทึกไม่สำเร็จ: ' + r.error, 'err', [], 9000);
      return false;
    } catch (e) { unbusy(); toast('บันทึกไม่สำเร็จ: ' + e.message, 'err'); return false; }
  }
  return deliver([{ name, bytes }]);
}
const MIME = { pdf: 'application/pdf', jpg: 'image/jpeg', png: 'image/png', tif: 'image/tiff', tiff: 'image/tiff', mtiff: 'image/tiff', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', txt: 'text/plain' };
// browser/iPhone: offer the share sheet (Save to Files, LINE, Mail…) or a plain download
async function deliver(list) {
  const files = list.map((f) => new File([f.bytes], f.name, { type: MIME[ext(f.name)] || 'application/octet-stream' }));
  const canShare = navigator.canShare && navigator.canShare({ files });
  if (!canShare || !(IS_TOUCH || PACKED)) { list.forEach((f) => download(f.bytes, f.name)); return true; }
  const r = await dialog({
    title: 'ไฟล์พร้อมแล้ว',
    body: `<p>${list.length === 1 ? esc(list[0].name) : list.length + ' ไฟล์'}</p><p>กด “บันทึก/แชร์” แล้วเลือก <b>บันทึกไปยังไฟล์ (Save to Files)</b> หรือส่งต่อทาง LINE / อีเมล</p>`,
    buttons: [['cancel', 'ยกเลิก'], ['dl', 'ดาวน์โหลด'], ['share', 'บันทึก / แชร์', true]],
  });
  if (!r.btn) return false;
  if (r.btn === 'dl') { list.forEach((f) => download(f.bytes, f.name)); return true; }
  try { await navigator.share({ files }); return true; }
  catch (e) { if (e.name !== 'AbortError') { toast('แชร์ไม่สำเร็จ ลองกด “ดาวน์โหลด” แทน', 'err'); } return false; }
}
function download(bytes, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes]));
  a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 60000);
}
// save many files at once: user picks folder + base name once
async function saveMany(frames, base, extn) {
  if (!host.host) return deliver(frames.map((f) => ({ name: cleanName(base + '_' + f.name), bytes: f.bytes })));
  const enc = new TextEncoder();
  const parts = [];
  for (const f of frames) {
    const nb = enc.encode(f.name);
    const h1 = new Uint8Array(4); new DataView(h1.buffer).setInt32(0, nb.length, true);
    const h2 = new Uint8Array(4); new DataView(h2.buffer).setInt32(0, f.bytes.length, true);
    parts.push(h1, nb, h2, f.bytes);
  }
  try {
    busy(`รอเลือกโฟลเดอร์สำหรับบันทึก ${frames.length} ไฟล์ …`);
    const dir = D && D.dir ? '&dir=' + encodeURIComponent(D.dir) : '';
    const r = await (await api(`/api/savemany?ext=${extn}&name=${encodeURIComponent(cleanName(base))}${dir}`, { method: 'POST', body: new Blob(parts) })).json();
    unbusy();
    if (r.ok) { toast(`บันทึก ${r.count} ไฟล์แล้วที่ ${r.dir}`, 'ok', [['เปิดโฟลเดอร์', () => api('/api/reveal', { method: 'POST', body: r.dir })]], 8000); return true; }
    if (r.error) toast('บันทึกไม่สำเร็จ: ' + r.error, 'err', [], 9000);
    return false;
  } catch (e) { unbusy(); toast('บันทึกไม่สำเร็จ: ' + e.message, 'err'); return false; }
}

async function makePdf(list) {
  try {
    const result = await buildPdf(list);
    if (result.rastered) toast(`มี ${result.rastered} หน้าจากไฟล์ที่ล็อกรหัสผ่าน ถูกบันทึกเป็นภาพ (คัดลอกข้อความไม่ได้)`, '', [], 8000);
    return result.bytes;
  } catch (e) { toast('สร้าง PDF ไม่สำเร็จ: ' + (e.message || e), 'err', [], 9000); return null; }
}
async function saveList(list, suggested, onSaved) {
  if (!list.length) return false;
  const bytes = await makePdf(list);
  return bytes ? saveBytes(bytes, suggested + '.pdf', 'pdf', onSaved) : false;
}

// Save: if the document came from a PDF file on disk -> ask "overwrite original" or "save as new file"
async function saveAll(forceAs = false) {
  if (!D.pages.length) return false;
  const d = D;
  let choice = 'new';
  if (host.host && d.path && !forceAs) {
    const r = await dialog({
      title: 'บันทึกเอกสาร',
      body: `<p>เอกสารนี้เปิดมาจากไฟล์</p><div class="pathbox">${esc(d.path)}</div><p style="margin-top:12px">ต้องการบันทึกแบบใด?</p>`,
      buttons: [['cancel', 'ยกเลิก'], ['new', 'บันทึกเป็นไฟล์ใหม่…'], ['over', 'บันทึกทับไฟล์เดิม', true]],
    });
    if (!r.btn) return false;
    choice = r.btn;
  }
  if (choice === 'over') {
    const bytes = await makePdf(d.pages);
    if (!bytes) return false;
    try {
      busy('กำลังบันทึกทับไฟล์เดิม …');
      const r = await (await api('/api/overwrite?path=' + encodeURIComponent(d.path), { method: 'POST', body: bytes })).json();
      unbusy();
      if (!r.ok) { toast('บันทึกทับไม่สำเร็จ: ' + r.error, 'err', [], 10000); return false; }
      toast('บันทึกทับไฟล์เดิมแล้ว: ' + r.path, 'ok', [['เปิดโฟลเดอร์', () => api('/api/reveal', { method: 'POST', body: r.path })]], 6000);
      d.dirty = false; refresh();
      return true;
    } catch (e) { unbusy(); toast('บันทึกทับไม่สำเร็จ: ' + e.message, 'err'); return false; }
  }
  const ok = await saveList(d.pages, d.name || 'เอกสาร', (path) => {
    d.path = path; d.dir = dirOf(path);
    d.name = baseName(path.split(/[\\/]/).pop());
  });
  if (ok) { d.dirty = false; refresh(); }
  return ok;
}

// ------------------------------------------------------------------ files in/out wiring
let pendingInsertAt = null;
async function pickInsert(at) {
  pendingInsertAt = at;
  const fs = await pickFiles('fileAdd', 'เลือกไฟล์ที่จะแทรก');
  if (fs) { pendingInsertAt = null; await insertFiles(fs, at); }
}

$('btnOpen').onclick = doOpen;
$('btnPick').onclick = doOpen;
$('btnPastePick').onclick = (e) => { e.stopPropagation(); pasteAny(0); };
$('btnAdd').onclick = () => pickInsert(insertionIndex());
$('btnBlank').onclick = () => addBlankPage();
$('btnDelete').onclick = () => deletePages(selPages().map((p) => p.uid));
$('btnRotL').onclick = () => rotatePages([...D.selected], -90);
$('btnRotR').onclick = () => rotatePages([...D.selected], 90);
$('btnUndo').onclick = undo; $('btnRedo').onclick = redo;
$('btnSave').onclick = () => saveAll();
$('btnSaveMenu').onclick = (e) => {
  const r = e.currentTarget.getBoundingClientRect();
  showMenu([
    { label: D.path ? 'บันทึกทับไฟล์เดิม' : 'บันทึก', icon: 'save', kbd: 'Ctrl+S', disabled: !D.pages.length, action: () => (D.path && host.host ? saveOverwriteDirect() : saveAll()) },
    { label: 'บันทึกเป็นไฟล์ใหม่…', icon: 'extract', kbd: 'Ctrl+Shift+S', disabled: !D.pages.length, action: () => saveAll(true) },
    { sep: true },
    { label: 'บันทึกเฉพาะหน้าที่เลือก…', icon: 'split', disabled: !D.selected.size, action: () => splitToFile(false) },
  ], r.right - 260, r.bottom + 4);
};
async function saveOverwriteDirect() {
  const p = D.path; // reuse the dialog-free path of saveAll
  const bytes = await makePdf(D.pages); if (!bytes) return;
  try {
    busy('กำลังบันทึกทับไฟล์เดิม …');
    const r = await (await api('/api/overwrite?path=' + encodeURIComponent(p), { method: 'POST', body: bytes })).json();
    unbusy();
    if (!r.ok) { toast('บันทึกทับไม่สำเร็จ: ' + r.error, 'err', [], 10000); return; }
    toast('บันทึกทับไฟล์เดิมแล้ว: ' + r.path, 'ok', [], 5000);
    D.dirty = false; refresh();
  } catch (e) { unbusy(); toast('บันทึกทับไม่สำเร็จ: ' + e.message, 'err'); }
}
$('viewOrganize').onclick = () => setMode('organize');
$('viewRead').onclick = () => setMode('read');
$('btnZoomIn').onclick = () => zoomStep(1);
$('btnZoomOut').onclick = () => zoomStep(-1);
$('btnFit').onclick = () => setZoom(0);
$('pageInput').onchange = () => goToPage(Math.max(1, Math.min(D.pages.length, +$('pageInput').value || 1)));
$('thumbSize').oninput = (e) => document.documentElement.style.setProperty('--thumb', e.target.value + 'px');

$('fileOpen').onchange = async (e) => {
  const fs = await Promise.all([...e.target.files].map(readFile)); e.target.value = '';
  await openPicked(fs);
};
$('fileAdd').onchange = async (e) => {
  const fs = await Promise.all([...e.target.files].map(readFile)); e.target.value = '';
  const at = pendingInsertAt; pendingInsertAt = null;
  await insertFiles(fs, at);
};

// drag & drop files from Windows Explorer
let dragDepth = 0;
const isFileDrag = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
function clearDropMarks() { document.querySelectorAll('.dropbefore,.dropafter,.tab.droptarget').forEach((el) => el.classList.remove('dropbefore', 'dropafter', 'droptarget')); }
function dropIndexFromEvent(e) {
  const card = e.target.closest && e.target.closest('.card');
  if (!card) return null;
  const r = card.getBoundingClientRect();
  const i = D.pages.findIndex((p) => p.uid === card.dataset.uid);
  return e.clientX < r.left + r.width / 2 ? { i, before: true, card } : { i: i + 1, before: false, card };
}
window.addEventListener('dragenter', (e) => {
  if (!isFileDrag(e)) return; e.preventDefault(); dragDepth++;
  if (D.pages.length) $('dropOverlay').classList.remove('hidden'); else $('dropcard').classList.add('over');
});
window.addEventListener('dragleave', (e) => {
  if (!isFileDrag(e)) return; dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) { $('dropOverlay').classList.add('hidden'); $('dropcard').classList.remove('over'); clearDropMarks(); }
});
window.addEventListener('dragover', (e) => {
  if (!isFileDrag(e)) return; e.preventDefault();
  clearDropMarks();
  if (e.target.closest && e.target.closest('#tabbar')) { e.target.closest('.tab')?.classList.add('droptarget'); return; }
  if (mode === 'organize') { const d = dropIndexFromEvent(e); if (d) d.card.classList.add(d.before ? 'dropbefore' : 'dropafter'); }
});
window.addEventListener('drop', async (e) => {
  if (!isFileDrag(e)) return; e.preventDefault();
  dragDepth = 0; $('dropOverlay').classList.add('hidden'); $('dropcard').classList.remove('over');
  const onTabbar = e.target.closest && e.target.closest('#tabbar');
  const tabEl = e.target.closest && e.target.closest('.tab');
  const d = mode === 'organize' ? dropIndexFromEvent(e) : null; clearDropMarks();
  const fs = await Promise.all([...e.dataTransfer.files].map(readFile));
  if (tabEl) { const t = docs.find((x) => x.id === tabEl.dataset.id); switchTo(t); await insertFiles(fs); }
  else if (onTabbar || !D.pages.length) await openPicked(fs);
  else await insertFiles(fs, d ? d.i : D.pages.length);
});

// keyboard
window.addEventListener('keydown', (e) => {
  if (modalOpen()) return;
  if (e.key === 'Escape' && ctxEl) { closeMenu(); return; }
  if (typingInField()) return;
  const c = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
  if (c && (k === 'c' || k === 'x') && !e.shiftKey && !e.altKey) { if (D.selected.size) { e.preventDefault(); menuCopy(k === 'x'); } }
  else if (c && k === 'v') { pasteHandled = false; setTimeout(() => { if (!pasteHandled) pasteAny(); }, 120); }
  else if (c && k === 's') { e.preventDefault(); saveAll(e.shiftKey); }
  else if (c && k === 'o') { e.preventDefault(); doOpen(); }
  else if (c && k === 't') { e.preventDefault(); switchTo(newDoc()); }
  else if (c && k === 'w') { e.preventDefault(); closeTab(D); }
  else if (c && e.key === 'Tab') { e.preventDefault(); const i = docs.indexOf(D); switchTo(docs[(i + (e.shiftKey ? -1 : 1) + docs.length) % docs.length]); }
  else if (c && k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
  else if (c && k === 'y') { e.preventDefault(); redo(); }
  else if (c && k === 'd') { e.preventDefault(); duplicatePages(); }
  else if (c && k === 'a' && mode === 'organize') { e.preventDefault(); selectAll(); }
  else if ((e.key === 'Delete' || e.key === 'Backspace') && mode === 'organize') { e.preventDefault(); if (D.selected.size) deletePages(selPages().map((p) => p.uid)); }
  else if (e.key === 'Escape') { D.selected.clear(); paintSelection(); }
  else if (c && (e.key === '=' || e.key === '+')) { e.preventDefault(); if (mode === 'read') zoomStep(1); }
  else if (c && e.key === '-') { e.preventDefault(); if (mode === 'read') zoomStep(-1); }
});
$('reader').addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); zoomStep(e.deltaY < 0 ? 1 : -1); } }, { passive: false });
window.addEventListener('resize', () => { if (mode === 'read' && zoom === 0) { clearTimeout(window.__rz); window.__rz = setTimeout(() => renderReader(), 200); } });
window.addEventListener('beforeunload', (e) => { if (docs.some((d) => d.dirty && d.pages.length)) { e.preventDefault(); e.returnValue = ''; } });

// expose for automated testing
window.__superpdf = { get annDefaults() { return annDefaults; }, setAnnTool, convertTiff, pasteAny, textToPdf, get D() { return D; }, docs, get pages() { return D.pages; }, openAsTabs, insertFiles, buildPdf, setMode, switchTo, get clip() { return clip; } };

if (IS_TOUCH) {
  document.querySelector('#dropcard h1').textContent = 'แตะเพื่อเลือกไฟล์';
  $('dropcard').addEventListener('click', (e) => { if (!e.target.closest('button')) doOpen(); });
}
D = newDoc();
refresh();
initHost();
