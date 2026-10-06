// Super PDF v1.9.3 — front-end (runs inside Edge app window served by SuperPDF.exe)
import * as pdfjsLib from './pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
const { PDFDocument, degrees } = window.PDFLib;

const $ = (id) => document.getElementById(id);
const TOKEN = new URLSearchParams(location.search).get('t') || '';
const IMG_EXT = ['jpg', 'jpeg', 'jfif', 'png', 'gif', 'bmp', 'webp', 'heic', 'heif'];
const TIFF_EXT = ['tif', 'tiff', 'mtiff'];
const WORD_EXT = ['doc', 'docx', 'docm', 'rtf', 'odt', 'txt', 'htm', 'html', 'xml', 'wpd', 'dot', 'dotx'];
const EXCEL_EXT = ['xls', 'xlsx', 'xlsm', 'xlsb', 'csv', 'ods'];
const PPT_EXT = ['ppt', 'pptx', 'pps', 'ppsx', 'odp'];
const A4 = [595.28, 841.89];
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// ------------------------------------------------------------------ state
const sources = new Map();   // id -> { id, name, bytes, pdf (pdfjs doc), password, blank }
const docs = [];             // tabs: { id, name, pages:[{uid,src,index,rot}], selected:Set, lastClicked, undo:[], redo:[], dirty, scroll:{} }
const APP_VERSION = '1.9.3';   // shown at the top-left (#appVer) — change here and in index.html when releasing
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
// v1.9.1: a rejected promise that no handler caught must never leave the busy overlay up silently
window.addEventListener('unhandledrejection', (e) => {
  try {
    const m = e.reason && (e.reason.message || String(e.reason)) || 'ไม่ทราบสาเหตุ';
    if (/ResizeObserver|AbortError/.test(m)) return;
    unbusy(); toast('เกิดข้อผิดพลาด: ' + m, 'err', [], 9000);
  } catch {}
});

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
  : { cMapUrl: './cmaps/', cMapPacked: true, standardFontDataUrl: './standard_fonts/' };

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
      sources.set(id, { id, name, bytes, pdf, password, born: Date.now() });
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

// v1.9.3: iPhone photos (.heic/.heif). Edge on Windows cannot show them, so they are decoded with libheif
// (web/heic/libheif-bundle.js, ~2 MB, loaded only the first time a HEIC picture is opened).
async function isHeic(blob) {
  const b = new Uint8Array(await blob.slice(0, 16).arrayBuffer());
  const t = String.fromCharCode(...b.slice(4, 12));
  return t.startsWith('ftyp') && /^(heic|heix|heim|heis|hevc|hevx|mif1|msf1|avci)$/.test(t.slice(4));
}
let heifLib = null;
function loadHeif() {
  if (!heifLib) heifLib = new Promise((res, rej) => {
    if (window.libheif) return res(window.libheif);
    const sc = document.createElement('script');
    sc.src = new URL('./heic/libheif-bundle.js', location.href).href;
    sc.onload = () => (window.libheif ? res(window.libheif) : rej(new Error('โหลดตัวอ่าน HEIC ไม่สำเร็จ')));
    sc.onerror = () => rej(new Error('โหลดตัวอ่าน HEIC ไม่สำเร็จ (ถ้าเป็น iPhone ให้เปิดเว็บขณะออนไลน์ 1 ครั้งก่อน)'));
    document.head.appendChild(sc);
  }).then((f) => (typeof f === 'function' ? f() : f)).catch((e) => { heifLib = null; throw e; });
  return heifLib;
}
async function heicToJpegBlob(blob) {
  const lib = await loadHeif();
  const dec = new lib.HeifDecoder();
  const imgs = dec.decode(new Uint8Array(await blob.arrayBuffer()));
  if (!imgs || !imgs.length) throw new Error('อ่านไฟล์ HEIC ไม่ได้');
  const im = imgs[0], w = im.get_width(), h = im.get_height();
  try {
    const id = await new Promise((res, rej) => im.display({ data: new Uint8ClampedArray(w * h * 4), width: w, height: h }, (d) => (d ? res(d) : rej(new Error('ถอดรหัสภาพ HEIC ไม่สำเร็จ')))));
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    c.getContext('2d').putImageData(new ImageData(id.data, w, h), 0, 0);
    const out = await canvasToBytes(c, 'image/jpeg', 0.92);
    c.width = c.height = 0;
    return new Blob([out], { type: 'image/jpeg' });
  } finally { for (const x of imgs) try { x.free && x.free(); } catch {} }
}
async function loadImageEl(blob) {
  try { return await loadImageElRaw(blob); }
  catch (e) { if (await isHeic(blob).catch(() => false)) return loadImageElRaw(await heicToJpegBlob(blob)); throw e; }
}
function loadImageElRaw(blob) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(blob);
    const im = new Image();
    // v1.9.1: the decoded bitmap stays valid after the URL is released — release it on both paths
    im.onload = () => { res(im); setTimeout(() => URL.revokeObjectURL(url), 0); };
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
  // v1.9.1: always settles — toBlob() may hand back null (huge canvas / low memory) and arrayBuffer() may reject
  return new Promise((res, rej) => {
    try {
      canvas.toBlob((b) => {
        if (!b) { rej(new Error('สร้างข้อมูลรูปภาพไม่สำเร็จ (ภาพใหญ่เกินไปหรือหน่วยความจำไม่พอ)')); return; }
        b.arrayBuffer().then((a) => res(new Uint8Array(a)), (e) => rej(e instanceof Error ? e : new Error(String(e || 'อ่านข้อมูลรูปภาพไม่สำเร็จ'))));
      }, type, q);
    } catch (e) { rej(e); }
  });
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
  gcSources();
}

// v1.9.1: release PDF sources / thumbnails nobody references any more (open tabs, their undo/redo history, clipboard).
// Sources younger than SRC_GRACE_MS are kept: they may be mid-import and not yet attached to a page.
const SRC_GRACE_MS = 30000;
let gcTimer = 0;
function gcSources() {
  const live = new Set();
  const addPages = (list) => { for (const p of list || []) live.add(p.src); };
  for (const d of docs) {
    addPages(d.pages);
    for (const s of d.undo || []) addPages(s.pages);
    for (const s of d.redo || []) addPages(s.pages);
  }
  if (clip) for (const it of clip.items || []) live.add(it.src);
  let young = false;
  const dead = new Set();
  for (const [id, s] of sources) {
    if (live.has(id)) continue;
    if (Date.now() - (s.born || 0) < SRC_GRACE_MS) { young = true; continue; }
    dead.add(id);
  }
  if (young && !gcTimer) gcTimer = setTimeout(() => { gcTimer = 0; gcSources(); }, SRC_GRACE_MS + 1000);
  if (!dead.size) return 0;
  for (const id of dead) {
    const s = sources.get(id);
    sources.delete(id); libCache.delete(id);
    try { const r = s.pdf && s.pdf.destroy && s.pdf.destroy(); if (r && r.catch) r.catch(() => {}); } catch {}
    s.bytes = null; s.pdf = null;
  }
  for (const [k, url] of [...thumbCache]) {
    if (dead.has(k.split(':')[0])) { thumbCache.delete(k); try { if (url) URL.revokeObjectURL(url); } catch {} }
  }
  for (let i = thumbQueue.length - 1; i >= 0; i--) if (dead.has(thumbQueue[i].p.src)) thumbQueue.splice(i, 1);
  return dead.size;
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
  if (!blob) throw new Error('thumb');
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
const modalOpen = () => !$('modal').classList.contains('hidden') || sigOverlayOpen();
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
const TYPE_EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/heic': 'heic', 'image/heif': 'heif', 'image/tiff': 'tiff', 'application/pdf': 'pdf' };
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
  window.__menuAt = Date.now();
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
window.addEventListener('mousedown', (e) => { if (ctxEl && !ctxEl.contains(e.target) && Date.now() - (window.__menuAt || 0) > 350) closeMenu(); }, true);
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
  { label: 'แบ่งตามขนาดไฟล์ (ไม่เกิน … MB)…', icon: 'split', disabled: !D.pages.length, action: () => splitBySize() },
];
const convertMenu = () => [
  { head: n() ? `แปลง ${n()} หน้าที่เลือก` : 'แปลงทุกหน้า' },
  { label: 'ย่อขนาดไฟล์ PDF…', icon: 'compress', disabled: !D.pages.length, action: () => compressDialog() },
  { sep: true },
  { label: 'รูปภาพ JPG…', icon: 'img', action: () => convertImages('jpg') },
  { label: 'รูปภาพ PNG…', icon: 'img', action: () => convertImages('png') },
  { label: 'TIFF (1 หน้า = 1 ไฟล์)…', icon: 'img', action: () => convertTiff(false) },
  { label: 'TIFF หลายหน้า (.tiff / .mtiff)…', icon: 'img', action: () => convertTiff(true) },
  { label: 'Word (.docx)…', icon: 'word', disabled: !host.host || !host.word, action: () => convertWord() },
  { label: 'Word จากเอกสารสแกน (อ่านภาษาไทย)…', icon: 'word', action: () => convertScanWord() },
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
    { label: 'ลงลายเซ็น / ตราประทับ…', icon: 'sign', action: () => { setMode('read'); setTimeout(() => { goToPage(idx + 1); setTimeout(() => sigOpenPanel(), 150); }, 80); } },
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
    { label: 'แก้ไขภาพ (ลบ/ย้ายวัตถุ, ยางลบ)…', icon: 'edit', action: () => openImageEditor(D.pages[idx]) },
    ...(host.paint ? [{ label: 'แก้ไขภาพด้วย Paint…', icon: 'paint', action: () => openImageEditor(D.pages[idx], { paint: true }) }] : []),
    { label: `ลบจุดสกปรก${n() > 1 ? ` (${n()} หน้า)` : ''}`, icon: 'clean', action: () => despecklePages(selPages()) },
    { label: 'ปรับภาพแบบถ่ายเอกสาร / ครอป', icon: 'scan', action: () => enhancePages(selPages()) },
    { label: 'พิมพ์หน้าที่เลือก…', icon: 'print', kbd: 'Ctrl+P', action: () => printDialog('sel') },
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
    { label: 'สแกนเอกสารเพิ่ม…', icon: 'scan', action: () => openScanner() },
    { label: 'พิมพ์…', icon: 'print', kbd: 'Ctrl+P', action: () => printDialog() },
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
    if (annEl2.classList.contains('annimg')) { showMenu(sigAnnMenu(), e.clientX, e.clientY); return; }
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
    const L2 = !card && rpage.querySelector('.annlayer');
    showMenu(L2 ? [...sigPointMenu(L2, e), { sep: true }, ...pageMenu(idx)] : pageMenu(idx), e.clientX, e.clientY);
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

// v1.9.2: split by file size — each part holds as many consecutive pages as possible without going over the limit
// (1 MB = 1,048,576 bytes, the same "MB" Windows Explorer shows). PDF sizes are not additive (fonts/images are shared),
// so the per-page sizes are only a first guess: every part is built for real and checked, then grown/shrunk.
let splitMB = 5;
const fmtMB = (b) => (b / 1048576).toFixed(b < 10 * 1048576 ? 2 : 1) + ' MB';
async function splitBySize() {
  if (!D.pages.length) return;
  const nsel = D.selected.size;
  const r = await dialog({
    title: 'แบ่งไฟล์ตามขนาด',
    body: `<div class="field"><label>ขนาดสูงสุดต่อไฟล์ (MB)</label><input name="mb" type="text" inputmode="decimal" value="${splitMB}"></div>` +
      (nsel ? `<div class="field"><label>หน้าที่จะแบ่ง</label><select name="scope"><option value="all">ทั้งเอกสาร (${D.pages.length} หน้า)</option><option value="sel">เฉพาะหน้าที่เลือก (${nsel} หน้า)</option></select></div>` : '') +
      `<p>รวมหน้าตามลำดับให้แต่ละไฟล์ได้หน้ามากที่สุดโดยขนาดไม่เกินที่กำหนด (เช่น 2, 4.5, 10)</p>`,
    focus: 'input',
  });
  if (!r.btn) return;
  const mb = parseFloat(String(r.values.mb || '').replace(',', '.'));
  if (!(mb > 0)) { toast('กรุณาใส่ขนาดเป็นตัวเลขที่มากกว่า 0 (หน่วย MB)', 'err'); return; }
  splitMB = mb;
  const limit = Math.floor(mb * 1048576);
  const pages = nsel && r.values.scope === 'sel' ? selPages() : D.pages.slice();
  const num = new Map(D.pages.map((p, i) => [p.uid, i + 1]));
  const frames = [], over = [];
  try {
    const est = [];
    for (let k = 0; k < pages.length; k++) {
      busy(`กำลังวัดขนาดแต่ละหน้า … ${k + 1}/${pages.length}`);
      est.push((await buildPdfInner([pages[k]], true)).bytes.length);
    }
    const sumEst = (a, b) => { let t = 0; for (let k = a; k < b; k++) t += est[k]; return t; };
    // last index (exclusive) of a run from `from` whose estimated size * ratio fits; always at least one page
    const fit = (from, ratio) => { let j = from + 1, t = est[from]; while (j < pages.length && (t + est[j]) * ratio <= limit) t += est[j++]; return j; };
    let i = 0;
    while (i < pages.length) {
      busy(`กำลังแบ่งไฟล์ … ส่วนที่ ${frames.length + 1} (ถึงหน้า ${i + 1}/${pages.length})`);
      let best = null, bestJ = i, j = fit(i, 1), ratio = 1;
      const tried = new Set();
      for (let tries = 0; tries < 12 && !tried.has(j); tries++) {
        tried.add(j);
        const bytes = (await buildPdfInner(pages.slice(i, j), true)).bytes;
        ratio = bytes.length / sumEst(i, j);
        if (bytes.length <= limit || j - i === 1) {
          if (j > bestJ) { best = bytes; bestJ = j; }
          if (bytes.length > limit || j >= pages.length) break;   // one page alone is bigger than the limit / nothing left
          const j2 = Math.max(j + 1, fit(i, ratio));               // room left: try more pages
          if (tried.has(j2)) break;
          j = j2;
        } else {
          let j2 = Math.min(j - 1, fit(i, ratio * 1.01));          // too big: shrink
          if (j2 <= bestJ) { if (best) break; j2 = Math.max(i + 1, j2); }
          j = j2;
        }
      }
      if (!best) { // safety net: step down one page at a time
        for (j = j - 1; j > i; j--) {
          const bytes = (await buildPdfInner(pages.slice(i, j), true)).bytes;
          if (bytes.length <= limit || j - i === 1) { best = bytes; bestJ = j; break; }
        }
      }
      const a = num.get(pages[i].uid), b = num.get(pages[bestJ - 1].uid);
      if (best.length > limit) over.push(a);
      frames.push({ name: `ส่วนที่${pad(frames.length + 1, 2)}_${bestJ - i > 1 ? `หน้า${pad(a)}-${pad(b)}` : `หน้า${pad(a)}`}.pdf`, bytes: best });
      i = bestJ;
    }
  } catch (e) { unbusy(); toast('แบ่งไฟล์ไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  if (frames.length === 1 && !over.length) {
    toast(`ไฟล์มีขนาด ${fmtMB(frames[0].bytes.length)} ไม่เกิน ${mb} MB อยู่แล้ว จึงไม่ต้องแบ่ง`, 'ok', [], 5000);
    return;
  }
  if (over.length) toast(`หน้า ${over.join(', ')} มีขนาดเกิน ${mb} MB แม้อยู่หน้าเดียว จึงแยกเป็นไฟล์เดี่ยว (ลองลดขนาดภาพก่อนแบ่ง)`, 'err', [], 8000);
  window.__lastSplit = frames.map((f) => ({ name: f.name, size: f.bytes.length }));
  await saveMany(frames, D.name, 'pdf');
}

// ------------------------------------------------------------------ v1.9.2 compress (ย่อขนาดไฟล์)
// Re-encodes the pictures inside the PDF (scans, photos) as JPEG at a lower resolution; text, vector drawings,
// black-and-white (CCITT/JBIG2) images and transparent pictures (signatures, text boxes) are left untouched.
const COMPRESS_PRESETS = [
  { key: 'high', label: 'คุณภาพสูง', dpi: 200, q: 0.8, note: '200 dpi · อ่านชัด เหมาะกับพิมพ์' },
  { key: 'mid', label: 'คุณภาพกลาง', dpi: 150, q: 0.65, note: '150 dpi · สมดุล เหมาะกับส่งอีเมล/อัปโหลด' },
  { key: 'low', label: 'คุณภาพต่ำ', dpi: 100, q: 0.5, note: '100 dpi · ไฟล์เล็กที่สุด เหมาะกับดูบนจอ' },
];
function cmpImages(doc) {
  const { PDFRawStream, PDFName, PDFArray, PDFNumber, PDFBool } = window.PDFLib;
  const ctx = doc.context, N = (x) => PDFName.of(x), look = (o) => (o && ctx.lookup(o));
  const out = [];
  for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream)) continue;
    const d = obj.dict;
    if (d.get(N('Subtype')) !== N('Image')) continue;
    if (d.get(N('SMask')) || d.get(N('Mask')) || d.get(N('Decode')) || look(d.get(N('ImageMask'))) === PDFBool.True) continue;
    let f = look(d.get(N('Filter')));
    if (f instanceof PDFArray) { if (f.size() !== 1) continue; f = look(f.get(0)); }
    const kind = f === N('DCTDecode') ? 'jpeg' : f === N('FlateDecode') ? 'flate' : null;
    if (!kind) continue;
    let cs = look(d.get(N('ColorSpace'))), comps = 0;
    if (cs === N('DeviceRGB')) comps = 3; else if (cs === N('DeviceGray')) comps = 1;
    else if (cs instanceof PDFArray && look(cs.get(0)) === N('ICCBased')) { const icc = look(cs.get(1)); const nn = icc && icc.dict && look(icc.dict.get(N('N'))); comps = nn instanceof PDFNumber ? nn.asNumber() : 0; }
    if (comps !== 1 && comps !== 3) continue;
    const num = (k) => { const v = look(d.get(N(k))); return v instanceof PDFNumber ? v.asNumber() : 0; };
    const w = num('Width'), h = num('Height'), bpc = num('BitsPerComponent') || 8;
    if (!w || !h || bpc !== 8) continue;
    let pred = 1, cols = w, colors = comps;
    const dp = look(d.get(N('DecodeParms')));
    if (dp && dp.get) { const g = (k) => { const v = look(dp.get(N(k))); return v instanceof PDFNumber ? v.asNumber() : 0; }; pred = g('Predictor') || 1; cols = g('Columns') || w; colors = g('Colors') || comps; }
    if (kind === 'flate' && (pred === 2 || (pred > 1 && (cols !== w || colors !== comps)))) continue;
    const len = obj.contents.length;
    if (len < 16 * 1024) continue; // tiny pictures are not worth it
    out.push({ ref, obj, kind, comps, w, h, pred, len });
  }
  return out;
}
function cmpUnpredict(raw, w, h, comps) { // PNG predictors (Predictor >= 10)
  const bpr = w * comps, out = new Uint8Array(bpr * h);
  for (let y = 0; y < h; y++) {
    const t = raw[y * (bpr + 1)], src = y * (bpr + 1) + 1, dst = y * bpr, up = dst - bpr;
    for (let x = 0; x < bpr; x++) {
      const a = x >= comps ? out[dst + x - comps] : 0, b = y ? out[up + x] : 0, c = (y && x >= comps) ? out[up + x - comps] : 0;
      let v = raw[src + x];
      if (t === 1) v += a; else if (t === 2) v += b; else if (t === 3) v += (a + b) >> 1;
      else if (t === 4) { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      out[dst + x] = v & 255;
    }
  }
  return out;
}
async function cmpDecode(im) {
  if (im.kind === 'jpeg') return await createImageBitmap(new Blob([im.obj.contents], { type: 'image/jpeg' }));
  let px = pako.inflate(im.obj.contents);
  if (im.pred >= 10) px = cmpUnpredict(px, im.w, im.h, im.comps);
  if (px.length < im.w * im.h * im.comps) throw new Error('short image data');
  const c = document.createElement('canvas'); c.width = im.w; c.height = im.h;
  const id = new ImageData(im.w, im.h), o = id.data;
  for (let i = 0, j = 0, n = im.w * im.h; i < n; i++, j += im.comps) {
    const r = px[j], g = im.comps === 3 ? px[j + 1] : r, b = im.comps === 3 ? px[j + 2] : r;
    o[i * 4] = r; o[i * 4 + 1] = g; o[i * 4 + 2] = b; o[i * 4 + 3] = 255;
  }
  c.getContext('2d').putImageData(id, 0, 0);
  return c;
}
// re-encode one picture; returns { bytes, w, h } or null when it would not get smaller
async function cmpEncode(im, preset, maxPx) {
  const src = await cmpDecode(im);
  const k = Math.min(1, maxPx / Math.max(im.w, im.h));
  const w = Math.max(1, Math.round(im.w * k)), h = Math.max(1, Math.round(im.h * k));
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, w, h);
  g.imageSmoothingEnabled = true; g.imageSmoothingQuality = 'high'; g.drawImage(src, 0, 0, w, h);
  if (src.close) src.close(); else src.width = src.height = 0;
  const bytes = await canvasToBytes(c, 'image/jpeg', preset.q);
  c.width = c.height = 0;
  return bytes.length < im.len * 0.92 ? { bytes, w, h } : null;
}
const cmpMaxPx = (doc, dpi) => Math.round(Math.max(...doc.getPages().map((pg) => { const { width, height } = pg.getSize(); return Math.max(width, height); })) / 72 * dpi);
async function compressPdfBytes(bytes, preset, onProgress) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
  const maxPx = cmpMaxPx(doc, preset.dpi);
  const imgs = cmpImages(doc);
  let done = 0;
  for (const im of imgs) {
    onProgress && onProgress(++done, imgs.length);
    let r = null; try { r = await cmpEncode(im, preset, maxPx); } catch { r = null; }
    if (!r) continue;
    const dict = { Type: 'XObject', Subtype: 'Image', Width: r.w, Height: r.h, BitsPerComponent: 8, ColorSpace: 'DeviceRGB', Filter: 'DCTDecode' };
    doc.context.assign(im.ref, doc.context.stream(r.bytes, dict));
  }
  return await doc.save({ useObjectStreams: true });
}
// estimate each preset from a sample of the pictures (fast), then show the choice
async function compressDialog() {
  if (!D.pages.length) return;
  let orig, ests = [];
  try {
    busy('กำลังประเมินขนาดไฟล์ …');
    orig = (await buildPdfInner(D.pages, true)).bytes; // whole document, with edits/text boxes/signatures
    const doc = await PDFDocument.load(orig, { ignoreEncryption: true, updateMetadata: false });
    const imgs = cmpImages(doc), imgTotal = imgs.reduce((t, x) => t + x.len, 0);
    const step = Math.max(1, Math.ceil(imgs.length / 6)), sample = imgs.filter((_, i) => i % step === 0).slice(0, 6);
    const sTotal = sample.reduce((t, x) => t + x.len, 0);
    for (const p of COMPRESS_PRESETS) {
      const maxPx = cmpMaxPx(doc, p.dpi);
      let sNew = 0;
      for (const im of sample) { let r = null; try { r = await cmpEncode(im, p, maxPx); } catch {} sNew += r ? r.bytes.length : im.len; }
      const ratio = sTotal ? sNew / sTotal : 1;
      ests.push(Math.round(orig.length - imgTotal + imgTotal * ratio));
    }
  } catch (e) { unbusy(); toast('ประเมินขนาดไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  const pct = (b) => { const x = Math.round((1 - b / orig.length) * 100); return x > 0 ? ` (เล็กลง ~${x}%)` : ' (ใกล้เคียงเดิม)'; };
  const r = await dialog({
    title: 'ย่อขนาดไฟล์ PDF',
    body: `<p>ขนาดปัจจุบัน <b>${fmtMB(orig.length)}</b> · ${D.pages.length} หน้า</p>` +
      COMPRESS_PRESETS.map((p, i) => `<label class="opt"><input type="radio" name="q" value="${p.key}" ${p.key === 'mid' ? 'checked' : ''}><span><b>${p.label}</b><span>ประมาณ <b>${fmtMB(ests[i])}</b>${pct(ests[i])}</span><small>${p.note}</small></span></label>`).join('') +
      `<p class="hint">ย่อเฉพาะรูปภาพในไฟล์ (เช่น หน้าสแกน รูปถ่าย) ข้อความยังค้นหา/คัดลอกได้เหมือนเดิม · ขนาดจริงอาจต่างจากค่าประมาณเล็กน้อย</p>`,
    buttons: [['cancel', 'ยกเลิก'], ['ok', 'ย่อและบันทึก…', true]],
  });
  if (!r.btn) return;
  const preset = COMPRESS_PRESETS.find((p) => p.key === r.values.q) || COMPRESS_PRESETS[1];
  let out;
  try {
    out = await compressPdfBytes(orig, preset, (i, n) => busy(`กำลังย่อขนาด (${preset.label}) … รูปที่ ${i}/${n}`));
  } catch (e) { unbusy(); toast('ย่อขนาดไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  window.__lastCompress = { preset: preset.key, before: orig.length, after: out.length, estimate: ests[COMPRESS_PRESETS.indexOf(preset)] };
  if (out.length >= orig.length * 0.97) {
    toast(`ไฟล์นี้ย่อได้ไม่มาก (${fmtMB(orig.length)} → ${fmtMB(out.length)}) เพราะส่วนใหญ่เป็นข้อความหรือภาพขาวดำที่บีบอัดไว้แล้ว`, '', [], 8000);
    if (out.length >= orig.length) return;
  }
  const saved = await saveBytes(out, D.name + '_ย่อขนาด', 'pdf');
  if (lastSave === 'ok') toast(`ย่อขนาดแล้ว: ${fmtMB(orig.length)} → ${fmtMB(out.length)} (เล็กลง ${Math.round((1 - out.length / orig.length) * 100)}%)`, 'ok', [], 8000);
  return saved;
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
  for (const en of entries) {
    if (en.isIntersecting) paintReaderPage(en.target);
    else if (en.target.dataset.painted) { // free memory of pages far away (important on iPhone)
      en.target.querySelectorAll('canvas').forEach((c) => { c.width = c.height = 0; c.remove(); });
      delete en.target.dataset.painted;
    }
  }
}, { root: $('reader'), rootMargin: '900px 900px' });

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
  // limit canvas size (iOS refuses canvases above ~16 megapixels)
  const want = +d.dataset.scale * (window.devicePixelRatio || 1);
  const w = parseFloat(d.style.width) / +d.dataset.scale, h = parseFloat(d.style.height) / +d.dataset.scale;
  const maxPx = IS_TOUCH ? 12e6 : 36e6;
  const sc = Math.min(want, Math.sqrt(maxPx / (w * h)));
  const c = await renderPageCanvas(p, sc, false);
  if (!d.dataset.painted || !d.isConnected) { c.width = c.height = 0; return; }
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

function setZoom(z, focus) {
  const r = $('reader');
  const old = parseFloat($('zoomLabel').textContent) / 100 || 1;
  // keep the point under the fingers / screen centre in place
  const rr = r.getBoundingClientRect();
  const fx = focus ? focus.x - rr.left : r.clientWidth / 2, fy = focus ? focus.y - rr.top : r.clientHeight / 2;
  const px = (r.scrollLeft + fx) / old, py = (r.scrollTop + fy) / old;
  zoom = z;
  return renderReader(true).then(() => {
    const k = (parseFloat($('zoomLabel').textContent) / 100 || z);
    r.scrollLeft = px * k - fx; r.scrollTop = py * k - fy;
  });
}

// pinch with two fingers to zoom the reader (iPhone / iPad / touch screens)
(() => {
  const R = $('reader'), W = $('readerPages');
  let pinch = null;
  const dist = (t) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
  const mid = (t) => ({ x: (t[0].clientX + t[1].clientX) / 2, y: (t[0].clientY + t[1].clientY) / 2 });
  R.addEventListener('touchstart', (e) => {
    if (e.touches.length !== 2 || annEditing) return;
    const m = mid(e.touches), rr = R.getBoundingClientRect();
    pinch = { d0: dist(e.touches), s0: parseFloat($('zoomLabel').textContent) / 100 || 1, m, k: 1 };
    W.style.transformOrigin = `${R.scrollLeft + m.x - rr.left}px ${R.scrollTop + m.y - rr.top}px`;
  }, { passive: true });
  R.addEventListener('touchmove', (e) => {
    if (!pinch || e.touches.length !== 2) return;
    e.preventDefault();
    const k = Math.max(0.25 / pinch.s0, Math.min(5 / pinch.s0, dist(e.touches) / pinch.d0));
    pinch.k = k; pinch.m = mid(e.touches);
    W.style.transform = `scale(${k})`;               // instant preview
  }, { passive: false });
  const end = () => {
    if (!pinch) return;
    const p = pinch; pinch = null;
    const z = Math.max(0.25, Math.min(5, p.s0 * p.k));
    if (Math.abs(p.k - 1) < 0.03) { W.style.transform = ''; return; }
    setZoom(z, p.m).then(() => { W.style.transform = ''; });
  };
  R.addEventListener('touchend', (e) => { if (e.touches.length < 2) end(); });
  R.addEventListener('touchcancel', end);
  // double-tap: fit ⇄ 200 %
  let lastTap = 0;
  R.addEventListener('touchend', (e) => {
    if (e.touches.length || e.changedTouches.length !== 1 || e.target.closest('.annot') || annTool) return;
    const now = Date.now(), t = e.changedTouches[0];
    if (now - lastTap < 300) { lastTap = 0; const cur = parseFloat($('zoomLabel').textContent) / 100 || 1; setZoom(zoom && cur > 1.2 ? 0 : Math.max(2, cur * 2), { x: t.clientX, y: t.clientY }); }
    else lastTap = now;
  });
})();
function zoomStep(dir) {
  const cur = parseFloat($('zoomLabel').textContent) / 100 || 1;
  const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5];
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
// v1.9.2: fonts shipped inside the program (web/fonts, @font-face in app.css) — same look on every PC and on iPhone
const BUNDLED_FONTS = {
  'TH Sarabun New': { r: 'fonts/THSarabunNew.ttf', b: 'fonts/THSarabunNew-Bold.ttf', i: 'fonts/THSarabunNew-Italic.ttf', bi: 'fonts/THSarabunNew-BoldItalic.ttf' },
};
const bundledFontsReady = (async () => {
  try {
    await Promise.all(Object.keys(BUNDLED_FONTS).flatMap((f) =>
      ['', 'bold ', 'italic ', 'italic bold '].map((st) => document.fonts.load(`${st}16px "${f}"`, 'กขAa'))));
  } catch {}
  try { if (typeof D !== 'undefined' && D && D.pages && D.pages.length) { renderAllAnnLayers(); refresh(); } } catch {}
})();
let annDefaults = { font: 'TH Sarabun New', size: 16, color: '#1f3a93', bold: false, italic: false };
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
  if (a.kind === 'img') { // signature / stamp picture
    const im = sigFinal(a);
    if (im) { g.save(); g.imageSmoothingEnabled = true; g.imageSmoothingQuality = 'high'; g.translate(a.x, a.y); if (a.r) g.rotate(a.r * Math.PI / 180); g.drawImage(im, 0, 0, a.w, a.h); g.restore(); }
    return null;
  }
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
  await bundledFontsReady;
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const vb = pg.getViewport({ scale: 1, rotation: pg.rotate });
  const R = pg.rotate || 0;
  for (const a of p.annots) {
    if (a.kind === 'img') { await embedSigAnn(out, pdfPage, a, vb, R); continue; }
    if (!a.text || !a.text.trim()) continue;
    const m = annMetrics(a), K = 4;
    const c = document.createElement('canvas'); c.width = Math.ceil(m.w * K); c.height = Math.ceil(m.h * K);
    const g = c.getContext('2d'); g.scale(K, K); paintAnn(g, { ...a, x: 0, y: 0, r: 0 });
    const img = await out.embedPng(await canvasToBytes(c, 'image/png'));
    c.width = c.height = 0;
    const rot = ((R - (a.r || 0)) % 360 + 360) % 360;
    const [px, py] = vb.convertToPdfPoint(...annLocal(a, 0, m.h));
    pdfPage.drawImage(img, { x: px, y: py, width: m.w, height: m.h, rotate: degrees(rot) });
    const font = await annPdfFont(out, a);
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
async function annPdfFont(out, a) { // real font file (for the invisible, searchable text layer)
  if (!window.fontkit) return null;
  const family = a.font, bf = BUNDLED_FONTS[family];
  if (!bf && !host.host) return null;
  const variant = (a.bold ? 'b' : '') + (a.italic ? 'i' : '') || 'r';
  const key = bf ? family + '|' + variant : family;   // v1.9.2: bundled fonts carry their own bold/italic files
  if (out.__fonts.has(key)) return out.__fonts.get(key);
  let f = null;
  try {
    if (!fontBytesCache.has(key)) {
      const r = bf ? await fetch(bf[variant]) : await fetch('/api/fontfile?family=' + encodeURIComponent(family), { headers: { 'X-Token': TOKEN } });
      fontBytesCache.set(key, r.ok ? new Uint8Array(await r.arrayBuffer()) : null);
    }
    const bytes = fontBytesCache.get(key);
    if (bytes) { out.registerFontkit(window.fontkit); f = await out.embedFont(bytes, { subset: true }); }
  } catch { f = null; }
  out.__fonts.set(key, f);
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
    if (a.kind === 'img') return sigAnnEl(a, s, !!(annSel && annSel.id === a.id));
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
  if ((findAnn(annSel) || {}).kind === 'img') { sigOpenProps(); return; }
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
  if (el && el.classList.contains('annimg')) { sigPointerDown(e, L, p, el); return; }
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
    } else if (annSel) { selectAnn(null); sigCloseProps(); }
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
  try { box.setPointerCapture(e.pointerId); } catch {}
  const R = $('reader');
  const maxX = L.offsetWidth / s - 8, maxY = L.offsetHeight / s - 8;
  let last = e, vy = 0, raf = 0;
  const place = (ev) => {
    const [mx, my] = annPoint(L, ev);
    if (!moved && Math.hypot(mx - sx, my - sy) * s < 3) return;
    if (!moved) { moved = true; pushUndo(); }
    a.x = Math.max(-4, Math.min(maxX, ox + (mx - sx)));
    a.y = Math.max(-4, Math.min(maxY, oy + (my - sy)));
    box.style.left = a.x * s + 'px'; box.style.top = a.y * s + 'px';
  };
  const tick = () => { // auto-scroll when the finger/mouse is near the top or bottom edge
    raf = 0;
    if (!vy) return;
    R.scrollTop += vy; place(last);
    raf = requestAnimationFrame(tick);
  };
  const move = (ev) => {
    if (ev.pointerId !== e.pointerId) return;
    ev.preventDefault();
    last = ev; place(ev);
    const rr = R.getBoundingClientRect(), edge = 56;
    vy = !moved ? 0 : ev.clientY > rr.bottom - edge ? Math.min(18, (ev.clientY - (rr.bottom - edge)) / 3 + 2)
      : ev.clientY < rr.top + edge ? -Math.min(18, ((rr.top + edge) - ev.clientY) / 3 + 2) : 0;
    if (vy && !raf) raf = requestAnimationFrame(tick);
  };
  const up = (ev) => {
    if (ev && ev.pointerId !== e.pointerId) return;
    vy = 0; if (raf) cancelAnimationFrame(raf);
    window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); window.removeEventListener('pointercancel', up);
    if (moved) { touchPage(p); refresh(); }
    else if (wasSel) startEdit();
  };
  window.addEventListener('pointermove', move, { passive: false }); window.addEventListener('pointerup', up); window.addEventListener('pointercancel', up);
});
$('readerPages').addEventListener('dblclick', (e) => { if (e.target.closest('.annot') && !annEditing) startEdit(); });

function deleteAnn() {
  const p = pageByUid(annSel?.uid); if (!p) return;
  pushUndo(); p.annots = p.annots.filter((a) => a.id !== annSel.id); touchPage(p); annSel = null; refresh();
}
function copyAnn(cut) {
  const a = findAnn(annSel); if (!a) return;
  annClip = { ann: { ...a }, marker: 'SuperPDF-text-' + (seq++) };
  try { navigator.clipboard.writeText(annClip.marker).catch(() => {}); } catch {}
  if (cut) deleteAnn();
  const what = a.kind === 'img' ? sigKindName(a.sk) : 'กล่องข้อความ';
  toast(cut ? `ตัด${what}แล้ว` : `คัดลอก${what}แล้ว — กด Ctrl+V เพื่อวาง (วางหน้าอื่น/แท็บอื่นได้)`, '', [], 2500);
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
  if (a && a.kind !== 'img') { pushUndo(); Object.assign(a, patch); touchPage(pageByUid(annSel.uid)); annEditing = false; refresh(); }
  updateAnnBar();
}
function setAnnTool(on) { annTool = on; if (on && mode !== 'read') setMode('read'); renderAllAnnLayers(); updateAnnBar(); }

// ---- format bar
function updateAnnBar() {
  const bar = $('annbar'); if (!bar) return;
  bar.classList.toggle('hidden', mode !== 'read' || !D.pages.length);
  const a0 = findAnn(annSel), a = (a0 && a0.kind !== 'img' && a0) || annDefaults;
  $('annTool').classList.toggle('active', annTool);
  if (mode !== 'read') { sigClosePanel(); sigCloseProps(); } else sigPropsSync();
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
  $('annHint').textContent = annTool ? 'คลิกบนหน้าเพื่อวางข้อความ' : (a0 && a0.kind === 'img') ? 'ลากเพื่อย้าย · ลากปุ่มมุมขวาล่างเพื่อย่อ-ขยาย/หมุน · ⋯ = สี/ความหนา' : has ? 'ลากเพื่อย้าย · คลิกอีกครั้ง/ดับเบิลคลิกเพื่อแก้ข้อความ' : IS_TOUCH ? '' : 'คลิกขวาบนหน้า = วางลายเซ็น/ตรา/วันที่ตรงจุดนั้น';
  $('annUndo').disabled = !D.undo.length; $('annRedo').disabled = !D.redo.length;
}
async function loadFontList() {
  let fams = [];
  if (host.host) { try { fams = await (await api('/api/fonts')).json(); } catch {} }
  const prefer = ['TH Sarabun New', 'TH SarabunPSK', 'Sarabun', 'Tahoma', 'Leelawadee UI', 'Leelawadee', 'Angsana New', 'AngsanaUPC', 'Cordia New', 'Browallia New', 'TH Niramit AS', 'TH Charmonman', 'Arial', 'Times New Roman', 'Calibri', 'Cambria', 'Segoe UI', 'Courier New', 'Thonburi'];
  const set = new Set([...fams, ...Object.keys(BUNDLED_FONTS)]);
  const top = fams.length ? prefer.filter((f) => set.has(f)) : [...Object.keys(BUNDLED_FONTS), 'Tahoma', 'Leelawadee UI', 'Thonburi', 'Sarabun', 'Arial', 'Times New Roman', 'Courier New'];
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
  if (!c && !e.altKey && k === 's' && !a) { e.preventDefault(); $('annSig').click(); return; }
  if (!a) { if (e.key === 'Escape' && annTool) setAnnTool(false); return; }
  let handled = true;
  if (c && k === 'c') copyAnn(false);
  else if (c && k === 'x') copyAnn(true);
  else if (c && k === 'd') { copyAnn(false); pasteAnn(); }
  else if (e.key === 'Delete' || e.key === 'Backspace') deleteAnn();
  else if (e.key === 'Enter' || e.key === 'F2') startEdit();
  else if (e.key === 'Escape') { selectAnn(null); sigCloseProps(); }
  else if (e.key.startsWith('Arrow')) {
    const d = e.shiftKey ? 10 : 1;
    pushUndo();
    if (e.key === 'ArrowLeft') a.x -= d; if (e.key === 'ArrowRight') a.x += d;
    if (e.key === 'ArrowUp') a.y -= d; if (e.key === 'ArrowDown') a.y += d;
    touchPage(pageByUid(annSel.uid)); refresh();
  } else handled = false;
  if (handled) { e.preventDefault(); e.stopImmediatePropagation(); }
}, true);

// ------------------------------------------------------------------ scanner (สแกนเอกสาร) — scanic + photocopy filters
const SCAN_W = 1654;                 // output width in px for a portrait A4 page (~200 dpi)
const scan = {
  queue: [],          // [{ canvas (source photo), replace? }]
  cur: null,          // { src canvas, corners, out canvas (straightened), rot }
  pages: [],          // accepted: [{ canvas, filter, dark, bright, thumb }]
  filter: 'bw', dark: 50, bright: 50,
  editor: null,
  replace: null,      // when enhancing existing pages: [page uids]
};
try { const s = JSON.parse(localStorage.getItem('superpdf.scan') || '{}'); if (s.filter) scan.filter = s.filter; } catch {}

function openScanner(opts = {}) {
  scan.queue = []; scan.cur = null; scan.pages = []; scan.replace = opts.replace || null;
  $('scanner').classList.remove('hidden');
  $('scanDoneLbl').textContent = scan.replace ? 'แทนที่หน้าเดิม' : 'เสร็จ';
  showScanStep('empty');
  renderScanThumbs();
  if (opts.images) { scan.queue.push(...opts.images); nextScanImage(); }
}
function closeScanner() {
  if (scan.editor) { try { scan.editor.destroy(); } catch {} scan.editor = null; }
  $('scanner').classList.add('hidden');
  scan.queue = []; scan.cur = null; scan.pages = [];
}
function showScanStep(step) {
  $('scanEmpty').classList.toggle('hidden', step !== 'empty');
  $('scanCrop').classList.toggle('hidden', step !== 'crop');
  $('scanCropBar').classList.toggle('hidden', step !== 'crop');
  $('scanPreview').classList.toggle('hidden', step !== 'filter');
  $('scanFilterBar').classList.toggle('hidden', step !== 'filter');
  const n = scan.pages.length;
  $('scanStep').textContent = step === 'crop' ? 'ขั้นที่ 1: ครอปและดึงให้ตรง' : step === 'filter' ? 'ขั้นที่ 2: ปรับภาพ' : (n ? `สแกนแล้ว ${n} หน้า` : '');
  $('scanDone').disabled = !n;
  $('scanDoneLbl').textContent = (scan.replace ? 'แทนที่หน้าเดิม' : 'เสร็จ') + (n ? ` (${n})` : '');
}

// photo file -> canvas (EXIF orientation applied by the browser), long side limited for memory
async function photoToCanvas(file, maxSide = 3000) {
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch { bmp = await loadImageEl(file); }
  const w0 = bmp.width || bmp.naturalWidth, h0 = bmp.height || bmp.naturalHeight;
  const k = Math.min(1, maxSide / Math.max(w0, h0));
  const c = document.createElement('canvas'); c.width = Math.round(w0 * k); c.height = Math.round(h0 * k);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  if (bmp.close) bmp.close();
  return c;
}
async function addScanFiles(files) {
  if (!files.length) return;
  busy('กำลังเปิดรูป …');
  try { for (const f of files) scan.queue.push({ canvas: await photoToCanvas(f) }); }
  catch (e) { toast('เปิดรูปไม่ได้: ' + e.message, 'err'); }
  unbusy();
  if (!scan.cur) nextScanImage();
}
function nextScanImage() {
  const item = scan.queue.shift();
  if (!item) { scan.cur = null; showScanStep('empty'); return; }
  startCrop(item);
}
const fullCorners = (c, inset = 0) => ({
  topLeft: { x: inset, y: inset }, topRight: { x: c.width - inset, y: inset },
  bottomRight: { x: c.width - inset, y: c.height - inset }, bottomLeft: { x: inset, y: c.height - inset },
});
async function detectCorners(c) {
  try {
    const r = await window.scanic.scanDocument(c, { mode: 'detect', maxProcessingDimension: 800 });
    if (r && r.success && r.corners) return r.corners;
  } catch {}
  return null;
}
async function startCrop(item) {
  scan.cur = { src: item.canvas, rot: 0, noCrop: !!item.noCrop, uid: item.uid };
  if (item.noCrop) { scan.cur.corners = fullCorners(item.canvas); return finishCrop(); }
  showScanStep('crop');
  busy('กำลังหาขอบกระดาษ …');
  const found = await detectCorners(item.canvas);
  unbusy();
  scan.cur.corners = found || fullCorners(item.canvas, Math.round(Math.min(item.canvas.width, item.canvas.height) * 0.05));
  mountEditor(scan.cur.corners);
  if (!found) toast('หาขอบกระดาษอัตโนมัติไม่พบ — กรุณาลากมุมเอง', '', [], 3500);
}
function mountEditor(corners) {
  if (scan.editor) { try { scan.editor.destroy(); } catch {} scan.editor = null; }
  const host = $('scanCrop'); host.innerHTML = '';
  scan.editor = window.scanic.createCornerEditor({
    container: host, image: scan.cur.src, corners,
    toolbar: { enabled: false }, magnifier: { enabled: true, zoom: 2.5, size: 120 },
    theme: { accent: '#3b82f6', mask: 'rgba(0,0,0,.45)', handleSize: 22 },
    onChange: (c) => { scan.cur.corners = c; },
  });
}
async function finishCrop() {
  const c = scan.cur;
  if (scan.editor) { c.corners = scan.editor.getCorners(); try { scan.editor.destroy(); } catch {} scan.editor = null; }
  busy('กำลังดึงเอกสารให้ตรง …');
  try {
    let out = c.src;
    const full = fullCorners(c.src);
    const isFull = ['topLeft', 'topRight', 'bottomRight', 'bottomLeft'].every((k) => Math.hypot(c.corners[k].x - full[k].x, c.corners[k].y - full[k].y) < 2);
    if (!isFull) {
      const r = await window.scanic.extractDocument(c.src, c.corners, { output: 'canvas' });
      if (!r || !r.success || !r.output) throw new Error(r && r.message || 'extract failed');
      out = r.output;
    }
    c.flat = normalizeSize(out);
  } catch (e) { unbusy(); toast('ดึงภาพไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  showScanStep('filter');
  setFilterChip();
  renderScanPreview();
}
// resize so the short side matches A4 width at ~200 dpi (keeps aspect ratio)
function normalizeSize(src) {
  const portrait = src.height >= src.width;
  const k = SCAN_W / (portrait ? src.width : src.height);
  const c = document.createElement('canvas');
  c.width = Math.round(src.width * k); c.height = Math.round(src.height * k);
  const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(src, 0, 0, c.width, c.height);
  return c;
}
function rotateCanvas(src, deg) {
  if (!deg) return src;
  const c = document.createElement('canvas');
  const sw = deg % 180 !== 0;
  c.width = sw ? src.height : src.width; c.height = sw ? src.width : src.height;
  const g = c.getContext('2d');
  g.translate(c.width / 2, c.height / 2); g.rotate(deg * Math.PI / 180); g.drawImage(src, -src.width / 2, -src.height / 2);
  return c;
}

// ---- photocopy filters
// background (paper brightness) map: block maximum -> dilate -> blur -> bilinear upsample; removes shadows & uneven light
function backgroundMap(ch, W, H, B = 8) {
  const w = Math.ceil(W / B), h = Math.ceil(H / B);
  let m = new Float32Array(w * h);
  for (let by = 0; by < h; by++) for (let bx = 0; bx < w; bx++) {
    let mx = 0;
    const y1 = Math.min(H, by * B + B), x1 = Math.min(W, bx * B + B);
    for (let y = by * B; y < y1; y++) { const row = y * W; for (let x = bx * B; x < x1; x++) { const v = ch[row + x]; if (v > mx) mx = v; } }
    m[by * w + bx] = mx;
  }
  const pass = (src, r, fn) => { // separable filter (max or mean)
    const t = new Float32Array(w * h), o = new Float32Array(w * h);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let a = fn === 'max' ? 0 : 0, n = 0;
      for (let k = -r; k <= r; k++) { const xx = Math.min(w - 1, Math.max(0, x + k)); const v = src[y * w + xx]; if (fn === 'max') { if (v > a) a = v; } else { a += v; n++; } }
      t[y * w + x] = fn === 'max' ? a : a / n;
    }
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let a = 0, n = 0;
      for (let k = -r; k <= r; k++) { const yy = Math.min(h - 1, Math.max(0, y + k)); const v = t[yy * w + x]; if (fn === 'max') { if (v > a) a = v; } else { a += v; n++; } }
      o[y * w + x] = fn === 'max' ? a : a / n;
    }
    return o;
  };
  m = pass(m, 3, 'max');
  m = pass(m, 4, 'mean');
  const out = new Float32Array(W * H);
  for (let y = 0; y < H; y++) {
    const fy = Math.min(h - 1.001, Math.max(0, (y + 0.5) / B - 0.5)), y0 = fy | 0, ty = fy - y0;
    for (let x = 0; x < W; x++) {
      const fx = Math.min(w - 1.001, Math.max(0, (x + 0.5) / B - 0.5)), x0 = fx | 0, tx = fx - x0;
      const i = y0 * w + x0;
      const a = m[i] + (m[i + 1] - m[i]) * tx, b = m[i + w] + (m[i + w + 1] - m[i + w]) * tx;
      out[y * W + x] = Math.max(40, a + (b - a) * ty);
    }
  }
  return out;
}
function applyScanFilter(src, filter, dark = 50, bright = 50) {
  const W = src.width, H = src.height;
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const g = c.getContext('2d'); g.drawImage(src, 0, 0);
  const img = g.getImageData(0, 0, W, H), d = img.data, N = W * H;
  const dk = (dark - 50) / 50, br = (bright - 50) / 50;   // -1 .. 1
  if (filter === 'orig') {
    const con = 1 + dk * 0.5, off = br * 50;
    for (let i = 0; i < d.length; i += 4) for (let k = 0; k < 3; k++) d[i + k] = (d[i + k] - 128) * con + 128 + off;
  } else if (filter === 'clean') {
    const bgs = [0, 1, 2].map((k) => { const ch = new Uint8Array(N); for (let i = 0, j = k; i < N; i++, j += 4) ch[i] = d[j]; return backgroundMap(ch, W, H); });
    const white = 236 - br * 20, black = 40 + dk * 40;
    for (let i = 0, j = 0; i < N; i++, j += 4) {
      let r = d[j] / bgs[0][i] * 255, gg = d[j + 1] / bgs[1][i] * 255, b = d[j + 2] / bgs[2][i] * 255;
      const l = (r + gg + b) / 3;
      r = l + (r - l) * 1.35; gg = l + (gg - l) * 1.35; b = l + (b - l) * 1.35;          // a bit more colour for ink / stamps
      d[j] = (r - black) * 255 / (white - black); d[j + 1] = (gg - black) * 255 / (white - black); d[j + 2] = (b - black) * 255 / (white - black);
    }
  } else {
    const lum = new Uint8Array(N);
    for (let i = 0, j = 0; i < N; i++, j += 4) lum[i] = (d[j] * 299 + d[j + 1] * 587 + d[j + 2] * 114) / 1000;
    const bg = backgroundMap(lum, W, H);
    if (filter === 'gray') {
      const white = 238 - br * 22, black = 60 + dk * 55;
      for (let i = 0, j = 0; i < N; i++, j += 4) {
        const n = lum[i] / bg[i] * 255;
        const v = (n - black) * 255 / (white - black);
        d[j] = d[j + 1] = d[j + 2] = v;
      }
    } else { // bw photocopy
      const T = 200 + dk * 32 + br * -18;                       // darker -> more ink kept
      for (let i = 0, j = 0; i < N; i++, j += 4) {
        const v = lum[i] / bg[i] * 255 < T ? 0 : 255;
        d[j] = d[j + 1] = d[j + 2] = v;
      }
    }
  }
  g.putImageData(img, 0, 0);
  return c;
}

let scanPrevTimer = 0;
function renderScanPreview() {
  clearTimeout(scanPrevTimer);
  scanPrevTimer = setTimeout(() => {
    const c = scan.cur; if (!c || !c.flat) return;
    const src = rotateCanvas(c.flat, c.rot);
    const k = Math.min(1, 900 / Math.max(src.width, src.height));   // fast preview
    const small = document.createElement('canvas'); small.width = Math.round(src.width * k); small.height = Math.round(src.height * k);
    small.getContext('2d').drawImage(src, 0, 0, small.width, small.height);
    const out = applyScanFilter(small, scan.filter, scan.dark, scan.bright);
    const cv = $('scanCanvas'); cv.width = out.width; cv.height = out.height;
    cv.getContext('2d').drawImage(out, 0, 0);
  }, 60);
}
function setFilterChip() {
  document.querySelectorAll('#scanFilters .fchip').forEach((b) => b.classList.toggle('active', b.dataset.f === scan.filter));
  $('scanDark').value = scan.dark; $('scanBright').value = scan.bright;
}
function acceptScanPage() {
  const c = scan.cur; if (!c || !c.flat) return;
  busy('กำลังปรับภาพ …');
  setTimeout(() => {
    const out = applyScanFilter(rotateCanvas(c.flat, c.rot), scan.filter, scan.dark, scan.bright);
    const t = document.createElement('canvas'); const k = 64 / out.height; t.width = Math.round(out.width * k); t.height = 64;
    t.getContext('2d').drawImage(out, 0, 0, t.width, t.height);
    scan.pages.push({ canvas: out, filter: scan.filter, thumb: t.toDataURL('image/jpeg', 0.7), uid: c.uid });
    try { localStorage.setItem('superpdf.scan', JSON.stringify({ filter: scan.filter })); } catch {}
    unbusy();
    renderScanThumbs();
    nextScanImage();
  }, 20);
}
function renderScanThumbs() {
  $('scanThumbs').replaceChildren(...scan.pages.map((p, i) => {
    const d = document.createElement('div'); d.className = 'sthumb';
    d.innerHTML = `<img src="${p.thumb}"><b>${i + 1}</b><button title="ลบ">✕</button>`;
    d.querySelector('button').onclick = () => { scan.pages.splice(i, 1); renderScanThumbs(); if (!scan.cur) showScanStep('empty'); else showScanStep(scan.cur.flat ? 'filter' : 'crop'); };
    return d;
  }));
  const hide = !!scan.replace;
  $('scanMoreCam').classList.toggle('hidden', hide); $('scanMorePick').classList.toggle('hidden', hide);
  if (!scan.cur) showScanStep('empty');
}

// accepted pages -> PDF (black & white pages stored as CCITT G4: tiny files)
async function scanPagesToPdf(pages) {
  const doc = await PDFDocument.create();
  const { PDFName, PDFNumber } = window.PDFLib;
  for (const p of pages) {
    const cv = p.canvas, W = cv.width, H = cv.height;
    const pw = p.wPt || 595.28 * (W <= H ? 1 : W / H), ph = p.hPt || pw * H / W;
    const page = doc.addPage([pw, ph]);
    if (p.filter === 'bw') {
      const px = cv.getContext('2d').getImageData(0, 0, W, H).data;
      const bits = new Uint8Array(W * H);
      for (let i = 0, j = 0; i < bits.length; i++, j += 4) bits[i] = px[j] < 128 ? 1 : 0;
      const g4 = TiffEnc.g4(bits, W, H);
      const ctx = doc.context;
      const dict = ctx.obj({
        Type: 'XObject', Subtype: 'Image', Width: W, Height: H, ColorSpace: 'DeviceGray', BitsPerComponent: 1,
        Filter: 'CCITTFaxDecode', DecodeParms: { K: -1, Columns: W, Rows: H, BlackIs1: false },
      });
      const raw = window.PDFLib.PDFRawStream.of(dict, g4);
      const ref = ctx.register(raw);
      const name = page.node.newXObject('Scan', ref);
      page.pushOperators(
        window.PDFLib.pushGraphicsState(),
        window.PDFLib.concatTransformationMatrix(pw, 0, 0, ph, 0, 0),
        window.PDFLib.drawObject(name),
        window.PDFLib.popGraphicsState(),
      );
    } else {
      const img = await doc.embedJpg(await canvasToBytes(cv, 'image/jpeg', p.filter === 'orig' ? 0.85 : 0.82));
      page.drawImage(img, { x: 0, y: 0, width: pw, height: ph });
    }
  }
  return doc.save();
}
async function finishScan() {
  if (!scan.pages.length) return;
  busy('กำลังสร้าง PDF …');
  let bytes;
  try { bytes = await scanPagesToPdf(scan.pages); }
  catch (e) { unbusy(); toast('สร้าง PDF ไม่สำเร็จ: ' + e.message, 'err', [], 9000); return; }
  unbusy();
  const now = new Date(), stamp = `${now.getFullYear()}${pad(now.getMonth() + 1, 2)}${pad(now.getDate(), 2)}_${pad(now.getHours(), 2)}${pad(now.getMinutes(), 2)}`;
  const nDone = scan.pages.length;
  if (scan.replace) {
    const src = await openPdfSource('สแกน_' + stamp + '.pdf', bytes);
    pushUndo();
    let lostText = false;
    scan.pages.forEach((sp, k) => {
      const i = D.pages.findIndex((p) => p.uid === sp.uid);
      if (i < 0) return;
      if (D.pages[i].annots && D.pages[i].annots.length) lostText = true;
      D.pages[i] = { uid: uid(), src: src.id, index: k, rot: 0, av: 1, annots: [] };
    });
    if (lostText) toast('ข้อความที่เพิ่มไว้ในหน้าที่ปรับภาพถูกนำออก (กด เลิกทำ เพื่อย้อนกลับ)', '', [['เลิกทำ', undo]], 7000);
    D.dirty = true;
    closeScanner();
    refresh();
    toast(`ปรับภาพ ${nDone} หน้าแล้ว`, 'ok', [], 2500);
    return;
  }
  const n = scan.pages.length;
  closeScanner();
  await insertOrOpen([{ name: 'สแกน_' + stamp + '.pdf', bytes }], insertionIndex());
  toast(`เพิ่มหน้าสแกน ${n} หน้าแล้ว`, 'ok', [], 2500);
}
// enhance existing pages (render → filter screen, no crop by default)
async function enhancePages(list) {
  if (!list.length) return;
  busy('กำลังเตรียมหน้า …');
  const imgs = [];
  for (const p of list) { const c = await renderPageCanvas(p, 200 / 72, false); imgs.push({ canvas: c, noCrop: true, uid: p.uid }); }
  unbusy();
  openScanner({ replace: list.map((p) => p.uid), images: imgs });
}

$('btnScan').onclick = () => openScanner();
$('btnScanPick').onclick = (e) => { e.stopPropagation(); openScanner(); };
$('scanClose').onclick = async () => {
  if (scan.pages.length) {
    const r = await dialog({ title: 'ปิดหน้าสแกน', body: `<p>มี ${scan.pages.length} หน้าที่สแกนไว้ยังไม่ได้ใส่ลงเอกสาร ต้องการทิ้งหรือไม่?</p>`, buttons: [['cancel', 'ยกเลิก'], ['drop', 'ทิ้ง'], ['keep', 'ใส่ลงเอกสาร', true]] });
    if (!r.btn) return;
    if (r.btn === 'keep') return finishScan();
  }
  closeScanner();
};
$('scanDone').onclick = finishScan;
for (const id of ['scanCamBtn', 'scanMoreCam']) $(id).onclick = () => $('scanCam').click();
for (const id of ['scanPickBtn', 'scanMorePick']) $(id).onclick = () => $('scanPick').click();
for (const id of ['scanCam', 'scanPick']) $(id).onchange = async (e) => { const fs = [...e.target.files]; e.target.value = ''; await addScanFiles(fs); };
$('scanAuto').onclick = async () => {
  busy('กำลังหาขอบกระดาษ …'); const f = await detectCorners(scan.cur.src); unbusy();
  if (f) { scan.cur.corners = f; mountEditor(f); } else toast('หาขอบอัตโนมัติไม่พบ', '', [], 2500);
};
$('scanFull').onclick = () => { scan.cur.corners = fullCorners(scan.cur.src); mountEditor(scan.cur.corners); };
$('scanCropOk').onclick = finishCrop;
$('scanRecrop').onclick = () => { const c = scan.cur; c.flat = null; showScanStep('crop'); mountEditor(c.corners); };
$('scanRotL').onclick = () => { scan.cur.rot = (scan.cur.rot + 270) % 360; renderScanPreview(); };
$('scanRotR').onclick = () => { scan.cur.rot = (scan.cur.rot + 90) % 360; renderScanPreview(); };
$('scanAccept').onclick = acceptScanPage;
document.querySelectorAll('#scanFilters .fchip').forEach((b) => { b.onclick = () => { scan.filter = b.dataset.f; setFilterChip(); renderScanPreview(); }; });
$('scanDark').oninput = (e) => { scan.dark = +e.target.value; renderScanPreview(); };
$('scanBright').oninput = (e) => { scan.bright = +e.target.value; renderScanPreview(); };

// ------------------------------------------------------------------ print (พิมพ์)
const IS_IOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
async function printDialog(preset) {
  if (!D.pages.length) { toast('ยังไม่มีเอกสารให้พิมพ์'); return; }
  const nSel = D.selected.size, total = D.pages.length;
  const r = await dialog({
    title: 'พิมพ์เอกสาร',
    body: `<div class="field"><label>หน้าที่จะพิมพ์</label>
      <label class="opt"><input type="radio" name="rng" value="all" ${preset !== 'sel' ? 'checked' : ''}><span>ทั้งเอกสาร<small>${total} หน้า</small></span></label>
      <label class="opt"><input type="radio" name="rng" value="sel" ${preset === 'sel' ? 'checked' : ''} ${nSel ? '' : 'disabled'}><span>เฉพาะหน้าที่เลือก<small>${nSel ? nSel + ' หน้า' : 'ยังไม่ได้เลือกหน้า'}</small></span></label>
      <label class="opt"><input type="radio" name="rng" value="range"><span>ช่วงหน้า<small><input name="from" type="text" inputmode="numeric" value="1" style="width:56px;height:30px"> ถึง <input name="to" type="text" inputmode="numeric" value="${total}" style="width:56px;height:30px"></small></span></label>
      </div>${IS_IOS ? '<p>iPhone/iPad: เลือกเครื่องพิมพ์ (AirPrint) ได้ในหน้าต่างถัดไป</p>' : ''}`,
    buttons: [['cancel', 'ยกเลิก'], ['ok', 'ถัดไป', true]],
  });
  if (!r.btn) return;
  let list = D.pages;
  if (r.values.rng === 'sel') list = selPages();
  if (r.values.rng === 'range') {
    const a = Math.max(1, Math.min(total, parseInt(r.values.from, 10) || 1)), b = Math.max(a, Math.min(total, parseInt(r.values.to, 10) || total));
    list = D.pages.slice(a - 1, b);
  }
  if (!list.length) return;
  if (IS_IOS) return printIOS(list);
  return printPdfFrame(list);
}
// Windows / desktop: print the real PDF (vector, sharp text) through the browser's PDF viewer
async function printPdfFrame(list) {
  let bytes;
  try { busy('กำลังเตรียมพิมพ์ …'); bytes = (await buildPdfInner(list, true)).bytes; }
  catch (e) { unbusy(); toast('เตรียมพิมพ์ไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  document.querySelectorAll('iframe.printframe').forEach((f) => f.remove());
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
  const f = document.createElement('iframe');
  f.className = 'printframe';
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none';
  f.onload = () => { setTimeout(() => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch { printImages(list); } }, 300); };
  f.src = url;
  document.body.appendChild(f);
  setTimeout(() => URL.revokeObjectURL(url), 10 * 60 * 1000);
}
// iPhone / iPad: render pages as images, then AirPrint via window.print(); or share the PDF → Print
async function renderPrintImages(list) {
  const area = $('printArea'); area.innerHTML = '';
  for (let i = 0; i < list.length; i++) {
    busy(`กำลังเตรียมพิมพ์ … หน้า ${i + 1}/${list.length}`);
    const c = await renderPageCanvas(list[i], (IS_TOUCH ? 150 : 200) / 72);
    const pb = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.9));
    c.width = c.height = 0;
    if (!pb) throw new Error('สร้างภาพสำหรับพิมพ์ไม่สำเร็จ (หน่วยความจำไม่พอ)');
    const im = new Image(); im.src = URL.createObjectURL(pb);
    area.appendChild(im);
  }
  await Promise.all([...area.images || area.querySelectorAll('img')].map((im) => im.decode().catch(() => {})));
  unbusy();
}
function doWindowPrint() {
  document.body.classList.add('printing');
  const done = () => { document.body.classList.remove('printing'); window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  window.print();
  setTimeout(done, 60000);
}
async function printImages(list) { await renderPrintImages(list); doWindowPrint(); }
async function printIOS(list) {
  try { await renderPrintImages(list); } catch (e) { unbusy(); toast('เตรียมพิมพ์ไม่สำเร็จ: ' + e.message, 'err'); return; }
  let pdfFile = null;
  try { pdfFile = new File([(await buildPdfInner(list, true)).bytes], (D.name || 'เอกสาร') + '.pdf', { type: 'application/pdf' }); } catch {}
  const canShare = pdfFile && navigator.canShare && navigator.canShare({ files: [pdfFile] });
  const r = await dialog({
    title: 'พร้อมพิมพ์แล้ว',
    body: `<p>${list.length} หน้า — กด “พิมพ์” แล้วเลือกเครื่องพิมพ์ (AirPrint)</p>${canShare ? '<p><small>ถ้าปุ่มพิมพ์ไม่ทำงาน ให้กด “ผ่านเมนูแชร์” แล้วเลือก “พิมพ์ (Print)”</small></p>' : ''}`,
    buttons: canShare ? [['cancel', 'ยกเลิก'], ['share', 'ผ่านเมนูแชร์'], ['print', 'พิมพ์', true]] : [['cancel', 'ยกเลิก'], ['print', 'พิมพ์', true]],
  });
  if (r.btn === 'print') doWindowPrint();
  else if (r.btn === 'share') { try { await navigator.share({ files: [pdfFile] }); } catch {} }
}
$('btnPrint').onclick = () => printDialog();

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
let lastSave = '';   // v1.9.1: result of the last saveBytes(): 'ok' | 'cancel' | 'error'
async function saveBytes(bytes, name, extn = 'pdf', onSaved = null) {
  lastSave = 'error';
  name = cleanName(name);
  if (host.host) {
    try {
      busy('รอเลือกตำแหน่งบันทึกไฟล์ …');
      const dir = D && D.dir ? '&dir=' + encodeURIComponent(D.dir) : '';
      const r = await (await api(`/api/save?ext=${extn}&name=${encodeURIComponent(name)}${dir}`, { method: 'POST', body: bytes })).json();
      unbusy();
      if (r.ok) {
        if (onSaved) onSaved(r.path);
        lastSave = 'ok';
        toast('บันทึกแล้ว: ' + r.path, 'ok', [
          ['เปิดไฟล์', () => api('/api/openpdf', { method: 'POST', body: r.path })],
          ['เปิดโฟลเดอร์', () => api('/api/reveal', { method: 'POST', body: r.path })],
        ], 8000);
        return true;
      }
      if (r.error) toast('บันทึกไม่สำเร็จ: ' + r.error, 'err', [], 9000); else lastSave = 'cancel';
      return false;
    } catch (e) { unbusy(); lastSave = 'error'; toast('บันทึกไม่สำเร็จ: ' + e.message, 'err'); return false; }
  }
  const okd = await deliver([{ name, bytes }]);
  lastSave = okd ? 'ok' : 'cancel';
  return okd;
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
    { label: 'ย่อขนาดไฟล์ (สูง / กลาง / ต่ำ)…', icon: 'compress', disabled: !D.pages.length, action: () => compressDialog() },
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
  // v1.9.1: read each file on its own; report files that cannot be read and continue with the rest
  const dropped = [...e.dataTransfer.files];
  const rs = await Promise.allSettled(dropped.map(readFile));
  const fs = [], bad = [];
  rs.forEach((r, k) => { if (r.status === 'fulfilled') fs.push(r.value); else bad.push(dropped[k].name || ('ไฟล์ที่ ' + (k + 1))); });
  if (bad.length) toast(`อ่านไฟล์ไม่ได้ ${bad.length} ไฟล์: ${bad.join(', ')}${fs.length ? ` — นำเข้าเฉพาะอีก ${fs.length} ไฟล์ที่อ่านได้` : ''}`, 'err', [], 10000);
  if (!fs.length) return;
  try {
  if (tabEl) { const t = docs.find((x) => x.id === tabEl.dataset.id); switchTo(t); await insertFiles(fs); }
  else if (onTabbar || !D.pages.length) await openPicked(fs);
  else await insertFiles(fs, d ? d.i : D.pages.length);
  } catch (err) { unbusy(); toast('นำเข้าไฟล์ไม่สำเร็จ: ' + (err.message || err), 'err', [], 9000); }
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
  else if (c && k === 'p') { e.preventDefault(); printDialog(); }
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
$('reader').addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); const cur = parseFloat($('zoomLabel').textContent) / 100 || 1; const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5]; const z = e.deltaY < 0 ? steps.find((x) => x > cur + 0.001) : [...steps].reverse().find((x) => x < cur - 0.001); if (z) setZoom(z, { x: e.clientX, y: e.clientY }); } }, { passive: false });
window.addEventListener('resize', () => { if (mode === 'read' && zoom === 0) { clearTimeout(window.__rz); window.__rz = setTimeout(() => renderReader(), 200); } });
window.addEventListener('beforeunload', (e) => { if (docs.some((d) => d.dirty && d.pages.length)) { e.preventDefault(); e.returnValue = ''; } });

// ------------------------------------------------------------------ v1.7 image editor (แก้ไขภาพ) — select object / delete / move, eraser, pen, despeckle, edit in Paint
const IE = {
  open: false, uid: null, cv: null, g: null, view: 1, tool: 'select', sel: null, float: null,
  undo: [], redo: [], size: 24, color: '#000000', strength: 40, group: 6, wPt: 0, hPt: 0,
  changed: false, ink: null, paint: null, clip: null, op: null, pointers: new Map(), pinch: null,
};
const IE_MAX_UNDO = IS_TOUCH ? 6 : 20;
const ieEl = (id) => document.getElementById(id);

// --- pixel helpers
const lumOf = (d, j) => (d[j] * 299 + d[j + 1] * 587 + d[j + 2] * 114) / 1000;
const isInkPx = (d, j) => { const l = lumOf(d, j); if (l < 165) return true; const mx = Math.max(d[j], d[j + 1], d[j + 2]), mn = Math.min(d[j], d[j + 1], d[j + 2]); return mx - mn > 70 && l < 235; };

// render a page in its own orientation (without the extra user rotation) so text boxes keep their coordinates
async function renderPageBase(p, maxPx) {
  const pg = await sources.get(p.src).pdf.getPage(p.index + 1);
  const b = pg.getViewport({ scale: 1, rotation: pg.rotate });
  const sc = Math.min(200 / 72, Math.sqrt(maxPx / (b.width * b.height)));
  const vp = pg.getViewport({ scale: sc, rotation: pg.rotate });
  const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const g = c.getContext('2d', { willReadFrequently: true }); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await pg.render({ canvasContext: g, viewport: vp }).promise;
  return { canvas: c, wPt: b.width, hPt: b.height };
}

// remove isolated small dark specks (scanner dust). Thai tone marks / dots next to letters are kept because they are not isolated.
function despeckleCanvas(cv, strength = 40, rect = null) { // two passes: specks next to other specks become isolated after the first pass
  let n = 0; for (let pass = 0; pass < 2; pass++) { const k = despecklePass(cv, strength, rect); n += k; if (!k) break; } return n;
}
function despecklePass(cv, strength, rect) {
  const W = cv.width, H = cv.height, g = cv.getContext('2d', { willReadFrequently: true });
  const x0 = rect ? Math.max(0, Math.floor(rect.x)) : 0, y0 = rect ? Math.max(0, Math.floor(rect.y)) : 0;
  const w = rect ? Math.min(W - x0, Math.ceil(rect.w)) : W, h = rect ? Math.min(H - y0, Math.ceil(rect.h)) : H;
  if (w < 3 || h < 3) return 0;
  const img = g.getImageData(x0, y0, w, h), d = img.data, N = w * h;
  const ink = new Uint8Array(N);
  for (let i = 0, j = 0; i < N; i++, j += 4) ink[i] = (lumOf(d, j) < 170 || (Math.max(d[j], d[j + 1], d[j + 2]) - Math.min(d[j], d[j + 1], d[j + 2]) > 80 && lumOf(d, j) < 230)) ? 1 : 0;
  // integral image of ink
  const S = new Int32Array((w + 1) * (h + 1));
  for (let y = 0; y < h; y++) { let row = 0; for (let x = 0; x < w; x++) { row += ink[y * w + x]; S[(y + 1) * (w + 1) + x + 1] = S[y * (w + 1) + x + 1] + row; } }
  const boxSum = (ax, ay, bx, by) => { ax = Math.max(0, ax); ay = Math.max(0, ay); bx = Math.min(w - 1, bx); by = Math.min(h - 1, by); return S[(by + 1) * (w + 1) + bx + 1] - S[ay * (w + 1) + bx + 1] - S[(by + 1) * (w + 1) + ax] + S[ay * (w + 1) + ax]; };
  const k = W / 1654, maxArea = Math.max(2, Math.round(k * k * (4 + strength * 1.2))), margin = Math.max(3, Math.round(W * 0.006));
  const lab = new Int32Array(N), stack = new Int32Array(N);
  let removed = 0, cur = 0;
  const kill = [];
  for (let s = 0; s < N; s++) {
    if (!ink[s] || lab[s]) continue;
    cur++; let sp = 0, area = 0, minx = w, miny = h, maxx = 0, maxy = 0, big = false;
    stack[sp++] = s; lab[s] = cur;
    while (sp) {
      const q = stack[--sp]; area++;
      const qx = q % w, qy = (q - qx) / w;
      if (qx < minx) minx = qx; if (qx > maxx) maxx = qx; if (qy < miny) miny = qy; if (qy > maxy) maxy = qy;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = qy + dy; if (ny < 0 || ny >= h) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = qx + dx; if (nx < 0 || nx >= w) continue;
          const nq = ny * w + nx; if (ink[nq] && !lab[nq]) { lab[nq] = cur; stack[sp++] = nq; }
        }
      }
      if (area > maxArea) big = true;
    }
    if (big) continue;
    if (boxSum(minx - margin, miny - margin, maxx + margin, maxy + margin) === area) kill.push(cur);
  }
  if (!kill.length) return 0;
  const ks = new Uint8Array(cur + 1); for (const c of kill) ks[c] = 1;
  for (let i = 0, j = 0; i < N; i++, j += 4) if (ks[lab[i]]) { d[j] = d[j + 1] = d[j + 2] = 255; d[j + 3] = 255; }
  g.putImageData(img, x0, y0);
  removed = kill.length;
  return removed;
}

// low-resolution ink map + dilation, used to pick an "object" (a stamp, emblem, a line of text …) with one click
function ieInkMap() {
  if (IE.ink && IE.ink.group === IE.group) return IE.ink;
  const cv = IE.cv, W = cv.width, H = cv.height;
  const f = Math.max(1, Math.round(W / 1000));
  const w = Math.ceil(W / f), h = Math.ceil(H / f);
  const d = IE.g.getImageData(0, 0, W, H).data;
  const ink = new Uint8Array(w * h);
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const j = (y * W + x) * 4; if (isInkPx(d, j)) ink[((y / f) | 0) * w + ((x / f) | 0)] = 1; }
  const r = Math.max(1, Math.round(IE.group * (W / 1654) / f));
  // separable dilation with running counts
  const tmp = new Uint8Array(w * h), dil = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) { let c = 0; const o = y * w; for (let x = -r; x < w + r; x++) { if (x + r < w && x + r >= 0) c += ink[o + x + r]; if (x - r - 1 >= 0 && x - r - 1 < w) c -= ink[o + x - r - 1]; if (x >= 0 && x < w) tmp[o + x] = c > 0 ? 1 : 0; } }
  for (let x = 0; x < w; x++) { let c = 0; for (let y = -r; y < h + r; y++) { if (y + r < h && y + r >= 0) c += tmp[(y + r) * w + x]; if (y - r - 1 >= 0 && y - r - 1 < h) c -= tmp[(y - r - 1) * w + x]; if (y >= 0 && y < h) dil[y * w + x] = c > 0 ? 1 : 0; } }
  IE.ink = { f, w, h, ink, dil, group: IE.group };
  return IE.ink;
}
// component under (px,py) → selection { x,y,w,h, mask (low-res, f) }
function iePickObject(px, py) {
  const M = ieInkMap(), { f, w, h, dil } = M;
  let sx = Math.floor(px / f), sy = Math.floor(py / f);
  if (sx < 0 || sy < 0 || sx >= w || sy >= h) return null;
  if (!dil[sy * w + sx]) { // tap near an object
    let best = null; const R = Math.round(14 / f) + 2;
    for (let rr = 1; rr <= R && !best; rr++) for (let dy = -rr; dy <= rr && !best; dy++) for (let dx = -rr; dx <= rr; dx++) {
      const x = sx + dx, y = sy + dy; if (x < 0 || y < 0 || x >= w || y >= h) continue;
      if (dil[y * w + x]) { best = [x, y]; break; }
    }
    if (!best) return null; [sx, sy] = best;
  }
  const comp = new Uint8Array(w * h), stack = new Int32Array(w * h);
  let sp = 0, minx = w, miny = h, maxx = 0, maxy = 0;
  stack[sp++] = sy * w + sx; comp[sy * w + sx] = 1;
  while (sp) {
    const q = stack[--sp], qx = q % w, qy = (q - qx) / w;
    if (qx < minx) minx = qx; if (qx > maxx) maxx = qx; if (qy < miny) miny = qy; if (qy > maxy) maxy = qy;
    if (qx > 0 && dil[q - 1] && !comp[q - 1]) { comp[q - 1] = 1; stack[sp++] = q - 1; }
    if (qx < w - 1 && dil[q + 1] && !comp[q + 1]) { comp[q + 1] = 1; stack[sp++] = q + 1; }
    if (qy > 0 && dil[q - w] && !comp[q - w]) { comp[q - w] = 1; stack[sp++] = q - w; }
    if (qy < h - 1 && dil[q + w] && !comp[q + w]) { comp[q + w] = 1; stack[sp++] = q + w; }
  }
  // fill holes (e.g. the inside of a round emblem / stamp)
  const bw = maxx - minx + 1, bh = maxy - miny + 1, out = new Uint8Array(bw * bh);
  const st2 = new Int32Array(bw * bh); sp = 0;
  const push = (x, y) => { const i = y * bw + x; if (out[i] || comp[(y + miny) * w + x + minx]) return; out[i] = 2; st2[sp++] = i; };
  for (let x = 0; x < bw; x++) { push(x, 0); push(x, bh - 1); }
  for (let y = 0; y < bh; y++) { push(0, y); push(bw - 1, y); }
  while (sp) { const i = st2[--sp], x = i % bw, y = (i - x) / bw; if (x > 0) push(x - 1, y); if (x < bw - 1) push(x + 1, y); if (y > 0) push(x, y - 1); if (y < bh - 1) push(x, y + 1); }
  for (let i = 0; i < out.length; i++) out[i] = out[i] === 2 ? 0 : 1;
  return { x: minx * f, y: miny * f, w: Math.min(IE.cv.width - minx * f, bw * f), h: Math.min(IE.cv.height - miny * f, bh * f), mask: out, mw: bw, mh: bh, f };
}
const ieMaskAt = (s, x, y) => !s.mask || s.mask[Math.min(s.mh - 1, Math.floor(y / s.f)) * s.mw + Math.min(s.mw - 1, Math.floor(x / s.f))];
function ieUnionSel(a, b) {
  if (!a) return b; if (!b) return a;
  const f = a.f || b.f || 1;
  const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y), x2 = Math.max(a.x + a.w, b.x + b.w), y2 = Math.max(a.y + a.h, b.y + b.h);
  const mw = Math.ceil((x2 - x) / f), mh = Math.ceil((y2 - y) / f), mask = new Uint8Array(mw * mh);
  for (const s of [a, b]) for (let my = 0; my < mh; my++) for (let mx = 0; mx < mw; mx++) {
    const px = x + mx * f - s.x, py = y + my * f - s.y;
    if (px >= 0 && py >= 0 && px < s.w && py < s.h && ieMaskAt(s, px, py)) mask[my * mw + mx] = 1;
  }
  return { x, y, w: x2 - x, h: y2 - y, mask, mw, mh, f };
}

// --- undo inside the editor
function iePushUndo() {
  IE.undo.push({ w: IE.cv.width, h: IE.cv.height, hPt: IE.hPt, data: IE.g.getImageData(0, 0, IE.cv.width, IE.cv.height) });
  if (IE.undo.length > IE_MAX_UNDO) IE.undo.shift();
  IE.redo = []; IE.changed = true; ieUpdateButtons();
}
function ieRestore(s) {
  if (IE.cv.width !== s.w || IE.cv.height !== s.h) { IE.cv.width = s.w; IE.cv.height = s.h; IE.hPt = s.hPt; }
  IE.g.putImageData(s.data, 0, 0); IE.ink = null; ieLayout();
}
const ieSnap = () => ({ w: IE.cv.width, h: IE.cv.height, hPt: IE.hPt, data: IE.g.getImageData(0, 0, IE.cv.width, IE.cv.height) });
function ieUndo() { ieDropFloat(true); if (!IE.undo.length) return; IE.redo.push(ieSnap()); ieRestore(IE.undo.pop()); ieSetSel(null); ieUpdateButtons(); }
function ieRedo() { ieCommit(); if (!IE.redo.length) return; IE.undo.push(ieSnap()); ieRestore(IE.redo.pop()); ieSetSel(null); ieUpdateButtons(); }
function ieUpdateButtons() {
  ieEl('ieUndo').disabled = !IE.undo.length && !IE.float; ieEl('ieRedo').disabled = !IE.redo.length;
  const has = !!(IE.sel || IE.float);
  ieEl('ieSelBar').classList.toggle('hidden', !has);
  ieEl('iePasteBtn').disabled = !IE.clip;
}
const ieChanged = () => { IE.ink = null; IE.changed = true; };

// --- open / close
async function openImageEditor(p, opts = {}) {
  if (!p) { toast('เลือกหน้าที่ต้องการแก้ไขก่อน'); return; }
  busy('กำลังเตรียมภาพ …');
  let r;
  try { r = await renderPageBase(p, IS_TOUCH ? 9e6 : 20e6); } catch (e) { unbusy(); toast('เปิดหน้าไม่สำเร็จ: ' + e.message, 'err'); return; }
  unbusy();
  const old = ieEl('ieCanvas'), cv = r.canvas; cv.id = 'ieCanvas'; old.replaceWith(cv);
  Object.assign(IE, { open: true, uid: p.uid, cv, g: cv.getContext('2d', { willReadFrequently: true }), wPt: r.wPt, hPt: r.hPt, sel: null, float: null, undo: [], redo: [], changed: false, ink: null, op: null });
  IE.pointers.clear(); IE.pinch = null;
  const i = D.pages.indexOf(p);
  ieEl('ieStep').textContent = `หน้า ${i + 1} จาก ${D.pages.length}` + (p.rot ? ' · แสดงแบบยังไม่หมุน' : '');
  ieEl('imged').classList.remove('hidden');
  ieEl('iePaint').classList.toggle('hidden', !host.paint);
  ieSetTool(IE.tool === 'select' || IE.tool === 'erase' || IE.tool === 'pen' ? IE.tool : 'select');
  ieSetSel(null); ieFit(); ieUpdateButtons(); ieHint();
  if (opts.paint) ieOpenPaint();
}
async function closeImageEditor(force = false) {
  if (!force && (IE.changed || IE.float)) {
    const r = await dialog({ title: 'ปิดหน้าแก้ไขภาพ', body: '<p>ภาพนี้ถูกแก้ไขแล้ว ต้องการนำไปใช้กับหน้าเอกสารหรือไม่?</p>', buttons: [['cancel', 'ยกเลิก'], ['drop', 'ทิ้งการแก้ไข'], ['apply', 'ใช้ภาพนี้', true]] });
    if (!r.btn) return;
    if (r.btn === 'apply') return ieApply();
  }
  ieStopPaint();
  IE.open = false; IE.undo = []; IE.redo = []; IE.ink = null; IE.float = null; IE.sel = null;
  ieEl('imged').classList.add('hidden');
  const c = ieEl('ieCanvas'); c.width = c.height = 1;
}
async function ieApply() {
  ieCommit();
  const p = pageByUid(IE.uid);
  if (!p) { toast('ไม่พบหน้าเดิมแล้ว', 'err'); return closeImageEditor(true); }
  if (!IE.changed) return closeImageEditor(true);
  busy('กำลังบันทึกภาพลงหน้า …');
  try { await replacePagesWithCanvases([{ uid: IE.uid, canvas: IE.cv, wPt: IE.wPt, hPt: IE.hPt }], 'แก้ไขภาพ'); }
  catch (e) { unbusy(); toast('บันทึกภาพไม่สำเร็จ: ' + e.message, 'err', [], 9000); return; }
  unbusy();
  closeImageEditor(true);
  toast('แก้ไขภาพหน้าแล้ว', 'ok', [['เลิกทำ', undo]], 4000);
}

// pages → image pages (bilevel pages are stored as CCITT G4, others JPEG). Rotation and added text boxes are kept.
function isBilevel(cv) {
  const W = cv.width, H = cv.height, d = cv.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data;
  let mid = 0, n = 0; const step = 4 * 3;
  for (let j = 0; j < d.length; j += step) { n++; const l = lumOf(d, j); const s = Math.max(d[j], d[j + 1], d[j + 2]) - Math.min(d[j], d[j + 1], d[j + 2]); if ((l > 60 && l < 195) || s > 45) mid++; }
  return mid / n < 0.004;
}
async function replacePagesWithCanvases(list, label) {
  const bytes = await scanPagesToPdf(list.map((it) => ({ canvas: it.canvas, filter: isBilevel(it.canvas) ? 'bw' : 'orig', wPt: it.wPt, hPt: it.hPt })));
  const now = new Date(), stamp = `${pad(now.getHours(), 2)}${pad(now.getMinutes(), 2)}${pad(now.getSeconds(), 2)}`;
  const src = await openPdfSource(`${label}_${stamp}.pdf`, bytes);
  pushUndo();
  list.forEach((it, k) => {
    const i = D.pages.findIndex((p) => p.uid === it.uid); if (i < 0) return;
    const old = D.pages[i];
    D.pages[i] = { uid: uid(), src: src.id, index: k, rot: old.rot, av: 1, annots: (old.annots || []).map(cloneAnn) };
    if (D.selected.has(old.uid)) { D.selected.delete(old.uid); D.selected.add(D.pages[i].uid); }
  });
  D.dirty = true;
  refresh();
}
async function despecklePages(list, strength = 40) {
  if (!list.length) { toast('เลือกหน้าก่อน'); return; }
  busy('กำลังลบจุดสกปรก …');
  const out = []; let total = 0;
  try {
    for (const p of list) {
      const r = await renderPageBase(p, IS_TOUCH ? 9e6 : 20e6);
      const n = despeckleCanvas(r.canvas, strength);
      if (n) { total += n; out.push({ uid: p.uid, canvas: r.canvas, wPt: r.wPt, hPt: r.hPt }); } else r.canvas.width = r.canvas.height = 0;
    }
    if (out.length) await replacePagesWithCanvases(out, 'ลบจุด');
  } catch (e) { unbusy(); toast('ลบจุดไม่สำเร็จ: ' + e.message, 'err', [], 8000); return; }
  unbusy();
  out.forEach((o) => { o.canvas.width = o.canvas.height = 0; });
  if (!total) toast('ไม่พบจุดสกปรกที่ต้องลบ', '', [], 3000);
  else toast(`ลบจุดสกปรก ${total.toLocaleString()} จุด ใน ${out.length} หน้า`, 'ok', [['เลิกทำ', undo]], 6000);
}

// --- view / zoom
function ieFit() {
  const st = ieEl('ieStage'), pad2 = IS_TOUCH ? 16 : 40;
  IE.view = Math.max(0.05, Math.min((st.clientWidth - pad2) / IE.cv.width, (st.clientHeight - pad2) / IE.cv.height, 2));
  ieLayout();
}
function ieZoom(f, cx, cy) {
  const st = ieEl('ieStage'), r = st.getBoundingClientRect();
  if (cx == null) { cx = r.left + st.clientWidth / 2; cy = r.top + st.clientHeight / 2; }
  const ip = ieToImg(cx, cy);
  IE.view = Math.max(0.05, Math.min(8, IE.view * f));
  ieLayout();
  const cr = IE.cv.getBoundingClientRect();
  st.scrollLeft += cr.left + ip.x * IE.view - cx; st.scrollTop += cr.top + ip.y * IE.view - cy;
}
function ieLayout() {
  const W = IE.cv.width * IE.view, H = IE.cv.height * IE.view;
  Object.assign(IE.cv.style, { width: W + 'px', height: H + 'px' });
  Object.assign(ieEl('ieWrap').style, { width: W + 'px', height: H + 'px' });
  ieEl('ieZoomLbl').textContent = Math.round(IE.view * IE.cv.width / IE.wPt * 100) + '%';
  ieDrawSel();
}
function ieToImg(cx, cy) { const r = IE.cv.getBoundingClientRect(); return { x: (cx - r.left) / IE.view, y: (cy - r.top) / IE.view }; }

// --- selection & floating layer
function ieSetSel(s) { IE.sel = s; ieDrawSel(); ieUpdateButtons(); ieHint(); }
function ieDrawSel() {
  const box = ieEl('ieSel'), s = IE.float || IE.sel;
  if (!s) { box.classList.add('hidden'); return; }
  box.classList.remove('hidden');
  Object.assign(box.style, { left: s.x * IE.view + 'px', top: s.y * IE.view + 'px', width: s.w * IE.view + 'px', height: s.h * IE.view + 'px' });
  box.classList.toggle('floating', !!IE.float);
  const c = box.querySelector('canvas');
  if (IE.float) {
    if (c.dataset.src !== IE.float.id) { c.width = IE.float.cv.width; c.height = IE.float.cv.height; c.getContext('2d').drawImage(IE.float.cv, 0, 0); c.dataset.src = IE.float.id; }
  } else if (s.mask) {
    const key = 'm' + s.x + ',' + s.y + ',' + s.mw + ',' + s.mh;
    if (c.dataset.src !== key) {
      c.width = s.mw; c.height = s.mh; const g = c.getContext('2d'), im = g.createImageData(s.mw, s.mh);
      for (let i = 0; i < s.mask.length; i++) if (s.mask[i]) { im.data[i * 4] = 20; im.data[i * 4 + 1] = 102; im.data[i * 4 + 2] = 224; im.data[i * 4 + 3] = 70; }
      g.putImageData(im, 0, 0); c.dataset.src = key;
    }
  } else { c.width = 1; c.height = 1; c.dataset.src = ''; }
}
let ieFloatSeq = 0;
function ieLift(copyOnly = false) { // selection → floating pixels (white made transparent); the original spot becomes white
  const s = IE.sel; if (!s || IE.float) return;
  const x = Math.max(0, Math.floor(s.x)), y = Math.max(0, Math.floor(s.y)), w = Math.min(IE.cv.width - x, Math.ceil(s.w)), h = Math.min(IE.cv.height - y, Math.ceil(s.h));
  if (w < 1 || h < 1) return;
  iePushUndo();
  const src = IE.g.getImageData(x, y, w, h), fd = new ImageData(w, h), sd = src.data, d = fd.data;
  for (let yy = 0; yy < h; yy++) for (let xx = 0; xx < w; xx++) {
    const j = (yy * w + xx) * 4;
    if (!ieMaskAt(s, xx, yy)) continue;
    const l = lumOf(sd, j), sat = Math.max(sd[j], sd[j + 1], sd[j + 2]) - Math.min(sd[j], sd[j + 1], sd[j + 2]);
    d[j] = sd[j]; d[j + 1] = sd[j + 1]; d[j + 2] = sd[j + 2];
    d[j + 3] = (l > 238 && sat < 25) ? 0 : (l > 215 && sat < 25 ? Math.round((238 - l) / 23 * 255) : 255);
    if (!copyOnly) { sd[j] = sd[j + 1] = sd[j + 2] = 255; }
  }
  if (!copyOnly) IE.g.putImageData(src, x, y);
  const fc = document.createElement('canvas'); fc.width = w; fc.height = h; fc.getContext('2d').putImageData(fd, 0, 0);
  IE.float = { cv: fc, x, y, w, h, id: 'f' + (++ieFloatSeq) };
  IE.sel = null; ieChanged(); ieDrawSel(); ieUpdateButtons();
}
function ieCommit() { // stamp the floating pixels back onto the image
  if (!IE.float) { if (IE.sel) ieSetSel(null); return; }
  const f = IE.float; IE.g.drawImage(f.cv, f.x, f.y, f.w, f.h);
  IE.float = null; ieChanged(); ieSetSel(null);
}
function ieDropFloat(stampBack) { if (!IE.float) return; if (stampBack) ieCommit(); else { IE.float = null; ieChanged(); ieSetSel(null); } }
function ieDelete() {
  if (IE.float) { IE.float = null; ieChanged(); ieSetSel(null); return; }
  const s = IE.sel; if (!s) return;
  iePushUndo();
  const x = Math.max(0, Math.floor(s.x)), y = Math.max(0, Math.floor(s.y)), w = Math.min(IE.cv.width - x, Math.ceil(s.w)), h = Math.min(IE.cv.height - y, Math.ceil(s.h));
  if (!s.mask) { IE.g.fillStyle = '#fff'; IE.g.fillRect(x, y, w, h); }
  else {
    const im = IE.g.getImageData(x, y, w, h), d = im.data;
    for (let yy = 0; yy < h; yy++) for (let xx = 0; xx < w; xx++) if (ieMaskAt(s, xx, yy)) { const j = (yy * w + xx) * 4; d[j] = d[j + 1] = d[j + 2] = 255; d[j + 3] = 255; }
    IE.g.putImageData(im, x, y);
  }
  ieChanged(); ieSetSel(null);
  toast('ลบแล้ว', '', [['เลิกทำ', ieUndo]], 2500);
}
function ieDuplicate() {
  if (!IE.float && !IE.sel) return;
  if (!IE.float) { ieLift(true); } else { iePushUndo(); IE.g.drawImage(IE.float.cv, IE.float.x, IE.float.y, IE.float.w, IE.float.h); }
  const off = Math.round(24 / Math.max(IE.view, 0.2));
  IE.float.x += off; IE.float.y += off; ieChanged(); ieDrawSel();
}
function ieCopy(cut = false) {
  if (!IE.float && !IE.sel) return;
  let f = IE.float, lifted = false; const prevSel = IE.sel;
  if (!f) { ieLift(!cut); f = IE.float; lifted = true; if (!f) return; }
  const c = document.createElement('canvas'); c.width = f.cv.width; c.height = f.cv.height; c.getContext('2d').drawImage(f.cv, 0, 0);
  IE.clip = { cv: c, w: f.w, h: f.h };
  if (cut) { IE.float = null; ieChanged(); ieSetSel(null); }
  else if (lifted) { IE.float = null; IE.undo.pop(); ieSetSel(prevSel); } // copy alone does not change the picture
  toast(cut ? 'ตัดแล้ว' : 'คัดลอกแล้ว — กด Ctrl+V เพื่อวาง', '', [], 1800);
  ieUpdateButtons();
}
function iePasteCanvas(c, w, h) {
  ieCommit(); iePushUndo();
  const maxW = IE.cv.width * 0.8, maxH = IE.cv.height * 0.8, k = Math.min(1, maxW / w, maxH / h);
  w *= k; h *= k;
  const st = ieEl('ieStage'), r = st.getBoundingClientRect(), mid = ieToImg(r.left + st.clientWidth / 2, r.top + st.clientHeight / 2);
  const x = Math.max(0, Math.min(IE.cv.width - w, mid.x - w / 2)), y = Math.max(0, Math.min(IE.cv.height - h, mid.y - h / 2));
  IE.float = { cv: c, x, y, w, h, id: 'f' + (++ieFloatSeq) };
  ieChanged(); ieDrawSel(); ieUpdateButtons(); ieHint();
}
function iePaste() { if (IE.clip) { const c = document.createElement('canvas'); c.width = IE.clip.cv.width; c.height = IE.clip.cv.height; c.getContext('2d').drawImage(IE.clip.cv, 0, 0); iePasteCanvas(c, IE.clip.w, IE.clip.h); } }
async function iePasteImageBlob(blob) {
  try {
    const bm = await createImageBitmap(blob);
    const c = document.createElement('canvas'); c.width = bm.width; c.height = bm.height; c.getContext('2d').drawImage(bm, 0, 0);
    // paper white → transparent so a pasted signature / stamp sits on the page naturally
    const g = c.getContext('2d'), im = g.getImageData(0, 0, c.width, c.height), d = im.data;
    for (let j = 0; j < d.length; j += 4) { const l = lumOf(d, j), s = Math.max(d[j], d[j + 1], d[j + 2]) - Math.min(d[j], d[j + 1], d[j + 2]); if (l > 238 && s < 25) d[j + 3] = 0; }
    g.putImageData(im, 0, 0);
    const scale = IE.cv.width / (IE.wPt * 200 / 72); // keep pasted images at ~200 dpi of the page
    iePasteCanvas(c, bm.width * scale, bm.height * scale);
  } catch { toast('วางรูปไม่สำเร็จ', 'err'); }
}

// --- tools
function ieSetTool(t) {
  if (t !== 'select') ieCommit();
  IE.tool = t;
  document.querySelectorAll('#imged [data-tool]').forEach((b) => b.classList.toggle('active', b.dataset.tool === t));
  ieEl('ieSizeWrap').classList.toggle('hidden', t === 'select');
  ieEl('ieColors').classList.toggle('hidden', t !== 'pen');
  ieEl('ieGroupWrap').classList.toggle('hidden', t !== 'select');
  ieEl('ieStage').dataset.tool = t;
  ieEl('ieCursor').classList.add('hidden');
  ieHint();
}
function ieHint(msg) {
  const el = ieEl('ieHint');
  if (IE.paint) { el.innerHTML = '<b>กำลังแก้ไขใน Paint</b> — แก้เสร็จกด บันทึก (Ctrl+S) ใน Paint แล้วภาพจะอัปเดตที่นี่อัตโนมัติ'; return; }
  if (msg) { el.textContent = msg; return; }
  const t = IE.tool;
  el.textContent = t === 'erase' ? 'ลากเพื่อลบให้เป็นพื้นขาว' : t === 'pen' ? 'ลากเพื่อเขียน' :
    (IE.float ? 'ลากเพื่อย้าย · ลากมุมขวาล่างเพื่อย่อ/ขยาย · คลิกที่ว่างเพื่อวางลง' : IE.sel ? (IS_TOUCH ? 'แตะปุ่มถังขยะเพื่อลบ · ลากเพื่อย้าย · ลากวงกลมมุมเพื่อย่อ/ขยาย' : 'กด Delete เพื่อลบ · ลากเพื่อย้าย · Shift+คลิก เพื่อเลือกเพิ่ม') :
      (IS_TOUCH ? 'แตะวัตถุ (เช่น ตรา โลโก้ ข้อความ) เพื่อเลือก หรือลากกรอบ · สองนิ้วเพื่อซูม/เลื่อน' : 'คลิกวัตถุ (เช่น ตรา โลโก้ ข้อความ) เพื่อเลือก หรือลากกรอบเลือกพื้นที่'));
}
function ieStroke(a, b) {
  const g = IE.g; g.save();
  g.lineCap = g.lineJoin = 'round';
  const k = IE.cv.width / 1654;
  g.lineWidth = Math.max(1, (IE.tool === 'pen' ? IE.size / 6 : IE.size) * k);
  g.strokeStyle = IE.tool === 'pen' ? IE.color : '#fff';
  g.beginPath(); g.moveTo(a.x, a.y); g.lineTo(b.x + 0.01, b.y); g.stroke(); g.restore();
}
function ieShowCursor(e) {
  const c = ieEl('ieCursor');
  if (IE.tool !== 'erase' || e.pointerType === 'touch') { c.classList.add('hidden'); return; }
  const wr = ieEl('ieWrap').getBoundingClientRect(), d = IE.size * (IE.cv.width / 1654) * IE.view;
  c.classList.remove('hidden');
  Object.assign(c.style, { width: d + 'px', height: d + 'px', left: e.clientX - wr.left - d / 2 + 'px', top: e.clientY - wr.top - d / 2 + 'px' });
}

// --- pointer handling (mouse / pen / touch, pinch to zoom)
function iePointerDown(e) {
  if (e.button > 0) return;
  const st = ieEl('ieStage');
  IE.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
  try { st.setPointerCapture(e.pointerId); } catch {}
  if (IE.pointers.size === 2) { // start pinch / two-finger pan
    if (IE.op && IE.op.kind === 'rect') ieEl('ieRect').classList.add('hidden');
    if (IE.op && IE.op.kind === 'stroke' && !IE.op.moved && IE.undo.length) { IE.op = null; }
    IE.op = null;
    const [a, b] = [...IE.pointers.values()];
    IE.pinch = { d: Math.hypot(a.x - b.x, a.y - b.y), view: IE.view, cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2 };
    return;
  }
  if (IE.pointers.size > 2) return;
  e.preventDefault();
  const P = ieToImg(e.clientX, e.clientY);
  if (IE.tool === 'erase' || IE.tool === 'pen') {
    ieCommit(); iePushUndo(); ieStroke(P, P); ieChanged();
    IE.op = { kind: 'stroke', last: P, moved: false };
    return;
  }
  // select tool
  const s = IE.float || IE.sel;
  if (s) {
    const hx = (s.x + s.w) * IE.view, hy = (s.y + s.h) * IE.view, px = P.x * IE.view, py = P.y * IE.view, hr = IS_TOUCH || e.pointerType === 'touch' ? 26 : 12;
    if (Math.abs(px - hx) < hr && Math.abs(py - hy) < hr) { if (!IE.float) ieLift(); const f = IE.float; if (f) IE.op = { kind: 'resize', start: P, w: f.w, h: f.h }; return; }
    if (P.x >= s.x && P.y >= s.y && P.x <= s.x + s.w && P.y <= s.y + s.h && !e.shiftKey) {
      IE.op = { kind: 'move', start: P, ox: s.x, oy: s.y, lifted: !!IE.float, moved: false };
      return;
    }
  }
  if (!e.shiftKey) ieCommit();
  IE.op = { kind: 'rect', start: P, cx: e.clientX, cy: e.clientY, moved: false, add: e.shiftKey };
}
function iePointerMove(e) {
  if (IE.pointers.has(e.pointerId)) IE.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
  if (IE.pinch && IE.pointers.size >= 2) {
    const [a, b] = [...IE.pointers.values()], st = ieEl('ieStage');
    const d = Math.hypot(a.x - b.x, a.y - b.y), cx = (a.x + b.x) / 2, cy = (a.y + b.y) / 2;
    const ip = ieToImg(IE.pinch.cx, IE.pinch.cy);
    IE.view = Math.max(0.05, Math.min(8, IE.pinch.view * d / IE.pinch.d));
    ieLayout();
    const cr = IE.cv.getBoundingClientRect();
    st.scrollLeft += cr.left + ip.x * IE.view - cx; st.scrollTop += cr.top + ip.y * IE.view - cy;
    IE.pinch.cx = cx; IE.pinch.cy = cy; IE.pinch.d = d; IE.pinch.view = IE.view;
    return;
  }
  ieShowCursor(e);
  const op = IE.op; if (!op) return;
  const P = ieToImg(e.clientX, e.clientY);
  if (op.kind === 'stroke') { ieStroke(op.last, P); op.last = P; op.moved = true; return; }
  if (op.kind === 'move') {
    const dx = P.x - op.start.x, dy = P.y - op.start.y;
    if (!op.moved && Math.hypot(dx, dy) * IE.view < 3) return;
    if (!op.moved) { op.moved = true; if (!IE.float) ieLift(); if (!IE.float) { IE.op = null; return; } }
    IE.float.x = op.ox + dx; IE.float.y = op.oy + dy; ieDrawSel(); return;
  }
  if (op.kind === 'resize' && IE.float) {
    const k = Math.max(0.05, Math.max((P.x - IE.float.x) / op.w, (P.y - IE.float.y) / op.h));
    IE.float.w = op.w * k; IE.float.h = op.h * k; ieDrawSel(); return;
  }
  if (op.kind === 'rect') {
    if (!op.moved && Math.hypot(e.clientX - op.cx, e.clientY - op.cy) < 5) return;
    op.moved = true;
    const x = Math.min(op.start.x, P.x), y = Math.min(op.start.y, P.y), w = Math.abs(P.x - op.start.x), h = Math.abs(P.y - op.start.y);
    Object.assign(ieEl('ieRect').style, { left: x * IE.view + 'px', top: y * IE.view + 'px', width: w * IE.view + 'px', height: h * IE.view + 'px' });
    ieEl('ieRect').classList.remove('hidden');
    op.rect = { x: Math.max(0, x), y: Math.max(0, y), w: Math.min(IE.cv.width, x + w) - Math.max(0, x), h: Math.min(IE.cv.height, y + h) - Math.max(0, y) };
  }
}
function iePointerUp(e) {
  IE.pointers.delete(e.pointerId);
  if (IE.pinch) { if (IE.pointers.size < 2) IE.pinch = null; IE.op = null; return; }
  const op = IE.op; IE.op = null; if (!op) return;
  if (op.kind === 'stroke') { ieChanged(); return; }
  if (op.kind === 'move' && !op.moved) { ieHint(); return; }
  if (op.kind === 'rect') {
    ieEl('ieRect').classList.add('hidden');
    let s = null;
    if (op.moved && op.rect && op.rect.w > 2 && op.rect.h > 2) s = { ...op.rect, mask: null };
    else {
      s = iePickObject(op.start.x, op.start.y);
      if (!s) { ieSetSel(op.add ? IE.sel : null); ieHint('ไม่พบวัตถุตรงนี้ — ลองคลิกบนตัววัตถุ หรือลากกรอบเลือกพื้นที่'); return; }
    }
    if (op.add && IE.sel) s = ieUnionSel(IE.sel.mask ? IE.sel : { ...IE.sel, mask: new Uint8Array(Math.ceil(IE.sel.w) * Math.ceil(IE.sel.h)).fill(1), mw: Math.ceil(IE.sel.w), mh: Math.ceil(IE.sel.h), f: 1 }, s.mask ? s : { ...s, mask: new Uint8Array(Math.ceil(s.w) * Math.ceil(s.h)).fill(1), mw: Math.ceil(s.w), mh: Math.ceil(s.h), f: 1 });
    ieSetSel(s);
  }
  ieHint();
}

// --- Edit in Paint (Windows exe only): the picture goes to Paint; every save in Paint comes back here automatically
async function ieOpenPaint() {
  if (!host.paint) return;
  ieCommit();
  if (IE.paint) ieStopPaint();
  busy('กำลังเปิด Paint …');
  try {
    const blob = await new Promise((res) => IE.cv.toBlob(res, 'image/png'));
    const r = await (await api('/api/paint/open', { method: 'POST', body: blob })).json();
    unbusy();
    if (!r.ok) { toast('เปิด Paint ไม่สำเร็จ: ' + (r.error || ''), 'err', [], 8000); return; }
    IE.paint = { id: r.id, v: 0, busy: false, timer: setInterval(iePaintPoll, 900) };
    ieEl('iePaint').classList.add('active');
    ieHint();
    toast('เปิดภาพใน Paint แล้ว — แก้ไขแล้วกด บันทึก (Ctrl+S) ใน Paint', 'ok', [], 6000);
  } catch (e) { unbusy(); toast('เปิด Paint ไม่สำเร็จ: ' + e.message + ' (ต้องใช้ Super PDF เวอร์ชัน 1.7 ขึ้นไป)', 'err', [], 8000); }
}
async function iePaintPoll() {
  const P = IE.paint; if (!P || P.busy || !IE.open) return;
  P.busy = true;
  try {
    const r = await (await api('/api/paint/poll?id=' + P.id)).json();
    if (r.ok && r.v > P.v) {
      P.v = r.v;
      const b = await (await api('/api/paint/file?id=' + P.id)).blob();
      const bm = await createImageBitmap(b);
      ieDropFloat(true);
      iePushUndo();
      if (bm.width !== IE.cv.width || bm.height !== IE.cv.height) { IE.cv.width = bm.width; IE.cv.height = bm.height; IE.hPt = IE.wPt * bm.height / bm.width; }
      IE.g.fillStyle = '#fff'; IE.g.fillRect(0, 0, IE.cv.width, IE.cv.height);
      IE.g.drawImage(bm, 0, 0);
      ieChanged(); ieSetSel(null); ieLayout();
      toast('รับภาพที่แก้ใน Paint แล้ว', 'ok', [], 2200);
    } else if (!r.ok) ieStopPaint();
  } catch {}
  P.busy = false;
}
function ieStopPaint() {
  const P = IE.paint; if (!P) return;
  clearInterval(P.timer); IE.paint = null;
  ieEl('iePaint').classList.remove('active');
  api('/api/paint/close?id=' + P.id).catch(() => {});
  ieHint();
}

// --- wiring
(() => {
  const st = ieEl('ieStage');
  st.addEventListener('pointerdown', iePointerDown);
  st.addEventListener('pointermove', iePointerMove);
  st.addEventListener('pointerup', iePointerUp);
  st.addEventListener('pointercancel', iePointerUp);
  st.addEventListener('pointerleave', () => ieEl('ieCursor').classList.add('hidden'));
  st.addEventListener('wheel', (e) => { if (e.ctrlKey || e.metaKey) { e.preventDefault(); ieZoom(e.deltaY < 0 ? 1.15 : 1 / 1.15, e.clientX, e.clientY); } }, { passive: false });
  st.addEventListener('contextmenu', (e) => {
    e.preventDefault(); e.stopPropagation();
    const P = ieToImg(e.clientX, e.clientY), s = IE.float || IE.sel;
    if (!s || P.x < s.x || P.y < s.y || P.x > s.x + s.w || P.y > s.y + s.h) { if (IE.tool === 'select') { ieCommit(); const o = iePickObject(P.x, P.y); if (o) ieSetSel(o); } }
    const has = !!(IE.float || IE.sel);
    showMenu([
      { label: 'ลบ', icon: 'trash', kbd: 'Delete', danger: true, disabled: !has, action: ieDelete },
      { label: 'ตัด', icon: 'cut', kbd: 'Ctrl+X', disabled: !has, action: () => ieCopy(true) },
      { label: 'คัดลอก', icon: 'copy', kbd: 'Ctrl+C', disabled: !has, action: () => ieCopy(false) },
      { label: 'วาง', icon: 'paste', kbd: 'Ctrl+V', disabled: !IE.clip, action: iePaste },
      { label: 'ทำสำเนา', icon: 'dup', kbd: 'Ctrl+D', disabled: !has, action: ieDuplicate },
      { sep: true },
      { label: 'ลบจุดสกปรก' + (IE.sel ? ' (ในกรอบ)' : ' (ทั้งหน้า)'), icon: 'clean', action: ieDespeckle },
      { label: 'เลือกทั้งหน้า', icon: 'selall', kbd: 'Ctrl+A', action: () => { ieCommit(); ieSetSel({ x: 0, y: 0, w: IE.cv.width, h: IE.cv.height, mask: null }); } },
    ], e.clientX, e.clientY);
  });
  ieEl('ieClose').onclick = () => closeImageEditor();
  ieEl('ieApply').onclick = ieApply;
  ieEl('ieUndo').onclick = ieUndo;
  ieEl('ieRedo').onclick = ieRedo;
  document.querySelectorAll('#imged [data-tool]').forEach((b) => { b.onclick = () => ieSetTool(b.dataset.tool); });
  ieEl('ieSize').oninput = (e) => { IE.size = +e.target.value; };
  ieEl('ieGroup').oninput = (e) => { IE.group = +e.target.value; IE.ink = null; };
  document.querySelectorAll('#ieColors .swatch').forEach((b) => { b.onclick = () => { IE.color = b.dataset.c; document.querySelectorAll('#ieColors .swatch').forEach((x) => x.classList.toggle('on', x === b)); }; });
  ieEl('ieZoomIn').onclick = () => ieZoom(1.25);
  ieEl('ieZoomOut').onclick = () => ieZoom(1 / 1.25);
  ieEl('ieZoomFit').onclick = ieFit;
  ieEl('ieDespeckle').onclick = ieDespeckle;
  ieEl('ieStrength').oninput = (e) => { IE.strength = +e.target.value; };
  ieEl('iePaint').onclick = () => (IE.paint ? ieStopPaint() : ieOpenPaint());
  ieEl('ieDelBtn').onclick = ieDelete;
  ieEl('ieDupBtn').onclick = ieDuplicate;
  ieEl('ieCopyBtn').onclick = () => ieCopy(false);
  ieEl('iePasteBtn').onclick = iePaste;
  ieEl('ieDoneBtn').onclick = () => { ieCommit(); };
  window.addEventListener('resize', () => { if (IE.open) ieDrawSel(); });

  // keyboard (only while the editor is open; blocks the document shortcuts underneath)
  let iePasteHandled = false;
  window.addEventListener('keydown', (e) => {
    if (!IE.open || modalOpen()) return;
    if (typingInField()) return;
    const c = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    let handled = true;
    if (e.key === 'Delete' || e.key === 'Backspace') ieDelete();
    else if (c && k === 'z') e.shiftKey ? ieRedo() : ieUndo();
    else if (c && k === 'y') ieRedo();
    else if (c && k === 'c') ieCopy(false);
    else if (c && k === 'x') ieCopy(true);
    else if (c && k === 'd') ieDuplicate();
    else if (c && k === 'a') { ieCommit(); ieSetSel({ x: 0, y: 0, w: IE.cv.width, h: IE.cv.height, mask: null }); }
    else if (c && k === 'v') { iePasteHandled = false; setTimeout(() => { if (!iePasteHandled) iePaste(); }, 150); handled = false; e.stopImmediatePropagation(); return; }
    else if (c && (e.key === '=' || e.key === '+')) ieZoom(1.25);
    else if (c && e.key === '-') ieZoom(1 / 1.25);
    else if (c && k === '0') ieFit();
    else if (c && k === 's') ieApply();
    else if (e.key === 'Enter') { if (IE.float || IE.sel) ieCommit(); else ieApply(); }
    else if (e.key === 'Escape') { if (IE.float) ieCommit(); else if (IE.sel) ieSetSel(null); else closeImageEditor(); }
    else if (!c && k === 'v') ieSetTool('select');
    else if (!c && k === 'e') ieSetTool('erase');
    else if (!c && k === 'p') ieSetTool('pen');
    else if (e.key.startsWith('Arrow') && (IE.sel || IE.float)) {
      if (!IE.float) ieLift(); if (!IE.float) return;
      const dd = (e.shiftKey ? 10 : 1) / Math.max(IE.view, 0.1);
      if (e.key === 'ArrowLeft') IE.float.x -= dd; if (e.key === 'ArrowRight') IE.float.x += dd;
      if (e.key === 'ArrowUp') IE.float.y -= dd; if (e.key === 'ArrowDown') IE.float.y += dd;
      ieDrawSel();
    } else if (!c) handled = false;
    if (handled) e.preventDefault();
    e.stopImmediatePropagation();
  }, true);
  for (const ev of ['copy', 'cut']) window.addEventListener(ev, (e) => { if (IE.open && !typingInField()) { e.stopImmediatePropagation(); e.preventDefault(); } }, true);
  window.addEventListener('paste', async (e) => {
    if (!IE.open || typingInField()) return;
    e.stopImmediatePropagation(); e.preventDefault();
    iePasteHandled = true;
    const items = [...(e.clipboardData?.items || [])];
    const img = items.find((it) => it.kind === 'file' && it.type.startsWith('image/'));
    if (img) { const f = img.getAsFile(); if (f) return iePasteImageBlob(f); }
    iePaste();
  }, true);
})();
function ieDespeckle() {
  const s = IE.float ? null : IE.sel;
  ieCommit();
  const before = ieSnap();
  const n = despeckleCanvas(IE.cv, IE.strength, s ? { x: s.x, y: s.y, w: s.w, h: s.h } : null);
  if (!n) { toast('ไม่พบจุดสกปรก (ลองเพิ่มความแรง)', '', [], 2500); return; }
  IE.undo.push(before); if (IE.undo.length > IE_MAX_UNDO) IE.undo.shift(); IE.redo = [];
  ieChanged(); ieUpdateButtons();
  toast(`ลบจุดสกปรก ${n.toLocaleString()} จุด`, 'ok', [['เลิกทำ', ieUndo]], 3500);
}
function imgEditTarget() {
  if (!D.pages.length) return null;
  if (mode === 'read') return currentReaderPage();
  const s = selPages(); return s[0] || null;
}
$('btnImgEdit').onclick = () => { const p = imgEditTarget(); if (!p) { toast(D.pages.length ? 'คลิกเลือกหน้าที่ต้องการแก้ไขก่อน' : 'ยังไม่มีเอกสาร'); return; } openImageEditor(p); };


// ------------------------------------------------------------------ v1.8 signatures & stamps (ลายเซ็น / ตราประทับ)
// Library items are transparent PNGs (ink only). Windows exe: files in %LOCALAPPDATA%\SuperPDF\signatures (via /api/sig/*);
// web / iPhone: IndexedDB of this site. Placed on a page as an annotation { kind:'img', sk, base, x, y, w, h, r, color, weight }
// and saved into the PDF as a transparent image on top of the original page (the page itself is not rasterised).
const SIG = {
  items: null,          // [{ id, kind:'sig'|'stamp', w, h, t, v }]
  thumbs: new Map(),    // id:v -> object URL
  base: new Map(),      // base key -> canvas (RGBA, transparent background)
  fin: new Map(),       // base|color|weight -> { c: canvas, url }
  tab: 'sig',
  target: null,         // { p, pt } where the next library item should go
  props: null,          // { undo: bool } while the colour / thickness sheet is open
};
const SIG_COLORS = [null, '#000000', '#1435c8', '#d92d20', '#ffffff', '#7a3b0c', '#f97316', '#eab308', '#067647', '#7c3aed'];
const sigKindName = (k) => (k === 'stamp' ? 'ตราประทับ' : 'ลายเซ็น');
const sigOverlayOpen = () => !!document.querySelector('.sigfull:not(.hidden)');

// ---- storage
const sigStore = (() => {
  const viaHost = () => !!(host.host && host.sig);
  const req = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  let dbp = null;
  const db = () => (dbp ||= new Promise((res, rej) => {
    const r = indexedDB.open('superpdf-signatures', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('items', { keyPath: 'id' });
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  }));
  const os = async (mode) => (await db()).transaction('items', mode).objectStore('items');
  let persisted = false;
  const persist = () => { if (persisted) return; persisted = true; try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch {} };
  async function hostIndex() {
    const names = await (await api('/api/sig/list')).json();
    let items = [];
    if (names.includes('index.json')) { try { items = (await (await api('/api/sig/get?name=index.json')).json()).items || []; } catch { items = []; } }
    return { names, items };
  }
  const hostSaveIndex = (items) => api('/api/sig/put?name=index.json', { method: 'POST', body: JSON.stringify({ app: 'Super PDF', format: 1, items }, null, 1) });
  const clean = (m) => ({ id: m.id, kind: m.kind === 'stamp' ? 'stamp' : 'sig', w: m.w | 0, h: m.h | 0, t: +m.t || Date.now(), v: +m.v || 1 });
  return {
    where() { return viaHost() ? 'ในเครื่องนี้ (โฟลเดอร์ %LOCALAPPDATA%\\SuperPDF\\signatures)' : 'ในเครื่องนี้ (ที่เก็บข้อมูลของแอปบนอุปกรณ์นี้)'; },
    async list() {
      let L;
      if (viaHost()) { const { names, items } = await hostIndex(); L = items.filter((it) => names.includes(it.id + '.png')); }
      else L = (await req((await os('readonly')).getAll())).map(({ png, ...m }) => m);
      return L.map(clean).sort((a, b) => b.t - a.t);
    },
    async png(id) {
      if (viaHost()) return await (await api('/api/sig/get?name=' + id + '.png')).blob();
      const r = await req((await os('readonly')).get(id));
      return r ? r.png : null;
    },
    async put(meta, blob) {
      meta = clean(meta);
      if (viaHost()) {
        if (blob) await api('/api/sig/put?name=' + meta.id + '.png', { method: 'POST', body: blob });
        const { items } = await hostIndex();
        await hostSaveIndex([...items.filter((x) => x.id !== meta.id), meta]);
        return;
      }
      if (!blob) blob = await this.png(meta.id);
      await req((await os('readwrite')).put({ ...meta, png: blob }));
      persist();
    },
    async del(id) {
      if (viaHost()) {
        const { items } = await hostIndex();
        await hostSaveIndex(items.filter((x) => x.id !== id));
        await api('/api/sig/del?name=' + id + '.png', { method: 'POST' });
        return;
      }
      await req((await os('readwrite')).delete(id));
    },
  };
})();
const sigNewId = () => 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

async function sigLoadLibrary(force = false) {
  if (SIG.items && !force) return SIG.items;
  try { SIG.items = await sigStore.list(); }
  catch (e) { SIG.items = []; toast('เปิดคลังลายเซ็นไม่ได้: ' + e.message, 'err'); }
  return SIG.items;
}
async function sigThumb(it) {
  const k = it.id + ':' + it.v;
  if (!SIG.thumbs.has(k)) {
    const b = await sigStore.png(it.id);
    SIG.thumbs.set(k, b ? URL.createObjectURL(b) : '');
  }
  return SIG.thumbs.get(k);
}
// library item -> base key (decoded canvas kept for the session; earlier versions stay valid for objects already placed)
async function sigEnsureBase(it) {
  const key = 'L' + it.id + 'v' + it.v;
  if (SIG.base.has(key)) return key;
  const url = await sigThumb(it);
  const im = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('อ่านภาพไม่ได้')); i.src = url; });
  const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight;
  c.getContext('2d').drawImage(im, 0, 0);
  SIG.base.set(key, c);
  return key;
}
async function sigAddToLibrary(canvas, kind) {
  const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
  const meta = { id: sigNewId(), kind, w: canvas.width, h: canvas.height, t: Date.now(), v: 1 };
  await sigStore.put(meta, blob);
  SIG.thumbs.set(meta.id + ':1', URL.createObjectURL(blob));
  const key = 'L' + meta.id + 'v1';
  SIG.base.set(key, canvas);
  await sigLoadLibrary(true);
  return SIG.items.find((x) => x.id === meta.id) || meta;
}

// ---- appearance: colour + stroke thickness, computed from the base image (cached)
function sigMorph(A, W, H, r, isMax) { // separable square max/min filter, radius r (px)
  if (r <= 0) return A;
  const T = new Float32Array(W * H), O = new Float32Array(W * H);
  for (let y = 0; y < H; y++) {
    const row = y * W;
    for (let x = 0; x < W; x++) {
      let v = isMax ? 0 : 1;
      const x0 = Math.max(0, x - r), x1 = Math.min(W - 1, x + r);
      for (let k = x0; k <= x1; k++) { const a = A[row + k]; if (isMax ? a > v : a < v) v = a; }
      T[row + x] = v;
    }
  }
  for (let x = 0; x < W; x++) {
    for (let y = 0; y < H; y++) {
      let v = isMax ? 0 : 1;
      const y0 = Math.max(0, y - r), y1 = Math.min(H - 1, y + r);
      for (let k = y0; k <= y1; k++) { const a = T[k * W + x]; if (isMax ? a > v : a < v) v = a; }
      O[y * W + x] = v;
    }
  }
  return O;
}
// typical stroke width (px) ≈ 2 · ink area / ink outline length; cached per picture
function sigStrokeWidth(src, A, W, H) {
  if (src.__sw) return src.__sw;
  let area = 0, edge = 0;
  for (let y = 1; y < H - 1; y++) for (let x = 1; x < W - 1; x++) {
    const i = y * W + x; if (A[i] < 0.5) continue;
    area++;
    if (A[i - 1] < 0.5 || A[i + 1] < 0.5 || A[i - W] < 0.5 || A[i + W] < 0.5) edge++;
  }
  src.__sw = Math.max(1.5, Math.min(40, edge ? 2 * area / edge : 3));
  return src.__sw;
}
function sigStyle(src, color, weight) {
  const W = src.width, H = src.height;
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const g = c.getContext('2d'); g.drawImage(src, 0, 0);
  if (!color && !weight) return c;
  const img = g.getImageData(0, 0, W, H), d = img.data, N = W * H;
  const A0 = new Float32Array(N);
  for (let i = 0; i < N; i++) A0[i] = d[i * 4 + 3] / 255;
  let A = A0;
  if (weight) {
    const unit = sigStrokeWidth(src, A0, W, H) * (weight > 0 ? 0.28 : 0.2);   // relative to this picture's own stroke width
    const r = Math.abs(weight) * unit, ri = Math.floor(r), fr = r - ri;
    const m1 = sigMorph(A0, W, H, ri, weight > 0), m2 = sigMorph(A0, W, H, ri + 1, weight > 0);
    A = new Float32Array(N);
    for (let i = 0; i < N; i++) {
      let a = m1[i] + (m2[i] - m1[i]) * fr;
      if (weight > 0) a = 1 - Math.pow(1 - a, 1 + 0.45 * weight);                    // thicker & darker
      else a = Math.max(a, A0[i] * 0.3) * (1 + 0.2 * weight);                       // thinner & lighter
      A[i] = a;
    }
  }
  let cr, cg, cb;
  if (color) { cr = parseInt(color.slice(1, 3), 16); cg = parseInt(color.slice(3, 5), 16); cb = parseInt(color.slice(5, 7), 16); }
  else { // keep the ink's own colours; pixels that only exist after thickening take the average ink colour
    let sr = 0, sg = 0, sb = 0, sw = 0;
    for (let i = 0; i < N; i++) { const w = A0[i]; if (w > 0.3) { sr += d[i * 4] * w; sg += d[i * 4 + 1] * w; sb += d[i * 4 + 2] * w; sw += w; } }
    cr = sw ? sr / sw : 0; cg = sw ? sg / sw : 0; cb = sw ? sb / sw : 0;
  }
  for (let i = 0, j = 0; i < N; i++, j += 4) {
    if (!color) { const k = Math.min(1, A0[i] * 2); d[j] = cr + (d[j] - cr) * k; d[j + 1] = cg + (d[j + 1] - cg) * k; d[j + 2] = cb + (d[j + 2] - cb) * k; }
    else { d[j] = cr; d[j + 1] = cg; d[j + 2] = cb; }
    d[j + 3] = Math.max(0, Math.min(255, Math.round(A[i] * 255)));
  }
  g.putImageData(img, 0, 0);
  return c;
}
const sigFinKey = (a) => a.base + '|' + (a.color || '') + '|' + (+(a.weight || 0)).toFixed(1);
function sigFinalEntry(a) {
  const k = sigFinKey(a);
  let e = SIG.fin.get(k);
  if (!e) {
    const b = SIG.base.get(a.base); if (!b) return null;
    e = { c: sigStyle(b, a.color, +(a.weight || 0)), url: '' };
    SIG.fin.set(k, e);
    if (SIG.fin.size > 80) { const first = SIG.fin.keys().next().value; if (first !== k) SIG.fin.delete(first); }
  }
  return e;
}
function sigFinal(a) { const e = sigFinalEntry(a); return e && e.c; }
function sigFinalUrl(a) { const e = sigFinalEntry(a); if (!e) return ''; if (!e.url) e.url = e.c.toDataURL('image/png'); return e.url; }
async function embedSigAnn(out, pdfPage, a, vb, R) {
  const im = sigFinal(a); if (!im) return;
  if (!out.__sigImgs) out.__sigImgs = new Map();
  const k = sigFinKey(a);
  let emb = out.__sigImgs.get(k);
  if (!emb) { emb = await out.embedPng(await canvasToBytes(im, 'image/png')); out.__sigImgs.set(k, emb); }
  const rot = ((R - (a.r || 0)) % 360 + 360) % 360;
  const [px, py] = vb.convertToPdfPoint(...annLocal(a, 0, a.h));
  pdfPage.drawImage(emb, { x: px, y: py, width: a.w, height: a.h, rotate: degrees(rot) });
}

// ---- placed objects on the reader page
const SIG_ICON = { del: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></svg>', more: '<svg viewBox="0 0 24 24"><circle cx="5" cy="12" r="1.3" fill="currentColor"/><circle cx="12" cy="12" r="1.3" fill="currentColor"/><circle cx="19" cy="12" r="1.3" fill="currentColor"/></svg>', rs: '<svg viewBox="0 0 24 24"><path d="M7 7l10 10M7 7h5M7 7v5M17 17h-5M17 17v-5"/></svg>' };
function sigAnnEl(a, s, selected) {
  const el = document.createElement('div');
  el.className = 'annot annimg' + (selected ? ' selected' : '');
  el.dataset.id = a.id;
  Object.assign(el.style, { left: a.x * s + 'px', top: a.y * s + 'px', width: a.w * s + 'px', height: a.h * s + 'px', transform: a.r ? `rotate(${a.r}deg)` : '', transformOrigin: '0 0' });
  const url = sigFinalUrl(a);
  if (url) { const im = new Image(); im.src = url; im.draggable = false; im.alt = ''; el.appendChild(im); }
  if (selected) {
    for (const h of ['del', 'more', 'rs']) {
      const b = document.createElement('span'); b.className = 'ah ah' + h; b.dataset.h = h; b.innerHTML = SIG_ICON[h];
      b.title = { del: 'ลบออกจากหน้า', more: 'ปรับสี / ความหนา', rs: 'ลากเพื่อย่อ-ขยาย และหมุน' }[h];
      el.appendChild(b);
    }
  }
  return el;
}
function sigSetBox(el, a, s) {
  Object.assign(el.style, { left: a.x * s + 'px', top: a.y * s + 'px', width: a.w * s + 'px', height: a.h * s + 'px', transform: a.r ? `rotate(${a.r}deg)` : '' });
}
const sigCenter = (a) => annLocal(a, a.w / 2, a.h / 2);
function sigPlaceCenter(a, c) { // move a so that its centre is at c (base pt), keeping a.r
  const t = (a.r || 0) * Math.PI / 180, hw = a.w / 2, hh = a.h / 2;
  a.x = c[0] - (hw * Math.cos(t) - hh * Math.sin(t));
  a.y = c[1] - (hw * Math.sin(t) + hh * Math.cos(t));
}
function sigPointerDown(e, L, p, el) {
  e.preventDefault();
  if (annEditing) document.activeElement.blur();
  const sel = { uid: p.uid, id: el.dataset.id };
  const h = e.target.closest('.ah') ? e.target.closest('.ah').dataset.h : null;
  if (h === 'del') { annSel = sel; sigCloseProps(); deleteAnn(); return; }
  if (h === 'more') { selectAnn(sel); sigOpenProps(); return; }
  if (!annSel || annSel.id !== sel.id) { selectAnn(sel); if (SIG.props) sigOpenProps(); }
  const a = findAnn(sel); if (!a) return;
  const s = +L.dataset.scale;
  const box = annEl(sel); if (!box) return;
  try { box.setPointerCapture(e.pointerId); } catch {}
  const [sx, sy] = annPoint(L, e);
  const o = { x: a.x, y: a.y, w: a.w, h: a.h, r: a.r || 0 };
  const C = sigCenter(a);
  const v0 = [sx - C[0], sy - C[1]], d0 = Math.hypot(...v0) || 1, ang0 = Math.atan2(v0[1], v0[0]);
  const bw = L.offsetWidth / s, bh = L.offsetHeight / s;
  let moved = false;
  const R = $('reader');
  let last = e, vy = 0, raf = 0;
  const place = (ev) => {
    const [mx, my] = annPoint(L, ev);
    if (!moved && Math.hypot(mx - sx, my - sy) * s < 3) return;
    if (!moved) { moved = true; pushUndo(); }
    if (h === 'rs') {
      const v = [mx - C[0], my - C[1]];
      const k = Math.hypot(...v) / d0;
      let r = o.r + (Math.atan2(v[1], v[0]) - ang0) * 180 / Math.PI;
      r = ((r % 360) + 360) % 360;
      const snap = Math.round(r / 90) * 90; if (Math.abs(r - snap) < 3) r = snap % 360;
      const nw = Math.max(10, Math.min(Math.max(bw, bh) * 1.5, o.w * k));
      a.w = nw; a.h = o.h * nw / o.w; a.r = r;
      sigPlaceCenter(a, C);
    } else {
      const c0 = [C[0] + (mx - sx), C[1] + (my - sy)];
      sigPlaceCenter(a, [Math.max(0, Math.min(bw, c0[0])), Math.max(0, Math.min(bh, c0[1]))]);
    }
    sigSetBox(box, a, s);
  };
  const tick = () => { raf = 0; if (!vy) return; R.scrollTop += vy; place(last); raf = requestAnimationFrame(tick); };
  const move = (ev) => {
    if (ev.pointerId !== e.pointerId) return;
    ev.preventDefault(); last = ev; place(ev);
    if (h === 'rs') return;
    const rr = R.getBoundingClientRect(), edge = 56;
    vy = !moved ? 0 : ev.clientY > rr.bottom - edge ? Math.min(18, (ev.clientY - (rr.bottom - edge)) / 3 + 2)
      : ev.clientY < rr.top + edge ? -Math.min(18, ((rr.top + edge) - ev.clientY) / 3 + 2) : 0;
    if (vy && !raf) raf = requestAnimationFrame(tick);
  };
  const up = (ev) => {
    if (ev && ev.pointerId !== e.pointerId) return;
    vy = 0; if (raf) cancelAnimationFrame(raf);
    window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); window.removeEventListener('pointercancel', up);
    if (moved) { touchPage(p); refresh(); }
  };
  window.addEventListener('pointermove', move, { passive: false }); window.addEventListener('pointerup', up); window.addEventListener('pointercancel', up);
}
function sigAnnMenu() {
  const a = findAnn(annSel);
  return [
    { label: 'ปรับสี / ความหนา…', icon: 'edit', kbd: 'Enter', action: sigOpenProps },
    { label: 'หมุน 90° ตามเข็ม', icon: 'rr', action: () => sigRotateSel(90) },
    { label: 'หมุน 90° ทวนเข็ม', icon: 'rl', action: () => sigRotateSel(-90) },
    { sep: true },
    { label: 'คัดลอก', icon: 'copy', kbd: 'Ctrl+C', action: () => copyAnn(false) },
    { label: 'ตัด', icon: 'cut', kbd: 'Ctrl+X', action: () => copyAnn(true) },
    { label: 'วาง', icon: 'paste', kbd: 'Ctrl+V', disabled: !annClip, action: () => pasteAnn() },
    { label: 'ทำสำเนา', icon: 'dup', kbd: 'Ctrl+D', action: () => { copyAnn(false); pasteAnn(); } },
    { sep: true },
    { label: `ลบ${a ? sigKindName(a.sk) : ''}ออกจากหน้า`, icon: 'trash', kbd: 'Delete', danger: true, action: () => { sigCloseProps(); deleteAnn(); } },
  ];
}
function sigRotateSel(d) {
  const a = findAnn(annSel); if (!a) return;
  pushUndo();
  const C = sigCenter(a); a.r = (((a.r || 0) + d) % 360 + 360) % 360; sigPlaceCenter(a, C);
  touchPage(pageByUid(annSel.uid)); refresh();
}

// ---- colour / thickness sheet
function sigOpenProps() {
  const a = findAnn(annSel); if (!a || a.kind !== 'img') return;
  sigClosePanel();
  SIG.props = { undo: false };
  $('sigWeight').value = +(a.weight || 0);
  sigPropsSync();
  $('sigProps').classList.remove('hidden');
  document.body.classList.add('sigsheet');
}
function sigCloseProps() { SIG.props = null; $('sigProps').classList.add('hidden'); if (!$('sigPanel') || $('sigPanel').classList.contains('hidden')) document.body.classList.remove('sigsheet'); }
function sigPropsSync() {
  const a = findAnn(annSel);
  if (!a || a.kind !== 'img') { if (SIG.props) sigCloseProps(); return; }
  const w = +(a.weight || 0);
  $('sigWeightVal').textContent = (w > 0 ? '+' : '') + w.toFixed(1);
  $('sigColors').querySelectorAll('button').forEach((b) => b.classList.toggle('on', (b.dataset.c || null) === (a.color || null)));
}
function sigApplyProps(patch) {
  const a = findAnn(annSel); if (!a || a.kind !== 'img') return;
  if (!SIG.props) SIG.props = { undo: false };
  if (!SIG.props.undo) { pushUndo(); SIG.props.undo = true; }
  const all = $('sigAll').checked;
  const targets = [];
  for (const p of D.pages) for (const x of p.annots || []) if (x === a || (all && x.kind === 'img' && x.sk === a.sk)) targets.push([p, x]);
  const touched = new Set();
  for (const [p, x] of targets) { Object.assign(x, patch); touched.add(p); }
  touched.forEach(touchPage);
  renderAllAnnLayers(); sigPropsSync(); updateStatus();
  clearTimeout(SIG.thumbT); SIG.thumbT = setTimeout(() => { if (mode === 'organize') renderGrid(); }, 400);
}

// ---- library panel (bottom sheet)
async function sigOpenPanel(kind, target) {
  if (!D.pages.length) { toast('เปิดเอกสารก่อน แล้วจึงลงลายเซ็น'); return; }
  if (mode !== 'read') { setMode('read'); await new Promise((r) => setTimeout(r, 120)); }
  if (annTool) setAnnTool(false);
  sigCloseProps();
  if (kind) SIG.tab = kind;
  SIG.target = target || null;
  $('sigPanel').classList.remove('hidden');
  document.body.classList.add('sigsheet');
  $('annSig').classList.add('active');
  await sigRenderPanel();
}
function sigClosePanel() {
  if (!$('sigPanel')) return;
  $('sigPanel').classList.add('hidden');
  if (!SIG.props) document.body.classList.remove('sigsheet');
  $('annSig').classList.remove('active');
  SIG.target = null;
}
async function sigRenderPanel() {
  $('sigPanel').querySelectorAll('.sigtabs button').forEach((b) => b.classList.toggle('on', b.dataset.k === SIG.tab));
  const list = $('sigList');
  const items = (await sigLoadLibrary()).filter((it) => it.kind === SIG.tab);
  const add = document.createElement('button');
  add.className = 'sigadd'; add.innerHTML = '<span>+</span>เพิ่ม'; add.title = 'เพิ่ม' + sigKindName(SIG.tab) + 'ใหม่: วาด / สแกน / นำเข้ารูป';
  add.onclick = (e) => { const r = add.getBoundingClientRect(); sigAddMenu(SIG.tab, r.left, r.top - 8, true); };
  const tiles = await Promise.all(items.map(async (it) => {
    const b = document.createElement('button');
    b.className = 'sigitem'; b.dataset.id = it.id; b.title = 'แตะเพื่อวางลงหน้า · คลิกขวา/กดค้างเพื่อแก้ไขหรือลบ';
    const im = new Image(); im.src = await sigThumb(it); im.alt = ''; im.draggable = false; b.appendChild(im);
    return b;
  }));
  list.replaceChildren(add, ...tiles);
  $('sigHint').textContent = items.length ? (SIG.target ? 'แตะเพื่อวางตรงจุดที่เลือกไว้' : 'แตะเพื่อวางลงหน้าที่กำลังอ่าน · คลิกขวา/กดค้างที่รายการเพื่อแก้ไขหรือลบ')
    : `ยังไม่มี${sigKindName(SIG.tab)}ในคลัง — กด “+ เพิ่ม” เพื่อวาด สแกน หรือนำเข้ารูป`;
}
function sigAddMenu(kind, x, y, above) {
  const items = [
    { head: 'เพิ่ม' + sigKindName(kind) },
    { label: 'วาด', icon: 'pen', action: () => sigOpenDraw(kind) },
    { label: IS_TOUCH ? 'สแกน (ถ่ายรูปจากกระดาษ)' : 'สแกน / ถ่ายรูป', icon: 'camera', action: () => sigPickImage(kind, true) },
    { label: 'นำเข้า (เลือกรูปในเครื่อง)', icon: 'img', action: () => sigPickImage(kind, false) },
  ];
  showMenu(items, x, y);
  if (above && ctxEl) ctxEl.style.top = Math.max(4, y - ctxEl.offsetHeight) + 'px';
}
function sigItemMenu(it, x, y) {
  showMenu([
    { label: 'วางลงหน้า', icon: 'check', action: () => sigPlace(it) },
    { label: 'แก้ไข (ตัดกรอบ / ลบส่วนเกิน)…', icon: 'erase', action: () => sigEditItem(it) },
    { label: it.kind === 'stamp' ? 'ย้ายไปแท็บลายเซ็น' : 'ย้ายไปแท็บตราประทับ', icon: 'tab', action: () => sigMoveItem(it) },
    { sep: true },
    { label: 'ลบออกจากคลัง…', icon: 'trash', danger: true, action: () => sigDeleteItem(it) },
  ], x, y);
}
async function sigDeleteItem(it) {
  const r = await dialog({ title: 'ลบออกจากคลัง', body: `<p>ยืนยันจะลบ${sigKindName(it.kind)}นี้ออกจากคลังหรือไม่?</p><p class="hint">${sigKindName(it.kind)}ที่วางไว้บนเอกสารแล้วจะยังอยู่ตามเดิม</p>`, buttons: [['cancel', 'ยกเลิก'], ['ok', 'ลบ', true]] });
  if (!r.btn) return;
  try { await sigStore.del(it.id); } catch (e) { toast('ลบไม่สำเร็จ: ' + e.message, 'err'); }
  await sigLoadLibrary(true); sigRenderPanel();
}
async function sigMoveItem(it) {
  try { await sigStore.put({ ...it, kind: it.kind === 'stamp' ? 'sig' : 'stamp' }, null); } catch (e) { toast('ย้ายไม่สำเร็จ: ' + e.message, 'err'); }
  await sigLoadLibrary(true); sigRenderPanel();
}
async function sigEditItem(it) {
  const url = await sigThumb(it);
  const im = await new Promise((res) => { const i = new Image(); i.onload = () => res(i); i.src = url; });
  const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight; c.getContext('2d').drawImage(im, 0, 0);
  sigOpenEditor(c, it.kind, { title: 'แก้ไข' + sigKindName(it.kind), keepAlpha: true, startClean: true }, async (out) => {
    const blob = await new Promise((r) => out.toBlob(r, 'image/png'));
    const meta = { ...it, w: out.width, h: out.height, v: (it.v || 1) + 1 };
    try { await sigStore.put(meta, blob); } catch (e) { toast('บันทึกไม่สำเร็จ: ' + e.message, 'err'); return; }
    SIG.thumbs.set(meta.id + ':' + meta.v, URL.createObjectURL(blob));
    SIG.base.set('L' + meta.id + 'v' + meta.v, out);
    await sigLoadLibrary(true);
    if (!$('sigPanel').classList.contains('hidden')) sigRenderPanel();
    toast('บันทึกการแก้ไขในคลังแล้ว (ชิ้นที่วางไว้บนหน้าแล้วไม่เปลี่ยน)', 'ok', [], 3500);
  });
}

// ---- place a library item on the page
function sigVisibleCenter(p) {
  const i = D.pages.indexOf(p);
  const L = document.querySelector(`.annlayer[data-i="${i}"]`);
  if (!L) return null;
  const R = $('reader').getBoundingClientRect(), r = L.parentElement.getBoundingClientRect();
  const top = Math.max(R.top, r.top), bot = Math.min(R.bottom, r.bottom);
  const cy = bot > top ? (top + bot) / 2 : r.top + r.height / 2;
  const cx = Math.max(r.left, Math.min(r.right, R.left + R.width / 2));
  return { L, pt: annPoint(L, { clientX: cx, clientY: cy }) };
}
async function sigPlace(it) {
  let key;
  try { key = await sigEnsureBase(it); } catch (e) { toast(e.message, 'err'); return; }
  const tg = SIG.target;
  let p = tg && D.pages.includes(tg.p) ? tg.p : currentReaderPage();
  if (!p) return;
  const vc = sigVisibleCenter(p);
  const pt = tg && tg.p === p ? tg.pt : vc ? vc.pt : null;
  const L = vc ? vc.L : null;
  const s = L ? +L.dataset.scale : 1;
  const bw = L ? L.offsetWidth / s : 595, bh = L ? L.offsetHeight / s : 842;
  const stamp = it.kind === 'stamp';
  let w = stamp ? 110 : 150, h = w * it.h / it.w;
  const maxH = stamp ? 110 : 64;
  if (h > maxH) { h = maxH; w = h * it.w / it.h; }
  if (w > bw * 0.7) { w = bw * 0.7; h = w * it.h / it.w; }
  const a = { id: annId(), kind: 'img', sk: it.kind, base: key, x: 0, y: 0, w, h, r: (360 - (p.rot || 0)) % 360, color: null, weight: 0 };
  sigPlaceCenter(a, pt || [bw / 2, bh / 2]);
  pushUndo();
  p.annots = [...(p.annots || []), a];
  touchPage(p);
  sigClosePanel();
  selectAnn({ uid: p.uid, id: a.id });
  refresh();
}

// ---- date & text at a point
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const TH_MONL = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
function sigDateFormats(d = new Date()) {
  const dd = String(d.getDate()).padStart(2, '0'), mm = String(d.getMonth() + 1).padStart(2, '0'), be = d.getFullYear() + 543;
  return [`${d.getDate()} ${TH_MON[d.getMonth()]} ${be}`, `${d.getDate()} ${TH_MONL[d.getMonth()]} ${be}`, `${dd}/${mm}/${be}`, `${dd}/${mm}/${String(be).slice(2)}`, `${dd}/${mm}/${d.getFullYear()}`];
}
function sigTextAt(p, pt, text) {
  const r = (360 - (p.rot || 0)) % 360;
  const [ax, ay] = annLocal({ x: pt[0], y: pt[1], r }, -annDefaults.size * ANN_PAD, -annDefaults.size * ANN_LH / 2);
  const a = newAnnAt(p, ax, ay, text);
  a.r = r;
  if (annTool) setAnnTool(false);
  selectAnn({ uid: p.uid, id: a.id }, !text);
  if (text) refresh();
}
function sigPointMenu(L, ev) {
  const p = D.pages[+L.dataset.i];
  const pt = annPoint(L, ev);
  return [
    { label: 'วางลายเซ็นตรงนี้', icon: 'sign', action: () => sigOpenPanel('sig', { p, pt }) },
    { label: 'วางตราประทับตรงนี้', icon: 'stamp', action: () => sigOpenPanel('stamp', { p, pt }) },
    { label: 'ใส่วันที่ตรงนี้', icon: 'cal', sub: sigDateFormats().map((f) => ({ label: f, action: () => sigTextAt(p, pt, f) })) },
    { label: 'เพิ่มข้อความตรงนี้', icon: 'type', action: () => sigTextAt(p, pt, '') },
  ];
}
// touch: press and hold on a page -> signature / stamp / date / text menu
(() => {
  let timer = 0, st = null;
  const cancel = () => { if (st) { st = null; clearTimeout(timer); } };
  $('readerPages').addEventListener('pointerdown', (e) => {
    if (e.pointerType !== 'touch' || mode !== 'read' || annTool) return;
    const L = e.target.closest('.annlayer');
    if (!L || e.target.closest('.annot')) return;
    st = { x: e.clientX, y: e.clientY, L };
    clearTimeout(timer);
    timer = setTimeout(() => {
      if (!st) return;
      const s0 = st; st = null;
      if (ctxEl && Date.now() - (window.__menuAt || 0) < 800) return; // the browser already opened our menu (contextmenu)
      try { navigator.vibrate && navigator.vibrate(12); } catch {}
      showMenu(sigPointMenu(s0.L, { clientX: s0.x, clientY: s0.y }), s0.x, s0.y);
    }, 520);
  });
  window.addEventListener('pointermove', (e) => { if (st && Math.hypot(e.clientX - st.x, e.clientY - st.y) > 10) cancel(); }, true);
  window.addEventListener('pointerup', cancel, true);
  window.addEventListener('pointercancel', cancel, true);
  $('reader').addEventListener('scroll', cancel, { passive: true });
  $('reader').addEventListener('touchstart', (e) => { if (e.touches.length > 1) cancel(); }, { passive: true });
})();

// library tiles: tap = place, right-click / long-press = menu
(() => {
  const list = $('sigList');
  let timer = 0, st = null, fired = false;
  list.addEventListener('pointerdown', (e) => {
    const b = e.target.closest('.sigitem'); if (!b) return;
    fired = false; st = { x: e.clientX, y: e.clientY, b };
    if (e.pointerType === 'touch') timer = setTimeout(() => { if (!st) return; fired = true; const it = SIG.items.find((x) => x.id === st.b.dataset.id); try { navigator.vibrate && navigator.vibrate(12); } catch {} if (it) sigItemMenu(it, st.x, st.y - 10); st = null; }, 520);
  });
  list.addEventListener('pointermove', (e) => { if (st && Math.hypot(e.clientX - st.x, e.clientY - st.y) > 10) { st = null; clearTimeout(timer); } });
  list.addEventListener('pointerup', () => { clearTimeout(timer); st = null; });
  list.addEventListener('click', (e) => {
    const b = e.target.closest('.sigitem'); if (!b) return;
    if (fired) { fired = false; return; }
    const it = SIG.items.find((x) => x.id === b.dataset.id); if (it) sigPlace(it);
  });
  list.addEventListener('contextmenu', (e) => {
    const b = e.target.closest('.sigitem'); if (!b) return;
    e.preventDefault(); e.stopPropagation();
    if (fired) return;
    const it = SIG.items.find((x) => x.id === b.dataset.id); if (it) sigItemMenu(it, e.clientX, e.clientY);
  });
})();

// ---- export / import the whole library (one file)
async function sigExport() {
  const items = await sigLoadLibrary(true);
  if (!items.length) { toast('คลังยังว่างอยู่'); return; }
  busy('กำลังเตรียมไฟล์คลัง …');
  try {
    const out = [];
    for (const it of items) {
      const b = await sigStore.png(it.id); if (!b) continue;
      const u8 = new Uint8Array(await b.arrayBuffer());
      let bin = ''; for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
      out.push({ ...it, png: btoa(bin) });
    }
    const json = JSON.stringify({ app: 'Super PDF', type: 'signature-library', format: 1, exported: new Date().toISOString(), items: out });
    unbusy();
    const d = new Date(), stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    await saveBytes(new TextEncoder().encode(json), `คลังลายเซ็น SuperPDF ${stamp}.json`, 'json');
  } catch (e) { unbusy(); toast('ส่งออกไม่สำเร็จ: ' + e.message, 'err'); }
}
async function sigImportFile(f) {
  let data;
  try { data = JSON.parse(await f.text()); } catch { toast('ไฟล์นี้ไม่ใช่ไฟล์คลังลายเซ็นของ Super PDF', 'err'); return; }
  const items = (data && data.type === 'signature-library' && Array.isArray(data.items)) ? data.items.filter((x) => x && x.id && x.png) : null;
  if (!items || !items.length) { toast('ไม่พบรายการในไฟล์นี้', 'err'); return; }
  const cur = await sigLoadLibrary(true);
  const r = await dialog({
    title: 'นำเข้าคลังลายเซ็น',
    body: `<p>พบ ${items.length} รายการ (ลายเซ็น ${items.filter((x) => x.kind !== 'stamp').length} · ตราประทับ ${items.filter((x) => x.kind === 'stamp').length})</p>` +
      (cur.length ? `<label class="radio"><input type="radio" name="m" value="merge" checked> รวมกับคลังเดิม (${cur.length} รายการ)</label><label class="radio"><input type="radio" name="m" value="replace"> แทนที่คลังเดิมทั้งหมด</label>` : ''),
    buttons: [['cancel', 'ยกเลิก'], ['ok', 'นำเข้า', true]],
  });
  if (!r.btn) return;
  busy('กำลังนำเข้า …');
  try {
    if (r.values.m === 'replace') for (const it of cur) await sigStore.del(it.id);
    for (const it of items) {
      const bin = atob(it.png), u8 = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      const id = /^[A-Za-z0-9_-]{1,40}$/.test(it.id) ? it.id : sigNewId();
      await sigStore.put({ id, kind: it.kind, w: it.w, h: it.h, t: it.t, v: (it.v || 1) }, new Blob([u8], { type: 'image/png' }));
    }
    SIG.thumbs.forEach((u) => u && URL.revokeObjectURL(u)); SIG.thumbs.clear();
    await sigLoadLibrary(true);
    unbusy();
    toast(`นำเข้าแล้ว ${items.length} รายการ`, 'ok', [], 3000);
    if (!$('sigPanel').classList.contains('hidden')) sigRenderPanel();
  } catch (e) { unbusy(); toast('นำเข้าไม่สำเร็จ: ' + e.message, 'err'); }
}
function sigLibMenu(x, y) {
  showMenu([
    { head: 'ที่เก็บ: ' + sigStore.where() },
    { label: 'ส่งออกคลังเป็นไฟล์ (สำรอง / ย้ายเครื่อง)…', icon: 'save', action: sigExport },
    { label: 'นำเข้าคลังจากไฟล์…', icon: 'open', action: () => $('sigImportFile').click() },
  ], x, y);
}

// ---- pick a picture (scan = camera on phones)
function sigPickImage(kind, camera) {
  const inp = $(camera && IS_TOUCH ? 'sigCam' : 'sigPick');
  inp.dataset.kind = kind;
  inp.click();
}
for (const id of ['sigCam', 'sigPick']) $(id).onchange = async (e) => {
  const f = e.target.files[0]; const kind = e.target.dataset.kind || 'sig'; e.target.value = '';
  if (!f) return;
  busy('กำลังเปิดรูป …');
  let c;
  try { c = await photoToCanvas(f, 2600); } catch (err) { unbusy(); toast('เปิดรูปไม่ได้: ' + err.message, 'err'); return; }
  unbusy();
  sigOpenEditor(c, kind, { title: 'เพิ่ม' + sigKindName(kind) }, async (out) => {
    const it = await sigAddToLibrary(out, kind);
    sigAfterCreate(it);
  });
};
$('sigImportFile').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) sigImportFile(f); };
function sigAfterCreate(it) { // new item goes to the library and straight onto the page (like CamScanner)
  if (mode === 'read' && D.pages.length) sigPlace(it);
  else toast(`บันทึก${sigKindName(it.kind)}ลงคลังแล้ว`, 'ok', [], 2500);
}

// ------------------------------------------------------------------ draw a signature (วาด)
const SD = { strokes: [], redo: [], cur: null, color: '#111111', width: 3.2, kind: 'sig', rot: false };
function sigOpenDraw(kind) {
  SD.kind = kind; SD.strokes = []; SD.redo = []; SD.cur = null;
  $('sigDrawTitle').textContent = 'สร้าง' + sigKindName(kind);
  $('sigDraw').classList.remove('hidden');
  sigDrawLayout();
  sigDrawPaint();
}
function sigDrawClose() { $('sigDraw').classList.add('hidden'); }
function sigDrawLayout() {
  const box = $('sigDrawBox');
  // phones held upright: turn the pad sideways (landscape) so there is room to sign — like CamScanner
  SD.rot = IS_TOUCH && innerHeight > innerWidth && innerWidth < 760;
  box.classList.toggle('rot', SD.rot);
  if (SD.rot) { box.style.width = innerHeight + 'px'; box.style.height = innerWidth + 'px'; }
  else { box.style.width = ''; box.style.height = ''; }
  const cv = $('sigPad'), dpr = window.devicePixelRatio || 1;
  const w = cv.clientWidth, h = cv.clientHeight;
  cv.width = Math.round(w * dpr); cv.height = Math.round(h * dpr);
  SD.dpr = dpr;
}
function sigPadPoint(e) { // pointer -> pad css px (handles the rotated layout)
  const cv = $('sigPad'), r = cv.getBoundingClientRect();
  if (!SD.rot) return [e.clientX - r.left, e.clientY - r.top];
  return [(e.clientY - r.top) * cv.clientWidth / r.height, (r.right - e.clientX) * cv.clientHeight / r.width];
}
function sigStrokePath(g, st) {
  const P = st.pts;
  g.strokeStyle = st.color; g.fillStyle = st.color; g.lineCap = 'round'; g.lineJoin = 'round';
  if (P.length === 1) { g.beginPath(); g.arc(P[0][0], P[0][1], P[0][2] / 2, 0, Math.PI * 2); g.fill(); return; }
  for (let i = 1; i < P.length; i++) {
    const a = P[i - 1], b = P[i], pa = P[i - 2] || a;
    const m0 = [(pa[0] + a[0]) / 2, (pa[1] + a[1]) / 2], m1 = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
    g.lineWidth = (a[2] + b[2]) / 2;
    g.beginPath(); g.moveTo(i === 1 ? a[0] : m0[0], i === 1 ? a[1] : m0[1]); g.quadraticCurveTo(a[0], a[1], m1[0], m1[1]);
    if (i === P.length - 1) g.lineTo(b[0], b[1]);
    g.stroke();
  }
}
function sigDrawPaint() {
  const cv = $('sigPad'), g = cv.getContext('2d');
  g.setTransform(1, 0, 0, 1, 0, 0); g.clearRect(0, 0, cv.width, cv.height);
  g.setTransform(SD.dpr, 0, 0, SD.dpr, 0, 0);
  for (const st of SD.strokes) sigStrokePath(g, st);
  if (SD.cur) sigStrokePath(g, SD.cur);
  const has = SD.strokes.length > 0;
  $('sigPadPh').classList.toggle('hidden', has || !!SD.cur);
  $('sigDrawOk').disabled = !has;
  $('sigDrawUndo').disabled = !has; $('sigDrawRedo').disabled = !SD.redo.length; $('sigDrawClear').disabled = !has;
}
(() => {
  const cv = $('sigPad');
  cv.addEventListener('pointerdown', (e) => {
    if (SD.cur) return;
    e.preventDefault();
    try { cv.setPointerCapture(e.pointerId); } catch {}
    const [x, y] = sigPadPoint(e);
    const w = SD.width * (e.pointerType === 'pen' && e.pressure ? 0.6 + e.pressure * 0.8 : 1);
    SD.cur = { color: SD.color, pts: [[x, y, w]], id: e.pointerId, t: performance.now(), lw: w };
    sigDrawPaint();
  });
  cv.addEventListener('pointermove', (e) => {
    const st = SD.cur; if (!st || e.pointerId !== st.id) return;
    e.preventDefault();
    const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
    for (const ev of evs.length ? evs : [e]) {
      const [x, y] = sigPadPoint(ev);
      const last = st.pts[st.pts.length - 1];
      const dist = Math.hypot(x - last[0], y - last[1]); if (dist < 0.8) continue;
      const now = performance.now(), v = dist / Math.max(1, now - st.t); st.t = now;
      let w = SD.width * Math.max(0.55, Math.min(1.25, 1.3 - v * 0.35));
      if (ev.pointerType === 'pen' && ev.pressure) w = SD.width * (0.5 + ev.pressure * 0.9);
      w = st.lw * 0.6 + w * 0.4; st.lw = w;
      st.pts.push([x, y, w]);
    }
    sigDrawPaint();
  });
  const end = (e) => {
    const st = SD.cur; if (!st || e.pointerId !== st.id) return;
    SD.cur = null; SD.strokes.push(st); SD.redo = []; sigDrawPaint();
  };
  cv.addEventListener('pointerup', end); cv.addEventListener('pointercancel', end);
  $('sigDrawCancel').onclick = sigDrawClose;
  $('sigDrawUndo').onclick = () => { if (SD.strokes.length) { SD.redo.push(SD.strokes.pop()); sigDrawPaint(); } };
  $('sigDrawRedo').onclick = () => { if (SD.redo.length) { SD.strokes.push(SD.redo.pop()); sigDrawPaint(); } };
  $('sigDrawClear').onclick = () => { if (SD.strokes.length) { SD.redo = []; SD.strokes = []; sigDrawPaint(); } };
  $('sigDrawPens').querySelectorAll('button').forEach((b) => b.onclick = () => {
    if (b.dataset.w) { SD.width = +b.dataset.w; $('sigDrawPens').querySelectorAll('[data-w]').forEach((x) => x.classList.toggle('on', x === b)); }
    if (b.dataset.c) { SD.color = b.dataset.c; $('sigDrawPens').querySelectorAll('[data-c]').forEach((x) => x.classList.toggle('on', x === b)); }
  });
  $('sigDrawOk').onclick = async () => {
    const out = sigDrawExport(); if (!out) return;
    sigDrawClose();
    const it = await sigAddToLibrary(out, SD.kind);
    sigAfterCreate(it);
  };
  window.addEventListener('resize', () => { if (!$('sigDraw').classList.contains('hidden')) { sigDrawLayout(); sigDrawPaint(); } });
})();
function sigDrawExport() {
  if (!SD.strokes.length) return null;
  let x0 = 1e9, y0 = 1e9, x1 = -1e9, y1 = -1e9;
  for (const st of SD.strokes) for (const [x, y, w] of st.pts) { x0 = Math.min(x0, x - w); y0 = Math.min(y0, y - w); x1 = Math.max(x1, x + w); y1 = Math.max(y1, y + w); }
  const pd = 6; x0 -= pd; y0 -= pd; x1 += pd; y1 += pd;
  const bw = x1 - x0, bh = y1 - y0;
  const k = Math.min(4, 1400 / bw, 900 / bh);
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(bw * k)); c.height = Math.max(1, Math.round(bh * k));
  const g = c.getContext('2d'); g.setTransform(k, 0, 0, k, -x0 * k, -y0 * k);
  for (const st of SD.strokes) sigStrokePath(g, st);
  return c;
}

// ------------------------------------------------------------------ crop + clean-up editor (ตัดกรอบ / ลบส่วนเกิน)
// step 1: crop rectangle on the photo (rotate 90°); step 2: the paper is removed (soft alpha from ink darkness) and the
// user erases leftovers: brush eraser, tap = remove a whole piece, drag a box = remove everything inside, auto despeckle.
const SE = {
  open: false, src: null, kind: 'sig', onDone: null, keepAlpha: false,
  step: 'crop', crop: null, cropUsed: null,
  W: 0, H: 0, rgb: null, alpha: null, col: null, er: null, undo: [], redo: [],
  tool: 'erase', size: 28, group: 6, thr: 50,
  view: { z: 1, tx: 0, ty: 0 }, ctx: null, img: null,
};
function sigOpenEditor(src, kind, opts, onDone) {
  Object.assign(SE, { open: true, src, kind, onDone, keepAlpha: !!opts.keepAlpha || sigHasAlpha(src), step: 'crop', er: null, undo: [], redo: [], cropUsed: null });
  SE.crop = SE.keepAlpha ? sigAlphaBox(src) : sigAutoCrop(src);
  $('sigEdTitle').textContent = opts.title || 'เพิ่ม' + sigKindName(kind);
  $('sigEdThrWrap').classList.toggle('hidden', SE.keepAlpha);
  $('sigEd').classList.remove('hidden');
  sigEdStep(opts.startClean ? 'clean' : 'crop');
}
function sigCloseEditor() { SE.open = false; $('sigEd').classList.add('hidden'); SE.src = SE.rgb = SE.alpha = SE.col = SE.er = null; SE.undo = []; SE.redo = []; }
function sigHasAlpha(c) {
  const g = c.getContext('2d'), d = g.getImageData(0, 0, c.width, c.height).data;
  let t = 0, n = 0;
  for (let i = 3; i < d.length; i += 4 * 7) { n++; if (d[i] < 200) t++; }
  return t / n > 0.03;
}
function sigAlphaBox(c) {
  const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data, W = c.width, H = c.height;
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) if (d[(y * W + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  if (x1 < 0) return { x: 0, y: 0, w: W, h: H };
  const p = Math.round(Math.max(W, H) * 0.02);
  x0 = Math.max(0, x0 - p); y0 = Math.max(0, y0 - p); x1 = Math.min(W, x1 + p + 1); y1 = Math.min(H, y1 + p + 1);
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}
// ink box on a small copy: paper removed, 1st..99th percentile of ink positions, + margin
function sigAutoCrop(src) {
  const k = Math.min(1, 700 / Math.max(src.width, src.height));
  const W = Math.max(1, Math.round(src.width * k)), H = Math.max(1, Math.round(src.height * k));
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const g = c.getContext('2d'); g.drawImage(src, 0, 0, W, H);
  const d = g.getImageData(0, 0, W, H).data, N = W * H, lum = new Uint8Array(N);
  for (let i = 0, j = 0; i < N; i++, j += 4) lum[i] = Math.min(d[j], d[j + 1], d[j + 2]);
  const bg = backgroundMap(lum, W, H, 6);
  const xs = [], ys = [];
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const i = y * W + x; if (1 - lum[i] / bg[i] > 0.3) { xs.push(x); ys.push(y); } }
  const full = { x: 0, y: 0, w: src.width, h: src.height };
  if (xs.length < 20) return full;
  xs.sort((a, b) => a - b); ys.sort((a, b) => a - b);
  const q = (A, f) => A[Math.min(A.length - 1, Math.max(0, Math.round(f * (A.length - 1))))];
  let x0 = q(xs, 0.01), x1 = q(xs, 0.99), y0 = q(ys, 0.01), y1 = q(ys, 0.99);
  const mx = (x1 - x0) * 0.08 + 6, my = (y1 - y0) * 0.15 + 6;
  x0 = Math.max(0, x0 - mx); y0 = Math.max(0, y0 - my); x1 = Math.min(W, x1 + mx); y1 = Math.min(H, y1 + my);
  if ((x1 - x0) * (y1 - y0) > 0.9 * W * H) return full;
  return { x: Math.round(x0 / k), y: Math.round(y0 / k), w: Math.round((x1 - x0) / k), h: Math.round((y1 - y0) / k) };
}

// ---- steps
function sigEdStep(step) {
  if (step === 'clean' && SE.crop.w < 4) return;
  SE.step = step;
  $('sigEd').dataset.step = step;
  $('sigEd').querySelectorAll('.sigsteps button').forEach((b) => b.classList.toggle('on', b.dataset.s === step));
  $('sigEdCropBar').classList.toggle('hidden', step !== 'crop');
  $('sigEdCleanBar').classList.toggle('hidden', step !== 'clean');
  $('sigEdUndo').classList.toggle('hidden', step !== 'clean'); $('sigEdRedo').classList.toggle('hidden', step !== 'clean');
  $('sigEdOkLbl').textContent = step === 'crop' ? 'ถัดไป' : 'เสร็จแล้ว';
  $('sigCropBox').classList.toggle('hidden', step !== 'crop');
  $('sigEdCanvas').classList.toggle('checker', step === 'clean');
  if (step === 'crop') {
    $('sigEdStep').textContent = 'ขั้นที่ 1/2 · ลากมุมหรือขอบของกรอบให้คลุมเฉพาะ' + sigKindName(SE.kind);
    const cv = $('sigEdCanvas'); cv.width = SE.src.width; cv.height = SE.src.height;
    cv.getContext('2d').drawImage(SE.src, 0, 0);
    sigEdFit(); sigCropDraw();
  } else {
    $('sigEdStep').textContent = 'ขั้นที่ 2/2 · ลบสิ่งที่ไม่ใช่' + sigKindName(SE.kind) + (IS_TOUCH ? '' : ' (ลายตารางหมากรุก = ส่วนที่โปร่งใส)');
    const key = JSON.stringify(SE.crop);
    if (SE.cropUsed !== key) { SE.cropUsed = key; sigEdBuild(); }
    const cv = $('sigEdCanvas'); cv.width = SE.W; cv.height = SE.H;
    SE.ctx = cv.getContext('2d'); SE.img = SE.ctx.createImageData(SE.W, SE.H);
    sigEdRender();
    sigEdFit(); sigEdTool(SE.tool);
  }
  sigEdUndoBtns();
}
function sigEdBuild() { // crop -> working image -> extract ink
  const maxS = IS_TOUCH ? 1100 : 1500;
  const k = Math.min(1, maxS / Math.max(SE.crop.w, SE.crop.h));
  const W = Math.max(1, Math.round(SE.crop.w * k)), H = Math.max(1, Math.round(SE.crop.h * k));
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const g = c.getContext('2d'); g.imageSmoothingQuality = 'high';
  g.drawImage(SE.src, SE.crop.x, SE.crop.y, SE.crop.w, SE.crop.h, 0, 0, W, H);
  SE.W = W; SE.H = H; SE.rgb = g.getImageData(0, 0, W, H).data;
  SE.er = new Uint8Array(W * H); SE.undo = []; SE.redo = [];
  SE.size = Math.round(Math.max(12, Math.min(60, Math.min(W, H) / 10)));
  sigEdExtract();
}
function sigEdExtract() {
  const W = SE.W, H = SE.H, N = W * H, d = SE.rgb;
  SE.alpha = new Float32Array(N); SE.col = new Uint8ClampedArray(N * 3);
  if (SE.keepAlpha) {
    for (let i = 0; i < N; i++) { SE.alpha[i] = d[i * 4 + 3] / 255; SE.col[i * 3] = d[i * 4]; SE.col[i * 3 + 1] = d[i * 4 + 1]; SE.col[i * 3 + 2] = d[i * 4 + 2]; }
    return;
  }
  const B = Math.max(6, Math.round(Math.min(W, H) / 70));
  const bg = [0, 1, 2].map((k) => { const ch = new Uint8Array(N); for (let i = 0, j = k; i < N; i++, j += 4) ch[i] = d[j]; return backgroundMap(ch, W, H, B); });
  const t0 = 0.05 + (100 - SE.thr) / 100 * 0.26, t1 = t0 + 0.30;
  for (let i = 0, j = 0; i < N; i++, j += 4) {
    const nr = Math.min(1, d[j] / bg[0][i]), ng = Math.min(1, d[j + 1] / bg[1][i]), nb = Math.min(1, d[j + 2] / bg[2][i]);
    const ink = 1 - Math.min(nr, ng, nb);
    let a = (ink - t0) / (t1 - t0); a = a <= 0 ? 0 : a >= 1 ? 1 : a * a * (3 - 2 * a);
    SE.alpha[i] = a;
    if (a > 0.01) { // un-mix the ink colour from the white paper
      const q = 1 - a;
      SE.col[i * 3] = Math.max(0, Math.min(1, (nr - q) / a)) * 255;
      SE.col[i * 3 + 1] = Math.max(0, Math.min(1, (ng - q) / a)) * 255;
      SE.col[i * 3 + 2] = Math.max(0, Math.min(1, (nb - q) / a)) * 255;
    }
  }
}
function sigEdRender(x0 = 0, y0 = 0, x1 = SE.W, y1 = SE.H) {
  x0 = Math.max(0, x0 | 0); y0 = Math.max(0, y0 | 0); x1 = Math.min(SE.W, Math.ceil(x1)); y1 = Math.min(SE.H, Math.ceil(y1));
  if (x1 <= x0 || y1 <= y0) return;
  const o = SE.img.data, W = SE.W;
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
    const i = y * W + x, j = i * 4;
    o[j] = SE.col[i * 3]; o[j + 1] = SE.col[i * 3 + 1]; o[j + 2] = SE.col[i * 3 + 2];
    o[j + 3] = SE.alpha[i] * (255 - SE.er[i]);
  }
  SE.ctx.putImageData(SE.img, 0, 0, x0, y0, x1 - x0, y1 - y0);
}
function sigEdResult() {
  const W = SE.W, H = SE.H;
  let x0 = W, y0 = H, x1 = -1, y1 = -1;
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const i = y * W + x; if (SE.alpha[i] * (255 - SE.er[i]) > 10) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } }
  if (x1 < 0) return null;
  const p = Math.round(Math.max(4, Math.max(x1 - x0, y1 - y0) * 0.02));
  x0 = Math.max(0, x0 - p); y0 = Math.max(0, y0 - p); x1 = Math.min(W - 1, x1 + p); y1 = Math.min(H - 1, y1 + p);
  const w = x1 - x0 + 1, h = y1 - y0 + 1;
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const g = c.getContext('2d'), im = g.createImageData(w, h), o = im.data;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y + y0) * W + x + x0, j = (y * w + x) * 4;
    const a = SE.alpha[i] * (255 - SE.er[i]);
    if (a < 3) continue;
    o[j] = SE.col[i * 3]; o[j + 1] = SE.col[i * 3 + 1]; o[j + 2] = SE.col[i * 3 + 2]; o[j + 3] = a;
  }
  g.putImageData(im, 0, 0);
  const k = Math.min(1, 1200 / Math.max(w, h));
  if (k >= 1) return c;
  const c2 = document.createElement('canvas'); c2.width = Math.round(w * k); c2.height = Math.round(h * k);
  const g2 = c2.getContext('2d'); g2.imageSmoothingQuality = 'high'; g2.drawImage(c, 0, 0, c2.width, c2.height);
  return c2;
}

// ---- view (fit / zoom / pan)
function sigEdFit() {
  const st = $('sigEdStage'), cv = $('sigEdCanvas');
  const pad = SE.step === 'crop' ? 28 : 16;
  const z = Math.min((st.clientWidth - pad * 2) / cv.width, (st.clientHeight - pad * 2) / cv.height, SE.step === 'crop' ? 4 : 8);
  SE.view = { z, tx: (st.clientWidth - cv.width * z) / 2, ty: (st.clientHeight - cv.height * z) / 2 };
  sigEdApplyView();
}
function sigEdApplyView() {
  const v = SE.view, cv = $('sigEdCanvas');
  cv.style.width = cv.width * v.z + 'px'; cv.style.height = cv.height * v.z + 'px';
  $('sigEdWrap').style.transform = `translate(${v.tx}px,${v.ty}px)`;
  $('sigEdZoomLbl').textContent = Math.round(v.z * 100) + '%';
  if (SE.step === 'crop') sigCropDraw();
}
function sigEdZoomAt(z, cx, cy) {
  const v = SE.view, st = $('sigEdStage').getBoundingClientRect();
  if (cx == null) { cx = st.left + st.width / 2; cy = st.top + st.height / 2; }
  z = Math.max(0.05, Math.min(16, z));
  const px = (cx - st.left - v.tx) / v.z, py = (cy - st.top - v.ty) / v.z;
  SE.view = { z, tx: cx - st.left - px * z, ty: cy - st.top - py * z };
  sigEdApplyView();
}
const sigEdImgPt = (e) => { const st = $('sigEdStage').getBoundingClientRect(), v = SE.view; return [(e.clientX - st.left - v.tx) / v.z, (e.clientY - st.top - v.ty) / v.z]; };

// ---- crop box
function sigCropDraw() {
  const b = $('sigCropBox'), v = SE.view, c = SE.crop;
  Object.assign(b.style, { left: v.tx + c.x * v.z + 'px', top: v.ty + c.y * v.z + 'px', width: c.w * v.z + 'px', height: c.h * v.z + 'px' });
}
function sigCropRotate() {
  const s = SE.src, c = SE.crop;
  SE.src = rotateCanvas(s, -90);
  SE.crop = { x: c.y, y: s.width - c.x - c.w, w: c.h, h: c.w };
  SE.cropUsed = null;
  sigEdStep('crop');
}

// ---- undo inside the editor (erase mask snapshots)
function sigEdPush() { SE.undo.push(SE.er.slice()); if (SE.undo.length > (IS_TOUCH ? 12 : 25)) SE.undo.shift(); SE.redo = []; sigEdUndoBtns(); }
function sigEdUndoBtns() { $('sigEdUndo').disabled = !SE.undo.length; $('sigEdRedo').disabled = !SE.redo.length; }
function sigEdUndo() { if (!SE.undo.length) return; SE.redo.push(SE.er); SE.er = SE.undo.pop(); sigEdRender(); sigEdUndoBtns(); }
function sigEdRedo() { if (!SE.redo.length) return; SE.undo.push(SE.er); SE.er = SE.redo.pop(); sigEdRender(); sigEdUndoBtns(); }

// ---- tools
function sigEdTool(t) {
  SE.tool = t;
  $('sigEdCleanBar').querySelectorAll('[data-t]').forEach((b) => b.classList.toggle('active', b.dataset.t === t));
  $('sigEdSizeLbl').textContent = t === 'erase' ? 'ขนาดยางลบ' : 'จับกลุ่ม';
  const inp = $('sigEdSize');
  if (t === 'erase') { inp.min = 4; inp.max = 140; inp.value = SE.size; }
  else { inp.min = 0; inp.max = 40; inp.value = SE.group; }
  $('sigEdStage').dataset.tool = t;
  $('sigEdHint').textContent = t === 'erase' ? 'ระบายทับส่วนที่ไม่ต้องการ · สองนิ้วหรือ Ctrl+ล้อเมาส์ = ซูม'
    : 'แตะที่ตัวอักษร/เส้นที่ไม่ต้องการเพื่อลบทั้งชิ้น · ลากเป็นกรอบเพื่อลบทุกอย่างในกรอบ';
}
function sigEdStamp(x, y, r) {
  const W = SE.W, H = SE.H, er = SE.er;
  const x0 = Math.max(0, Math.floor(x - r - 1)), x1 = Math.min(W - 1, Math.ceil(x + r + 1));
  const y0 = Math.max(0, Math.floor(y - r - 1)), y1 = Math.min(H - 1, Math.ceil(y + r + 1));
  for (let yy = y0; yy <= y1; yy++) for (let xx = x0; xx <= x1; xx++) {
    const dd = Math.hypot(xx - x, yy - y);
    if (dd <= r + 0.5) { const v = dd <= r - 0.5 ? 255 : Math.round((r + 0.5 - dd) * 255); const i = yy * W + xx; if (v > er[i]) er[i] = v; }
  }
}
function sigEdEraseSeg(a, b, r) {
  const L = Math.hypot(b[0] - a[0], b[1] - a[1]), n = Math.max(1, Math.ceil(L / Math.max(1, r / 3)));
  for (let k = 0; k <= n; k++) sigEdStamp(a[0] + (b[0] - a[0]) * k / n, a[1] + (b[1] - a[1]) * k / n, r);
  sigEdRender(Math.min(a[0], b[0]) - r - 2, Math.min(a[1], b[1]) - r - 2, Math.max(a[0], b[0]) + r + 2, Math.max(a[1], b[1]) + r + 2);
}
const sigInk = (i) => SE.alpha[i] * (255 - SE.er[i]) > 50;
// tap: remove the piece of ink under the finger (pieces closer than "group" px count as one)
function sigEdPick(x, y) {
  const W = SE.W, H = SE.H, N = W * H, g = Math.round(SE.group * Math.max(1, Math.min(W, H) / 400));
  let ink = new Float32Array(N);
  for (let i = 0; i < N; i++) ink[i] = sigInk(i) ? 1 : 0;
  const dil = g ? sigMorph(ink, W, H, g, true) : ink;
  // nearest ink near the tap
  const R = Math.round(Math.max(8, Math.min(W, H) * 0.04));
  let best = -1, bd = 1e9;
  for (let yy = Math.max(0, Math.round(y) - R); yy <= Math.min(H - 1, Math.round(y) + R); yy++)
    for (let xx = Math.max(0, Math.round(x) - R); xx <= Math.min(W - 1, Math.round(x) + R); xx++) {
      const i = yy * W + xx; if (!dil[i]) continue;
      const dd = (xx - x) ** 2 + (yy - y) ** 2; if (dd < bd) { bd = dd; best = i; }
    }
  if (best < 0) { toast('ตรงนี้ไม่มีหมึก — แตะให้ตรงตัวอักษรหรือเส้นที่ต้องการลบ', '', [], 2200); return; }
  const seen = new Uint8Array(N), stack = [best]; seen[best] = 1;
  let bx0 = W, by0 = H, bx1 = 0, by1 = 0, cnt = 0;
  while (stack.length) {
    const i = stack.pop(), px = i % W, py = (i / W) | 0;
    cnt++;
    if (px < bx0) bx0 = px; if (px > bx1) bx1 = px; if (py < by0) by0 = py; if (py > by1) by1 = py;
    for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
      const qx = px + dx, qy = py + dy; if (qx < 0 || qy < 0 || qx >= W || qy >= H) continue;
      const q = qy * W + qx; if (!seen[q] && dil[q]) { seen[q] = 1; stack.push(q); }
    }
  }
  sigEdPush();
  // erase everything visible inside the piece (soft edges of the strokes included)
  const er = SE.er, gx = g + 2;
  for (let yy = Math.max(0, by0 - gx); yy <= Math.min(H - 1, by1 + gx); yy++) for (let xx = Math.max(0, bx0 - gx); xx <= Math.min(W - 1, bx1 + gx); xx++) {
    const i = yy * W + xx;
    if (seen[i]) { er[i] = 255; continue; }
    // halo pixels next to the piece
    if (SE.alpha[i] > 0) for (let dy = -2; dy <= 2 && er[i] < 255; dy++) for (let dx = -2; dx <= 2; dx++) {
      const qx = xx + dx, qy = yy + dy; if (qx < 0 || qy < 0 || qx >= W || qy >= H) continue;
      if (seen[qy * W + qx]) { er[i] = 255; break; }
    }
  }
  sigEdRender(bx0 - gx, by0 - gx, bx1 + gx + 1, by1 + gx + 1);
  ink = null;
}
function sigEdEraseRect(x0, y0, x1, y1) {
  [x0, x1] = [Math.max(0, Math.min(x0, x1)) | 0, Math.min(SE.W, Math.ceil(Math.max(x0, x1)))];
  [y0, y1] = [Math.max(0, Math.min(y0, y1)) | 0, Math.min(SE.H, Math.ceil(Math.max(y0, y1)))];
  if (x1 - x0 < 2 || y1 - y0 < 2) return;
  sigEdPush();
  for (let y = y0; y < y1; y++) SE.er.fill(255, y * SE.W + x0, y * SE.W + x1);
  sigEdRender(x0, y0, x1, y1);
}
// remove isolated small specks (dust, paper texture); dots that belong to the signature are near other ink and stay
function sigEdDespeckle() {
  const W = SE.W, H = SE.H, N = W * H;
  const ink = new Uint8Array(N); for (let i = 0; i < N; i++) ink[i] = SE.alpha[i] * (255 - SE.er[i]) > 60 ? 1 : 0;
  const I = new Uint32Array((W + 1) * (H + 1));
  for (let y = 0; y < H; y++) { let s = 0; for (let x = 0; x < W; x++) { s += ink[y * W + x]; I[(y + 1) * (W + 1) + x + 1] = I[y * (W + 1) + x + 1] + s; } }
  const sum = (x0, y0, x1, y1) => I[y1 * (W + 1) + x1] - I[y0 * (W + 1) + x1] - I[y1 * (W + 1) + x0] + I[y0 * (W + 1) + x0];
  // small pieces are specks unless a big piece of writing is near them (a dot of the signature, a Thai tone mark …)
  const minA = Math.max(6, (Math.min(W, H) / 70) ** 2), M = Math.round(Math.max(6, Math.min(W, H) * 0.06));
  const lab = new Int32Array(N); let nl = 0, removed = 0;
  const comps = [];
  for (let s = 0; s < N; s++) {
    if (!ink[s] || lab[s]) continue;
    nl++; const st = [s]; let cnt = 0; lab[s] = nl;
    let x0 = W, y0 = H, x1 = 0, y1 = 0;
    while (st.length) {
      const i = st.pop(), px = i % W, py = (i / W) | 0; cnt++;
      if (px < x0) x0 = px; if (px > x1) x1 = px; if (py < y0) y0 = py; if (py > y1) y1 = py;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const qx = px + dx, qy = py + dy; if (qx < 0 || qy < 0 || qx >= W || qy >= H) continue;
        const q = qy * W + qx; if (ink[q] && !lab[q]) { lab[q] = nl; st.push(q); }
      }
    }
    comps.push({ l: nl, n: cnt, box: [x0, y0, x1, y1] });
  }
  const big = new Uint8Array(nl + 1);
  for (const c of comps) if (c.n > minA) big[c.l] = 1;
  I.fill(0);
  for (let y = 0; y < H; y++) { let s = 0; for (let x = 0; x < W; x++) { const l = lab[y * W + x]; s += l && big[l] ? 1 : 0; I[(y + 1) * (W + 1) + x + 1] = I[y * (W + 1) + x + 1] + s; } }
  const kill = [];
  for (const c of comps) {
    if (big[c.l]) continue;
    const [x0, y0, x1, y1] = c.box;
    if (sum(Math.max(0, x0 - M), Math.max(0, y0 - M), Math.min(W, x1 + M + 1), Math.min(H, y1 + M + 1)) > 0) continue;
    kill.push(c.box);
  }
  if (!kill.length) { toast('ไม่พบจุดสกปรกที่แยกอยู่โดด ๆ', '', [], 2200); return; }
  sigEdPush();
  for (const [x0, y0, x1, y1] of kill) {
    for (let y = Math.max(0, y0 - 2); y <= Math.min(H - 1, y1 + 2); y++) for (let x = Math.max(0, x0 - 2); x <= Math.min(W - 1, x1 + 2); x++) SE.er[y * W + x] = 255;
    removed++;
  }
  sigEdRender();
  toast(`ลบจุดสกปรก ${removed} จุด`, 'ok', [], 2000);
}

// ---- pointer handling on the stage
(() => {
  const st = $('sigEdStage');
  const pts = new Map();
  let act = null;   // { kind:'erase'|'pick'|'crop'|'pan'|'pinch', ... }
  const cur = $('sigEdCursor');
  const showCursor = (e) => {
    if (SE.step !== 'clean' || SE.tool !== 'erase' || e.pointerType === 'touch') { cur.classList.add('hidden'); return; }
    const r = st.getBoundingClientRect(), d = SE.size * 2;
    Object.assign(cur.style, { left: e.clientX - r.left - SE.size + 'px', top: e.clientY - r.top - SE.size + 'px', width: d + 'px', height: d + 'px' });
    cur.classList.remove('hidden');
  };
  st.addEventListener('pointerleave', () => cur.classList.add('hidden'));
  st.addEventListener('pointerdown', (e) => {
    if (!SE.open) return;
    e.preventDefault();
    try { st.setPointerCapture(e.pointerId); } catch {}
    pts.set(e.pointerId, e);
    if (pts.size === 2) { // two fingers: pinch-zoom / pan (cancel the current stroke)
      if (act && act.kind === 'erase' && act.pushed) { sigEdUndo(); SE.redo.pop(); sigEdUndoBtns(); }
      $('sigEdRect').classList.add('hidden');
      const [a, b] = [...pts.values()];
      act = { kind: 'pinch', d0: Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY), z0: SE.view.z, m0: [(a.clientX + b.clientX) / 2, (a.clientY + b.clientY) / 2], v0: { ...SE.view } };
      return;
    }
    if (pts.size > 2) return;
    const [x, y] = sigEdImgPt(e);
    if (SE.step === 'crop') {
      const h = e.target.closest('.sch'); const c = SE.crop;
      act = { kind: 'crop', h: h ? h.dataset.h : (x > c.x && x < c.x + c.w && y > c.y && y < c.y + c.h ? 'move' : 'new'), sx: x, sy: y, c0: { ...c } };
      return;
    }
    if (e.button === 1 || e.button === 2 || (e.pointerType === 'mouse' && e.shiftKey && false)) { act = { kind: 'pan', sx: e.clientX, sy: e.clientY, v0: { ...SE.view } }; return; }
    if (SE.tool === 'erase') { act = { kind: 'erase', last: [x, y], pushed: false, r: SE.size / SE.view.z }; sigEdPush(); act.pushed = true; sigEdEraseSeg([x, y], [x, y], act.r); }
    else act = { kind: 'pick', sx: x, sy: y, cx: e.clientX, cy: e.clientY, moved: false };
  });
  st.addEventListener('pointermove', (e) => {
    showCursor(e);
    if (!pts.has(e.pointerId)) return;
    pts.set(e.pointerId, e);
    if (!act) return;
    if (act.kind === 'pinch' && pts.size >= 2) {
      const [a, b] = [...pts.values()];
      const d = Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY), m = [(a.clientX + b.clientX) / 2, (a.clientY + b.clientY) / 2];
      const r = st.getBoundingClientRect(), v0 = act.v0;
      const z = Math.max(0.05, Math.min(16, act.z0 * d / act.d0));
      const px = (act.m0[0] - r.left - v0.tx) / v0.z, py = (act.m0[1] - r.top - v0.ty) / v0.z;
      SE.view = { z, tx: m[0] - r.left - px * z, ty: m[1] - r.top - py * z };
      sigEdApplyView();
      return;
    }
    if (act.kind === 'pan') { SE.view = { ...act.v0, tx: act.v0.tx + e.clientX - act.sx, ty: act.v0.ty + e.clientY - act.sy }; sigEdApplyView(); return; }
    const [x, y] = sigEdImgPt(e);
    if (act.kind === 'erase') { const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e]; for (const ev of evs.length ? evs : [e]) { const p = sigEdImgPt(ev); sigEdEraseSeg(act.last, p, act.r); act.last = p; } return; }
    if (act.kind === 'pick') {
      if (!act.moved && Math.hypot(e.clientX - act.cx, e.clientY - act.cy) < 6) return;
      act.moved = true; act.ex = x; act.ey = y;
      const v = SE.view, R = $('sigEdRect');
      Object.assign(R.style, { left: v.tx + Math.min(act.sx, x) * v.z + 'px', top: v.ty + Math.min(act.sy, y) * v.z + 'px', width: Math.abs(x - act.sx) * v.z + 'px', height: Math.abs(y - act.sy) * v.z + 'px' });
      R.classList.remove('hidden');
      return;
    }
    if (act.kind === 'crop') {
      const c = { ...act.c0 }, dx = x - act.sx, dy = y - act.sy, W = SE.src.width, H = SE.src.height, m = 8 / SE.view.z;
      if (act.h === 'move') { c.x = Math.max(0, Math.min(W - c.w, c.x + dx)); c.y = Math.max(0, Math.min(H - c.h, c.y + dy)); }
      else if (act.h === 'new') { c.x = Math.max(0, Math.min(act.sx, x)); c.y = Math.max(0, Math.min(act.sy, y)); c.w = Math.min(W, Math.max(act.sx, x)) - c.x; c.h = Math.min(H, Math.max(act.sy, y)) - c.y; }
      else {
        let x0 = c.x, y0 = c.y, x1 = c.x + c.w, y1 = c.y + c.h;
        if (act.h.includes('w')) x0 = Math.max(0, Math.min(x1 - m * 3, x0 + dx));
        if (act.h.includes('e')) x1 = Math.min(W, Math.max(x0 + m * 3, x1 + dx));
        if (act.h.includes('n')) y0 = Math.max(0, Math.min(y1 - m * 3, y0 + dy));
        if (act.h.includes('s')) y1 = Math.min(H, Math.max(y0 + m * 3, y1 + dy));
        Object.assign(c, { x: x0, y: y0, w: x1 - x0, h: y1 - y0 });
      }
      SE.crop = { x: Math.round(c.x), y: Math.round(c.y), w: Math.max(1, Math.round(c.w)), h: Math.max(1, Math.round(c.h)) };
      sigCropDraw();
    }
  });
  const up = (e) => {
    if (!pts.has(e.pointerId)) return;
    pts.delete(e.pointerId);
    if (!act) return;
    if (act.kind === 'pinch') { if (!pts.size) act = null; return; }
    if (act.kind === 'pick') {
      $('sigEdRect').classList.add('hidden');
      if (act.moved) sigEdEraseRect(act.sx, act.sy, act.ex, act.ey); else sigEdPick(act.sx, act.sy);
    }
    if (act.kind === 'crop' && act.h === 'new' && SE.crop.w < 10 && SE.crop.h < 10) { SE.crop = act.c0; sigCropDraw(); }
    act = null;
  };
  st.addEventListener('pointerup', up); st.addEventListener('pointercancel', up);
  st.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); });
  st.addEventListener('wheel', (e) => {
    if (!SE.open) return;
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) sigEdZoomAt(SE.view.z * (e.deltaY < 0 ? 1.15 : 1 / 1.15), e.clientX, e.clientY);
    else if (SE.step === 'clean') { SE.view.tx -= e.shiftKey ? e.deltaY : e.deltaX; SE.view.ty -= e.shiftKey ? 0 : e.deltaY; sigEdApplyView(); }
  }, { passive: false });
})();

// ---- wiring
$('sigEdClose').onclick = sigCloseEditor;
$('sigEdUndo').onclick = sigEdUndo; $('sigEdRedo').onclick = sigEdRedo;
$('sigEd').querySelectorAll('.sigsteps button').forEach((b) => b.onclick = () => sigEdStep(b.dataset.s));
$('sigEdRot').onclick = sigCropRotate;
$('sigEdAuto').onclick = () => { SE.crop = SE.keepAlpha ? sigAlphaBox(SE.src) : sigAutoCrop(SE.src); sigCropDraw(); };
$('sigEdFull').onclick = () => { SE.crop = { x: 0, y: 0, w: SE.src.width, h: SE.src.height }; sigCropDraw(); };
$('sigEdOk').onclick = async () => {
  if (SE.step === 'crop') { sigEdStep('clean'); return; }
  const out = sigEdResult();
  if (!out) { toast('ไม่เหลือส่วนที่เป็น' + sigKindName(SE.kind) + 'แล้ว — กดเลิกทำ หรือลดการลบ', 'err'); return; }
  const fn = SE.onDone; sigCloseEditor();
  busy('กำลังบันทึกลงคลัง …');
  try { await fn(out); } catch (e) { toast('บันทึกไม่สำเร็จ: ' + e.message, 'err'); }
  unbusy();
};
$('sigEdCleanBar').querySelectorAll('[data-t]').forEach((b) => b.onclick = () => sigEdTool(b.dataset.t));
$('sigEdSize').oninput = (e) => { if (SE.tool === 'erase') SE.size = +e.target.value; else SE.group = +e.target.value; };
$('sigEdThr').oninput = (e) => { SE.thr = +e.target.value; clearTimeout(SE.thrT); SE.thrT = setTimeout(() => { sigEdExtract(); sigEdRender(); }, 60); };
$('sigEdSpeck').onclick = sigEdDespeckle;
$('sigEdReset').onclick = () => { if (SE.er.some((v) => v)) { sigEdPush(); SE.er = new Uint8Array(SE.W * SE.H); sigEdRender(); } };
$('sigEdZoomIn').onclick = () => sigEdZoomAt(SE.view.z * 1.25);
$('sigEdZoomOut').onclick = () => sigEdZoomAt(SE.view.z / 1.25);
$('sigEdZoomFit').onclick = sigEdFit;
window.addEventListener('resize', () => { if (SE.open) sigEdFit(); });

// ---- panel / props wiring
$('annSig').onclick = () => ($('sigPanel').classList.contains('hidden') ? sigOpenPanel() : sigClosePanel());
$('sigPanelClose').onclick = sigClosePanel;
$('sigMenuBtn').onclick = (e) => { const r = e.currentTarget.getBoundingClientRect(); sigLibMenu(r.left, r.top - 8); if (ctxEl) ctxEl.style.top = Math.max(4, r.top - ctxEl.offsetHeight - 6) + 'px'; };
$('sigPanel').querySelectorAll('.sigtabs button').forEach((b) => b.onclick = () => { SIG.tab = b.dataset.k; sigRenderPanel(); });
$('sigColors').innerHTML = SIG_COLORS.map((c) => c ? `<button data-c="${c}" style="background:${c}" title="${c}"></button>` : '<button class="orig" title="สีเดิม"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M6.5 17.5l11-11"/></svg></button>').join('');
$('sigColors').querySelectorAll('button').forEach((b) => b.onclick = () => sigApplyProps({ color: b.dataset.c || null }));
$('sigWeight').oninput = (e) => sigApplyProps({ weight: Math.round(+e.target.value * 10) / 10 });
$('sigAll').onchange = () => { const a = findAnn(annSel); if ($('sigAll').checked && a && a.kind === 'img') sigApplyProps({ color: a.color, weight: a.weight || 0 }); };
$('sigPropsOk').onclick = sigCloseProps;

// keyboard while the full-screen editors are open
window.addEventListener('keydown', (e) => {
  if (!sigOverlayOpen() || !$('modal').classList.contains('hidden')) return;
  const c = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
  const drawOpen = !$('sigDraw').classList.contains('hidden');
  let handled = true;
  if (e.key === 'Escape') { drawOpen ? sigDrawClose() : sigCloseEditor(); }
  else if (c && k === 'z' && !e.shiftKey) drawOpen ? $('sigDrawUndo').click() : sigEdUndo();
  else if (c && (k === 'y' || (k === 'z' && e.shiftKey))) drawOpen ? $('sigDrawRedo').click() : sigEdRedo();
  else if (e.key === 'Enter' || (c && k === 's')) drawOpen ? (!$('sigDrawOk').disabled && $('sigDrawOk').click()) : $('sigEdOk').click();
  else if (!drawOpen && SE.step === 'clean' && !c && k === 'e') sigEdTool('erase');
  else if (!drawOpen && SE.step === 'clean' && !c && k === 'v') sigEdTool('pick');
  else if (!drawOpen && c && (e.key === '=' || e.key === '+')) sigEdZoomAt(SE.view.z * 1.25);
  else if (!drawOpen && c && e.key === '-') sigEdZoomAt(SE.view.z / 1.25);
  else if (!drawOpen && c && e.key === '0') sigEdFit();
  else if (c || e.key === 'Delete' || e.key === 'Backspace' || e.key.startsWith('Arrow')) { /* swallow document shortcuts */ }
  else handled = false;
  if (handled) { e.preventDefault(); e.stopImmediatePropagation(); }
}, true);


// ------------------------------------------------------------------ v1.9 scanned pages -> Word (Thai OCR, offline)
// worker: web/ocr/ocr-worker.js (onnxruntime-web, PP-OCRv5 detection + Thai recognition). Page structure rules and the
// .docx writer live here. Form-like pages (many ruled boxes) are inserted as pictures; words the model was unsure of
// are highlighted yellow so they are quick to check.
const OCRW = { worker: null, ready: null, seq: 0, wait: new Map() };
function ocrBase() { return new URL(PACKED ? './ocr/' : './ocr/', location.href).href; }
function ocrWorker() {
  if (OCRW.ready) return OCRW.ready;
  OCRW.ready = new Promise((resolve, reject) => {
    let w;
    try { w = new Worker(ocrBase() + 'ocr-worker.js', { type: 'module' }); } catch (e) { reject(new Error('เบราว์เซอร์นี้เปิดตัวอ่านข้อความไม่ได้: ' + e.message)); return; }
    OCRW.worker = w;
    w.onmessage = (e) => {
      const m = e.data;
      if (m.type === 'ready') resolve(w);
      else if (m.type === 'error' && m.id == null) { reject(new Error(m.message)); }
      else if (OCRW.wait.has(m.id)) { const p = OCRW.wait.get(m.id); OCRW.wait.delete(m.id); m.type === 'error' ? p.rej(new Error(m.message)) : p.res(m); }
    };
    w.onerror = (e) => reject(new Error(e.message || 'โหลดตัวอ่านข้อความไม่สำเร็จ'));
    w.postMessage({ type: 'init', base: ocrBase() });
  });
  OCRW.ready.catch(() => { OCRW.ready = null; if (OCRW.worker) OCRW.worker.terminate(); OCRW.worker = null; });
  return OCRW.ready;
}
async function ocrCanvas(canvas) {
  const w = await ocrWorker();
  const g = canvas.getContext('2d');
  const img = g.getImageData(0, 0, canvas.width, canvas.height);
  const id = ++OCRW.seq;
  return new Promise((res, rej) => {
    OCRW.wait.set(id, { res, rej });
    w.postMessage({ type: 'page', id, w: img.width, h: img.height, data: img.data.buffer }, [img.data.buffer]);
  });
}

// ---- text helpers (a "txt" is { t, low } where low marks uncertain characters with '1')
const OW_TH = (ch) => ch >= '\u0e00' && ch <= '\u0e7f';
const owT = (t, low) => ({ t, low: low || '0'.repeat(t.length) });
function owJoin(parts, sep = ' ') { // like ' '.join, keeps the uncertainty mask
  const out = { t: '', low: '' };
  for (const p of parts) {
    if (!p.t) continue;
    if (out.t) { out.t += sep; out.low += '0'.repeat(sep.length); }
    out.t += p.t; out.low += p.low;
  }
  return out;
}
function owJoinLines(parts) { // Thai wrapped lines join without a space
  const out = { t: '', low: '' };
  for (const p of parts) {
    if (!p.t) continue;
    if (out.t) { const sp = OW_TH(out.t[out.t.length - 1]) && OW_TH(p.t[0]) ? '' : ' '; out.t += sp; out.low += '0'.repeat(sp.length); }
    out.t += p.t; out.low += p.low;
  }
  return out;
}
// amounts: digits are reliable, separators are not ('1.77.23', '7,000,000,00', '51,732,008:24', '12;000.00')
function owFixNumbers(x) {
  const re = /(^|[^\d/])(\(?)(\d[\d.,:;]*[.,:;])(\d{2})(\)?)(?![\d/])/g;
  let out = '', low = '', last = 0, m;
  while ((m = re.exec(x.t))) {
    const body = m[3], groups = body.slice(0, -1).split(/[.,:;]/);
    const start = m.index + m[1].length, end = m.index + m[0].length;
    let fixed = null;
    if (!(groups.slice(1).some((g) => g.length > 3) || (groups[0].length > 3 && groups.length > 1))) {
      const ip = groups.join('');
      if (/^\d{1,13}$/.test(ip)) fixed = m[2] + Number(ip).toLocaleString('en-US') + '.' + m[4] + m[5];
    }
    const orig = x.t.slice(start, end);
    out += x.t.slice(last, start); low += x.low.slice(last, start);
    if (fixed && fixed !== orig) { out += fixed; low += '1'.repeat(fixed.length); }      // repaired -> ask the user to look
    else { out += orig; low += x.low.slice(start, end); }
    last = end;
  }
  out += x.t.slice(last); low += x.low.slice(last);
  return { t: out, low };
}

// ---- page structure (rows, amount columns, tables, paragraphs)
const OW_AMT = /^\(?-?\d{1,3}(?:,\d{3})*\.\d{2}\)?$|^\(?\d+\.\d{2}\)?$/;
const OW_DASH = /^[-–—]$/;
const owIsAmt = (t) => OW_AMT.test(t) || OW_DASH.test(t);
const OW_STAMP = /^[A-Z0-9 .,()\-:]{2,24}$/;
function owTokens(res) {
  const out = [];
  for (const l of res.lines) {
    let x = owFixNumbers(owT(l.t, l.low));
    // '1,989, 159.84' -> one amount (remove the space and its mask char)
    const rm = [];
    x.t.replace(/(?<=\d,)\s+(?=\d{3})|(?<=\d)\s+(?=[.,]\d{2}\b)/g, (s, i) => { rm.push([i, s.length]); return s; });
    for (let k = rm.length - 1; k >= 0; k--) { const [i, n] = rm[k]; x = { t: x.t.slice(0, i) + x.t.slice(i + n), low: x.low.slice(0, i) + x.low.slice(i + n) }; }
    const tt = x.t.trim(); if (!tt) continue;
    const lead = x.t.indexOf(tt); x = { t: tt, low: x.low.slice(lead, lead + tt.length) };
    if (l.c < 0.7 && !/\d[\d,]*\.\d{2}/.test(x.t)) continue;                       // stamp / signature scribbles
    if (OW_STAMP.test(x.t) && /[A-Z]{2}/.test(x.t) && l.c < 0.97) continue;
    const parts = x.t.split(' ');
    if (parts.length > 1 && parts.some(owIsAmt)) {
      const tot = parts.reduce((s, p) => s + p.length, 0) + parts.length - 1, w = l.b[2] - l.b[0];
      let px = l.b[0], pos = 0;
      for (const p of parts) {
        const pw = w * p.length / tot;
        out.push({ ...owT(p, x.low.slice(pos, pos + p.length)), b: [px, l.b[1], px + pw, l.b[3]] });
        px += pw + w / tot; pos += p.length + 1;
      }
    } else out.push({ ...x, b: l.b });
  }
  return out;
}
const owMed = (a) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[s.length >> 1] : 0; };
const owPct = (a, q) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[Math.min(s.length - 1, Math.max(0, Math.round(q * (s.length - 1))))] : 0; };
function owRows(items, mh) {
  items = [...items].sort((a, b) => (a.b[1] + a.b[3]) - (b.b[1] + b.b[3]));
  const rows = []; let cur = [];
  for (const it of items) {
    const cy = (it.b[1] + it.b[3]) / 2;
    if (cur.length && Math.abs(cy - cur.reduce((s, c) => s + (c.b[1] + c.b[3]) / 2, 0) / cur.length) > mh * 0.55) { rows.push(cur); cur = []; }
    cur.push(it);
  }
  if (cur.length) rows.push(cur);
  return rows.map((r) => r.sort((a, b) => a.b[0] - b.b[0]));
}
function owPageBlocks(res) {
  const W = res.w, H = res.h;
  const items = owTokens(res);
  if (!items.length) return [];
  const mh = owMed(items.map((it) => it.b[3] - it.b[1]));
  const rows = owRows(items, mh);
  const edges = rows.flat().filter((it) => owIsAmt(it.t)).map((it) => it.b[2]).sort((a, b) => a - b);
  let cl = [];
  for (const e of edges) { if (cl.length && e - cl[cl.length - 1][cl[cl.length - 1].length - 1] < W * 0.03) cl[cl.length - 1].push(e); else cl.push([e]); }
  const cols = cl.filter((c) => c.length >= 3).map(owMed);
  const colOf = (it) => {
    if (!cols.length) return null;
    let j = 0; cols.forEach((c, k) => { if (Math.abs(it.b[2] - c) < Math.abs(it.b[2] - cols[j])) j = k; });
    return Math.abs(it.b[2] - cols[j]) < W * 0.05 ? j : null;
  };
  const info = rows.map((r) => {
    const amts = r.filter((it) => owIsAmt(it.t)).map((it) => [colOf(it), it]).filter(([j]) => j != null);
    const text = r.filter((it) => !amts.some(([, a]) => a === it));
    return { r, amts, text, x0: Math.min(...r.map((it) => it.b[0])), x1: Math.max(...r.map((it) => it.b[2])), y0: Math.min(...r.map((it) => it.b[1])) };
  });
  const isAmtRow = info.map((x) => x.amts.length > 0);
  const inT = info.map(() => false);
  if (isAmtRow.filter(Boolean).length >= 3 && cols.length) {
    const labRight = Math.min(...cols) - W * 0.12;
    isAmtRow.forEach((v, i) => { if (v) inT[i] = true; });
    const nearAmtBelow = (i) => { for (let k = i + 1; k < Math.min(info.length, i + 4); k++) if (isAmtRow[k]) return true; return false; };
    for (let i = 0; i < info.length; i++) {          // header rows above the amounts
      if (inT[i] || !nearAmtBelow(i)) continue;
      const r = info[i], t = r.r.map((it) => it.t).join(' ');
      if (r.x0 > labRight || (/25\d\d/.test(t) && r.x1 - r.x0 < W * 0.7 && r.x0 > W * 0.3)) inT[i] = 'h';
    }
    for (let i = 0; i < info.length; i++) {          // section titles inside the table
      if (inT[i]) continue;
      let prev = i > 0 && inT[i - 1] === 'h';
      for (let k = Math.max(0, i - 3); k < i; k++) if (inT[k]) prev = true;
      if (prev && nearAmtBelow(i) && info[i].x1 - info[i].x0 < W * 0.55) inT[i] = true;
    }
  }
  const textRows = info.filter((_, i) => !inT[i]);
  const left = textRows.length ? owPct(textRows.map((x) => x.x0), 0.2) : 0;
  const right = textRows.length ? owPct(textRows.map((x) => x.x1), 0.9) : W;
  const full = Math.max(1, right - left);
  const blocks = []; let para = [];
  const flush = () => { if (para.length) { blocks.push({ type: 'p', x: owJoinLines(para) }); para = []; } };
  const mincol = cols.length ? Math.min(...cols) : W;
  let i = 0;
  while (i < info.length) {
    const x = info[i];
    if (inT[i]) {
      flush();
      const tb = { type: 'table', cols: cols.length, header: [], rows: [] };
      const tr = []; for (let k = i; k < info.length && inT[k]; k++) if (inT[k] === true) tr.push(info[k].x0);
      const labX0 = tr.length ? Math.min(...tr) : x.x0;
      while (i < info.length && inT[i]) {
        const y = info[i];
        if (inT[i] === 'h') {
          const cells = cols.map(() => owT(''));
          for (const it of y.r) for (const word of (it.t.includes('หมายเหตุ') ? it.t.split(' ') : [it.t])) {
            if (word === 'หมายเหตุ') continue;
            const cx = (it.b[0] + it.b[2]) / 2; let j = 0;
            cols.forEach((c, k) => { if (Math.abs(cx - c + W * 0.04) < Math.abs(cx - cols[j] + W * 0.04)) j = k; });
            cells[j] = owJoin([cells[j], word === it.t ? it : owT(word)]);
          }
          tb.header.push(cells);
        } else {
          const vals = cols.map(() => owT(''));
          for (const [j, it] of y.amts) vals[j] = it;
          const txt = [...y.text]; let note = '';
          const last = txt[txt.length - 1];
          if (last && /^\d{1,2}$/.test(last.t) && last.b[0] > W * 0.4) note = txt.pop().t;
          else if (last) {
            const m = last.t.match(/^(.*\D)\s+(\d{1,2})$/);
            if (m && last.b[2] > W * 0.45 && last.b[2] < mincol - W * 0.08) { txt[txt.length - 1] = { ...last, t: m[1], low: last.low.slice(0, m[1].length) }; note = m[2]; }
          }
          const lab = owJoin(txt.filter((it) => it.b[2] < mincol - W * 0.03 || !/^\d{1,2}$/.test(it.t)));
          const lvl = Math.max(0, Math.min(3, Math.round((y.x0 - labX0) / (W * 0.022))));
          const bold = !y.amts.length || /^รวม/.test(lab.t.replace(/\s/g, '')) || /^กำไร/.test(lab.t);
          tb.rows.push({ label: lab, note, vals, indent: lvl, bold });
        }
        i++;
      }
      tb.hasNote = tb.rows.some((r) => r.note);
      blocks.push(tb); continue;
    }
    const t = owJoin(x.r);
    const cx = (x.x0 + x.x1) / 2, wrel = (x.x1 - x.x0) / full;
    if (Math.abs(cx - W / 2) < W * 0.05 && wrel < 0.75 && x.x0 - left > W * 0.06) { flush(); blocks.push({ type: 'center', x: t, bold: x.y0 < H * 0.25 }); i++; continue; }
    if (x.x0 - left > W * 0.33) { flush(); blocks.push({ type: 'right', x: t }); i++; continue; }
    const ind = x.x0 - left > W * 0.025, rel = (x.x1 - left) / full;
    const nx = i + 1 < info.length && !inT[i + 1] ? info[i + 1] : null;
    const nxtInd = !!nx && nx.x0 - left > W * 0.025 && nx.x0 - left < W * 0.2;
    if (ind && para.length) flush();
    if (!ind && rel < 0.55 && t.t.length < 70 && (nxtInd || !nx || (i + 1 < info.length && inT[i + 1])) && !para.length) { blocks.push({ type: 'h', x: t }); i++; continue; }
    para.push(t);
    if (rel < 0.6) flush();
    i++;
  }
  flush();
  return blocks;
}
function owIsForm(res) {
  const nAmt = res.lines.reduce((s, l) => s + (l.t.match(/\d{1,3}(?:,\d{3})*\.\d{2}/g) || []).length, 0);
  const scale = 1654 / res.w;                       // thresholds were measured at 200 dpi on A4
  return res.vl * scale > 3000 || (res.lines.length > 100 && nAmt < 0.1 * res.lines.length);
}

// ---- minimal .docx writer (WordprocessingML + stored zip)
const OW_FONT = 'TH Sarabun New';
const owEsc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function owRuns(x, o = {}) {
  const sz = o.size || 32, out = [];
  const rpr = (low) => `<w:rPr><w:rFonts w:ascii="${OW_FONT}" w:hAnsi="${OW_FONT}" w:cs="${OW_FONT}"/>${o.bold ? '<w:b/><w:bCs/>' : ''}${low ? '<w:highlight w:val="yellow"/>' : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`;
  let k = 0;
  while (k < x.t.length) {
    // uncertain spans are widened to the whole Thai syllable cluster around them
    const lowAt = (j) => x.low[j] === '1';
    let j = k; const lw = lowAt(k);
    while (j < x.t.length && lowAt(j) === lw) j++;
    out.push(`<w:r>${rpr(lw)}<w:t xml:space="preserve">${owEsc(x.t.slice(k, j))}</w:t></w:r>`);
    k = j;
  }
  return out.join('') || `<w:r>${rpr(false)}<w:t></w:t></w:r>`;
}
function owPara(x, o = {}) {
  const ppr = [];
  if (o.keepNext) ppr.push('<w:keepNext/>');
  if (o.pageBreakBefore) ppr.push('<w:pageBreakBefore/>');
  ppr.push(`<w:spacing w:before="${o.before || 0}" w:after="${o.after == null ? 60 : o.after}"/>`);
  if (o.firstLine || o.left) ppr.push(`<w:ind${o.left ? ` w:left="${o.left}"` : ''}${o.firstLine ? ` w:firstLine="${o.firstLine}"` : ''}/>`);
  if (o.align) ppr.push(`<w:jc w:val="${o.align}"/>`);
  return `<w:p><w:pPr>${ppr.join('')}</w:pPr>${x ? owRuns(x, o) : ''}</w:p>`;
}
const OW_TW = 11906 - 2 * 1134;
function owTable(b) {
  const n = b.cols, note = b.hasNote;
  const amtW = Math.min(2200, Math.round(OW_TW * 0.2)), noteW = note ? 1250 : 0, labW = OW_TW - amtW * n - noteW;
  const widths = [labW, ...(note ? [noteW] : []), ...Array(n).fill(amtW)];
  const bd = (top, bottom) => `<w:tcBorders><w:top w:val="${top || 'nil'}" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="nil"/><w:bottom w:val="${bottom || 'nil'}" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="nil"/></w:tcBorders>`;
  const cell = (x, w, o = {}) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${bd(o.top, o.bottom)}</w:tcPr>${owPara(x, { align: o.align, left: o.indent, bold: o.bold, size: 30, after: 0 })}</w:tc>`;
  const rows = [];
  b.header.forEach((h, k) => {
    const last = k === b.header.length - 1, bot = last ? 'single' : null;
    rows.push(`<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(null, labW, { bottom: bot })}${note ? cell(owT(last ? 'หมายเหตุ' : ''), noteW, { align: 'center', bottom: bot, bold: true }) : ''}${h.map((c) => cell(c, amtW, { align: 'right', bold: true, bottom: bot })).join('')}</w:tr>`);
  });
  b.rows.forEach((r, k) => {
    const total = r.bold && r.vals.some((v) => v.t), lastRow = k === b.rows.length - 1;
    rows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(r.label, labW, { bold: r.bold, indent: r.indent * 360 })}${note ? cell(owT(r.note || ''), noteW, { align: 'center' }) : ''}${r.vals.map((v) => cell(v, amtW, { align: 'right', top: total ? 'single' : null, bottom: total && lastRow ? 'double' : null })).join('')}</w:tr>`);
  });
  return `<w:tbl><w:tblPr><w:tblW w:w="${OW_TW}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
}
function owImage(rid, idx, wpx, hpx) { // inline picture, fitted into the text area
  const maxW = OW_TW * 635, maxH = (16838 - 2 * 1134) * 635 * 0.97;
  const k = Math.min(maxW / wpx, maxH / hpx), cx = Math.round(wpx * k), cy = Math.round(hpx * k);
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${idx}" name="หน้า ${idx}" descr="หน้าแบบฟอร์มจากต้นฉบับ (แทรกเป็นภาพ)"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${idx}" name="page${idx}.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}
function owDocx(pages) { // pages: [{ kind:'form', jpg:Uint8Array, w, h } | { kind:'text', blocks }]
  const body = [], media = [];
  pages.forEach((pg, pi) => {
    const pb = pi > 0;
    if (pg.kind === 'form') {
      media.push(pg.jpg);
      const rid = 'rIdImg' + media.length;
      body.push(pb ? owPara(null, { pageBreakBefore: true, after: 0 }) : '');
      body.push(owImage(rid, media.length, pg.w, pg.h));
      return;
    }
    let first = true;
    const P = (x, o) => { const s = owPara(x, { ...o, pageBreakBefore: pb && first }); first = false; return s; };
    if (!pg.blocks.length) body.push(P(owT(''), {}));
    for (const b of pg.blocks) {
      if (b.type === 'center') body.push(P(b.x, { align: 'center', bold: b.bold }));
      else if (b.type === 'right') body.push(P(b.x, { left: Math.round(OW_TW * 0.5) }));
      else if (b.type === 'h') body.push(P(b.x, { bold: true, before: 120, after: 40, keepNext: true }));
      else if (b.type === 'p') body.push(P(b.x, { align: 'thaiDistribute', firstLine: 720, after: 80 }));
      else if (b.type === 'table') {
        if (first) body.push(P(owT(''), { after: 0 }));
        body.push(owTable(b)); body.push(owPara(owT(''), { after: 0 }));
      }
    }
  });
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${OW_FONT}" w:hAnsi="${OW_FONT}" w:cs="${OW_FONT}" w:eastAsia="${OW_FONT}"/><w:sz w:val="32"/><w:szCs w:val="32"/><w:lang w:val="en-US" w:bidi="th-TH"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${media.map((_, k) => `<Relationship Id="rIdImg${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page${k + 1}.jpg"/>`).join('')}</Relationships>`;
  const files = [
    ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', doc], ['word/styles.xml', styles], ['word/_rels/document.xml.rels', rels],
    ...media.map((m, k) => [`word/media/page${k + 1}.jpg`, m]),
  ];
  return owZip(files);
}
const OW_CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function owCrc(u8) { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = OW_CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function owZip(files) { // deflate (pako) for xml, store for jpg
  const enc = new TextEncoder(), parts = [], central = [];
  let off = 0;
  for (const [name, data] of files) {
    const raw = typeof data === 'string' ? enc.encode(data) : data;
    const nm = enc.encode(name), crc = owCrc(raw);
    const comp = typeof data === 'string' ? pako.deflateRaw(raw) : raw, method = typeof data === 'string' ? 8 : 0;
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, method, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, comp.length, true); lh.setUint32(22, raw.length, true); lh.setUint16(26, nm.length, true);
    parts.push(new Uint8Array(lh.buffer), nm, comp);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, method, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, comp.length, true); ch.setUint32(24, raw.length, true); ch.setUint16(28, nm.length, true); ch.setUint32(42, off, true);
    central.push(new Uint8Array(ch.buffer), nm);
    off += 30 + nm.length + comp.length;
  }
  const csize = central.reduce((s, p) => s + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, csize, true); end.setUint32(16, off, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((s, p) => s + p.length, 0)); let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

// ---- the command
async function convertScanWord() {
  const list = convList(); if (!list.length) return;
  const r = await dialog({
    title: 'แปลงเอกสารสแกนเป็น Word (อ่านภาษาไทย)',
    body: `<p>${list.length} หน้า → ไฟล์ Word 1 ไฟล์ · ทำงานในเครื่องนี้ทั้งหมด ไม่ส่งเอกสารออกไปภายนอก</p>
      <ul class="hint"><li>ย่อหน้า หัวข้อ และตารางตัวเลข (เช่น งบการเงิน) จะแก้ไขได้ใน Word</li>
      <li>หน้าที่เป็นแบบฟอร์มช่องกรอกจำนวนมาก (เช่น ภ.ง.ด.) จะแทรกเป็นภาพหน้าเดิม</li>
      <li>คำที่ระบบอ่านไม่มั่นใจจะ<b style="background:#ff0">ไฮไลต์สีเหลือง</b> — ควรตรวจทานก่อนใช้</li></ul>
      <p class="hint">ใช้เวลาประมาณหน้าละ 10–30 วินาที${IS_TOUCH ? ' (บนมือถืออาจนานกว่านี้)' : ''} · ครั้งแรกต้องโหลดตัวอ่านข้อความประมาณ 16 MB</p>
      <label class="radio"><input type="checkbox" name="noform"> แปลงทุกหน้าเป็นข้อความ (ไม่แทรกหน้าแบบฟอร์มเป็นภาพ)</label>`,
    buttons: [['cancel', 'ยกเลิก'], ['ok', 'เริ่มแปลง', true]],
  });
  if (!r.btn) return;
  const noForm = !!r.values.noform;
  let cancelled = false;
  const t0 = Date.now();
  const show = (k, msg) => {
    busy(msg);
    let c = $('busy').querySelector('.ocrcancel');
    if (!c) { c = document.createElement('button'); c.className = 'tbtn ocrcancel'; c.textContent = 'ยกเลิก'; c.onclick = () => { cancelled = true; c.disabled = true; c.textContent = 'กำลังยกเลิก…'; }; $('busy').querySelector('.busybox').appendChild(c); }
  };
  const done = () => { const c = $('busy').querySelector('.ocrcancel'); if (c) c.remove(); unbusy(); };
  try {
    show(0, 'กำลังเตรียมตัวอ่านข้อความภาษาไทย … (ครั้งแรกอาจใช้เวลาสักครู่)');
    await ocrWorker();
    const out = []; let low = 0, words = 0;
    for (let k = 0; k < list.length; k++) {
      if (cancelled) throw new Error('ยกเลิกแล้ว');
      const el = Math.round((Date.now() - t0) / 1000), eta = k ? Math.round((Date.now() - t0) / 1000 / k * (list.length - k)) : null;
      show(k, `กำลังอ่านหน้า ${k + 1}/${list.length} …${eta != null ? ` (เหลือประมาณ ${eta >= 60 ? Math.round(eta / 60) + ' นาที' : eta + ' วินาที'})` : ''}`);
      const c = await renderPageCanvas(list[k], 200 / 72);
      const jpg = await canvasToBytes(c, 'image/jpeg', 0.85);
      const res = await ocrCanvas(c);
      if (!noForm && owIsForm(res)) out.push({ kind: 'form', jpg, w: c.width, h: c.height });
      else {
        const blocks = owPageBlocks(res);
        for (const l of res.lines) { words++; if (l.low.includes('1')) low++; }
        out.push({ kind: 'text', blocks });
      }
      c.width = c.height = 0;
    }
    show(0, 'กำลังเขียนไฟล์ Word …');
    const bytes = owDocx(out);
    done();
    const forms = out.filter((p) => p.kind === 'form').length;
    const saved = await saveBytes(bytes, D.name + ' (OCR).docx', 'docx');
    if (!saved) {
      // v1.9.1: never claim success when the file was not written
      toast(lastSave === 'cancel'
        ? `อ่านข้อความครบ ${list.length} หน้าแล้ว แต่ยังไม่ได้บันทึกไฟล์ Word เพราะยกเลิกการบันทึก — ต้องสั่งแปลงใหม่เพื่อสร้างไฟล์`
        : `อ่านข้อความครบ ${list.length} หน้าแล้ว แต่บันทึกไฟล์ Word ไม่สำเร็จ — ไม่มีไฟล์ใหม่ ลองสั่งแปลงอีกครั้งหรือเลือกตำแหน่งอื่น`, lastSave === 'cancel' ? '' : 'err', [], 12000);
      return;
    }
    toast(`แปลงเสร็จ ${list.length} หน้า${forms ? ` (แทรกเป็นภาพ ${forms} หน้า)` : ''} · มีบรรทัดที่ควรตรวจทาน ${low}/${words} บรรทัด (ไฮไลต์สีเหลือง)`, 'ok', [], 8000);
  } catch (e) { done(); if (!cancelled) toast('แปลงไม่สำเร็จ: ' + e.message, 'err', [], 9000); else toast('ยกเลิกการแปลงแล้ว'); }
}


// expose for automated testing
window.__superpdf = { isHeic, heicToJpegBlob, loadImageEl, fileToPages, compressDialog, compressPdfBytes, cmpImages, COMPRESS_PRESETS, APP_VERSION, splitBySize, buildPdfInner, bundledFontsReady, BUNDLED_FONTS, IE, openImageEditor, closeImageEditor, iePickObject, ieDelete, ieApply, despeckleCanvas, despecklePages, ieSetSel, ieLift, ieCommit, ieOpenPaint, scan, openScanner, applyScanFilter, scanPagesToPdf, printDialog, enhancePages, get annDefaults() { return annDefaults; }, setAnnTool, convertTiff, pasteAny, textToPdf, get D() { return D; }, docs, get pages() { return D.pages; }, openAsTabs, insertFiles, buildPdf, setMode, switchTo, get clip() { return clip; }, SIG, SE, SD, sigStore, sigOpenPanel, sigPlace, sigAddToLibrary, sigOpenEditor, sigOpenDraw, sigEdStep, sigEdPick, sigEdEraseRect, sigEdDespeckle, sigEdResult, sigExport, sigImportFile, sigApplyProps, sigOpenProps, sigPointMenu, sigFinal, convertScanWord, ocrCanvas, owPageBlocks, owDocx, owIsForm, owFixNumbers, renderPageCanvas, get annSel() { return annSel; }, selectAnn, findAnn, undo, redo, gcSources, sources, thumbCache, canvasToBytes, saveBytes, loadImageEl, closeTab, doCopy, doPaste, get lastSave() { return lastSave; }, get host() { return host; }, set host(v) { host = v; } };

if (IS_TOUCH) {
  document.querySelector('#dropcard h1').textContent = 'แตะเพื่อเลือกไฟล์';
  $('dropcard').addEventListener('click', (e) => { if (!e.target.closest('button')) doOpen(); });
}
D = newDoc();
refresh();
initHost();
