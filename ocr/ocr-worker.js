// Super PDF v1.9 — Thai OCR worker (PP-OCRv5 mobile detection + Thai recognition, onnxruntime-web, single thread)
// in : {type:'init', base}            -> {type:'ready'} | {type:'error'}
//      {type:'page', id, w, h, data}  (RGBA Uint8ClampedArray, transferred)
// out: {type:'page', id, w, h, skew, vl, lines:[{b:[x0,y0,x1,y1], t, c, low}]}   low = '0'/'1' per char (uncertain)
import * as ort from './ort.wasm.min.mjs';

let det = null, rec = null, chars = null;
const post = (m, tr) => self.postMessage(m, tr || []);

async function init(base) {
  ort.env.wasm.numThreads = 1;
  ort.env.wasm.proxy = false;
  ort.env.wasm.wasmPaths = base;
  // the wasm binary ships gzip-compressed (3.7 MB instead of 14 MB); inflate it here
  const gz = new Uint8Array(await (await fetch(base + 'ort-wasm-simd-threaded.wasm.gz')).arrayBuffer());
  ort.env.wasm.wasmBinary = await gunzip(gz);
  const opt = { executionProviders: ['wasm'], graphOptimizationLevel: 'all' };
  det = await ort.InferenceSession.create(await bytes(base + 'det.onnx'), opt);
  rec = await ort.InferenceSession.create(await bytes(base + 'rec_th.onnx'), opt);
  const dict = (await (await fetch(base + 'dict_th.txt')).text()).split('\n').map((s) => s.replace(/\r$/, '')).filter((s) => s !== '');
  chars = ['', ...dict, ' '];
  await loadDict(base);
}
const bytes = async (u) => { const r = await fetch(u); if (!r.ok) throw new Error('โหลด ' + u.split('/').pop() + ' ไม่ได้ (' + r.status + ')'); return new Uint8Array(await r.arrayBuffer()); };
async function gunzip(u8) {
  if (typeof DecompressionStream !== 'undefined') {
    const ds = new Blob([u8]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Uint8Array(await new Response(ds).arrayBuffer());
  }
  throw new Error('เบราว์เซอร์นี้ไม่รองรับการคลายไฟล์ (DecompressionStream)');
}


// ---------------- Thai dictionary check (PyThaiNLP words_th, CC0): highlight non-words, fix one well-known confusion
let DICT = null, MAXL = 0;
const COMB = new Set([...'ัิีึืฺุู็่้๊๋์ํ']);
const CONF = ['ซชข', 'ดต', 'คศ', 'ถภ', 'บป', 'พฟ', 'ผฝ', 'ิีึื', 'ุู', '่้๊๋', 'ำา', 'ฮอ', 'ฆม', 'ฎฏ', 'ญฌ'];
const SUB = new Map();
for (const g of CONF) for (const c of g) SUB.set(c, [...(SUB.get(c) || []), ...[...g].filter((x) => x !== c)]);
const DEL = new Set([...'่้๊๋็์ิีึื']), INS = ['่', '้'];
async function loadDict(base) {
  try {
    const gz = new Uint8Array(await (await fetch(base + 'words_th.txt.gz')).arrayBuffer());
    const words = new TextDecoder().decode(await gunzip(gz)).split('\n').map((w) => w.trim()).filter(Boolean);
    DICT = new Set(words); MAXL = Math.max(...words.map((w) => w.length));
  } catch { DICT = null; }
}
function segWords(s) { // DP: fewest characters outside dictionary words, then fewest words
  const n = s.length, INF = 1e9, best = new Array(n + 1).fill(null).map(() => [INF, 0]), back = new Array(n + 1);
  best[0] = [0, 0];
  const better = (a, b) => a[0] < b[0] || (a[0] === b[0] && a[1] < b[1]);
  for (let i = 0; i < n; i++) {
    if (best[i][0] === INF) continue;
    if (!COMB.has(s[i])) for (let j = i + 1; j <= Math.min(n, i + MAXL); j++) {
      if ((j === n || !COMB.has(s[j])) && DICT.has(s.slice(i, j))) { const c = [best[i][0], best[i][1] + 1]; if (better(c, best[j])) { best[j] = c; back[j] = [i, 0]; } }
    }
    let j = i + 1; while (j < n && COMB.has(s[j])) j++;
    const c = [best[i][0] + (j - i), best[i][1] + 1];
    if (better(c, best[j])) { best[j] = c; back[j] = [i, 1]; }
  }
  const mask = new Array(n).fill(0), words = []; let j = n;
  while (j > 0) { const [i, u] = back[j]; if (u) for (let k = i; k < j; k++) mask[k] = 1; else words.push([i, j]); j = i; }
  return { cost: best[n][0], mask, words };
}
function candidates(s, a, b) {
  const out = new Set(), lo = Math.max(0, a - 1), hi = Math.min(s.length, b + 1);
  for (let k = lo; k < hi; k++) {
    for (const r of SUB.get(s[k]) || []) out.add(s.slice(0, k) + r + s.slice(k + 1));
    if (DEL.has(s[k])) out.add(s.slice(0, k) + s.slice(k + 1));
    if (!COMB.has(s[k]) || 'ิีึืุูั'.includes(s[k])) for (const t of INS) if (k + 1 === s.length || !'่้๊๋'.includes(s[k + 1])) out.add(s.slice(0, k + 1) + t + s.slice(k + 1));
  }
  return out;
}
function checkRunOnce(s, prev, nxt) {
  const { cost, mask } = segWords(s);
  if (!cost) return { t: s, mask };
  if (s.length <= 2 || prev.endsWith('.') || nxt.startsWith('.')) return { t: s, mask };   // abbreviations: ม.ค. พ.ศ.
  const spans = []; for (let i = 0; i < s.length;) { if (mask[i]) { let j = i; while (j < s.length && mask[j]) j++; spans.push([i, j]); i = j; } else i++; }
  for (const [a, b] of spans) {
    if (b - a > 6) continue;
    const scored = [];
    for (const c of candidates(s, a, b)) {
      const r = segWords(c);
      if (r.cost > cost - (b - a)) continue;
      const d = c.length - s.length, lo = a, hi = Math.min(b + Math.max(0, d), c.length);
      const cover = r.words.filter(([i, j]) => i <= lo && j >= hi && j > lo).map(([i, j]) => j - i);
      if (!cover.length || Math.max(...cover) < 4) continue;
      scored.push([Math.max(...cover), c]);
    }
    if (!scored.length) continue;
    scored.sort((x, y) => y[0] - x[0] || (x[1] < y[1] ? -1 : 1));
    if (scored.length > 1 && scored[1][0] === scored[0][0] && scored[1][1] !== scored[0][1]) continue;   // ambiguous
    const t = scored[0][1], d = t.length - s.length;
    return { t, changed: true, a: Math.max(0, a - 1), b: b + 1, d };
  }
  return { t: s, mask };
}
function checkRun(s, prev, nxt) { // apply single fixes repeatedly (several errors in one run); remember edited spots
  let cur = s, touched = new Array(s.length).fill(0), changed = false;
  for (let it = 0; it < 6; it++) {
    const r = checkRunOnce(cur, prev, nxt);
    if (!r.changed) break;
    changed = true;
    const mid = new Array(Math.max(0, r.b + r.d - r.a)).fill(1);
    touched = [...touched.slice(0, r.a), ...mid, ...touched.slice(r.b)];
    cur = r.t;
  }
  const m = segWords(cur).mask;
  return { t: cur, mask: m.map((v, k) => (v || touched[k] ? 1 : 0)), changed };
}
function spell(t, low) { // -> { t, low } ; low: '1' = please check (not a dictionary word / corrected / very uncertain)
  if (!DICT) return { t, low };
  const parts = t.split(/([\u0e00-\u0e7f]+)/);
  let out = '', fl = '', pos = 0;
  parts.forEach((part, k) => {
    if (!part) return;
    const ol = low.slice(pos, pos + part.length); pos += part.length;
    if (/^[\u0e00-\u0e7f]/.test(part)) {
      const pv = parts[k - 1] || '', nx = parts[k + 1] || '';
      if (part.length <= 3 && (pv.endsWith('.') || nx.startsWith('.'))) { out += part; fl += ol; return; }   // abbreviations: พ.ศ. ม.ค. ภ.ง.ด.
      const r = checkRun(part, pv, nx);
      out += r.t;
      fl += r.changed ? r.mask.join('') : r.mask.map((v, i) => (v || ol[i] === '1' ? '1' : '0')).join('');
    } else { out += part; fl += ol; }
  });
  return { t: out, low: fl };
}

// ---------------- image helpers (RGBA in, planar BGR float out — the Paddle models were trained on BGR)
function resizeRGBA(src, sw, sh, dw, dh) { // bilinear
  const out = new Uint8ClampedArray(dw * dh * 4), fx = sw / dw, fy = sh / dh;
  for (let y = 0; y < dh; y++) {
    const sy = Math.min(sh - 1, Math.max(0, (y + 0.5) * fy - 0.5)), y0 = Math.floor(sy), y1 = Math.min(sh - 1, y0 + 1), wy = sy - y0;
    for (let x = 0; x < dw; x++) {
      const sx = Math.min(sw - 1, Math.max(0, (x + 0.5) * fx - 0.5)), x0 = Math.floor(sx), x1 = Math.min(sw - 1, x0 + 1), wx = sx - x0;
      const a = (y0 * sw + x0) * 4, b = (y0 * sw + x1) * 4, c = (y1 * sw + x0) * 4, d = (y1 * sw + x1) * 4, o = (y * dw + x) * 4;
      for (let k = 0; k < 3; k++) out[o + k] = (src[a + k] * (1 - wx) + src[b + k] * wx) * (1 - wy) + (src[c + k] * (1 - wx) + src[d + k] * wx) * wy;
      out[o + 3] = 255;
    }
  }
  return out;
}
function cropRGBA(src, sw, sh, x0, y0, w, h) {
  const out = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    const sy = Math.min(sh - 1, Math.max(0, y0 + y));
    for (let x = 0; x < w; x++) {
      const sx = Math.min(sw - 1, Math.max(0, x0 + x)), s = (sy * sw + sx) * 4, o = (y * w + x) * 4;
      out[o] = src[s]; out[o + 1] = src[s + 1]; out[o + 2] = src[s + 2]; out[o + 3] = 255;
    }
  }
  return out;
}
function rotateRGBA(src, w, h, deg) { // rotate about the centre, same size, edge pixels replicated (cv2 BORDER_REPLICATE)
  const t = -deg * Math.PI / 180, c = Math.cos(t), s = Math.sin(t), cx = w / 2, cy = h / 2;
  const out = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const dx = x - cx, dy = y - cy;
    let sx = c * dx - s * dy + cx, sy = s * dx + c * dy + cy;
    sx = Math.min(w - 1.001, Math.max(0, sx)); sy = Math.min(h - 1.001, Math.max(0, sy));
    const x0 = sx | 0, y0 = sy | 0, wx = sx - x0, wy = sy - y0;
    const a = (y0 * w + x0) * 4, b = a + 4, cc = a + w * 4, d = cc + 4, o = (y * w + x) * 4;
    for (let k = 0; k < 3; k++) out[o + k] = (src[a + k] * (1 - wx) + src[b + k] * wx) * (1 - wy) + (src[cc + k] * (1 - wx) + src[d + k] * wx) * wy;
    out[o + 3] = 255;
  }
  return out;
}
function toTensorBGR(rgba, w, h, mean, std, out, off) {
  const N = w * h;
  for (let i = 0; i < N; i++) {
    const r = rgba[i * 4] / 255, g = rgba[i * 4 + 1] / 255, b = rgba[i * 4 + 2] / 255;
    out[off + i] = (b - mean[0]) / std[0]; out[off + N + i] = (g - mean[1]) / std[1]; out[off + 2 * N + i] = (r - mean[2]) / std[2];
  }
}

// ---------------- text-line detection (DB post-process on axis-aligned components)
async function detect(rgba, W, H) {
  const r = Math.min(1, 1600 / Math.max(W, H));
  const w = Math.max(32, Math.round(W * r / 32) * 32), h = Math.max(32, Math.round(H * r / 32) * 32);
  const small = resizeRGBA(rgba, W, H, w, h);
  const x = new Float32Array(3 * w * h);
  toTensorBGR(small, w, h, [0.485, 0.456, 0.406], [0.229, 0.224, 0.225], x, 0);
  const res = await det.run({ [det.inputNames[0]]: new ort.Tensor('float32', x, [1, 3, h, w]) });
  const p = res[det.outputNames[0]].data;
  const sx = W / w, sy = H / h;
  // connected components of p > 0.3 (8-neighbour)
  const lab = new Int32Array(w * h), boxes = [];
  const stack = new Int32Array(w * h);
  let n = 0;
  for (let s = 0; s < w * h; s++) {
    if (p[s] <= 0.3 || lab[s]) continue;
    n++; let sp = 0; stack[sp++] = s; lab[s] = n;
    let x0 = w, y0 = h, x1 = 0, y1 = 0, cnt = 0, sum = 0, mx = 0, my = 0, mxx = 0, myy = 0, mxy = 0;
    while (sp) {
      const i = stack[--sp], px = i % w, py = (i / w) | 0;
      cnt++; sum += p[i]; mx += px; my += py; mxx += px * px; myy += py * py; mxy += px * py;
      if (px < x0) x0 = px; if (px > x1) x1 = px; if (py < y0) y0 = py; if (py > y1) y1 = py;
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const qx = px + dx, qy = py + dy; if (qx < 0 || qy < 0 || qx >= w || qy >= h) continue;
        const q = qy * w + qx; if (!lab[q] && p[q] > 0.3) { lab[q] = n; stack[sp++] = q; }
      }
    }
    const bw = x1 - x0 + 1, bh = y1 - y0 + 1;
    if (Math.min(bw, bh) < 3 || cnt < 4) continue;
    if (sum / cnt < 0.55) continue;
    // orientation of the component (for deskew)
    const cxm = mx / cnt, cym = my / cnt, vxx = mxx / cnt - cxm * cxm, vyy = myy / cnt - cym * cym, vxy = mxy / cnt - cxm * cym;
    const ang = 0.5 * Math.atan2(2 * vxy, vxx - vyy) * 180 / Math.PI;
    const d = (bw * bh) * 1.5 / (2 * (bw + bh));            // unclip distance (area * ratio / perimeter)
    const X0 = Math.max(0, x0 - d), Y0 = Math.max(0, y0 - d), X1 = Math.min(w, x1 + 1 + d), Y1 = Math.min(h, y1 + 1 + d);
    if (Math.min(X1 - X0, Y1 - Y0) < 5) continue;
    boxes.push({ b: [X0 * sx, Y0 * sy, X1 * sx, Y1 * sy], ang, len: bw * sx, hh: bh * sy, pr: [X0, Y0, X1, Y1] });
  }
  return { boxes: splitTall(boxes, p, w, h, sx, sy) };
}
function splitTall(boxes, p, w, h, sx, sy) { // two touching Thai lines detected as one tall box -> split at the weakest row
  if (boxes.length < 3) return boxes;
  const hs = boxes.map((b) => b.b[3] - b.b[1]).sort((a, b) => a - b), mh = hs[hs.length >> 1];
  const out = [];
  for (const b of boxes) {
    const bh = b.b[3] - b.b[1];
    if (bh < 1.55 * mh) { out.push(b); continue; }
    const [X0, Y0, X1, Y1] = b.pr.map(Math.round);
    const prof = [];
    for (let y = Y0; y < Y1; y++) { let s = 0; for (let x = X0; x < X1; x++) s += p[y * w + x]; prof.push(s / Math.max(1, X1 - X0)); }
    const n = Math.round(bh / mh), cuts = [];
    for (let k = 1; k < n; k++) {
      const c = Math.floor(prof.length * k / n), r = Math.max(2, Math.floor(prof.length / (3 * n)));
      let best = c, bv = 1e9;
      for (let i = Math.max(0, c - r); i < Math.min(prof.length, c + r); i++) if (prof[i] < bv) { bv = prof[i]; best = i; }
      cuts.push(best);
    }
    const ys = [0, ...cuts, prof.length];
    for (let k = 0; k + 1 < ys.length; k++) {
      if (ys[k + 1] - ys[k] < 4) continue;
      out.push({ ...b, b: [b.b[0], (Y0 + ys[k]) * sy, b.b[2], (Y0 + ys[k + 1]) * sy] });
    }
  }
  return out;
}
function skewAngle(boxes) { // length-weighted median of long boxes' orientation
  const a = boxes.filter((b) => b.len > 4 * b.hh && b.len > 80).map((b) => [b.ang, b.len]).sort((x, y) => x[0] - y[0]);
  if (a.length < 3) return 0;
  const tot = a.reduce((s, x) => s + x[1], 0); let c = 0;
  for (const [ang, l] of a) { c += l; if (c >= tot / 2) return ang; }
  return 0;
}

// ---------------- recognition (CTC)
async function recognize(rgba, W, H, boxes) {
  const crops = boxes.map((bx) => {
    let [x0, y0, x1, y1] = bx.b; const hh = y1 - y0;
    y0 -= hh * 0.04; y1 += hh * 0.03;                         // keep Thai upper / lower marks
    x0 = Math.max(0, Math.round(x0)); y0 = Math.max(0, Math.round(y0)); x1 = Math.min(W, Math.round(x1)); y1 = Math.min(H, Math.round(y1));
    const w = Math.max(1, x1 - x0), h = Math.max(1, y1 - y0);
    return { data: cropRGBA(rgba, W, H, x0, y0, w, h), w, h };
  });
  const order = crops.map((_, i) => i).sort((a, b) => crops[a].w / crops[a].h - crops[b].w / crops[b].h);
  const res = new Array(crops.length);
  const HH = 48;
  for (let s = 0; s < order.length; s += 8) {
    const ids = order.slice(s, s + 8);
    const maxr = Math.max(...ids.map((i) => crops[i].w / crops[i].h));
    const WW = Math.ceil(HH * Math.max(maxr, 320 / 48));
    const x = new Float32Array(ids.length * 3 * HH * WW);   // zero = mid-grey after normalisation (same padding as Paddle)
    ids.forEach((i, k) => {
      const c = crops[i], w = Math.min(WW, Math.ceil(HH * c.w / c.h));
      const r = resizeRGBA(c.data, c.w, c.h, w, HH);
      const plane = HH * WW, off = k * 3 * plane;
      for (let yy = 0; yy < HH; yy++) for (let xx = 0; xx < w; xx++) {
        const j = (yy * w + xx) * 4, o = yy * WW + xx;
        x[off + o] = (r[j + 2] / 255 - 0.5) / 0.5; x[off + plane + o] = (r[j + 1] / 255 - 0.5) / 0.5; x[off + 2 * plane + o] = (r[j] / 255 - 0.5) / 0.5;
      }
    });
    const out = await rec.run({ [rec.inputNames[0]]: new ort.Tensor('float32', x, [ids.length, 3, HH, WW]) });
    const t = out[rec.outputNames[0]], [B, T, C] = t.dims, d = t.data;
    ids.forEach((i, k) => {
      let prev = 0, txt = '', low = '', ps = [], pp = [];
      for (let tt = 0; tt < T; tt++) {
        const off = (k * T + tt) * C; let best = 0, bv = -1;
        for (let c = 0; c < C; c++) if (d[off + c] > bv) { bv = d[off + c]; best = c; }
        if (best !== 0 && best !== prev) { const ch = chars[best] || ''; txt += ch; low += (bv < 0.3 ? '1' : '0').repeat(ch.length); ps.push(bv); pp.push(+bv.toFixed(3)); }
        prev = best;
      }
      res[i] = { t: txt, c: ps.length ? ps.reduce((a, b) => a + b, 0) / ps.length : 0, low, pp };
    });
  }
  return res;
}

// vertical ruling lines (forms have many): adaptive threshold, vertical dark runs >= 1.5% of the page height
function verticalLines(rgba, W, H) {
  const g = new Uint8Array(W * H);
  for (let i = 0; i < W * H; i++) g[i] = (rgba[i * 4] * 0.299 + rgba[i * 4 + 1] * 0.587 + rgba[i * 4 + 2] * 0.114) | 0;
  const I = new Float64Array((W + 1) * (H + 1));
  for (let y = 0; y < H; y++) { let s = 0; for (let x = 0; x < W; x++) { s += g[y * W + x]; I[(y + 1) * (W + 1) + x + 1] = I[y * (W + 1) + x + 1] + s; } }
  const R = 12, minRun = Math.round(H * 0.015);
  let total = 0;
  for (let x = 0; x < W; x++) {
    let run = 0;
    const x0 = Math.max(0, x - R), x1 = Math.min(W, x + R + 1);
    for (let y = 0; y <= H; y++) {
      let dark = false;
      if (y < H) {
        const y0 = Math.max(0, y - R), y1 = Math.min(H, y + R + 1);
        const m = (I[y1 * (W + 1) + x1] - I[y0 * (W + 1) + x1] - I[y1 * (W + 1) + x0] + I[y0 * (W + 1) + x0]) / ((x1 - x0) * (y1 - y0));
        dark = g[y * W + x] < m - 15;
      }
      if (dark) run++; else { if (run >= minRun) total += run; run = 0; }
    }
  }
  return total;
}

async function page(m) {
  let { w: W, h: H, data } = m;
  let rgba = new Uint8ClampedArray(data);
  const vl = verticalLines(rgba, W, H);
  let { boxes } = await detect(rgba, W, H);
  let skew = skewAngle(boxes), skew2 = null;
  if (Math.abs(skew) > 0.25) {
    // rotate, then measure again: keep the straighter page (guards against a wrong estimate or direction)
    const rot = rotateRGBA(rgba, W, H, -skew), d2 = await detect(rot, W, H);
    skew2 = skewAngle(d2.boxes);
    if (Math.abs(skew2) < Math.abs(skew) * 0.6) { rgba = rot; boxes = d2.boxes; }
    else {
      const rot2 = rotateRGBA(rgba, W, H, skew), d3 = await detect(rot2, W, H), s3 = skewAngle(d3.boxes);
      if (Math.abs(s3) < Math.abs(skew) * 0.6) { rgba = rot2; boxes = d3.boxes; skew2 = s3; skew = -skew; }
      else skew = 0;
    }
  }
  const texts = await recognize(rgba, W, H, boxes);
  const hs = boxes.map((b) => b.b[3] - b.b[1]).sort((a, b) => a - b), medh = hs.length ? hs[hs.length >> 1] : 0;
  const lines = [];
  boxes.forEach((bx, i) => {
    const { t, c, low, pp } = texts[i], hh = bx.b[3] - bx.b[1];
    if (c < 0.5 && t.trim().length <= 3) return;
    if (hh < 0.45 * medh && c < 0.8) return;
    if (hh < 0.85 * medh && c < 0.6) return;
    if (!t.trim()) return;
    const sp = spell(t, low);
    lines.push({ b: bx.b, t: sp.t, c, low: sp.low, raw: t });
  });
  post({ type: 'page', id: m.id, w: W, h: H, skew, skew2, vl, lines });
}

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === 'init') { await init(m.base); post({ type: 'ready' }); }
    else if (m.type === 'page') await page(m);
  } catch (err) { post({ type: 'error', id: m.id, message: String(err && err.message || err) }); }
};
