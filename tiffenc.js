// Super PDF — TIFF encoder (single & multi-page). Colour/Gray: LZW, Black&White: CCITT Group 4.
(function () {
  // ---------------- bit writer (MSB first)
  function BitWriter(cap) { this.buf = new Uint8Array(cap || 65536); this.pos = 0; this.acc = 0; this.n = 0; }
  BitWriter.prototype.grow = function () { const b = new Uint8Array(this.buf.length * 2); b.set(this.buf); this.buf = b; };
  BitWriter.prototype.write = function (value, bits) {
    for (let i = bits - 1; i >= 0; i--) {
      this.acc = (this.acc << 1) | ((value >>> i) & 1);
      if (++this.n === 8) { if (this.pos >= this.buf.length) this.grow(); this.buf[this.pos++] = this.acc; this.acc = 0; this.n = 0; }
    }
  };
  BitWriter.prototype.writeStr = function (s) { for (let i = 0; i < s.length; i++) this.write(s.charCodeAt(i) - 48, 1); };
  BitWriter.prototype.finish = function () { if (this.n) this.write(0, 8 - this.n); return this.buf.subarray(0, this.pos); };

  // ---------------- LZW (TIFF flavour, as libtiff)
  function lzw(data) {
    const w = new BitWriter(Math.max(65536, data.length >> 2));
    const HS = 1 << 15, keys = new Int32Array(HS), vals = new Int16Array(HS);
    let next, width, maxcode;
    const reset = () => { keys.fill(-1); next = 258; width = 9; maxcode = 511; };
    reset(); w.write(256, 9);
    if (!data.length) { w.write(257, 9); return w.finish(); }
    let prefix = data[0];
    for (let i = 1; i < data.length; i++) {
      const c = data[i], key = (prefix << 8) | c;
      let h = ((key * 2654435761) >>> 17) & (HS - 1);
      while (keys[h] !== -1 && keys[h] !== key) h = (h + 1) & (HS - 1);
      if (keys[h] === key) { prefix = vals[h]; continue; }
      w.write(prefix, width);
      keys[h] = key; vals[h] = next++;
      if (next === 4094) { w.write(256, width); reset(); }
      else if (next > maxcode) { width++; maxcode = (1 << width) - 1; }
      prefix = c;
    }
    w.write(prefix, width);
    // account for the entry the decoder adds for the last code
    const n2 = next + 1;
    if (n2 === 4094) { w.write(256, width); width = 9; }
    else if (n2 > maxcode) width++;
    w.write(257, width);
    return w.finish();
  }

  // ---------------- CCITT G4 tables
  const WT = ['00110101','000111','0111','1000','1011','1100','1110','1111','10011','10100','00111','01000','001000','000011','110100','110101','101010','101011','0100111','0001100','0001000','0010111','0000011','0000100','0101000','0101011','0010011','0100100','0011000','00000010','00000011','00011010','00011011','00010010','00010011','00010100','00010101','00010110','00010111','00101000','00101001','00101010','00101011','00101100','00101101','00000100','00000101','00001010','00001011','01010010','01010011','01010100','01010101','00100100','00100101','01011000','01011001','01011010','01011011','01001010','01001011','00110010','00110011','00110100'];
  const WM = ['11011','10010','010111','0110111','00110110','00110111','01100100','01100101','01101000','01100111','011001100','011001101','011010010','011010011','011010100','011010101','011010110','011010111','011011000','011011001','011011010','011011011','010011000','010011001','010011010','011000','010011011'];
  const BT = ['0000110111','010','11','10','011','0011','0010','00011','000101','000100','0000100','0000101','0000111','00000100','00000111','000011000','0000010111','0000011000','0000001000','00001100111','00001101000','00001101100','00000110111','00000101000','00000010111','00000011000','000011001010','000011001011','000011001100','000011001101','000001101000','000001101001','000001101010','000001101011','000011010010','000011010011','000011010100','000011010101','000011010110','000011010111','000001101100','000001101101','000011011010','000011011011','000001010100','000001010101','000001010110','000001010111','000001100100','000001100101','000001010010','000001010011','000000100100','000000110111','000000111000','000000100111','000000101000','000001011000','000001011001','000000101011','000000101100','000001011010','000001100110','000001100111'];
  const BM = ['0000001111','000011001000','000011001001','000001011011','000000110011','000000110100','000000110101','0000001101100','0000001101101','0000001001010','0000001001011','0000001001100','0000001001101','0000001110010','0000001110011','0000001110100','0000001110101','0000001110110','0000001110111','0000001010010','0000001010011','0000001010100','0000001010101','0000001011010','0000001011011','0000001100100','0000001100101'];
  const XM = ['00000001000','00000001100','00000001101','000000010010','000000010011','000000010100','000000010101','000000010110','000000010111','000000011100','000000011101','000000011110','000000011111'];
  const VCODE = { '-3': '0000010', '-2': '000010', '-1': '010', '0': '1', '1': '011', '2': '000011', '3': '0000011' };

  function putRun(w, run, black) {
    const T = black ? BT : WT, M = black ? BM : WM;
    while (run >= 2560 + 64) { w.writeStr(XM[12]); run -= 2560; }
    if (run >= 64) {
      const m = Math.floor(run / 64) * 64;
      w.writeStr(m <= 1728 ? M[m / 64 - 1] : XM[(m - 1792) / 64]);
      run -= m;
    }
    w.writeStr(T[run]);
  }

  // bits: Uint8Array of 0 (white) / 1 (black), width*height
  function g4(bits, width, height) {
    const w = new BitWriter(65536);
    let ref = new Uint8Array(width); // all white
    const nextChange = (line, pos) => { // first x > pos with line[x] != line[x-1] (line[-1] = white)
      let x = pos + 1;
      if (x <= 0) { if (line[0]) return 0; x = 1; }
      const c = line[x - 1];
      while (x < width && line[x] === c) x++;
      return x;
    };
    const colorAt = (line, x) => (x < 0 || x >= width ? 0 : line[x]);
    for (let y = 0; y < height; y++) {
      const cur = bits.subarray(y * width, (y + 1) * width);
      let a0 = -1, color = 0;
      while (a0 < width) {
        const a1 = nextChange(cur, a0);
        let b1 = nextChange(ref, a0);
        while (b1 < width && colorAt(ref, b1) === color) b1 = nextChange(ref, b1);
        const b2 = b1 < width ? nextChange(ref, b1) : width;
        if (b2 < a1) { w.writeStr('0001'); a0 = b2; continue; }
        const d = a1 - b1;
        if (d >= -3 && d <= 3) { w.writeStr(VCODE[d]); a0 = a1; color ^= 1; continue; }
        const a2 = a1 < width ? nextChange(cur, a1) : width;
        w.writeStr('001');
        putRun(w, a1 - Math.max(a0, 0), color === 1);
        putRun(w, a2 - a1, color === 0);
        a0 = a2;
      }
      ref = cur;
    }
    w.writeStr('000000000001000000000001'); // EOFB
    return w.finish();
  }

  // ---------------- page preparation from canvas
  // mode: 'rgb' | 'gray' | 'bw'
  function fromCanvas(canvas, mode, dpi, threshold) {
    const W = canvas.width, H = canvas.height;
    const px = canvas.getContext('2d').getImageData(0, 0, W, H).data;
    let data, strip;
    if (mode === 'rgb') {
      data = new Uint8Array(W * H * 3);
      for (let i = 0, j = 0; i < px.length; i += 4) { data[j++] = px[i]; data[j++] = px[i + 1]; data[j++] = px[i + 2]; }
      strip = lzw(data);
    } else {
      const gray = new Uint8Array(W * H);
      for (let i = 0, j = 0; i < px.length; i += 4, j++) gray[j] = (px[i] * 299 + px[i + 1] * 587 + px[i + 2] * 114) / 1000;
      if (mode === 'gray') strip = lzw(gray);
      else {
        const t = threshold || 160;
        for (let j = 0; j < gray.length; j++) gray[j] = gray[j] < t ? 1 : 0;
        strip = g4(gray, W, H);
      }
    }
    return { width: W, height: H, mode, dpi: dpi || 150, strip };
  }

  // ---------------- TIFF container
  function encode(pages) {
    const TYPE = { SHORT: 3, LONG: 4, RATIONAL: 5 };
    let size = 8;
    for (const p of pages) size += p.strip.length + (p.strip.length & 1) + 2 + 16 * 12 + 4 + 64;
    const out = new Uint8Array(size);
    const dv = new DataView(out.buffer);
    out[0] = 0x49; out[1] = 0x49; dv.setUint16(2, 42, true);
    let off = 8, prevNextPtr = 4;
    pages.forEach((p, pi) => {
      const stripOff = off;
      out.set(p.strip, off); off += p.strip.length; if (off & 1) off++;
      const spp = p.mode === 'rgb' ? 3 : 1;
      const tags = [
        [254, TYPE.LONG, 1, pages.length > 1 ? 2 : 0],
        [256, TYPE.LONG, 1, p.width],
        [257, TYPE.LONG, 1, p.height],
        [258, TYPE.SHORT, spp, p.mode === 'bw' ? 1 : 8],
        [259, TYPE.SHORT, 1, p.mode === 'bw' ? 4 : 5],
        [262, TYPE.SHORT, 1, p.mode === 'bw' ? 0 : p.mode === 'gray' ? 1 : 2],
        [273, TYPE.LONG, 1, stripOff],
        [277, TYPE.SHORT, 1, spp],
        [278, TYPE.LONG, 1, p.height],
        [279, TYPE.LONG, 1, p.strip.length],
        [282, TYPE.RATIONAL, 1, p.dpi],
        [283, TYPE.RATIONAL, 1, p.dpi],
        [284, TYPE.SHORT, 1, 1],
        [296, TYPE.SHORT, 1, 2],
        [297, TYPE.SHORT, 2, [pi, pages.length]],
      ];
      const ifdOff = off;
      dv.setUint32(prevNextPtr, ifdOff, true);
      let extra = ifdOff + 2 + tags.length * 12 + 4;
      dv.setUint16(off, tags.length, true); off += 2;
      for (const [tag, type, count, val] of tags) {
        dv.setUint16(off, tag, true); dv.setUint16(off + 2, type, true); dv.setUint32(off + 4, count, true);
        if (type === TYPE.RATIONAL) {
          dv.setUint32(off + 8, extra, true);
          dv.setUint32(extra, Math.round(val * 100), true); dv.setUint32(extra + 4, 100, true); extra += 8;
        } else if (tag === 258 && count === 3) {
          dv.setUint32(off + 8, extra, true);
          for (let k = 0; k < 3; k++) dv.setUint16(extra + k * 2, 8, true); extra += 6; if (extra & 1) extra++;
        } else if (tag === 297) {
          dv.setUint16(off + 8, val[0], true); dv.setUint16(off + 10, val[1], true);
        } else if (type === TYPE.SHORT) dv.setUint16(off + 8, val, true);
        else dv.setUint32(off + 8, val, true);
        off += 12;
      }
      prevNextPtr = off;
      dv.setUint32(off, 0, true); off += 4;
      off = extra; if (off & 1) off++;
    });
    return out.subarray(0, off);
  }

  self.TiffEnc = { encode, fromCanvas, lzw, g4 };
})();
