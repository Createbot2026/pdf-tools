// Core PDF operations (pdf-lib). Works in browser (window.PDFLib) and Node (require).
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('pdf-lib'));
  else root.PPTCore = factory(root.PDFLib);
})(typeof self !== 'undefined' ? self : this, function (PDFLib) {
  const { PDFDocument, StandardFonts, rgb, degrees } = PDFLib;
  // Don't rely on constructor.name: it is mangled in the minified browser build.
  function fieldType(f) {
    const L = PDFLib;
    if (f instanceof L.PDFTextField) return 'TextField';
    if (f instanceof L.PDFCheckBox) return 'CheckBox';
    if (f instanceof L.PDFDropdown) return 'Dropdown';
    if (f instanceof L.PDFOptionList) return 'OptionList';
    if (f instanceof L.PDFRadioGroup) return 'RadioGroup';
    if (f instanceof L.PDFSignature) return 'Signature';
    if (f instanceof L.PDFButton) return 'Button';
    return 'Unknown';
  }
  const ENCRYPTED_MSG = 'This PDF is password-protected or encrypted, so it can\'t be edited. Open it in Chrome, enter the password, choose Print → "Save as PDF", then use the saved copy.';
  function encryptedError() { const e = new Error(ENCRYPTED_MSG); e.code = 'ENCRYPTED'; return e; }
  async function load(b) {
    let doc;
    try { doc = await PDFDocument.load(b, { ignoreEncryption: true }); }
    catch (e) { if (/encrypt|password/i.test(e.message || '')) throw encryptedError(); throw new Error('Could not read this PDF (it may be damaged): ' + (e.message || e)); }
    if (doc.isEncrypted) throw encryptedError();
    return doc;
  }
  // Page rotation in degrees (0/90/180/270) as set by /Rotate.
  const rot = p => (((p.getRotation().angle || 0) % 360) + 360) % 360;
  // Rotate vector (dx,dy) counter-clockwise by deg.
  function rv(dx, dy, deg) { const a = deg * Math.PI / 180, c = Math.round(Math.cos(a)), s = Math.round(Math.sin(a)); return [dx * c - dy * s, dx * s + dy * c]; }
  // Unicode font (Noto Sans subset: Latin, Greek, Cyrillic, symbols). Browser: fetched; Node: read from disk.
  async function getUnicodeFont(doc, fontBytes) {
    const fk = (typeof fontkit !== 'undefined') ? fontkit : (typeof require === 'function' ? require('@pdf-lib/fontkit') : null);
    if (!fk || !fontBytes) return null;
    doc.registerFontkit(fk);
    return doc.embedFont(fontBytes, { subset: true });
  }

  // Parse "1-3,5,7-" into 0-based indices (pageCount needed for open ranges)
  function parseRanges(spec, pageCount) {
    const out = [];
    for (const part of String(spec).split(',').map(s => s.trim()).filter(Boolean)) {
      const m = part.match(/^(\d*)\s*-\s*(\d*)$/);
      let a, b;
      if (m) { a = m[1] ? +m[1] : 1; b = m[2] ? +m[2] : pageCount; }
      else if (/^\d+$/.test(part)) { a = b = +part; }
      else throw new Error('Invalid page range: ' + part);
      if (a < 1 || b > pageCount || a > b) throw new Error('Page range out of bounds: ' + part);
      for (let i = a; i <= b; i++) out.push(i - 1);
    }
    return out;
  }

  async function merge(buffers) {
    const out = await PDFDocument.create();
    for (const b of buffers) {
      const src = await load(b);
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach(p => out.addPage(p));
    }
    return out.save();
  }

  // order: array of {index, rotate} or numbers (0-based); pages not listed are dropped
  async function reorder(buffer, order) {
    const src = await load(buffer);
    const out = await PDFDocument.create();
    const items = order.map(o => typeof o === 'number' ? { index: o, rotate: 0 } : o);
    const pages = await out.copyPages(src, items.map(i => i.index));
    pages.forEach((p, k) => {
      const r = items[k].rotate || 0;
      if (r) p.setRotation(degrees((p.getRotation().angle + r) % 360));
      out.addPage(p);
    });
    return out.save();
  }

  // Extract pages by range spec into one PDF
  async function extract(buffer, spec) {
    const src = await load(buffer);
    return reorder(buffer, parseRanges(spec, src.getPageCount()));
  }

  // Split into multiple PDFs: groups separated by ';' or every page if spec === 'each'
  async function split(buffer, spec) {
    const src = await load(buffer);
    const n = src.getPageCount();
    const groups = spec === 'each'
      ? src.getPageIndices().map(i => [i])
      : String(spec).split(';').map(s => parseRanges(s, n)).filter(g => g.length);
    const results = [];
    for (const g of groups) results.push(await reorder(buffer, g));
    return results;
  }

  function sniffImage(bytes) {
    const u = new Uint8Array(bytes);
    if (u[0] === 0x89 && u[1] === 0x50) return 'png';
    if (u[0] === 0xff && u[1] === 0xd8) return 'jpg';
    throw new Error('Unsupported image (use JPG or PNG)');
  }

  // images: array of bytes; pageSize: 'fit' or 'a4'
  async function imagesToPdf(images, pageSize = 'fit') {
    const doc = await PDFDocument.create();
    for (const bytes of images) {
      const u8 = Uint8Array.from(new Uint8Array(bytes));
      const img = sniffImage(u8) === 'png' ? await doc.embedPng(u8) : await doc.embedJpg(u8);
      if (pageSize === 'a4') {
        const [W, H] = [595.28, 841.89], m = 36;
        const s = Math.min((W - 2 * m) / img.width, (H - 2 * m) / img.height, 1);
        const w = img.width * s, h = img.height * s;
        doc.addPage([W, H]).drawImage(img, { x: (W - w) / 2, y: (H - h) / 2, width: w, height: h });
      } else {
        doc.addPage([img.width, img.height]).drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
      }
    }
    return doc.save();
  }

  function hexToRgb(hex) {
    const h = (hex || '#000000').replace('#', '');
    return rgb(parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255);
  }

  // items: [{page (0-based), x, y (PDF user-space points of the first baseline start), text, size, color, note:boolean}]
  // Text is drawn upright as the page is displayed, including on rotated pages.
  // opts.fontBytes: TTF bytes for a Unicode font. Characters the font can't draw (e.g. emoji) are removed;
  // they're reported in opts.stripped (array) if given.
  async function addText(buffer, items, opts = {}) {
    const doc = await load(buffer);
    let font = null;
    try { font = await getUnicodeFont(doc, opts.fontBytes); } catch (e) { font = null; }
    if (!font) font = await doc.embedFont(StandardFonts.Helvetica);
    const cs = new Set(font.getCharacterSet());
    const stripped = opts.stripped || [];
    const clean = t => [...String(t)].filter(ch => { if (ch === '\n') return true; const ok = cs.has(ch.codePointAt(0)); if (!ok && !/[\uFE0F\u200D]/.test(ch)) stripped.push(ch); return ok; }).join('');
    const pages = doc.getPages();
    for (const it of items) {
      const p = pages[it.page];
      if (!p) throw new Error('No page ' + (it.page + 1));
      const R = rot(p);
      const size = it.size || 12;
      const lines = clean(it.text).split('\n');
      const at = (dx, dy) => { const [a, b] = rv(dx, dy, R); return { x: it.x + a, y: it.y + b }; };
      if (it.note) {
        const w = Math.max(...lines.map(l => font.widthOfTextAtSize(l, size))) + 12;
        const h = lines.length * size * 1.2 + 10;
        p.drawRectangle({ ...at(-6, -h + size + 4), width: w, height: h, rotate: degrees(R), color: rgb(1, 0.95, 0.6), borderColor: rgb(0.85, 0.7, 0.1), borderWidth: 1 });
      }
      lines.forEach((l, i) => { if (l) p.drawText(l, { ...at(0, -i * size * 1.2), size, font, rotate: degrees(R), color: hexToRgb(it.color) }); });
    }
    return doc.save();
  }

  async function listFields(buffer) {
    const doc = await load(buffer);
    return doc.getForm().getFields().map(f => {
      const type = fieldType(f);
      const o = { name: f.getName(), type };
      try {
        if (type === 'TextField') o.value = f.getText() || '';
        else if (type === 'CheckBox') o.value = f.isChecked();
        else if (type === 'Dropdown' || type === 'OptionList') { o.options = f.getOptions(); o.value = (f.getSelected() || [])[0] || ''; }
        else if (type === 'RadioGroup') { o.options = f.getOptions(); o.value = f.getSelected() || ''; }
      } catch (e) {}
      return o;
    });
  }

  async function fillForm(buffer, values, flatten = false) {
    const doc = await load(buffer);
    const form = doc.getForm();
    for (const [name, v] of Object.entries(values)) {
      const f = form.getField(name);
      const t = 'PDF' + fieldType(f);
      if (t === 'PDFTextField') f.setText(String(v));
      else if (t === 'PDFCheckBox') v ? f.check() : f.uncheck();
      else if (t === 'PDFDropdown' || t === 'PDFOptionList' || t === 'PDFRadioGroup') { if (v) f.select(String(v)); }
    }
    if (flatten) form.flatten();
    return doc.save();
  }

  // Place a PNG signature image. pos: {page, x, y, width} in PDF points
  async function addImageStamp(buffer, pngBytes, pos) {
    const doc = await load(buffer);
    const img = await doc.embedPng(Uint8Array.from(new Uint8Array(pngBytes)));
    const p = doc.getPages()[pos.page];
    if (!p) throw new Error('No page ' + (pos.page + 1));
    const w = pos.width || 150, h = img.height * (w / img.width);
    if (pos.cx != null) {
      // Centre on (cx, cy) in user space, upright as displayed (handles /Rotate).
      const R = rot(p); const [a, b] = rv(-w / 2, -h / 2, R);
      p.drawImage(img, { x: pos.cx + a, y: pos.cy + b, width: w, height: h, rotate: degrees(R) });
    } else p.drawImage(img, { x: pos.x, y: pos.y, width: w, height: h });
    return doc.save();
  }

  async function pageCount(buffer) { return (await load(buffer)).getPageCount(); }
  async function pageSizes(buffer) { return (await load(buffer)).getPages().map(p => p.getSize()); }

  // Add page numbers. opts: {start, format ('n'|'n_of_N'|'Page n'), position ('br'|'bc'|'bl'|'tr'|'tc'|'tl'), size, color, margin}
  async function addPageNumbers(buffer, opts = {}) {
    const doc = await load(buffer);
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const pages = doc.getPages();
    const start = opts.start == null ? 1 : +opts.start;
    const format = opts.format || 'n';
    const pos = opts.position || 'bc';
    const size = opts.size || 10;
    const color = hexToRgb(opts.color || '#333333');
    const margin = opts.margin == null ? 28 : +opts.margin;
    const N = pages.length;
    pages.forEach((p, i) => {
      const n = start + i;
      let text = String(n);
      if (format === 'n_of_N') text = n + ' / ' + (start + N - 1);
      else if (format === 'Page n') text = 'Page ' + n;
      const { width: W, height: H } = p.getSize();
      const tw = font.widthOfTextAtSize(text, size);
      let x, y;
      if (pos[0] === 't') y = H - margin;
      else y = margin;
      if (pos[1] === 'l') x = margin;
      else if (pos[1] === 'r') x = W - margin - tw;
      else x = (W - tw) / 2;
      p.drawText(text, { x, y, size, font, color });
    });
    return doc.save();
  }

  // Diagonal or corner text watermark. opts: {text, opacity, size, color, angle, pages ('all'|spec)}
  async function addWatermark(buffer, opts = {}) {
    const doc = await load(buffer);
    const font = await doc.embedFont(StandardFonts.HelveticaBold);
    const text = String(opts.text || 'DRAFT');
    const size = opts.size || 48;
    const opacity = opts.opacity == null ? 0.25 : +opts.opacity;
    const color = hexToRgb(opts.color || '#888888');
    const angle = opts.angle == null ? 45 : +opts.angle;
    const pages = doc.getPages();
    const indices = opts.pages && opts.pages !== 'all'
      ? parseRanges(opts.pages, pages.length)
      : pages.map((_, i) => i);
    const tw = font.widthOfTextAtSize(text, size);
    for (const i of indices) {
      const p = pages[i];
      const { width: W, height: H } = p.getSize();
      p.drawText(text, {
        x: (W - tw * Math.cos(angle * Math.PI / 180)) / 2,
        y: (H - tw * Math.sin(angle * Math.PI / 180)) / 2,
        size, font, color, opacity, rotate: degrees(angle)
      });
    }
    return doc.save();
  }

  // Mixed merge: each item is {type:'pdf', bytes} or {type:'image', bytes}. Images become pages.
  async function mixedMerge(items, pageSize = 'a4') {
    const out = await PDFDocument.create();
    for (const it of items) {
      if (it.type === 'pdf') {
        const src = await load(it.bytes);
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(p => out.addPage(p));
      } else {
        const u8 = Uint8Array.from(new Uint8Array(it.bytes));
        const img = sniffImage(u8) === 'png' ? await out.embedPng(u8) : await out.embedJpg(u8);
        if (pageSize === 'a4') {
          const [W, H] = [595.28, 841.89], m = 36;
          const s = Math.min((W - 2 * m) / img.width, (H - 2 * m) / img.height, 1);
          const w = img.width * s, h = img.height * s;
          out.addPage([W, H]).drawImage(img, { x: (W - w) / 2, y: (H - h) / 2, width: w, height: h });
        } else {
          out.addPage([img.width, img.height]).drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        }
      }
    }
    return out.save();
  }

  return { ENCRYPTED_MSG, parseRanges, merge, reorder, extract, split, imagesToPdf, addText, listFields, fillForm, addImageStamp, pageCount, pageSizes, addPageNumbers, addWatermark, mixedMerge };
});
