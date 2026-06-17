// SPDX-License-Identifier: 0BSD
/*
 * textToDoc(input[, template]) -> Uint8Array : write a Word 97-2003 binary .doc.
 * `input` is either a plain string or a styled model — an array of paragraphs
 * { runs: [{ text, b, i, u, strike, size, font, color }], kind: 'p'|'cell'|'rowEnd' }
 * as produced by docToText.model(). `template` is an optional blank .doc
 * (Uint8Array / ArrayBuffer / Buffer) saved by a real word processor; when
 * omitted, a small bundled skeleton (embedded below) is used.
 *
 * Why a template: a from-scratch .doc round-trips through lenient parsers but
 * real word processors reject it — they need a valid stylesheet, section table
 * and character/paragraph property tables. Rather than synthesise all of those
 * (Apache POI doesn't either), we reuse them from a genuine blank document and
 * swap in the body text + piece table + freshly built CHPX/PAPX FKP pages.
 * Clean-room from [MS-CFB] (container) + [MS-DOC] (FIB / CLX / FKP / sprms).
 *
 * Writes back: paragraphs with alignment (sprmPJc) and lists (bullet/numbered
 * with nesting, via sprmPIlfo/sprmPIlvl against the skeleton's built-in list
 * definitions), character formatting (bold/italic/underline/strike/size/colour/
 * font as CHPX sprms — fonts the skeleton lacks are appended to its SttbfFfn),
 * tables (cell marks + sprmPFInTable / sprmPFTtp / sprmTDefTable with borders),
 * and inline images (PNG/JPEG as an OfficeArt picture in a Data stream, sized +
 * placed). Not yet: paragraph spacing/indentation, hyperlink URLs.
 *
 * Verification: round-tripped through docToText (.model re-reads the table
 * cells) AND the independent word-extractor (test/styled.test.js / writer.test.js),
 * and confirmed to open as real bold text + a real table in SoftMaker TextMaker,
 * driven via its COM automation (scripts/read-with-textmaker.ps1).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.textToDoc = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SECTOR = 512, FREESECT = 0xFFFFFFFF, ENDOFCHAIN = 0xFFFFFFFE, FATSECT = 0xFFFFFFFD;

  function u16(b, o, v) { b[o] = v & 0xFF; b[o + 1] = (v >> 8) & 0xFF; }
  function u32(b, o, v) { b[o] = v & 0xFF; b[o + 1] = (v >> 8) & 0xFF; b[o + 2] = (v >> 16) & 0xFF; b[o + 3] = (v >>> 24) & 0xFF; }
  function toU8(x) {
    if (x instanceof Uint8Array) return x;
    if (typeof ArrayBuffer !== 'undefined' && x instanceof ArrayBuffer) return new Uint8Array(x);
    if (x && x.buffer) return new Uint8Array(x.buffer, x.byteOffset, x.byteLength);
    return new Uint8Array(x);
  }

  function b64ToU8(s) {
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(s, 'base64'));
    var bin = atob(s), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }
  var _tpl = null;
  function defaultTemplate() { return _tpl || (_tpl = b64ToU8(TEMPLATE_B64)); }

  // ---- minimal [MS-CFB] reader: name -> stream bytes (root level) ----------
  function readCfb(bytes) {
    var b = toU8(bytes), dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    if (b.length < 512 || dv.getUint32(0, true) !== 0xE011CFD0) return null;
    var ssz = 1 << dv.getUint16(30, true), mssz = 1 << dv.getUint16(32, true);
    var nFat = dv.getUint32(44, true), dirStart = dv.getUint32(48, true);
    var miniCutoff = dv.getUint32(56, true), miniFatStart = dv.getUint32(60, true);
    function sectorOff(s) { return SECTOR + s * ssz; }
    // DIFAT -> FAT sector list (109 in header; deeper DIFAT unsupported but rare for our files)
    var fatSecs = [];
    for (var i = 0; i < 109 && fatSecs.length < nFat; i++) { var s = dv.getUint32(76 + i * 4, true); if (s !== FREESECT && s !== ENDOFCHAIN) fatSecs.push(s); }
    var fat = new Uint32Array(fatSecs.length * (ssz / 4));
    for (i = 0; i < fatSecs.length; i++) for (var j = 0; j < ssz / 4; j++) fat[i * (ssz / 4) + j] = dv.getUint32(sectorOff(fatSecs[i]) + j * 4, true);
    function chain(start, fatTable) { var out = [], s = start, guard = 0; while (s !== ENDOFCHAIN && s !== FREESECT && guard++ < 1e6) { out.push(s); s = fatTable[s]; } return out; }
    function readRegular(start, size) {
      var secs = chain(start, fat), out = new Uint8Array(secs.length * ssz);
      for (var k = 0; k < secs.length; k++) out.set(b.subarray(sectorOff(secs[k]), sectorOff(secs[k]) + ssz), k * ssz);
      return out.subarray(0, size);
    }
    // directory
    var dirBytes = readRegular(dirStart, 1 << 30), entries = [];
    for (i = 0; i + 128 <= dirBytes.length; i += 128) {
      var type = dirBytes[i + 0x42]; if (type === 0) continue;
      var nameLen = dirBytes[i] | (dirBytes[i + 1] << 8) ? (new DataView(dirBytes.buffer, dirBytes.byteOffset).getUint16(i + 0x40, true)) : 0;
      var name = '';
      for (var c = 0; c + 1 < nameLen; c += 2) { var ch = dirBytes[i + c] | (dirBytes[i + c + 1] << 8); if (ch) name += String.fromCharCode(ch); }
      var ddv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
      entries.push({ name: name, type: type, start: ddv.getUint32(i + 0x74, true), size: ddv.getUint32(i + 0x78, true) });
    }
    var root = entries.filter(function (e) { return e.type === 5; })[0];
    // mini stream (root's stream) + mini FAT, for small streams
    var miniStream = root ? readRegular(root.start, root.size) : new Uint8Array(0);
    var miniFat = null;
    if (miniFatStart !== ENDOFCHAIN) {
      var mfSecs = chain(miniFatStart, fat); miniFat = new Uint32Array(mfSecs.length * (ssz / 4));
      for (i = 0; i < mfSecs.length; i++) for (j = 0; j < ssz / 4; j++) miniFat[i * (ssz / 4) + j] = dv.getUint32(sectorOff(mfSecs[i]) + j * 4, true);
    }
    function readMini(start, size) {
      var secs = chain(start, miniFat), out = new Uint8Array(secs.length * mssz);
      for (var k = 0; k < secs.length; k++) out.set(miniStream.subarray(secs[k] * mssz, secs[k] * mssz + mssz), k * mssz);
      return out.subarray(0, size);
    }
    var byName = {};
    entries.forEach(function (e) {
      if (e.type !== 2) return;
      byName[e.name] = e.size < miniCutoff ? readMini(e.start, e.size) : readRegular(e.start, e.size);
    });
    return { byName: byName };
  }

  // ---- [MS-CFB] writer (streams padded >= 4096 -> regular FAT only) --------
  function writeHeader(b, fatSectors, dirStart) {
    var sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for (var i = 0; i < 8; i++) b[i] = sig[i];
    u16(b, 24, 0x003E); u16(b, 26, 0x0003); u16(b, 28, 0xFFFE); u16(b, 30, 0x0009); u16(b, 32, 0x0006);
    u32(b, 44, fatSectors); u32(b, 48, dirStart); u32(b, 56, 0x00001000);
    u32(b, 60, ENDOFCHAIN); u32(b, 64, 0); u32(b, 68, ENDOFCHAIN); u32(b, 72, 0);
    for (var d = 0; d < 109; d++) u32(b, 76 + d * 4, d < fatSectors ? d : FREESECT);
  }
  function writeDir(b, idx, name, type, start, size, left, right, child) {
    var o = idx * 128, n;
    for (n = 0; n < name.length && n < 31; n++) u16(b, o + n * 2, name.charCodeAt(n));
    u16(b, o + 0x40, name.length ? (name.length + 1) * 2 : 0);
    b[o + 0x42] = type; b[o + 0x43] = 1;
    u32(b, o + 0x44, left < 0 ? FREESECT : left);
    u32(b, o + 0x48, right < 0 ? FREESECT : right);
    u32(b, o + 0x4C, child < 0 ? FREESECT : child);
    u32(b, o + 0x74, start); u32(b, o + 0x78, size >>> 0);
  }
  function buildCfb(streams) {
    var pad = streams.map(function (s) {
      var len = Math.ceil(Math.max(s.data.length, 4096) / SECTOR) * SECTOR;
      var d = new Uint8Array(len); d.set(s.data);
      return { name: s.name, data: d, size: len, sectors: len / SECTOR };
    });
    // CFB directory is a tree ordered by (uppercased name length, then name);
    // a right-only chain in that order is a valid (if unbalanced) tree to walk.
    pad.sort(function (a, b) { var an = a.name.toUpperCase(), bn = b.name.toUpperCase(); return an.length - bn.length || (an < bn ? -1 : an > bn ? 1 : 0); });
    var dirSectors = Math.ceil((1 + pad.length) / 4);
    var nonFat = dirSectors + pad.reduce(function (a, s) { return a + s.sectors; }, 0);
    var fatSectors = 1; while (Math.ceil((nonFat + fatSectors) / 128) > fatSectors) fatSectors++;
    var total = fatSectors + nonFat, dirStart = fatSectors, p = dirStart + dirSectors;
    pad.forEach(function (s) { s.start = p; p += s.sectors; });
    var fat = new Uint8Array(fatSectors * SECTOR);
    for (var i = 0; i < fatSectors * 128; i++) u32(fat, i * 4, FREESECT);
    for (i = 0; i < fatSectors; i++) u32(fat, i * 4, FATSECT);
    for (i = 0; i < dirSectors; i++) u32(fat, (dirStart + i) * 4, i < dirSectors - 1 ? dirStart + i + 1 : ENDOFCHAIN);
    pad.forEach(function (s) { for (var k = 0; k < s.sectors; k++) u32(fat, (s.start + k) * 4, k < s.sectors - 1 ? s.start + k + 1 : ENDOFCHAIN); });
    var dir = new Uint8Array(dirSectors * SECTOR);
    writeDir(dir, 0, 'Root Entry', 5, ENDOFCHAIN, 0, -1, -1, pad.length ? 1 : -1);
    pad.forEach(function (s, idx) { writeDir(dir, idx + 1, s.name, 2, s.start, s.size, -1, idx < pad.length - 1 ? idx + 2 : -1, -1); });
    for (var e = 1 + pad.length; e < dirSectors * 4; e++) writeDir(dir, e, '', 0, 0, 0, -1, -1, -1);
    var file = new Uint8Array(SECTOR * (1 + total));
    writeHeader(file, fatSectors, dirStart);
    file.set(fat, SECTOR); file.set(dir, SECTOR + dirStart * SECTOR);
    pad.forEach(function (s) { file.set(s.data, SECTOR + s.start * SECTOR); });
    return file;
  }

  // ---- character properties -> CHPX grpprl ([MS-DOC] sprms, inverse of the
  // reader's parseChpGrpprl). ToggleOperands are written as 0x01 (on).
  function sprm(b, code, operand) { b.push(code & 0xFF, (code >> 8) & 0xFF); for (var i = 0; i < operand.length; i++) b.push(operand[i] & 0xFF); }
  function chpxGrpprl(r, ftc) {
    var b = [];
    if (r.b) sprm(b, 0x0835, [1]);            // sprmCFBold
    if (r.i) sprm(b, 0x0836, [1]);            // sprmCFItalic
    if (r.strike) sprm(b, 0x0837, [1]);       // sprmCFStrike
    if (r.u) sprm(b, 0x2A3E, [1]);            // sprmCKul = 1 (single underline)
    if (r.size) { var hp = Math.round(r.size * 2); sprm(b, 0x4A43, [hp, hp >> 8]); }      // sprmCHps (half-points)
    if (r.color != null) sprm(b, 0x6870, [r.color, r.color >> 8, r.color >> 16, 0]);     // sprmCCv (COLORREF)
    if (ftc != null) sprm(b, 0x4A4F, [ftc, ftc >> 8]);   // sprmCRgFtc0 (font index into SttbfFfn)
    return b;
  }
  // A Chpx in an FKP: [cb][grpprl]. null => use the FKP default (rgb = 0).
  function chpxBlob(r, ftc) { var g = chpxGrpprl(r, ftc); if (!g.length) return null; var u = new Uint8Array(g.length + 1); u[0] = g.length; u.set(g, 1); return u; }

  // ---- font table (SttbfFfn) -----------------------------------------------
  // Read the skeleton's font names (FFNs are self-sized via cbFfnM1; name is
  // UTF-16 at offset 40). Returns the existing names + where to bump the count.
  function readFonts(tbl, fc, lcb) {
    var dvT = new DataView(tbl.buffer, tbl.byteOffset, tbl.byteLength);
    var p = fc, cDataOff = fc;
    if (dvT.getUint16(p, true) === 0xFFFF) { p += 2; cDataOff = p; }
    var cData = dvT.getUint16(p, true); p += 4;            // cData (2) + cbExtra (2)
    var names = [], end = fc + lcb;
    for (var i = 0; i < cData && p < end; i++) {
      var ffnEnd = p + tbl[p] + 1, name = '';
      for (var q = p + 40; q + 1 < ffnEnd && q + 1 <= end; q += 2) { var ch = tbl[q] | (tbl[q + 1] << 8); if (!ch) break; name += String.fromCharCode(ch); }
      names.push(name); p = ffnEnd;
    }
    // `used` = bytes through the last FFN; append new fonts there (not after any
    // trailing slack within lcb, which would desync the sequential read).
    return { names: names, cData: cData, cDataOff: cDataOff - fc, used: p - fc, bytes: tbl.slice(fc, fc + lcb) };
  }
  // Build one FFN: cbFfnM1, flags (fTrueType), wWeight, then the UTF-16 name+NUL.
  function buildFfn(name) {
    name = name.slice(0, 100);
    var size = 40 + (name.length + 1) * 2, u = new Uint8Array(size);
    u[0] = size - 1; u[1] = 0x04; u16(u, 2, 400);          // cbFfnM1, fTrueType, wWeight=normal
    for (var i = 0; i < name.length; i++) u16(u, 40 + i * 2, name.charCodeAt(i));
    return u;
  }

  // ---- inline pictures -----------------------------------------------------
  // Replicates SoftMaker TextMaker's own inline-picture layout (reverse-engineered
  // from a reference doc): in the Data stream, a PICF (mm=0x64) + OfficeArt
  // SpContainer (picture-frame shape + OPT + anchor) + an inline FBSE whose blip
  // (msofbtBlipPNG/JPEG) holds rgbUid(16) + tag(1) + the image bytes. The drawing
  // group defaults (DggInfo) go in the table stream at FibRgFcLcb #50. A 0x01
  // picture char in the text carries sprmCFSpec + sprmCPicLocation (Data offset).
  var _pic = null;
  function picParts() { return _pic || (_pic = { picf: b64ToU8(PIC_PICF), sp: b64ToU8(PIC_SP), fbse: b64ToU8(PIC_FBSE), blip: b64ToU8(PIC_BLIP), dgg: b64ToU8(PIC_DGG) }); }
  function imgDims(b) {
    if (b[0] === 0x89) return { w: (b[16] * 0x1000000 + (b[17] << 16) + (b[18] << 8) + b[19]), h: (b[20] * 0x1000000 + (b[21] << 16) + (b[22] << 8) + b[23]) };
    if (b[0] === 0xFF && b[1] === 0xD8) for (var p = 2; p + 9 < b.length;) { if (b[p] !== 0xFF) { p++; continue; } var m = b[p + 1]; if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: (b[p + 5] << 8) | b[p + 6], w: (b[p + 7] << 8) | b[p + 8] }; p += 2 + ((b[p + 2] << 8) | b[p + 3]); }
    return null;
  }
  // Bytes for one picture in the Data stream (PICF + SpContainer + FBSE + blip).
  function imagePart(img) {
    var M = picParts(), bytes = img.bytes, n = bytes.length;
    var isJpeg = /jpe?g/i.test(img.mime || '') || (bytes[0] === 0xFF && bytes[1] === 0xD8);
    var picf = M.picf.slice(), sp = M.sp, fbse = M.fbse.slice(), blip = M.blip.slice();
    if (isJpeg) { u16(blip, 0, 0x46A << 4); u16(blip, 2, 0xF01D); fbse[8] = fbse[9] = 5; } // msofbtBlipJPEG / btWin32=JPEG
    u32(blip, 4, 17 + n);                 // blip recLen = rgbUid(16) + tag(1) + image
    u32(fbse, 4, 61 + n);                 // FBSE recLen = 36 fields + 8 blip header + 17 + n
    u32(fbse, 28, 25 + n);                // FBSE.size = blip record (8 + 17 + n)
    u32(fbse, 36, 0);                     // FBSE.foDelay = 0 (blip is inline)
    var lcb = 68 + sp.length + 8 + (61 + n);
    u32(picf, 0, lcb);
    // Display size: prefer the original's PICF dimensions (img.dxa/dya twips);
    // otherwise derive from the image's own pixels (15 twips/px @96dpi, capped).
    var w = img.dxa, h = img.dya, dim = imgDims(bytes);
    if (!(w > 0 && h > 0) && dim && dim.w && dim.h) { w = Math.min(Math.round(dim.w * 15), 9000); h = Math.round(w * dim.h / dim.w); }
    if (w > 0 && h > 0) { u16(picf, 0x1C, w); u16(picf, 0x1E, h); u16(picf, 0x20, 1000); u16(picf, 0x22, 1000); }
    return concat([picf, sp, fbse, blip, bytes], lcb);
  }

  // A PapxInFkp (cb=0 form): [0][cb'][GrpPrlAndIstd], where GrpPrlAndIstd is
  // istd (2 bytes, 0 = Normal) + grpprl, zero-padded to 2*cb' bytes.
  function papxInFkp(grpprl) {
    var body = [0, 0].concat(grpprl), cbq = Math.ceil(body.length / 2);
    var u = new Uint8Array(2 + cbq * 2); u[1] = cbq;
    for (var i = 0; i < body.length; i++) u[2 + i] = body[i] & 0xFF;
    return u;
  }
  // sprmTDefTable (0xD608) operand: itcMac, then rgdxaCenter (ncols+1 signed
  // twips, equal columns across the text width) and rgTc80 (one 20-byte cell
  // descriptor each — zero-filled = borderless). spra=6 => 2-byte length prefix.
  function tDefTableSprm(ncols) {
    var width = 9000, op = [ncols & 0xFF];
    for (var i = 0; i <= ncols; i++) { var x = Math.round(i * width / ncols) & 0xFFFF; op.push(x & 0xFF, (x >> 8) & 0xFF); }
    var brc = [6, 1, 0, 0];                 // Brc80: 3/4pt single line, auto colour, on all 4 sides
    for (i = 0; i < ncols; i++) {
      op.push(0, 0, 0, 0);                  // TC80: tcgrf=0, wWidth=0 (auto)
      for (var s = 0; s < 4; s++) op.push(brc[0], brc[1], brc[2], brc[3]); // brcTop/Left/Bottom/Right
    }
    return [0x08, 0xD6, op.length & 0xFF, (op.length >> 8) & 0xFF].concat(op);
  }
  var SPRM_FINTABLE = [0x16, 0x24, 0x01];  // sprmPFInTable = 1
  var SPRM_FTTP = [0x17, 0x24, 0x01];      // sprmPFTtp = 1 (table terminator paragraph)

  // Pack runs/paragraphs into 512-byte FKP pages. items: [{fc0, fc1, blob}],
  // contiguous (item[k].fc1 === item[k+1].fc0). blob is a Chpx ([cb][grpprl],
  // null => default) or a PapxInFkp. entrySize = bytes per run after rgfc: 1 for
  // CHPX rgb, 13 for PAPX BxPap (its first byte is the word offset; PHE stays 0).
  // Returns [{page, fc0, fc1}]; the caller places pages and builds the PlcfBte.
  function packFkp(items, entrySize) {
    function evenLen(b) { return b ? (b.length + 1) & ~1 : 0; }   // word-aligned blob span
    var LIMIT = SECTOR - 2;                                       // reserve byte 510 (gap) + 511 (crun)
    var pages = [], i = 0;
    while (i < items.length) {
      var group = [], blobBytes = 0;
      while (i < items.length) {
        var bl = evenLen(items[i].blob);
        var n = group.length + 1, overhead = 4 * (n + 1) + entrySize * n;
        if (group.length && overhead + blobBytes + bl > LIMIT) break;
        group.push(items[i]); blobBytes += bl; i++;
      }
      // Place blobs from byte 510 downward at exact even offsets (pos stays even,
      // so off>>1 is exact and no unaccounted alignment gap can overflow the page).
      var pg = new Uint8Array(SECTOR), crun = group.length, pos = SECTOR - 2, off = [];
      for (var j = crun - 1; j >= 0; j--) {
        var b = group[j].blob;
        if (b && b.length) { pos -= evenLen(b); pg.set(b, pos); off[j] = pos; } else off[j] = 0;
      }
      for (var k = 0; k <= crun; k++) u32(pg, k * 4, k < crun ? group[k].fc0 : group[crun - 1].fc1);
      var bx = 4 * (crun + 1);
      for (k = 0; k < crun; k++) pg[bx + k * entrySize] = off[k] >> 1;
      pg[SECTOR - 1] = crun;
      pages.push({ page: pg, fc0: group[0].fc0, fc1: group[crun - 1].fc1 });
    }
    return pages;
  }

  // Normalise input to model paragraphs. Accepts a plain string, an array of
  // model paragraphs, or a docToText.model() story ({ body: [...] }).
  function normRun(r) {
    if (r.image && r.image.bytes) return { image: r.image };
    if (r.ftnRef != null) return { ftnRef: r.ftnRef };   // footnote-reference anchor (no text)
    if (r.endRef != null) return { endRef: r.endRef };   // endnote-reference anchor (no text)
    if (r.comRef != null) return { comRef: r.comRef };   // comment-reference anchor (no text)
    if (r.tbxRef != null) return { tbxRef: r.tbxRef };   // text-box (drawn-object) anchor (no text)
    var n = { text: String(r.text == null ? '' : r.text), b: !!r.b, i: !!r.i, u: !!r.u, strike: !!r.strike, size: r.size || null, font: r.font || null, color: r.color == null ? null : r.color };
    if (r.url) n.url = String(r.url);
    return n;
  }
  // One model paragraph -> a normalized writer paragraph. List paragraphs map to
  // the skeleton's built-in LFOs: bullet -> ilfo 2, number -> ilfo 1.
  function mapPara(p) {
    var li = p.list, ilfo = li ? (li.kind === 'number' ? 1 : 2) : (p.ilfo || 0), ilvl = li ? (li.ilvl || 0) : (p.ilvl || 0);
    return { runs: (p.runs || []).map(normRun), kind: p.kind || 'p', align: p.align || 0, ilfo: ilfo, ilvl: ilvl, pp: p.pp || null };
  }
  // Non-body stories from a model object (the footnotes array, etc.), mapped.
  function storyParas(input, key) {
    return (input && !Array.isArray(input) && Array.isArray(input[key])) ? input[key].map(mapPara) : [];
  }
  function toParagraphs(input) {
    if (input && !Array.isArray(input) && Array.isArray(input.body)) input = input.body;
    if (Array.isArray(input)) return input.map(mapPara);
    var s = String(input == null ? '' : input).replace(/\r\n?|\n/g, '\n');
    return s.split('\n').map(function (line) { return { runs: line ? [normRun({ text: line })] : [], kind: 'p' }; });
  }

  function textToDoc(input, template) {
    var cfb = readCfb(template || defaultTemplate());
    if (!cfb) throw new Error('template is not a valid .doc (CFB) file');
    var paras = toParagraphs(input);
    var ftnParas = storyParas(input, 'footnotes');   // footnote-document paragraphs (one per footnote)
    var ftnRefCps = [];                               // body CPs of the footnote-reference chars
    var hdrParas = storyParas(input, 'header');       // page header paragraphs (odd-page)
    var ftrParas = storyParas(input, 'footer');       // page footer paragraphs (odd-page)
    var ednParas = storyParas(input, 'endnotes');     // endnote-document paragraphs (one per endnote)
    var ednRefCps = [];                               // body CPs of the endnote-reference chars
    var comParas = storyParas(input, 'annotations');  // comment (annotation) paragraphs — matches docToText.model().annotations
    if (!comParas.length) comParas = storyParas(input, 'comments'); // friendly alias
    var comRefCps = [];                               // body CPs of the comment-reference chars
    var tbxParas = storyParas(input, 'textboxes');    // text-box paragraphs (the single box's content)
    var tbxAnchorCps = [];                            // body CPs of the text-box drawn-object anchors (0x08)

    var wd = cfb.byName['WordDocument'];
    var dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    var flags = dv.getUint16(10, true), tableName = (flags >> 9) & 1 ? '1Table' : '0Table';
    var tbl = cfb.byName[tableName];
    var csw = dv.getUint16(32, true), rgLwStart = 34 + csw * 2 + 2, cslw = dv.getUint16(rgLwStart - 2, true);
    var fcLcbStart = rgLwStart + cslw * 4 + 2;
    function pairFc(i) { return dv.getUint32(fcLcbStart + i * 8, true); }
    function pairLcb(i) { return dv.getUint32(fcLcbStart + i * 8 + 4, true); }

    var T = (wd.length + 1) & ~1;                  // text offset in WordDocument (even)
    function fcAt(charIdx) { return T + charIdx * 2; }
    // PapxInFkp blobs: Normal paragraph, in-table cell, and (per row) the table
    // terminator paragraph carrying the column definition.
    var PNORMAL = papxInFkp([]);
    var PCELL = papxInFkp(SPRM_FINTABLE);
    function papxTtp(ncols) { return papxInFkp(SPRM_FINTABLE.concat(SPRM_FTTP, tDefTableSprm(ncols))); }
    // Normal paragraph PAPX with optional list membership (sprmPIlvl 0x260A +
    // sprmPIlfo 0x460B, referencing the skeleton's list tables) and alignment
    // (sprmPJc80 0x2403: 0 left / 1 centre / 2 right / 3 justify).
    function papxForP(par) {
      var g = [];
      if (par.ilfo) { g.push(0x0A, 0x26, par.ilvl & 0xFF); g.push(0x0B, 0x46, par.ilfo & 0xFF, (par.ilfo >> 8) & 0xFF); }
      if (par.align) g.push(0x03, 0x24, par.align & 0xFF);
      // Spacing/indentation (twips): each sprm carries a signed 16-bit operand
      // (first-line indent and "exact" line spacing can be negative).
      var pp = par.pp;
      if (pp) {
        function s16(lo, hi, v) { g.push(lo, hi, v & 0xFF, (v >> 8) & 0xFF); }
        if (pp.indL) s16(0x0F, 0x84, pp.indL);                 // sprmPDxaLeft
        if (pp.indR) s16(0x0E, 0x84, pp.indR);                 // sprmPDxaRight
        if (pp.ind1) s16(0x11, 0x84, pp.ind1);                 // sprmPDxaLeft1 (first line / hanging)
        if (pp.spB) s16(0x13, 0xA4, pp.spB);                   // sprmPDyaBefore
        if (pp.spA) s16(0x14, 0xA4, pp.spA);                   // sprmPDyaAfter
        if (pp.line) { g.push(0x12, 0x64, pp.line & 0xFF, (pp.line >> 8) & 0xFF, pp.lineMult & 0xFF, (pp.lineMult >> 8) & 0xFF); } // sprmPDyaLine (LSPD)
      }
      return g.length ? papxInFkp(g) : PNORMAL;
    }

    // Font table: reuse the skeleton's fonts, append any the model introduces,
    // and reference each run's font via sprmCRgFtc0. (If the skeleton has no
    // SttbfFfn, fonts are skipped and text uses the Normal style's typeface.)
    var ffnFc = pairFc(15), ffnLcb = pairLcb(15);
    var ffnInfo = (ffnLcb >= 6 && ffnFc + ffnLcb <= tbl.length) ? readFonts(tbl, ffnFc, ffnLcb) : null;
    var fontMap = {}, newFfns = [], nextFtc = ffnInfo ? ffnInfo.cData : 0;
    if (ffnInfo) ffnInfo.names.forEach(function (n, i) { if (fontMap[n.toLowerCase()] == null) fontMap[n.toLowerCase()] = i; });
    function ftcFor(name) {
      if (!ffnInfo || !name) return null;
      var key = name.toLowerCase();
      if (fontMap[key] != null) return fontMap[key];
      fontMap[key] = nextFtc; newFfns.push(buildFfn(name)); return nextFtc++;
    }

    // Build the character stream with per-run CHPX and per-paragraph PAPX. Table
    // cells end in a cell mark (0x07) and each row ends with an empty terminator
    // paragraph; every cell/terminator paragraph is flagged in-table.
    var codes = [], chpxRuns = [], papxParas = [], paraStart = 0;
    function emitRun(run) {
      if (run.ftnRef != null) return emitFootnoteRef(run);
      if (run.endRef != null) return emitEndnoteRef(run);
      if (run.comRef != null) return emitCommentRef(run);
      if (run.tbxRef != null) return emitTextboxRef(run);
      if (!run.text) return;
      if (run.url) return emitHyperlink(run);
      var s0 = codes.length;
      for (var i = 0; i < run.text.length; i++) codes.push(run.text.charCodeAt(i));
      chpxRuns.push({ s: s0, e: codes.length, blob: chpxBlob(run, ftcFor(run.font)) });
    }
    // Emit a "special" character (sprmCFSpec) — used for the footnote auto-number
    // / reference char (0x02), which editors treat as a live mark, not a glyph.
    function emitSpecial(code) {
      codes.push(code);
      var g = []; sprm(g, 0x0855, [1]); // sprmCFSpec
      var blob = new Uint8Array(g.length + 1); blob[0] = g.length; blob.set(g, 1);
      chpxRuns.push({ s: codes.length - 1, e: codes.length, blob: blob });
    }
    function emitFootnoteRef() { ftnRefCps.push(codes.length); emitSpecial(0x02); }
    function emitEndnoteRef() { ednRefCps.push(codes.length); emitSpecial(0x02); }
    function emitCommentRef() { comRefCps.push(codes.length); emitSpecial(0x05); }
    function emitTextboxRef() { tbxAnchorCps.push(codes.length); emitSpecial(0x08); }
    // Hyperlink run -> a HYPERLINK field: begin mark (0x13) + a hidden instruction
    // (HYPERLINK "addr") + separator (0x14) + the visible result text + end (0x15).
    // The begin/sep/end positions are recorded for the field PLC (PlcfFldMom) so
    // real editors treat it as a live, clickable field.
    var fieldMarks = [];
    // A field mark (0x13/0x14/0x15) is a "special" character — like the picture
    // char — so its CHPX must set sprmCFSpec, otherwise editors treat it as a
    // literal control char and show the raw instruction instead of a live field.
    function emitFieldMark(code) {
      fieldMarks.push({ cp: codes.length, kind: code });
      codes.push(code);
      var g = []; sprm(g, 0x0855, [1]); // sprmCFSpec
      var blob = new Uint8Array(g.length + 1); blob[0] = g.length; blob.set(g, 1);
      chpxRuns.push({ s: codes.length - 1, e: codes.length, blob: blob });
    }
    function emitHyperlink(run) {
      var instr = ' HYPERLINK "' + String(run.url).replace(/["\r\n]/g, '') + '" ';
      emitFieldMark(0x13);
      var bs = codes.length;
      for (var i = 0; i < instr.length; i++) codes.push(instr.charCodeAt(i));
      chpxRuns.push({ s: bs, e: codes.length, blob: null });
      emitFieldMark(0x14);
      emitRun({ text: run.text, b: run.b, i: run.i, u: run.u, strike: run.strike, size: run.size, font: run.font, color: run.color });
      emitFieldMark(0x15);
    }
    function emitMark(code) { codes.push(code); chpxRuns.push({ s: codes.length - 1, e: codes.length, blob: null }); }
    function endPara(blob) { papxParas.push({ s: paraStart, e: codes.length, blob: blob }); paraStart = codes.length; }
    // Inline pictures: append the picture bytes to the Data stream and emit a
    // 0x01 char whose CHPX points at it (sprmCPicLocation) and marks it special.
    var dataParts = [], dataLen = 0;
    function emitImage(img) {
      if (!img || !img.bytes || !img.bytes.length) return;
      var off = dataLen, part = imagePart(img); dataParts.push(part); dataLen += part.length;
      codes.push(0x01);
      var g = []; sprm(g, 0x6A03, [off, off >> 8, off >> 16, off >> 24]); sprm(g, 0x0855, [1]); // sprmCPicLocation + sprmCFSpec
      var blob = new Uint8Array(g.length + 1); blob[0] = g.length; blob.set(g, 1);
      chpxRuns.push({ s: codes.length - 1, e: codes.length, blob: blob });
    }
    var cellsInRow = 0;
    for (var pi = 0; pi < paras.length; pi++) {
      var par = paras[pi];
      for (var ri = 0; ri < par.runs.length; ri++) { var run = par.runs[ri]; if (run.image) emitImage(run.image); else emitRun(run); }
      if (par.kind === 'cell') { emitMark(0x07); endPara(PCELL); cellsInRow++; }
      else if (par.kind === 'rowEnd') {
        emitMark(0x07); endPara(PCELL); cellsInRow++;       // final cell of the row
        emitMark(0x07); endPara(papxTtp(cellsInRow));       // empty row-terminator paragraph
        cellsInRow = 0;
      } else { emitMark(0x0D); endPara(papxForP(par)); cellsInRow = 0; }  // normal paragraph (alignment + list)
    }
    if (!codes.length || codes[codes.length - 1] !== 0x0D) { emitMark(0x0D); endPara(PNORMAL); }
    var ccpText = codes.length;   // body ends here; non-body stories follow in CP space

    // ---- Footnote / endnote documents: one paragraph per note, each
    // [0x02 ref][text][0x0D] then a trailing mark. The txt PLC records boundaries
    // [0, ...ends..., ccp+2]; the ref PLC the body positions + an FRD each.
    function appendNotes(noteParas, refCps, refChar) {
      if (!noteParas.length || !refCps.length) return { txtCps: null, refCps: [] };
      var start = codes.length, n = Math.min(noteParas.length, refCps.length), txtCps = [0];
      for (var i = 0; i < n; i++) {
        emitSpecial(refChar || 0x02);                       // note auto-number / annotation mark
        var rs = noteParas[i].runs || [];
        for (var j = 0; j < rs.length; j++) if (!rs[j].image && rs[j].ftnRef == null && rs[j].endRef == null && rs[j].comRef == null) emitRun(rs[j]);
        emitMark(0x0D); endPara(PNORMAL);
        txtCps.push(codes.length - start);                  // cumulative end of note i
      }
      emitMark(0x0D); endPara(PNORMAL);                     // note-document end marker
      txtCps.push((codes.length - start) + 2);              // trailing phantom CP (observed +2)
      return { txtCps: txtCps, refCps: refCps.slice(0, n) };
    }
    var ftn = appendNotes(ftnParas, ftnRefCps); ftnRefCps = ftn.refCps; var ftnTxtCps = ftn.txtCps;
    var ccpFtn = codes.length - ccpText;

    // ---- Header document: PlcfHdd stories. 0-5 are footnote/endnote separators
    // (left empty), then per section: even/odd/first header, even/odd/first footer.
    // We place the header in the odd-page header (story 7) and footer in the odd
    // footer (story 9), all others empty, plus a trailing paragraph. ccpHdd counts
    // these; the final PlcfHdd CP carries the observed +2 phantom.
    var hddStart = codes.length, hddCps = null;
    if (hdrParas.length || ftrParas.length) {
      hddCps = [];
      function emitStoryParas(ps) { for (var i = 0; i < ps.length; i++) { var rs = ps[i].runs || []; for (var j = 0; j < rs.length; j++) if (!rs[j].image && rs[j].ftnRef == null && rs[j].endRef == null && rs[j].comRef == null) emitRun(rs[j]); emitMark(0x0D); endPara(PNORMAL); } }
      for (var hk = 0; hk < 13; hk++) {
        hddCps.push(codes.length - hddStart);          // start CP of story hk
        if (hk === 7 && hdrParas.length) emitStoryParas(hdrParas);
        else if (hk === 9 && ftrParas.length) emitStoryParas(ftrParas);
        else if (hk === 12) { emitMark(0x0D); endPara(PNORMAL); }   // trailing story
      }
      hddCps.push((codes.length - hddStart) + 2);      // lim = ccpHdd + 2 (observed phantom)
    }
    var ccpHdd = codes.length - hddStart;

    // ---- Comment (annotation) document: same shape, ref char 0x05, placed BEFORE
    // endnotes in CP order. PlcfandRef carries a 30-byte ATRD per comment.
    var atnStart = codes.length;
    var atn = appendNotes(comParas, comRefCps, 0x05); comRefCps = atn.refCps; var atnTxtCps = atn.txtCps;
    var ccpAtn = codes.length - atnStart;

    // ---- Endnote document: same shape as footnotes, after the comment document.
    var ednStart = codes.length;
    var edn = appendNotes(ednParas, ednRefCps); ednRefCps = edn.refCps; var ednTxtCps = edn.txtCps;
    var ccpEdn = codes.length - ednStart;

    // ---- Text-box document: the single box's content + an empty paragraph
    // (one region) then a trailing mark, mirroring the reference. PlcftxbxTxt
    // records [0, contentEnd, ccpTxbx]; the shape lives in TBX_DGG, anchored by a
    // body 0x08 char + an FSPA (PlcfspaMom).
    var txbxStart = codes.length, tbxTxtCps = null;
    if (tbxParas.length && tbxAnchorCps.length) {
      // Strip trailing empty paragraphs (the reader returns the box story with its
      // structural marks); we re-add a canonical empty + trailing below, so the
      // round-trip is stable rather than growing ccpTxbx each pass.
      var content = tbxParas.slice();
      while (content.length && !(content[content.length - 1].runs || []).some(function (r) { return r.text || r.image; })) content.pop();
      for (var ti = 0; ti < content.length; ti++) { var trs = content[ti].runs || []; for (var tj = 0; tj < trs.length; tj++) if (!trs[tj].image && trs[tj].tbxRef == null) emitRun(trs[tj]); emitMark(0x0D); endPara(PNORMAL); }
      emitMark(0x0D); endPara(PNORMAL);                 // empty paragraph closes the box content
      var tbxContentEnd = codes.length - txbxStart;
      emitMark(0x0D); endPara(PNORMAL);                 // trailing mark
      tbxTxtCps = [0, tbxContentEnd, codes.length - txbxStart];
      tbxAnchorCps = tbxAnchorCps.slice(0, 1);          // the embedded drawing holds one shape
    } else { tbxAnchorCps = []; }
    var ccpTxbx = codes.length - txbxStart;
    var ccp = codes.length, textBytes = ccp * 2;

    // ---- new WordDocument: skeleton WD + text + CHPX page(s) + PAPX page(s) --
    function pad512(n) { return (n + 511) & ~511; }
    var parts = [wd], totalLen = wd.length;
    function placeAt(off, arr) { while (totalLen < off) { parts.push(new Uint8Array(off - totalLen)); totalLen = off; } parts.push(arr); totalLen += arr.length; return off; }
    var textBuf = new Uint8Array(textBytes);
    for (var c = 0; c < ccp; c++) u16(textBuf, c * 2, codes[c]);
    placeAt(T, textBuf);
    function placePages(items, entrySize) {
      var pages = packFkp(items, entrySize), pns = [], fcs = [];
      pages.forEach(function (pg) { var pn = pad512(totalLen) / 512; placeAt(pn * 512, pg.page); pns.push(pn); fcs.push(pg.fc0); });
      fcs.push(pages[pages.length - 1].fc1);
      return { pns: pns, fcs: fcs };
    }
    var chpx = placePages(chpxRuns.map(function (r) { return { fc0: fcAt(r.s), fc1: fcAt(r.e), blob: r.blob }; }), 1);
    var papx = placePages(papxParas.map(function (p) { return { fc0: fcAt(p.s), fc1: fcAt(p.e), blob: p.blob }; }), 13);
    var newWd = concat(parts, totalLen);
    u32(newWd, rgLwStart + 3 * 4, ccpText);        // ccpText (body only)
    u32(newWd, rgLwStart + 4 * 4, ccpFtn);         // ccpFtn (footnote document)
    u32(newWd, rgLwStart + 5 * 4, ccpHdd);         // ccpHdd (header document)
    u32(newWd, rgLwStart + 7 * 4, ccpAtn);         // ccpAtn (comment document)
    u32(newWd, rgLwStart + 8 * 4, ccpEdn);         // ccpEdn (endnote document)
    u32(newWd, rgLwStart + 9 * 4, ccpTxbx);        // ccpTxbx (text-box document)
    u32(newWd, rgLwStart + 0 * 4, newWd.length);   // cbMac

    // ---- new 1Table: skeleton table + appended CLX / PlcfBteChpx / PlcfBtePapx
    // (+ a grown SttbfFfn when the model introduced new fonts) ----------------
    var clx = new Uint8Array(21);
    clx[0] = 0x02; u32(clx, 1, 16); u32(clx, 5, 0); u32(clx, 9, ccp); u16(clx, 13, 0); u32(clx, 15, T); u16(clx, 19, 0);
    function plcfBte(p) { var b = new Uint8Array(p.fcs.length * 4 + p.pns.length * 4); for (var i = 0; i < p.fcs.length; i++) u32(b, i * 4, p.fcs[i]); for (i = 0; i < p.pns.length; i++) u32(b, p.fcs.length * 4 + i * 4, p.pns[i]); return b; }
    // PlcfFldMom: a PLC over the field-mark CPs (begin/sep/end) with a 2-byte Fld
    // each. Fld byte 1 mirrors a real editor: begin -> flt 0x58 (HYPERLINK),
    // separator -> 0xFF, end -> 0x84 (fHasSep | fResultDirty). Without this PLC,
    // editors show the raw "HYPERLINK ..." instruction instead of a live link.
    function plcfldBytes(marks) {
      if (!marks.length) return null;
      var n = marks.length, b = new Uint8Array((n + 1) * 4 + n * 2);
      for (var i = 0; i < n; i++) u32(b, i * 4, marks[i].cp);
      u32(b, n * 4, marks[n - 1].cp + 1);
      var fb = (n + 1) * 4;
      for (i = 0; i < n; i++) { var k = marks[i].kind; b[fb + i * 2] = k; b[fb + i * 2 + 1] = k === 0x13 ? 0x58 : (k === 0x14 ? 0xFF : 0x84); }
      return b;
    }
    var newFfn = null;
    if (ffnInfo && newFfns.length) {
      var head = ffnInfo.bytes.slice(0, ffnInfo.used);
      var extra = newFfns.reduce(function (a, f) { return a + f.length; }, 0);
      newFfn = concat([head].concat(newFfns), head.length + extra);
      u16(newFfn, ffnInfo.cDataOff, ffnInfo.cData + newFfns.length);   // bump cData
    }
    var blocks = [], off = tbl.length;
    function append(bytes) { var o = off; blocks.push(bytes); off += bytes.length; return o; }
    var clxOff = append(clx), pbteCOff = append(plcfBte(chpx)), pbtePOff = append(plcfBte(papx));
    var pbteC = blocks[1], pbteP = blocks[2];
    var ffnOff = newFfn ? append(newFfn) : -1;
    // A text box needs the floating-shape drawing (TBX_DGG, which also carries the
    // group defaults pictures rely on); otherwise pictures use their own defaults.
    var dggBytes = tbxAnchorCps.length ? b64ToU8(TBX_DGG) : (dataLen ? picParts().dgg : null);
    var dggOff = dggBytes ? append(dggBytes) : -1;   // OfficeArt drawing (DggInfo)
    var fldBytes = plcfldBytes(fieldMarks);
    var fldOff = fldBytes ? append(fldBytes) : -1;   // PlcfFldMom (hyperlink field positions)
    // Note PLCs (footnotes #2/#3, endnotes #46/#47): refBytes = N body ref CPs +
    // doc-end lim + N FRDs (nAuto); txtBytes = the note-document text boundaries.
    var docEndLim = ccp + 1; // one past the last CP across all stories
    function refBytesOf(refCps) {
      var nF = refCps.length, b = new Uint8Array((nF + 1) * 4 + nF * 2);
      for (var i = 0; i < nF; i++) u32(b, i * 4, refCps[i]);
      u32(b, nF * 4, docEndLim);
      for (i = 0; i < nF; i++) u16(b, (nF + 1) * 4 + i * 2, i + 1);
      return b;
    }
    function txtBytesOf(txtCps) { var b = new Uint8Array(txtCps.length * 4); for (var i = 0; i < txtCps.length; i++) u32(b, i * 4, txtCps[i]); return b; }
    // PlcfandRef (#4) carries a 30-byte ATRD per comment: xstUsrInitials (Xst, cch
    // + up to 10 UTF-16 chars) + ibst (owner index) + 6 bytes (ak/grfbmc/lTagBkmk).
    function atrdRefBytes(refCps) {
      var nF = refCps.length, b = new Uint8Array((nF + 1) * 4 + nF * 30), base = (nF + 1) * 4;
      for (var i = 0; i < nF; i++) u32(b, i * 4, refCps[i]);
      u32(b, nF * 4, docEndLim);
      for (i = 0; i < nF; i++) { var o = base + i * 30; u16(b, o, 1); u16(b, o + 2, 0x43); u32(b, o + 26, 0xFFFFFFFF); } // initials "C", ibst 0, lTagBkmk -1
      return b;
    }
    var ftnRefBytes = (ftnRefCps.length && ftnTxtCps) ? refBytesOf(ftnRefCps) : null;
    var ftnTxtBytes = (ftnRefCps.length && ftnTxtCps) ? txtBytesOf(ftnTxtCps) : null;
    var ednRefBytes = (ednRefCps.length && ednTxtCps) ? refBytesOf(ednRefCps) : null;
    var ednTxtBytes = (ednRefCps.length && ednTxtCps) ? txtBytesOf(ednTxtCps) : null;
    var atnRefBytes = (comRefCps.length && atnTxtCps) ? atrdRefBytes(comRefCps) : null;
    var atnTxtBytes = (comRefCps.length && atnTxtCps) ? txtBytesOf(atnTxtCps) : null;
    var ftnRefOff = ftnRefBytes ? append(ftnRefBytes) : -1;
    var ftnTxtOff = ftnTxtBytes ? append(ftnTxtBytes) : -1;
    var ednRefOff = ednRefBytes ? append(ednRefBytes) : -1;
    var ednTxtOff = ednTxtBytes ? append(ednTxtBytes) : -1;
    var atnRefOff = atnRefBytes ? append(atnRefBytes) : -1;
    var atnTxtOff = atnTxtBytes ? append(atnTxtBytes) : -1;
    // PlcfHdd (#11): the header-document story boundaries (CP array, no data).
    var hddBytes = null;
    if (hddCps) { hddBytes = new Uint8Array(hddCps.length * 4); for (var hi = 0; hi < hddCps.length; hi++) u32(hddBytes, hi * 4, hddCps[hi]); }
    var hddOff = hddBytes ? append(hddBytes) : -1;
    // Text-box PLCs: PlcfspaMom (#40) = anchor CP + doc-end lim + one FSPA (the
    // shape id + bounding box); PlcftxbxTxt (#56) = text boundaries + FTXBXS
    // (lid = shape id); PlcfTxbxBkd (#75) = break descriptors. Values mirror the
    // one-box reference; the shape itself is the embedded TBX_DGG (spid 1025).
    var spaOff = -1, txbTxtOff = -1, txbBkdOff = -1, spaLen = 0, txbTxtLen = 0, txbBkdLen = 0;
    if (tbxAnchorCps.length && tbxTxtCps) {
      var fspa = new Uint8Array(26);
      u32(fspa, 0, 1025); u32(fspa, 4, 1746); u32(fspa, 8, 11462); u32(fspa, 12, 4014); u32(fspa, 16, 12074); u16(fspa, 20, 0x4a); u32(fspa, 22, 1);
      var spa = new Uint8Array(8 + 26); u32(spa, 0, tbxAnchorCps[0]); u32(spa, 4, docEndLim); spa.set(fspa, 8);
      spaOff = append(spa); spaLen = spa.length;
      var nc = tbxTxtCps.length, txt = new Uint8Array(nc * 4 + 2 * 22);
      for (var ci = 0; ci < nc; ci++) u32(txt, ci * 4, tbxTxtCps[ci]);
      var fb0 = nc * 4; u32(txt, fb0, 1); u32(txt, fb0 + 10, 0xFFFFFFFF); u32(txt, fb0 + 14, 1025); // FTXBXS[0]: cTxbx, reserved=-1, lid
      u32(txt, fb0 + 22, 0xFFFFFFFF);                                                               // FTXBXS[1]: trailing
      txbTxtOff = append(txt); txbTxtLen = txt.length;
      var bkd = new Uint8Array(nc * 4 + 2 * 6);
      for (ci = 0; ci < nc; ci++) u32(bkd, ci * 4, tbxTxtCps[ci]);
      var bb = nc * 4; bkd[bb + 5] = 0x08; bkd[bb + 6] = 0xFF; bkd[bb + 7] = 0xFF;                  // BKD[0]=..0x0800, BKD[1]=0xffff..
      txbBkdOff = append(bkd); txbBkdLen = bkd.length;
    }
    var newTbl = concat([tbl].concat(blocks), off);
    // extend the section table's last CP so the section spans the whole text
    var sedFc = pairFc(6), sedLcb = pairLcb(6);
    if (sedLcb >= 8) u32(newTbl, sedFc + ((sedLcb - 4) / 16) * 4, ccpText);
    function setPair(idx, fc, lcb) { u32(newWd, fcLcbStart + idx * 8, fc); u32(newWd, fcLcbStart + idx * 8 + 4, lcb); }
    setPair(33, clxOff, clx.length);     // Clx
    setPair(12, pbteCOff, pbteC.length); // PlcfBteChpx
    setPair(13, pbtePOff, pbteP.length); // PlcfBtePapx
    if (ffnOff >= 0) setPair(15, ffnOff, newFfn.length);  // SttbfFfn (grown)
    if (dggOff >= 0) setPair(50, dggOff, dggBytes.length); // fcDggInfo (drawing group)
    if (spaOff >= 0) setPair(40, spaOff, spaLen);     // PlcfspaMom (shape anchors)
    if (txbTxtOff >= 0) setPair(56, txbTxtOff, txbTxtLen); // PlcftxbxTxt
    if (txbBkdOff >= 0) setPair(75, txbBkdOff, txbBkdLen); // PlcfTxbxBkd
    if (fldOff >= 0) setPair(16, fldOff, fldBytes.length); // PlcfFldMom (hyperlink fields)
    if (ftnRefOff >= 0) setPair(2, ftnRefOff, ftnRefBytes.length); // PlcffndRef
    if (ftnTxtOff >= 0) setPair(3, ftnTxtOff, ftnTxtBytes.length); // PlcffndTxt
    if (ednRefOff >= 0) setPair(46, ednRefOff, ednRefBytes.length); // PlcfendRef
    if (ednTxtOff >= 0) setPair(47, ednTxtOff, ednTxtBytes.length); // PlcfendTxt
    if (atnRefOff >= 0) setPair(4, atnRefOff, atnRefBytes.length); // PlcfandRef
    if (atnTxtOff >= 0) setPair(5, atnTxtOff, atnTxtBytes.length); // PlcfandTxt
    if (hddOff >= 0) setPair(11, hddOff, hddBytes.length); // PlcfHdd (headers/footers)

    var streams = [{ name: 'WordDocument', data: newWd }, { name: tableName, data: newTbl }];
    if (dataLen) streams.push({ name: 'Data', data: concat(dataParts, dataLen) });
    return buildCfb(streams);
  }

  function concat(arrs, len) { var out = new Uint8Array(len), p = 0; arrs.forEach(function (a) { out.set(a, p); p += a.length; }); return out; }

  textToDoc.readCfb = readCfb;   // exposed for tests/tools
  textToDoc.buildCfb = buildCfb; // exposed for tooling (building the skeleton)

  // Image-independent inline-picture machinery, reverse-engineered from a
  // TextMaker-saved reference (see scripts/embed-template.js notes): the PICF,
  // the OfficeArt SpContainer, the FBSE header+fields, the blip header
  // (msofbtBlipPNG + rgbUid + tag), and the drawing-group defaults (DggInfo).
  var PIC_PICF = "JLwAAEQAZADAA8ADAAAAAAAAAAAAAAAAAAAAAEA4QDiqAKoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";
  var PIC_SP = "DwAE8N4AAACyBArwCAAAAAEEAAAACgAA8wEL8LoAAAAEAAAAAAB/AIABwAGBAAAAAACCAAAAAACDAAAAAACEAAAAAACFAAAAAACHAAAAAACIAAAAAACJAAAAAACKAAAAAAC/AAAACgAAAQAAAAABAQAAAAACAQAAAAADAQAAAAAEQQEAAAAIAQAAAQAJAQAAAAAKAQAAAQA/AQAABgC/AQAAEAD/AQAACAA/AgAAAgC/AgAACAA/AwgACACEAwAAAACFAwAAAACGAwAAAACHAwAAAAC/AyEAIwAAABDwBAAAAAAAAIA=";
  var PIC_FBSE = "YgAH8PK6AAAGBpMLHgABnC5BlDk91UJLk7H/AM66AAABAAAARAAAAAAAAAA=";
  var PIC_BLIP = "AG4e8Ma6AACTCx4AAZwuQZQ5PdVCS5Ox/w==";
  var PIC_DGG = "DwAA8HQBAAAAAAbwGAAAAAEEAAACAAAAAQAAAAEAAAABAAAAAQAAAOMCC/A0AQAABAAAAAAAfwAAAMABgQDoigAAggDoigAAgwDoigAAhADoigAAhQAAAAAAhwAAAAAAiAAAAAAAiQAAAAAAigAAAAAAvwAAAAoAwgACAAAAwwAAACQAxAAAAAEAxcAgAAAA/wAAR///PwEAAAAAfwEAAAAAgAEAAAAAgQH///8AggEAAAEAgwEAAAAAhAEAAAEAiwEAAFoAvwEQABAAwAEAAAAAywGcMQAAzQEAAAAAzgEAAAAA0AEAAAAA0QEAAAAA0gEBAAAA0wEBAAAA1AEBAAAA1QEBAAAA/wEIAAgAPwIAAAIAvwIAAAgA/wIAAAAAPwMAAAAAhAO/XQEAhQO/XQEAhgO/XQEAhwO/XQEAvwMBACMAVABpAG0AZQBzACAATgBlAHcAIABSAG8AbQBhAG4AAABAAB7xEAAAAP//AAAAAP8AgICAAPcAABAADwAC8BABAAAQAAjwCAAAAAEAAAAABAAADwAD8DAAAAAPAATwKAAAAAEACfAQAAAAAAAAAAAAAAAAAAAAAAAAAAIACvAIAAAAAAQAAAUAAAAPAATwwAAAABIACvAIAAAAAQQAAAAMAADDAQvwqAAAAH8AAADAAYEA6IoAAIIA6IoAAIMA6IoAAIQA6IoAAIUAAAAAAIcAAAAAAIgAAAAAAIkAAAAAAIoAAAAAAL8AAAAKAIABAAAAAIEB////AIIBAAABAIMBAAAAAIQBAAABAIsBAAC0AL8BEAAQAP8BAAAIAD8CAAACAL8CAAAIAAQDCQAAAD8DAQABAIQDv10BAIUDv10BAIYDv10BAIcDv10BAL8DAQAjAA==";
  // Floating text-box drawing (DggInfo): group defaults + one TextBox shape
  // (spid 1025, ClientTextbox lTxid 0x10000). Reverse-engineered from a one-box
  // reference (samples/detailed-sample.doc). The shape is fixed; the box's text
  // lives in the textbox story (PlcftxbxTxt) and the body 0x08 anchor + FSPA
  // (PlcfspaMom) tie it to a CP. Used verbatim when the model has a text box.
  var TBX_DGG = "DwAA8HQBAAAAAAbwGAAAAAIEAAACAAAAAgAAAAEAAAABAAAAAgAAAOMCC/A0AQAABAAAAAAAfwAAAMABgQDoigAAggDoigAAgwDoigAAhADoigAAhQAAAAAAhwAAAAAAiAAAAAAAiQAAAAAAigAAAAAAvwACAAoAwgACAAAAwwAAACQAxAAAAAEAxcAgAAAA/wAAR///PwEAAAAAfwEAAAAAgAEAAAAAgQH///8AggEAAAEAgwEAAAAAhAEAAAEAiwEAAFoAvwEQABAAwAEAAAAAywGcMQAAzQEAAAAAzgEAAAAA0AEAAAAA0QEAAAAA0gEBAAAA0wEBAAAA1AEBAAAA1QEBAAAA/wEIAAgAPwIAAAIAvwIAAAgA/wIAAAAAPwMAAAAAhAO/XQEAhQO/XQEAhgO/XQEAhwO/XQEAvwMBACMAVABpAG0AZQBzACAATgBlAHcAIABSAG8AbQBhAG4AAABAAB7xEAAAAP//AAAAAP8AgICAAPcAABAADwAC8FoBAAAQAAjwCAAAAAIAAAABBAAADwAD8EIBAAAPAATwKAAAAAEACfAQAAAAAAAAAAAAAAAAAAAAAAAAAAIACvAIAAAAAAQAAAUAAAAPAATwCgEAAKIMCvAIAAAAAQQAAAAKAABzAQvwqAAAAAQAAAAAAH8AAADAAYAAAAABAIEA6IoAAIIA6IoAAIMA6IoAAIQA6IoAAIUAAAAAAIcAAAAAAIgAAAAAAIkAAAAAAIoAAAAAAL8AAgAKAL8BAAAQAP8BAAAIAD8CAAACAL8CAAAIAIDDHgAAAIQDv10BAIUDv10BAIYDv10BAIcDv10BAL8DAQAjAFAAbwBsAGUAIAB0AGUAawBzAHQAbwB3AGUAMQAAAFMAIvEeAAAAjwMAAAAAkAMBAAAAkQMAAAAAkgMBAAAAvwMAAACAAAAQ8AQAAAAAAAAAAAAR8AQAAAABAAAAAAAN8AQAAAAAAAEA";

  // Bundled blank-document skeleton (a real, app-saved empty .doc, stripped to
  // its WordDocument + 1Table streams) that textToDoc() injects text into by
  // default. Structural/empty — carries no document content. See README.
  var TEMPLATE_B64 = "0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7/CQAGAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAEAAA/v///wAAAAD+////AAAAAAAAAAD////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9/////v///wMAAAAEAAAABQAAAAYAAAAHAAAACAAAAAkAAAAKAAAACwAAAAwAAAANAAAA/v///w8AAAAQAAAAEQAAABIAAAATAAAAFAAAABUAAAD+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUB//////////8BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAMQBUAGEAYgBsAGUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAgH/////AgAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAABgAAAAAAABXAG8AcgBkAEQAbwBjAHUAbQBlAG4AdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGgACAf///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQgQbABIAAQALAQ8ABwAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYgAAAPH/AgBiAAwEAAAAAAAAAAAGAE4AbwByAG0AYQBsAAAACwAAACokATEkAIAkAAAwAENKGABLSAEAT0oFAFBKBgBRSgUAbUgJBHNICQRuSAQIdEgECF9IOQReSgcAYUoYAF4AAQDxAAIBXgAMBAAAAAAAAAAACQBIAGUAYQBkAGkAbgBnACAAMQAAACQAAQAKJgALRgEADcYFAAEYAwAPhBgDEYRQ/kAmAF6EGANghFD+DgA1CAFDSiQAYUokAFwIAGIAAgDxAAIBYgAMBAAAAAAAAAAACQBIAGUAYQBkAGkAbgBnACAAMgAAACgAAgAKJgELRgEADcYFAAGoAwAPhKgDEYTA/ROkyABAJgFehKgDYITA/Q4ANQgBQ0ogAGFKIABcCABkAAMA8QACAWQADAQAAAAAAAAAAAkASABlAGEAZABpAG4AZwAgADMAAAAoAAMACiYCC0YBAA3GBQABOAQAD4Q4BBGEMP0TpIwAQCYCXoQ4BGCEMP0PADUIAUIqD3Bof39/AFwIAAAAAAAAAAAAAAAAAABEAEFA8v+hAEQADAQAAAAAAAAAABYARABlAGYAYQB1AGwAdAAgAFAAYQByAGEAZwByAGEAcABoACAARgBvAG4AdAAAAAAAAAAAAAAAAABGAP4PAQACAUYADAQAAAAAAAAAAAcASABlAGEAZABpAG4AZwAAAA0ADwAGJAETpPAAFKR4AAAQAENKHABPSggAUUoIAGFKHAA4AEIAAQACATgADAQAAAAAAAAAAAkAQgBvAGQAeQAgAFQAZQB4AHQAAAAMABAAEmQgAQEAFKSMAAAAJAAvAAEBEgEkAAwEAAAAAAAAAAAEAEwAaQBzAHQAAAACABEAAAA8ACIAAQAiATwADAQAAAAAAAAAAAcAQwBhAHAAdABpAG8AbgAAAA0AEgAMJAETpHgAFKR4AAAGADYIAV0IACoA/g8BADIBKgAMBAAAAAAAAAAABQBJAG4AZABlAHgAAAAFABMADCQBAAAARAD+DwEAQgFEAAwEAAAAAAAAAAAKAFEAdQBvAHQAYQB0AGkAbwBuAHMAAAAWABQADoQ3Ag+ENwIUpBsBXYQ3Al6ENwIAADoAPgDxAAIBOgAMBAAAAAAAAAAABQBUAGkAdABsAGUAAAAIABUAAyQBYSQBDgA1CAFDSjgAYUo4AFwIAD4ASgDxAAIBPgAMBAAAAAAAAAAACABTAHUAYgB0AGkAdABsAGUAAAAMABYAAyQBE6Q8AGEkAQgAQ0okAGFKJAA8AP4PAQByATwADAQAAAAAAAAAAA4AVABhAGIAbABlACAAQwBvAG4AdABlAG4AdABzAAAABQAXAAwkAQAAADYA/k+iAIEBNgAMBAAAAAAAAAAABwBCAHUAbABsAGUAdABzAAAAEABPSgQAUUoEAFBKBABeSgQASgBVQKIAkQFKAAwEAAAAAAAAAAAJAEgAeQBwAGUAcgBsAGkAbgBrAAAAIAA+KgFCKgltSP8Ac0j/AHBoAAB/AG5I/wB0SP8AX0j/AFoAVkCiAKEBWgAMBAAAAAAAAAAAEQBGAG8AbABsAG8AdwBlAGQASAB5AHAAZQByAGwAaQBuAGsAAAAgAD4qAUIqDW1I/wBzSP8AcGh/AAAAbkj/AHRI/wBfSP8AAAAAAAEAAAAEAAAOAAAAAP////8ACAAAAggAAAUAAAAACAAAAggAAAYAAAAPAADwdAEAAAAABvAYAAAAAQQAAAIAAAABAAAAAQAAAAEAAAABAAAA4wIL8DQBAAAEAAAAAAB/AAAAwAGBAOiKAACCAOiKAACDAOiKAACEAOiKAACFAAAAAACHAAAAAACIAAAAAACJAAAAAACKAAAAAAC/AAAACgDCAAIAAADDAAAAJADEAAAAAQDFwCAAAAD/AABH//8/AQAAAAB/AQAAAACAAQAAAACBAf///wCCAQAAAQCDAQAAAACEAQAAAQCLAQAAWgC/ARAAEADAAQAAAADLAZwxAADNAQAAAADOAQAAAADQAQAAAADRAQAAAADSAQEAAADTAQEAAADUAQEAAADVAQEAAAD/AQgACAA/AgAAAgC/AgAACAD/AgAAAAA/AwAAAACEA79dAQCFA79dAQCGA79dAQCHA79dAQC/AwEAIwBUAGkAbQBlAHMAIABOAGUAdwAgAFIAbwBtAGEAbgAAAEAAHvEQAAAA//8AAAAA/wCAgIAA9wAAEAAPAALwEAEAABAACPAIAAAAAQAAAAAEAAAPAAPwMAAAAA8ABPAoAAAAAQAJ8BAAAAAAAAAAAAAAAAAAAAAAAAAAAgAK8AgAAAAABAAABQAAAA8ABPDAAAAAEgAK8AgAAAABBAAAAAwAAMMBC/CoAAAAfwAAAMABgQDoigAAggDoigAAgwDoigAAhADoigAAhQAAAAAAhwAAAAAAiAAAAAAAiQAAAAAAigAAAAAAvwAAAAoAgAEAAAAAgQH///8AggEAAAEAgwEAAAAAhAEAAAEAiwEAALQAvwEQABAA/wEAAAgAPwIAAAIAvwIAAAgABAMJAAAAPwMBAAEAhAO/XQEAhQO/XQEAhgO/XQEAhwO/XQEAvwMBACMAAgBfxTFqX8Uxav8P/w//D/8P/w//D/8P/w//DwAAYMUxamDFMWr/D/8P/w//D/8P/w//D/8P/w8AAAEAAAD/AAAAAAAAAAAAAAIAAAAAAAAAAAMIAAAPhGgBEYQAAFMqAAAAAQAAAP8AAAAAAAAAAAAAAgAAAAAAAAAAAwgBAA+EaAERhAAAUyoAAAABAAAA/wAAAAAAAAAAAAACAAAAAAAAAAADCAIAD4RoARGEAABTKgAAAAEAAAD/AAAAAAAAAAAAAAIAAAAAAAAAAAMIAwAPhGgBEYQAAFMqAAAAAQAAAP8AAAAAAAAAAAAAAgAAAAAAAAAAAwgEAA+EaAERhAAAUyoAAAABAAAA/wAAAAAAAAAAAAACAAAAAAAAAAADCAUAD4RoARGEAABTKgAAAAEAAAD/AAAAAAAAAAAAAAIAAAAAAAAAAAMIBgAPhGgBEYQAAFMqAAAAAQAAAP8AAAAAAAAAAAAAAgAAAAAAAAAAAwgHAA+EaAERhAAAUyoAAAABAAAA/wAAAAAAAAAAAAACAAAAAAAAAAADCAgAD4RoARGEAABTKgAAAAEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIAAAPhGgBEYQAAFMqAE9KAQBRSgEAAQC38AEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIAQAPhNACEYQAAFMqAE9KBABRSgQAAQDmJQEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIAgAPhDgEEYQAAFMqAE9KBABRSgQAAQCqJQEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIAwAPhKAFEYQAAFMqAE9KAQBRSgEAAQC38AEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIBAAPhAgHEYQAAFMqAE9KBABRSgQAAQDmJQEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIBQAPhHAIEYQAAFMqAE9KBABRSgQAAQCqJQEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIBgAPhNgJEYQAAFMqAE9KAQBRSgEAAQC38AEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsIBwAPhEALEYQAAFMqAE9KBABRSgQAAQDmJQEAAAAXAAAAAAAAAAAAAAAAAAAAAAAAAAsICAAPhKgMEYQAAFMqAE9KBABRSgQAAQCqJQIAAABfxTFqAAAAAAAAAAAAAAAAYMUxagAAAAAAAAAAAAAAAP////////////8CAAAAEgBMAGkAcwB0AGEAIABuAHUAbQBlAHIAbwB3AGEAbgBhACAAMQASAEwAaQBzAHQAYQAgAG4AdQBtAGUAcgBvAHcAYQBuAGEAIAAyAAIQAAAAAAAAAAEAAAAAAAAIAAAAAAsAAABHFpAB7gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVABpAG0AZQBzACAATgBlAHcAIABSAG8AbQBhAG4AAAA1FpABAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUwB5AG0AYgBvAGwAAAAzJpAB7gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQQByAGkAYQBsAAAANQaQAQAAAgEGAAMBAQEBAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFMAaQBtAFMAdQBuAAAAPQaQAQAAAgsGBAICAgICBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAE8AcABlAG4AUwB5AG0AYgBvAGwAAABJFpABAAACAgYDBQQFAgMEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATABpAGIAZQByAGEAdABpAG8AbgAgAFMAZQByAGkAZgAAAE8GkAEAAAILBgQCAgICAgQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEAHIAbwBpAGQAIABTAGEAbgBzACAARgBhAGwAbABiAGEAYwBrAAAAOQaQAQAAAgsGBAICAgICBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEYAcgBlAGUAUwBhAG4AcwAAAEcmkAEAAAILBgQCAgICAgQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABMAGkAYgBlAHIAYQB0AGkAbwBuACAAUwBhAG4AcwAAAD8mkAEAAAILBgQCAgICAgQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABEAGUAagBhAFYAdQAgAFMAYQBuAHMAAAA7BpABAAACCwYEAgICAgIEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATwBwAGUAbgAgAFMAYQBuAHMAAAACAAQAAQiNGAAAxQIAAKkBAAAAAMkSWGfzheZHAAABAAIAAAAAAAAAAAAAAAAAAQABAAAABACDkAEAAAAAAAAAAAAAAAEAAQAAAAEAAAAAAAAAIYMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADwAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA0gQAAAAAAAAAAf////8H/////wf/////BAIAAAABAQD/////LwwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbAQAAGwEAAQAACCwAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAQDrvuu+urq6ugAABAD/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//xIAAAAAAAAAAAAAAAAAAAAAAAUAUgBvAG0AYQBuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7KUBAVVACQQAAAASvwAAAAAAADAAAAAAAAgAAAIIAAAOAFRNIDIwMTAgAAAAAAAAAAAAAAAAAAAAAAAACQQWACoOAABfxTFqX8UxagEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8PAAUAAAABAAAA//8PAAYAAAABAAAA//8PAAAAAAAAAAAAAAAAAAAAAACIAAAAAAAeCQAAAAAAAB4JAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4JAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIJAAAMAAAAPgkAAAwAAAAAAAAAAAAAAOgPAADCAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAqhIAAFgCAADyFQAANAAAANMPAAAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABKCQAAlQIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3wsAAHYDAABVDwAALAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgQ8AAFIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACCAAA8QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHDUIgUIqAUNKFQBPSgkAUUoJAHBoAAAAAF5KCQABAAgAAAIIAAD9AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARUAAAEoABew0AIYsNACH7CCLiCwxkEhsG4EIrBuBCOQbgQkkG4EJbAAAEuw/v8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";
  return textToDoc;
});
