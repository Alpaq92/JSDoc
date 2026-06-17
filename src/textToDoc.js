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
 * Writes back: paragraphs, character formatting (bold/italic/underline/strike/
 * size/colour as CHPX sprms) and tables (cell marks + sprmPFInTable / sprmPFTtp
 * / sprmTDefTable with borders). Not yet: per-run fonts (text uses the Normal
 * style's font) or embedded images.
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
  function chpxGrpprl(r) {
    var b = [];
    if (r.b) sprm(b, 0x0835, [1]);            // sprmCFBold
    if (r.i) sprm(b, 0x0836, [1]);            // sprmCFItalic
    if (r.strike) sprm(b, 0x0837, [1]);       // sprmCFStrike
    if (r.u) sprm(b, 0x2A3E, [1]);            // sprmCKul = 1 (single underline)
    if (r.size) { var hp = Math.round(r.size * 2); sprm(b, 0x4A43, [hp, hp >> 8]); }      // sprmCHps (half-points)
    if (r.color != null) sprm(b, 0x6870, [r.color, r.color >> 8, r.color >> 16, 0]);     // sprmCCv (COLORREF)
    return b;
  }
  // A Chpx in an FKP: [cb][grpprl]. null => use the FKP default (rgb = 0).
  function chpxBlob(r) { var g = chpxGrpprl(r); if (!g.length) return null; var u = new Uint8Array(g.length + 1); u[0] = g.length; u.set(g, 1); return u; }

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
    return { text: String(r.text == null ? '' : r.text), b: !!r.b, i: !!r.i, u: !!r.u, strike: !!r.strike, size: r.size || null, font: r.font || null, color: r.color == null ? null : r.color };
  }
  function toParagraphs(input) {
    if (input && !Array.isArray(input) && Array.isArray(input.body)) input = input.body;
    if (Array.isArray(input)) return input.map(function (p) { return { runs: (p.runs || []).map(normRun), kind: p.kind || 'p' }; });
    var s = String(input == null ? '' : input).replace(/\r\n?|\n/g, '\n');
    return s.split('\n').map(function (line) { return { runs: line ? [normRun({ text: line })] : [], kind: 'p' }; });
  }

  function textToDoc(input, template) {
    var cfb = readCfb(template || defaultTemplate());
    if (!cfb) throw new Error('template is not a valid .doc (CFB) file');
    var paras = toParagraphs(input);

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
    // Build the character stream with per-run CHPX and per-paragraph PAPX. Table
    // cells end in a cell mark (0x07) and each row ends with an empty terminator
    // paragraph; every cell/terminator paragraph is flagged in-table.
    var codes = [], chpxRuns = [], papxParas = [], paraStart = 0;
    function emitRun(run) {
      if (!run.text) return;
      var s0 = codes.length;
      for (var i = 0; i < run.text.length; i++) codes.push(run.text.charCodeAt(i));
      chpxRuns.push({ s: s0, e: codes.length, blob: chpxBlob(run) });
    }
    function emitMark(code) { codes.push(code); chpxRuns.push({ s: codes.length - 1, e: codes.length, blob: null }); }
    function endPara(blob) { papxParas.push({ s: paraStart, e: codes.length, blob: blob }); paraStart = codes.length; }
    var cellsInRow = 0;
    for (var pi = 0; pi < paras.length; pi++) {
      var par = paras[pi];
      for (var ri = 0; ri < par.runs.length; ri++) emitRun(par.runs[ri]);
      if (par.kind === 'cell') { emitMark(0x07); endPara(PCELL); cellsInRow++; }
      else if (par.kind === 'rowEnd') {
        emitMark(0x07); endPara(PCELL); cellsInRow++;       // final cell of the row
        emitMark(0x07); endPara(papxTtp(cellsInRow));       // empty row-terminator paragraph
        cellsInRow = 0;
      } else { emitMark(0x0D); endPara(PNORMAL); cellsInRow = 0; }  // normal paragraph
    }
    if (!codes.length || codes[codes.length - 1] !== 0x0D) { emitMark(0x0D); endPara(PNORMAL); }
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
    u32(newWd, rgLwStart + 3 * 4, ccp);            // ccpText
    u32(newWd, rgLwStart + 0 * 4, newWd.length);   // cbMac

    // ---- new 1Table: skeleton table + CLX + PlcfBteChpx + PlcfBtePapx -------
    var clx = new Uint8Array(21);
    clx[0] = 0x02; u32(clx, 1, 16); u32(clx, 5, 0); u32(clx, 9, ccp); u16(clx, 13, 0); u32(clx, 15, T); u16(clx, 19, 0);
    function plcfBte(p) { var b = new Uint8Array(p.fcs.length * 4 + p.pns.length * 4); for (var i = 0; i < p.fcs.length; i++) u32(b, i * 4, p.fcs[i]); for (i = 0; i < p.pns.length; i++) u32(b, p.fcs.length * 4 + i * 4, p.pns[i]); return b; }
    var pbteC = plcfBte(chpx), pbteP = plcfBte(papx);
    var clxOff = tbl.length, pbteCOff = clxOff + clx.length, pbtePOff = pbteCOff + pbteC.length;
    var newTbl = concat([tbl, clx, pbteC, pbteP], tbl.length + clx.length + pbteC.length + pbteP.length);
    // extend the section table's last CP so the section spans the whole text
    var sedFc = pairFc(6), sedLcb = pairLcb(6);
    if (sedLcb >= 8) u32(newTbl, sedFc + ((sedLcb - 4) / 16) * 4, ccp);
    function setPair(idx, fc, lcb) { u32(newWd, fcLcbStart + idx * 8, fc); u32(newWd, fcLcbStart + idx * 8 + 4, lcb); }
    setPair(33, clxOff, clx.length);     // Clx
    setPair(12, pbteCOff, pbteC.length); // PlcfBteChpx
    setPair(13, pbtePOff, pbteP.length); // PlcfBtePapx

    return buildCfb([{ name: 'WordDocument', data: newWd }, { name: tableName, data: newTbl }]);
  }

  function concat(arrs, len) { var out = new Uint8Array(len), p = 0; arrs.forEach(function (a) { out.set(a, p); p += a.length; }); return out; }

  textToDoc.readCfb = readCfb;   // exposed for tests/tools
  textToDoc.buildCfb = buildCfb; // exposed for tooling (building the skeleton)

  // Bundled blank-document skeleton (a real, app-saved empty .doc, stripped to
  // its WordDocument + 1Table streams) that textToDoc() injects text into by
  // default. Structural/empty — carries no document content. See README.
  var TEMPLATE_B64 = "0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAP7/CQAGAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAEAAA/v///wAAAAD+////AAAAAAAAAAD////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9/////v///wMAAAAEAAAABQAAAAYAAAAHAAAACAAAAAkAAAAKAAAACwAAAAwAAAANAAAA/v///w8AAAAQAAAAEQAAABIAAAATAAAAFAAAABUAAAD+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUB//////////8BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAMQBUAGEAYgBsAGUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAgH/////AgAAAP////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAABgAAAAAAABXAG8AcgBkAEQAbwBjAHUAbQBlAG4AdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGgACAf///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQgQbABIAAQALAQ8ABwAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYgAAAPH/AgBiAAwEAAAAAAAAAAAGAE4AbwByAG0AYQBsAAAACwAAACokATEkAIAkAAAwAENKGABLSAEAT0oFAFBKBgBRSgUA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bgAgAFMAYQBuAHMAAAACAAQAAQiNGAAAxQIAAKkBAAAAAMkSWGfzheZHAAABAAIAAAAAAAAAAAAAAAAAAQABAAAABACDkAEAAAAAAAAAAAAAAAEAAQAAAAEAAAAAAAAAIYMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADwAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA0gQAAAAAAAAAAf////8H/////wf/////BAIAAAABAQD/////LwwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbAQAAGwEAAQAACCwAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAQDrvuu+urq6ugAABAD/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//xIAAAAAAAAAAAAAAAAAAAAAAAUAUgBvAG0AYQBuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7KUBAVVACQQAAAASvwAAAAAAADAAAAAAAAgAAAIIAAAOAFRNIDIwMTAgAAAAAAAAAAAAAAAAAAAAAAAACQQWACoOAABfxTFqX8UxagEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//8PAAUAAAABAAAA//8PAAYAAAABAAAA//8PAAAAAAAAAAAAAAAAAAAAAACIAAAAAAAeCQAAAAAAAB4JAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB4JAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADIJAAAMAAAAPgkAAAwAAAAAAAAAAAAAAOgPAADCAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAqhIAAFgCAADyFQAANAAAANMPAAAVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABKCQAAlQIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3wsAAHYDAABVDwAALAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgQ8AAFIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACCAAA8QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHDUIgUIqAUNKFQBPSgkAUUoJAHBoAAAAAF5KCQABAAgAAAIIAAD9AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARUAAAEoABew0AIYsNACH7CCLiCwxkEhsG4EIrBuBCOQbgQkkG4EJbAAAEuw/v8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";
  return textToDoc;
});
