// SPDX-License-Identifier: 0BSD
/*
 * textToDoc(text[, template]) -> Uint8Array : write a Word 97-2003 binary .doc
 * containing `text`. `template` is an optional blank .doc (Uint8Array /
 * ArrayBuffer / Buffer) saved by a real word processor; when omitted, a small
 * bundled skeleton (embedded below) is used.
 *
 * Why a template: a from-scratch .doc round-trips through lenient parsers but
 * real word processors reject it — they need a valid stylesheet, section table
 * and character/paragraph property tables. Rather than synthesise all of those
 * (Apache POI doesn't either), we reuse them from a genuine blank document and
 * only swap in the body text + piece table + property-table FC ranges, building
 * fresh CHPX/PAPX FKP pages. Clean-room from [MS-CFB] (container) + [MS-DOC]
 * (FIB / CLX / FKP).
 *
 * v1 scope: body text + paragraph breaks (UTF-16), Normal-style formatting.
 * No styling write-back yet.
 *
 * Verification: round-tripped through docToText AND the independent
 * word-extractor (test/writer.test.js), and confirmed to open cleanly in a real
 * word processor — SoftMaker TextMaker — driven via its COM automation
 * (scripts/read-with-textmaker.ps1), which is what flushed out the structures
 * the lenient parsers didn't require.
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

  // ---- FKP page builders ([MS-DOC]) ----------------------------------------
  // A CHPX FKP page: rgfc[crun+1], rgb[crun] (word offsets to a Chpx), crun@511.
  // We use one run over the whole text, reusing the template's blank-para Chpx.
  function chpxFkpPage(fc0, fc1, chpxBytes) {
    var pg = new Uint8Array(SECTOR), pos = (SECTOR - 1 - chpxBytes.length) & ~1; // even, clear of crun@511
    pg.set(chpxBytes, pos);            // [cb][grpprl]
    u32(pg, 0, fc0); u32(pg, 4, fc1);  // rgfc[0..1]
    pg[8] = pos >> 1;                  // rgb[0] = word offset of the Chpx
    pg[511] = 1;                       // crun
    return pg;
  }
  // A PAPX FKP page: rgfc[crun+1], rgbx[crun] (BxPap, 13 bytes; first byte =
  // word offset to a PapxInFkp), crun@511. One run per paragraph, all sharing
  // the template's blank-para PapxInFkp. Returns null if the runs don't fit.
  function papxFkpPage(fcs, papxBytes) {
    var crun = fcs.length - 1;
    var need = 4 * (crun + 1) + 13 * crun + papxBytes.length + 1;
    if (need > SECTOR) return null;
    var pg = new Uint8Array(SECTOR), pos = (SECTOR - 1 - papxBytes.length) & ~1; // even, clear of crun@511
    pg.set(papxBytes, pos);
    for (var i = 0; i <= crun; i++) u32(pg, i * 4, fcs[i]);
    var bx = 4 * (crun + 1);
    for (i = 0; i < crun; i++) pg[bx + i * 13] = pos >> 1;  // BxPap.bOffset; PHE left 0
    pg[511] = crun;
    return pg;
  }

  function normalize(text) {
    text = String(text == null ? '' : text).replace(/\r\n?|\n/g, '\r');
    if (!text || text.charCodeAt(text.length - 1) !== 0x0D) text += '\r';
    return text;
  }

  function textToDoc(text, template) {
    var cfb = readCfb(template || defaultTemplate());
    if (!cfb) throw new Error('template is not a valid .doc (CFB) file');
    text = normalize(text);
    var ccp = text.length, textBytes = ccp * 2;

    var wd = cfb.byName['WordDocument'];
    var dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    var flags = dv.getUint16(10, true), tableName = (flags >> 9) & 1 ? '1Table' : '0Table';
    var tbl = cfb.byName[tableName];
    // FIB layout
    var csw = dv.getUint16(32, true), rgLwStart = 34 + csw * 2 + 2, cslw = dv.getUint16(rgLwStart - 2, true);
    var fcLcbStart = rgLwStart + cslw * 4 + 2;
    function pairFc(i) { return dv.getUint32(fcLcbStart + i * 8, true); }
    function pairLcb(i) { return dv.getUint32(fcLcbStart + i * 8 + 4, true); }
    // Clean default formatting, resolved against the template's (reused) stylesheet:
    // an empty Chpx (inherit Normal's character props) and a Normal-style (istd 0)
    // PapxInFkp. We deliberately do NOT copy the template's blank-paragraph grpprls,
    // which can carry stray props (bold/odd font) from whatever app saved it.
    var chpxBytes = new Uint8Array([0]);            // cb=0 -> no sprms -> default
    var papxBytes = new Uint8Array([0, 1, 0, 0]);   // cb=0, cb'=1, istd=0 (Normal)

    // ---- new WordDocument: template's WD + text + CHPX page + PAPX page(s) --
    function pad512(n) { return (n + 511) & ~511; }
    var T = (wd.length + 1) & ~1;                 // new text offset (even)
    var parts = [wd];
    var totalLen = wd.length;
    function place(arr) { parts.push(arr); var off = totalLen; totalLen += arr.length; return off; }
    function placeAt(off, arr) { while (totalLen < off) { parts.push(new Uint8Array(off - totalLen)); totalLen = off; } return place(arr); }
    // text
    var textBuf = new Uint8Array(textBytes);
    for (var i = 0; i < ccp; i++) u16(textBuf, i * 2, text.charCodeAt(i));
    placeAt(T, textBuf);
    // paragraph boundary FCs (each ends after a 0x0D or final mark)
    var bounds = [T];
    for (i = 0; i < ccp; i++) if (text.charCodeAt(i) === 0x0D) bounds.push(T + (i + 1) * 2);
    if (bounds[bounds.length - 1] !== T + textBytes) bounds.push(T + textBytes);
    // CHPX page (one run over all text)
    var chpxPn = pad512(totalLen) / 512;
    placeAt(chpxPn * 512, chpxFkpPage(T, T + textBytes, chpxBytes));
    // PAPX pages (split paragraphs across pages as needed)
    var papxPages = [], perPage = Math.max(1, Math.floor((SECTOR - papxBytes.length - 1 - 4) / 17));
    var papxPnList = [], papxFcs = [];
    for (var start = 0; start < bounds.length - 1; start += perPage) {
      var slice = bounds.slice(start, Math.min(start + perPage + 1, bounds.length));
      var page = papxFkpPage(slice, papxBytes);
      var ppn = pad512(totalLen) / 512;
      placeAt(ppn * 512, page);
      papxPnList.push(ppn); papxFcs.push(slice[0]);
    }
    papxFcs.push(T + textBytes);
    var newWd = concat(parts, totalLen);
    // patch FIB: ccpText, cbMac
    var ndv = new DataView(newWd.buffer, newWd.byteOffset, newWd.byteLength);
    u32(newWd, rgLwStart + 3 * 4, ccp);           // ccpText
    u32(newWd, rgLwStart + 0 * 4, newWd.length);  // cbMac

    // ---- new 1Table: template's table + appended CLX + PlcfBteChpx/Papx -----
    var add = [];
    // CLX: Pcdt with one piece
    var clx = new Uint8Array(21);
    clx[0] = 0x02; u32(clx, 1, 16); u32(clx, 5, 0); u32(clx, 9, ccp);
    u16(clx, 13, 0); u32(clx, 15, T); u16(clx, 19, 0);
    var clxOff = tbl.length;
    add.push(clx);
    // PlcfBteChpx: 2 FCs + 1 pn
    var pbteC = new Uint8Array(12);
    u32(pbteC, 0, T); u32(pbteC, 4, T + textBytes); u32(pbteC, 8, chpxPn);
    var pbteCOff = clxOff + clx.length;
    add.push(pbteC);
    // PlcfBtePapx: (k+1) FCs + k pns
    var k = papxPnList.length, pbteP = new Uint8Array((k + 1) * 4 + k * 4);
    for (i = 0; i <= k; i++) u32(pbteP, i * 4, papxFcs[i]);
    for (i = 0; i < k; i++) u32(pbteP, (k + 1) * 4 + i * 4, papxPnList[i]);
    var pbtePOff = pbteCOff + pbteC.length;
    add.push(pbteP);
    var newTbl = concat([tbl].concat(add), tbl.length + clx.length + pbteC.length + pbteP.length);
    // extend the section table's last CP to cover the whole text (else Word
    // treats only the first paragraph as in-section).
    var sedFc = pairFc(6), sedLcb = pairLcb(6);
    if (sedLcb >= 8) u32(newTbl, sedFc + ((sedLcb - 4) / 16) * 4, ccp);
    // repoint FIB fc/lcb pairs to the appended structures
    function setPair(idx, fc, lcb) { u32(newWd, fcLcbStart + idx * 8, fc); u32(newWd, fcLcbStart + idx * 8 + 4, lcb); }
    setPair(33, clxOff, clx.length);               // Clx
    setPair(12, pbteCOff, pbteC.length);           // PlcfBteChpx
    setPair(13, pbtePOff, pbteP.length);           // PlcfBtePapx

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
