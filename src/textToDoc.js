// SPDX-License-Identifier: 0BSD
/*
 * textToDoc(text) -> Uint8Array : writes a minimal Word 97-2003 binary .doc
 * (OLE2 / Compound File) containing the given body text. Clean-room from the
 * same open specs as the reader: [MS-CFB] for the container, [MS-DOC] for the
 * FIB + piece table.
 *
 * v1 scope: main body text + paragraph breaks, stored as UTF-16. No character
 * styling yet (that means writing CHPX/PAPX/STSH — a follow-on). The output is
 * the inverse of docToText(): docToText(textToDoc(s)) === s (plus Word's
 * required trailing paragraph mark).
 *
 * Verification: there is no Word/LibreOffice in this environment, so the output
 * is validated by reading it back with TWO independent parsers — our docToText
 * and the unrelated word-extractor (see test/writer.test.js). Both extracting
 * the right text is strong evidence the file is well-formed.
 *
 * Container note: every stream is padded to >= 4096 bytes so it lives in the
 * regular FAT (not the mini-FAT), which keeps the writer small.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.textToDoc = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SECTOR = 512;
  var FREESECT = 0xFFFFFFFF, ENDOFCHAIN = 0xFFFFFFFE, FATSECT = 0xFFFFFFFD;

  function u16(b, o, v) { b[o] = v & 0xFF; b[o + 1] = (v >> 8) & 0xFF; }
  function u32(b, o, v) { b[o] = v & 0xFF; b[o + 1] = (v >> 8) & 0xFF; b[o + 2] = (v >> 16) & 0xFF; b[o + 3] = (v >>> 24) & 0xFF; }

  // ---- [MS-DOC] WordDocument stream: FIB (900 bytes) then the UTF-16 text ----
  var FIB_LEN = 900;

  function writeFib(b, ccp, textBytes, fcClx, lcbClx) {
    u16(b, 0, 0xA5EC);                 // wIdent
    u16(b, 2, 0x00C1);                 // nFib = Word 97
    u16(b, 6, 0x0409);                 // lid = en-US
    u16(b, 10, 0x0200);                // flags: fWhichTblStm (bit 9) -> "1Table"
    u16(b, 12, 0x00BF);                // nFibBack
    u16(b, 32, 0x000E);                // csw = 14 (count of fibRgW words)
    u16(b, 62, 0x0016);                // cslw = 22 (count of fibRgLw longs)
    u32(b, 64 + 0 * 4, FIB_LEN + textBytes); // fibRgLw[0] = cbMac (used size)
    u32(b, 64 + 3 * 4, ccp);           // fibRgLw[3] = ccpText
    u16(b, 152, 0x005D);               // cbRgFcLcb = 93 (FibRgFcLcb97 pairs)
    u32(b, 154 + 33 * 8, fcClx);       // pair #33: fcClx
    u32(b, 154 + 33 * 8 + 4, lcbClx);  //           lcbClx
    // cswNew (offset 898) = 0 -> no fibRgCswNew, FIB ends at 900.
  }

  function buildWordDocument(text) {
    var ccp = text.length, wd = new Uint8Array(FIB_LEN + ccp * 2);
    for (var i = 0; i < ccp; i++) u16(wd, FIB_LEN + i * 2, text.charCodeAt(i));
    return wd; // FIB filled in by caller once fcClx/lcbClx are known
  }

  // ---- [MS-DOC] table stream: a CLX = Pcdt with a single piece -------------
  function buildClx(ccp) {
    var clx = new Uint8Array(21);
    clx[0] = 0x02;                     // Pcdt
    u32(clx, 1, 16);                   // lcb of the PlcPcd that follows
    u32(clx, 5, 0);                    // aCP[0] = 0
    u32(clx, 9, ccp);                  // aCP[1] = ccpText
    u16(clx, 13, 0);                   // Pcd: flags
    u32(clx, 15, FIB_LEN);             // Pcd: FcCompressed (fc=FIB_LEN, fCompressed=0 -> UTF-16)
    u16(clx, 19, 0);                   // Pcd: prm
    return clx;
  }

  // ---- [MS-CFB] compound file container ------------------------------------
  function writeHeader(b, fatSectors, dirStart) {
    var sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for (var i = 0; i < 8; i++) b[i] = sig[i];
    u16(b, 24, 0x003E);                // minor version
    u16(b, 26, 0x0003);                // major version (3 -> 512-byte sectors)
    u16(b, 28, 0xFFFE);                // byte order (little-endian)
    u16(b, 30, 0x0009);                // sector shift (2^9 = 512)
    u16(b, 32, 0x0006);                // mini sector shift (2^6 = 64)
    u32(b, 40, 0);                     // number of directory sectors (0 for v3)
    u32(b, 44, fatSectors);            // number of FAT sectors
    u32(b, 48, dirStart);              // first directory sector
    u32(b, 56, 0x00001000);            // mini stream cutoff (4096)
    u32(b, 60, ENDOFCHAIN);            // first mini-FAT sector (none)
    u32(b, 64, 0);                     // number of mini-FAT sectors
    u32(b, 68, ENDOFCHAIN);            // first DIFAT sector (none)
    u32(b, 72, 0);                     // number of DIFAT sectors
    for (var d = 0; d < 109; d++) u32(b, 76 + d * 4, d < fatSectors ? d : FREESECT);
  }

  function writeDir(b, idx, name, type, start, size, left, right, child) {
    var o = idx * 128, n;
    for (n = 0; n < name.length && n < 31; n++) u16(b, o + n * 2, name.charCodeAt(n));
    u16(b, o + 0x40, (name.length + 1) * 2); // name byte length incl. null terminator
    b[o + 0x42] = type;                       // 5 = root storage, 2 = stream, 0 = unused
    b[o + 0x43] = 1;                          // colour = black
    u32(b, o + 0x44, left < 0 ? FREESECT : left);
    u32(b, o + 0x48, right < 0 ? FREESECT : right);
    u32(b, o + 0x4C, child < 0 ? FREESECT : child);
    u32(b, o + 0x74, start);                  // starting sector
    u32(b, o + 0x78, size >>> 0);             // stream size (low)
  }

  function buildCfb(streams) {
    // Pad each stream to >= 4096 and a multiple of a sector (so it uses the FAT).
    var pad = streams.map(function (s) {
      var len = Math.ceil(Math.max(s.data.length, 4096) / SECTOR) * SECTOR;
      var d = new Uint8Array(len); d.set(s.data);
      return { name: s.name, data: d, size: len, sectors: len / SECTOR };
    });
    var dirSectors = Math.ceil((1 + pad.length) / 4);
    var streamSecs = pad.reduce(function (a, s) { return a + s.sectors; }, 0);
    var nonFat = dirSectors + streamSecs;
    var fatSectors = 1;
    while (Math.ceil((nonFat + fatSectors) / 128) > fatSectors) fatSectors++;
    var total = fatSectors + nonFat;

    var dirStart = fatSectors, p = dirStart + dirSectors;
    pad.forEach(function (s) { s.start = p; p += s.sectors; });

    // FAT
    var fat = new Uint8Array(fatSectors * SECTOR);
    for (var i = 0; i < fatSectors * 128; i++) u32(fat, i * 4, FREESECT);
    for (i = 0; i < fatSectors; i++) u32(fat, i * 4, FATSECT);
    for (i = 0; i < dirSectors; i++) u32(fat, (dirStart + i) * 4, i < dirSectors - 1 ? dirStart + i + 1 : ENDOFCHAIN);
    pad.forEach(function (s) {
      for (var k = 0; k < s.sectors; k++) u32(fat, (s.start + k) * 4, k < s.sectors - 1 ? s.start + k + 1 : ENDOFCHAIN);
    });

    // Directory: Root storage + one entry per stream, linked as a right-chain tree.
    var dir = new Uint8Array(dirSectors * SECTOR);
    writeDir(dir, 0, 'Root Entry', 5, ENDOFCHAIN, 0, -1, -1, pad.length ? 1 : -1);
    pad.forEach(function (s, idx) {
      writeDir(dir, idx + 1, s.name, 2, s.start, s.size, -1, idx < pad.length - 1 ? idx + 2 : -1, -1);
    });
    for (var e = 1 + pad.length; e < dirSectors * 4; e++) writeDir(dir, e, '', 0, 0, 0, -1, -1, -1);

    var file = new Uint8Array(SECTOR * (1 + total));
    writeHeader(file, fatSectors, dirStart);
    file.set(fat, SECTOR);                                  // FAT occupies sector 0
    file.set(dir, SECTOR + dirStart * SECTOR);
    pad.forEach(function (s) { file.set(s.data, SECTOR + s.start * SECTOR); });
    return file;
  }

  function textToDoc(text) {
    text = String(text == null ? '' : text).replace(/\r\n?|\n/g, '\r');
    if (!text || text.charCodeAt(text.length - 1) !== 0x0D) text += '\r'; // Word wants a final para mark
    var ccp = text.length;
    var wd = buildWordDocument(text);
    var clx = buildClx(ccp);
    writeFib(wd, ccp, ccp * 2, 0 /* fcClx */, clx.length /* lcbClx */);
    return buildCfb([{ name: 'WordDocument', data: wd }, { name: '1Table', data: clx }]);
  }

  return textToDoc;
});
