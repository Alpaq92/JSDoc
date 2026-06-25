// SPDX-License-Identifier: 0BSD
/*
 * style-list.test.js — list membership resolved through the paragraph STYLE.
 *
 * Word doesn't always put a paragraph's list reference (sprmPIlfo / sprmPIlvl)
 * in the paragraph's own PAPX. Its built-in numbered headings — and many real
 * documents — carry the ilfo/ilvl in the LINKED STYLE's PAPX (UpxPapx) instead,
 * so a paragraph with no direct list sprm is still a list item. The reader must
 * resolve that through the stylesheet, the same way it already resolves a
 * style's character formatting.
 *
 * This builds a minimal but spec-valid .doc whose single paragraph references a
 * paragraph style (istd 1) and carries NO direct list sprm; the ilfo/ilvl live
 * only in that style's UpxPapx. The byte layout is written independently from
 * the reader (straight from [MS-DOC]), so a passing round-trip is real evidence.
 *
 * It also confirms the validity guard: an out-of-range style ilfo (0x07FF, the
 * "not really in a list" sentinel Word writes into some built-in styles) is NOT
 * treated as a list, so style resolution can't manufacture spurious markers.
 *
 * Plus an optional check against test06.doc (a real Word doc of style-numbered
 * headings) when the gitignored oracle corpus is present locally.
 *
 * Run: node test/style-list.test.js   (exit 0 = all passed)
 */
'use strict';

var fs = require('fs');
var path = require('path');
var docToText = require('../src/docToText.js');

var failures = 0;
function check(name, cond) {
  if (cond) console.log('  ok   ' + name);
  else { console.log('  FAIL ' + name); failures++; }
}

// ---- CFB / [MS-DOC] byte layout (independent of the reader) ----------------
var SECTOR = 512, ENDOFCHAIN = 0xFFFFFFFE, FREESECT = 0xFFFFFFFF, FATSECT = 0xFFFFFFFD, NOSTREAM = 0xFFFFFFFF;
var SID_FAT = 0, SID_DIR = 1, SID_MINIFAT = 2, SID_MINISTREAM = 3, SID_WD = 4, WD_SECTORS = 8;
var TOTAL = SID_WD + WD_SECTORS;                 // 12 sectors (+1 header)
function sectorBase(sid) { return SECTOR + sid * SECTOR; }

// Table-stream (1Table) layout: CLX, STSH, PlcfBtePapx, PlfLst, PlfLfo.
var T_CLX = 0x00, T_STSH = 0x20, T_PAPX = 0x50, T_LST = 0x60, T_LFO = 0xA0, T_LEN = 0xB4;

// One paragraph "Hi\r" (compressed cp1252) at WordDocument byte 0x600, styled
// with paragraph style istd 1. `styleIlfo` is the ilfo the style's UpxPapx puts
// on the paragraph (1 = a real list; 0x07FF = the not-in-a-list sentinel).
function buildTable(styleIlfo) {
  var t = new Uint8Array(T_LEN), dv = new DataView(t.buffer);
  // CLX: a bare Pcdt + PlcPcd (2 CPs, 1 PCD); the piece is compressed at WD 0x600.
  t[T_CLX] = 0x02; dv.setUint32(T_CLX + 1, 16, true);
  dv.setUint32(T_CLX + 5, 0, true); dv.setUint32(T_CLX + 9, 3, true);          // CPs 0, 3
  dv.setUint32(T_CLX + 15, (0x600 * 2) | 0x40000000, true);                    // PCD.fc: fCompressed | byteOffset 0x600

  // STSH: cbStshi + STSHI{cstd=2, cbBase=10} + rgStd[ empty slot 0, list style 1 ].
  dv.setUint16(T_STSH + 0, 4, true);             // cbStshi (only cstd + cbBase are read)
  dv.setUint16(T_STSH + 2, 2, true);             // STSHI.cstd
  dv.setUint16(T_STSH + 4, 10, true);            // STSHI.cbSTDBaseInFile
  dv.setUint16(T_STSH + 6, 0, true);             // style 0: cbStd = 0 (skipped)
  var std = T_STSH + 8;
  dv.setUint16(std + 0, 28, true);               // style 1: cbStd = 28
  var s = std + 2;                               // STD body (STDF base is cbBase=10 bytes)
  dv.setUint16(s + 2, (0xFFF << 4) | 1, true);   // word: stk=1 (paragraph), istdBase=0xFFF (no parent)
  dv.setUint16(s + 4, 2, true);                  // cupx = 2 (UpxPapx then UpxChpx)
  // name (cch=0 + chTerm) spans s+10..s+14; UPX0 = UpxPapx at s+14
  dv.setUint16(s + 14, 9, true);                 // UpxPapx cbUpx = 9 (leading istd 2 + grpprlPapx 7)
  dv.setUint16(s + 16, 1, true);                 // UpxPapx leading istd (ignored by the reader)
  t[s + 18] = 0x0A; t[s + 19] = 0x26; t[s + 20] = 0x00;                        // sprmPIlvl = 0
  t[s + 21] = 0x0B; t[s + 22] = 0x46;                                          // sprmPIlfo ...
  dv.setUint16(s + 23, styleIlfo, true);         // ... operand (the list this style joins)
  dv.setUint16(s + 26, 0, true);                 // UPX1 = UpxChpx, cbUpx = 0 (no char props)

  // PlcfBtePapx: 2 FCs + 1 PN; the lone PAPX FKP page sits at WordDocument byte 0x400 (pn 2).
  dv.setUint32(T_PAPX + 0, 0x600, true); dv.setUint32(T_PAPX + 4, 0x603, true);
  dv.setUint32(T_PAPX + 8, 2, true);             // PN

  // PlfLst: one simple list (lsid), one bullet LVL (nfc 23, glyph U+2022).
  dv.setUint16(T_LST + 0, 1, true);              // cLst
  var lstf = T_LST + 2;
  dv.setInt32(lstf + 0, 0x12345678, true);       // LSTF.lsid
  t[lstf + 26] = 1;                              // fSimpleList -> 1 level
  var lvl = lstf + 28;
  dv.setInt32(lvl + 0, 1, true);                 // LVLF.iStartAt
  t[lvl + 4] = 23;                               // LVLF.nfc = 23 (bullet)
  t[lvl + 24] = 0; t[lvl + 25] = 0;              // cbGrpprlChpx, cbGrpprlPapx
  dv.setUint16(lvl + 28, 1, true);               // number-text cch = 1
  dv.setUint16(lvl + 30, 0x2022, true);          // the bullet glyph

  // PlfLfo: one LFO mapping ilfo 1 -> the LST above (by lsid).
  dv.setUint32(T_LFO + 0, 1, true);              // cLfo
  dv.setInt32(T_LFO + 4, 0x12345678, true);      // LFO.lsid
  return t;
}

function buildDoc(styleIlfo) {
  var file = new Uint8Array(SECTOR * (TOTAL + 1)), dv = new DataView(file.buffer);
  var sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  for (var i = 0; i < 8; i++) file[i] = sig[i];
  dv.setUint16(26, 0x0003, true); dv.setUint16(28, 0xFFFE, true);
  dv.setUint16(30, 0x0009, true); dv.setUint16(32, 0x0006, true);
  dv.setUint32(44, 1, true); dv.setUint32(48, SID_DIR, true);
  dv.setUint32(56, 0x1000, true); dv.setUint32(60, SID_MINIFAT, true);
  dv.setUint32(64, 1, true); dv.setUint32(68, ENDOFCHAIN, true); dv.setUint32(72, 0, true);
  dv.setUint32(76, SID_FAT, true);
  for (var d = 1; d < 109; d++) dv.setUint32(76 + d * 4, FREESECT, true);
  // FAT: WordDocument occupies sectors 4..11.
  var fb = sectorBase(SID_FAT), fat = new Array(128).fill(FREESECT);
  fat[SID_FAT] = FATSECT; fat[SID_DIR] = ENDOFCHAIN; fat[SID_MINIFAT] = ENDOFCHAIN; fat[SID_MINISTREAM] = ENDOFCHAIN;
  for (var w = 0; w < WD_SECTORS; w++) fat[SID_WD + w] = (w === WD_SECTORS - 1) ? ENDOFCHAIN : SID_WD + w + 1;
  for (i = 0; i < 128; i++) dv.setUint32(fb + i * 4, fat[i] >>> 0, true);
  // mini-FAT: 1Table spans mini-sectors 0 -> 1 -> 2.
  var mb = sectorBase(SID_MINIFAT);
  for (i = 0; i < 128; i++) dv.setUint32(mb + i * 4, FREESECT, true);
  dv.setUint32(mb + 0, 1, true); dv.setUint32(mb + 4, 2, true); dv.setUint32(mb + 8, ENDOFCHAIN, true);
  // directory
  var dbase = sectorBase(SID_DIR);
  writeDir(dv, dbase + 0, 'Root Entry', 5, SID_MINISTREAM, 512, 1);
  writeDir(dv, dbase + 128, 'WordDocument', 2, SID_WD, WD_SECTORS * SECTOR, NOSTREAM);
  writeDir(dv, dbase + 256, '1Table', 2, 0, T_LEN, NOSTREAM);
  file.set(buildTable(styleIlfo), sectorBase(SID_MINISTREAM));   // table stream lives in the mini stream
  writeWd(file, dv);
  return file;
}

function writeDir(dv, off, name, type, start, size, child) {
  for (var i = 0; i < name.length; i++) dv.setUint16(off + i * 2, name.charCodeAt(i), true);
  dv.setUint16(off + 64, (name.length + 1) * 2, true);
  dv.setUint8(off + 66, type); dv.setUint8(off + 67, 1);
  dv.setUint32(off + 68, NOSTREAM, true); dv.setUint32(off + 72, NOSTREAM, true);
  dv.setUint32(off + 76, child >>> 0, true);
  dv.setUint32(off + 116, start >>> 0, true); dv.setUint32(off + 120, size >>> 0, true);
}

function writeWd(file, dv) {
  var b = sectorBase(SID_WD);
  dv.setUint16(b + 0, 0xA5EC, true); dv.setUint16(b + 2, 0x00C1, true);
  dv.setUint16(b + 10, 0x0200, true);            // fWhichTblStm = 1 -> 1Table
  dv.setUint16(b + 0x20, 0x000E, true);          // csw
  dv.setUint16(b + 0x3E, 0x0016, true);          // cslw
  dv.setUint32(b + 0x4C, 3, true);               // ccpText = 3 ("Hi\r")
  dv.setUint16(b + 0x98, 0x005D, true);          // cbRgFcLcb (93 FibRgFcLcb pairs)
  var fib = b + 0x9A;
  function pairAt(idx, fc, lcb) { dv.setUint32(fib + idx * 8, fc, true); dv.setUint32(fib + idx * 8 + 4, lcb, true); }
  pairAt(1, T_STSH, 38);                         // #1  STSH
  pairAt(13, T_PAPX, 12);                        // #13 PlcfBtePapx
  pairAt(33, T_CLX, 21);                         // #33 CLX
  pairAt(73, T_LST, 62);                         // #73 PlfLst
  pairAt(74, T_LFO, 20);                         // #74 PlfLfo
  // PAPX FKP page (WD 0x400): one paragraph, BxPap -> PapxInFkp with istd 1 and NO grpprl
  // (no direct list sprm) — list membership must come from the style.
  var fkp = b + 0x400;
  file[fkp + 511] = 1;                           // crun
  dv.setUint32(fkp + 0, 0x600, true); dv.setUint32(fkp + 4, 0x603, true); // rgfc[2]
  file[fkp + 8] = 0x08;                          // BxPap.bOffset (words) -> PapxInFkp at fkp + 0x10
  file[fkp + 0x10] = 2;                          // PapxInFkp cb = 2 -> GrpPrlAndIstd = istd(2) + slack
  dv.setUint16(fkp + 0x11, 1, true);             // istd = 1
  file[b + 0x600] = 0x48; file[b + 0x601] = 0x69; file[b + 0x602] = 0x0D;   // "Hi" + paragraph mark
}

console.log('docToText — paragraph-style list membership');

// 1. A paragraph whose list reference lives only in its style is a list item.
var listed = buildDoc(1);
var lm = docToText.model(listed).body;
check('exactly one body paragraph', lm.length === 1);
check('paragraph is recognized as a list item (ilfo from style)', !!lm[0].list && lm[0].list.ilfo === 1 && lm[0].list.ilvl === 0);
check('list kind + synthesized bullet marker', !!lm[0].list && lm[0].list.kind === 'bullet' && lm[0].list.marker === '•');
check('marker appears in docToText() plain text', docToText(listed) === '• Hi\n');

// 2. An out-of-range style ilfo (0x07FF sentinel) is NOT a list — no spurious marker.
var sentinel = buildDoc(0x07FF);
var sm = docToText.model(sentinel).body;
check('sentinel style ilfo (0x07FF) is not treated as a list', sm.length === 1 && !sm[0].list);
check('no marker synthesized for the sentinel', docToText(sentinel) === 'Hi\n');

// 3. Real-doc check (style-numbered headings) when the gitignored corpus is present.
var real = path.join(__dirname, 'fixtures', 'test06.doc');
if (fs.existsSync(real)) {
  var buf = fs.readFileSync(real);
  var body = docToText.model(buf).body;
  var lists = body.filter(function (p) { return p.list; });
  check('test06.doc: style-based list items are recognized (was 0)', lists.length > 0);
  check('test06.doc: recognized items carry a synthesized marker', lists.every(function (p) { return !!p.list.marker; }));
  check('test06.doc: a marker reaches the plain-text output', docToText(buf).indexOf('•') !== -1);
} else {
  console.log('  skip test06.doc real-doc check (fixture not present)');
}

console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)');
process.exit(failures === 0 ? 0 : 1);
