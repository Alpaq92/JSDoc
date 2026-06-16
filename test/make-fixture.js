// SPDX-License-Identifier: 0BSD
/*
 * make-fixture.js — builds a minimal but spec-valid Word 97-2003 .doc in
 * memory, for offline, known-answer testing of docToText().
 *
 * It is written independently from the reader (it lays out the [MS-CFB] /
 * [MS-DOC] byte structures directly) so that a passing round-trip is real
 * evidence, not a shared bug. It deliberately exercises:
 *   - CFB v3 container: header DIFAT, a FAT chain (WordDocument, 8 sectors),
 *     the directory, and the mini-FAT / mini-stream path (1Table, 50 bytes).
 *   - A CLX with a leading Prc (must be skipped) then a Pcdt / PlcPcd.
 *   - Three pieces: compressed cp1252, compressed cp1252, uncompressed UTF-16LE.
 *   - Windows-1252 smart quotes, a UTF-16-only char, a field (code dropped,
 *     result kept), a cell mark, and paragraph marks.
 */
'use strict';

var SECTOR = 512;
var ENDOFCHAIN = 0xFFFFFFFE;
var FREESECT = 0xFFFFFFFF;
var FATSECT = 0xFFFFFFFD;
var NOSTREAM = 0xFFFFFFFF;

// Sector layout (sector N starts at file offset (N+1)*512):
//   0 FAT | 1 Directory | 2 mini-FAT | 3 mini-stream | 4..11 WordDocument
var SID_FAT = 0, SID_DIR = 1, SID_MINIFAT = 2, SID_MINISTREAM = 3, SID_WD = 4;
var WD_SECTORS = 8;                       // 8 * 512 = 4096 bytes (>= mini cutoff)
var TOTAL_SECTORS = SID_WD + WD_SECTORS;  // 12

function sectorBase(sid) { return SECTOR + sid * SECTOR; }

function buildDoc() {
  var file = new Uint8Array(SECTOR * (TOTAL_SECTORS + 1)); // +1 for header
  var dv = new DataView(file.buffer);

  writeHeader(dv);
  writeFat(dv);
  writeDirectory(dv);
  writeMiniFat(dv);

  var clx = buildClx();                // the piece table (50 bytes)
  file.set(clx, sectorBase(SID_MINISTREAM)); // 1Table -> mini sector 0

  writeWordDocument(file, dv);

  return file;
}

function writeHeader(dv) {
  var sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  for (var i = 0; i < 8; i++) dv.setUint8(i, sig[i]);
  dv.setUint16(24, 0x003E, true); // minor version
  dv.setUint16(26, 0x0003, true); // major version (v3)
  dv.setUint16(28, 0xFFFE, true); // byte order (little-endian)
  dv.setUint16(30, 0x0009, true); // sector shift -> 512
  dv.setUint16(32, 0x0006, true); // mini sector shift -> 64
  dv.setUint32(44, 1, true);          // number of FAT sectors
  dv.setUint32(48, SID_DIR, true);    // first directory sector
  dv.setUint32(56, 0x1000, true);     // mini stream cutoff (4096)
  dv.setUint32(60, SID_MINIFAT, true);// first mini-FAT sector
  dv.setUint32(64, 1, true);          // number of mini-FAT sectors
  dv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector (none)
  dv.setUint32(72, 0, true);          // number of DIFAT sectors
  // Header DIFAT: slot 0 -> FAT at sector 0, rest free.
  dv.setUint32(76, SID_FAT, true);
  for (var d = 1; d < 109; d++) dv.setUint32(76 + d * 4, FREESECT, true);
}

function writeFat(dv) {
  var b = sectorBase(SID_FAT);
  var fat = new Array(128).fill(FREESECT);
  fat[SID_FAT] = FATSECT;
  fat[SID_DIR] = ENDOFCHAIN;
  fat[SID_MINIFAT] = ENDOFCHAIN;
  fat[SID_MINISTREAM] = ENDOFCHAIN;
  for (var s = 0; s < WD_SECTORS; s++) {
    fat[SID_WD + s] = (s === WD_SECTORS - 1) ? ENDOFCHAIN : SID_WD + s + 1;
  }
  for (var i = 0; i < 128; i++) dv.setUint32(b + i * 4, fat[i] >>> 0, true);
}

function writeMiniFat(dv) {
  var b = sectorBase(SID_MINIFAT);
  for (var i = 0; i < 128; i++) dv.setUint32(b + i * 4, FREESECT, true);
  dv.setUint32(b, ENDOFCHAIN, true); // 1Table occupies a single mini sector
}

function writeDirectory(dv) {
  var b = sectorBase(SID_DIR);
  // 0: Root Entry (type 5) -> owns the mini stream (sector 3, 64 bytes).
  writeDirEntry(dv, b + 0 * 128, 'Root Entry', 5, SID_MINISTREAM, 64, 1);
  // 1: WordDocument (type 2) -> FAT chain starting at sector 4, 4096 bytes.
  writeDirEntry(dv, b + 1 * 128, 'WordDocument', 2, SID_WD, WD_SECTORS * SECTOR, NOSTREAM);
  // 2: 1Table (type 2) -> mini sector 0, 50 bytes (the CLX).
  writeDirEntry(dv, b + 2 * 128, '1Table', 2, 0, 50, NOSTREAM);
}

function writeDirEntry(dv, off, name, type, start, size, child) {
  for (var i = 0; i < name.length; i++) {
    dv.setUint16(off + i * 2, name.charCodeAt(i), true);
  }
  dv.setUint16(off + 64, (name.length + 1) * 2, true); // name length incl. NUL
  dv.setUint8(off + 66, type);
  dv.setUint8(off + 67, 1); // color = black
  dv.setUint32(off + 68, NOSTREAM, true); // left sibling
  dv.setUint32(off + 72, NOSTREAM, true); // right sibling
  dv.setUint32(off + 76, child >>> 0, true); // child
  dv.setUint32(off + 116, start >>> 0, true); // starting sector
  dv.setUint32(off + 120, size >>> 0, true);  // stream size (low)
  dv.setUint32(off + 124, 0, true);           // stream size (high)
}

// ----- the three text pieces ------------------------------------------------
// Piece A (compressed cp1252), 15 chars:  Hello, “World”\n
var PIECE_A = [0x48,0x65,0x6C,0x6C,0x6F,0x2C,0x20, 0x93,
               0x57,0x6F,0x72,0x6C,0x64, 0x94, 0x0D];
// Piece B (compressed cp1252), 16 chars: <field PAGE = "1">Tab<cell>End\n
var PIECE_B = [0x13, 0x50,0x41,0x47,0x45, 0x14, 0x31, 0x15,
               0x54,0x61,0x62, 0x07, 0x45,0x6E,0x64, 0x0D];
// Piece C (uncompressed UTF-16LE), 12 chars: "Unicode: " + é + π + \n
var PIECE_C_TEXT = 'Unicode: éπ\r';

var WD_OFF_A = 0x400, WD_OFF_B = 0x420, WD_OFF_C = 0x440;

// What docToText() should return for this fixture.
var EXPECTED = 'Hello, “World”\n' + '1Tab\tEnd\n' + 'Unicode: éπ\n';

function writeWordDocument(file, dv) {
  var b = sectorBase(SID_WD);

  // --- FibBase ---
  dv.setUint16(b + 0, 0xA5EC, true); // wIdent
  dv.setUint16(b + 2, 0x00C1, true); // nFib (Word 97)
  dv.setUint16(b + 10, 0x0200, true);// flags: fWhichTblStm=1 -> 1Table

  // --- FIB variable parts (counts that put fcClx at the canonical 0x1A2) ---
  dv.setUint16(b + 0x20, 0x000E, true); // csw
  dv.setUint16(b + 0x3E, 0x0016, true); // cslw
  dv.setUint32(b + 0x4C, 43, true);     // ccpText = fibRgLw[3]
  dv.setUint16(b + 0x98, 0x005D, true); // cbRgFcLcb
  dv.setUint32(b + 0x1A2, 0, true);     // fcClx
  dv.setUint32(b + 0x1A6, 50, true);    // lcbClx

  // --- piece data ---
  for (var i = 0; i < PIECE_A.length; i++) file[b + WD_OFF_A + i] = PIECE_A[i];
  for (var j = 0; j < PIECE_B.length; j++) file[b + WD_OFF_B + j] = PIECE_B[j];
  for (var k = 0; k < PIECE_C_TEXT.length; k++) {
    dv.setUint16(b + WD_OFF_C + k * 2, PIECE_C_TEXT.charCodeAt(k), true);
  }
}

function buildClx() {
  var clx = new Uint8Array(50);
  var dv = new DataView(clx.buffer);
  // Prc (skipped by the reader): clxt=1, cbGrpprl=2, grpprl=2 bytes.
  clx[0] = 0x01; dv.setInt16(1, 2, true); clx[3] = 0; clx[4] = 0;
  // Pcdt: clxt=2, lcb=40, then PlcPcd (40 bytes) at offset 10.
  clx[5] = 0x02; dv.setUint32(6, 40, true);
  var p = 10;
  // CPs: 0, 15, 31, 43
  [0, 15, 31, 43].forEach(function (cp, idx) { dv.setUint32(p + idx * 4, cp, true); });
  var pcd = p + 16;
  // PCD.fc (FcCompressed): bit 30 = fCompressed; compressed fc = byteOffset*2.
  dv.setUint32(pcd + 0 * 8 + 2, (WD_OFF_A * 2) | 0x40000000, true); // A compressed
  dv.setUint32(pcd + 1 * 8 + 2, (WD_OFF_B * 2) | 0x40000000, true); // B compressed
  dv.setUint32(pcd + 2 * 8 + 2, WD_OFF_C, true);                    // C uncompressed
  return clx;
}

module.exports = { buildDoc: buildDoc, EXPECTED: EXPECTED };
