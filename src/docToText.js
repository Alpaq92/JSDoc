// SPDX-License-Identifier: 0BSD
/*
 * docToText — pure-JavaScript text extractor for legacy Microsoft Word
 * 97-2003 binary ".doc" files (the OLE2 / Compound File Binary container).
 *
 * Clean-room implementation written from Microsoft's free, openly published
 * specifications:
 *   - [MS-CFB]  Compound File Binary File Format
 *               https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/
 *   - [MS-DOC]  Word (.doc) Binary File Format
 *               https://learn.microsoft.com/openspecs/office_file_formats/ms-doc/
 * No code was read, ported, or translated from GPL tools (e.g. catdoc/antiword).
 * File formats and the algorithms that read them are not copyrightable, so this
 * from-spec implementation may be licensed permissively (0BSD).
 *
 * Public API — a single pure function, no DOM and no network:
 *
 *     docToText(input) -> string | null
 *
 *   input : ArrayBuffer | Uint8Array | Node Buffer of a .doc file.
 *   returns: the extracted main-body text, or null when the file is
 *            unsupported or unreadable (not a CFB, Word 6/95 or older,
 *            encrypted/obfuscated, or any parse error). null is the signal
 *            for the host to fall back to its download / handoff path.
 *
 * Scope (lossy by design, like a plain-text/RTF view): main document body
 * text and paragraph breaks only. No fonts, images, or styles. Tables collapse
 * to tab/newline text. Headers, footers, and footnotes live in separate CP
 * ranges and are out of scope for v1 (see extractText() for where they plug in).
 * Tracked-change *deletions* are kept (the deleted text is still in the main CP
 * range); dropping them would require parsing character-level revision marks
 * (sprmCFRMarkDel via the CHPX bin table) — a documented v2 extension.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.docToText = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---- [MS-CFB] constants -------------------------------------------------
  var CFB_SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  var ENDOFCHAIN = 0xFFFFFFFE;
  var FREESECT = 0xFFFFFFFF;
  // (FATSECT 0xFFFFFFFD and DIFSECT 0xFFFFFFFC never appear as chain links.)

  // Byte -> Unicode mapping for compressed (8-bit) text, transcribed verbatim
  // from the table in [MS-DOC] "FcCompressed": a compressed character is
  // its own code point (Latin-1 identity), EXCEPT the bytes listed here. Note
  // this is *not* full Windows-1252 — 0x80, 0x8E and 0x9E are NOT remapped by
  // the spec (they stay U+0080/U+008E/U+009E). Indices below are byte - 0x80.
  var FC_COMPRESSED_MAP = [
    0x0080, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, // 80-87
    0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x008E, 0x008F, // 88-8F
    0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, // 90-97
    0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x009E, 0x0178  // 98-9F
  ];

  // ------------------------------------------------------------------------
  // Public entry point
  // ------------------------------------------------------------------------
  function docToText(input) {
    try {
      var bytes = toUint8(input);
      if (!bytes || bytes.length < 512) return null;

      var cfb = parseCfb(bytes);
      if (!cfb) return null;

      var wordDocument = cfb.byName['WordDocument'];
      if (!wordDocument) return null;
      var wd = cfb.getStream(wordDocument);

      return parseWord(wd, cfb);
    } catch (e) {
      return null; // any failure -> graceful handoff
    }
  }

  // ------------------------------------------------------------------------
  // Layer 1 — [MS-CFB] OLE2 container
  // ------------------------------------------------------------------------
  function parseCfb(bytes) {
    var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

    for (var i = 0; i < 8; i++) {
      if (bytes[i] !== CFB_SIGNATURE[i]) return null; // not a compound file
    }

    var sectorShift = dv.getUint16(30, true);
    var miniSectorShift = dv.getUint16(32, true);
    var sectorSize = 1 << sectorShift;       // 512 (v3) or 4096 (v4)
    var miniSectorSize = 1 << miniSectorShift; // 64
    if (sectorSize !== 512 && sectorSize !== 4096) return null;

    var firstDirSector = dv.getUint32(48, true);
    var miniStreamCutoff = dv.getUint32(56, true);   // usually 4096
    var firstMiniFatSector = dv.getUint32(60, true);
    var numMiniFatSectors = dv.getUint32(64, true);
    var firstDifatSector = dv.getUint32(68, true);
    var numDifatSectors = dv.getUint32(72, true);

    function sectorOffset(sid) { return (sid + 1) * sectorSize; }

    // -- Collect FAT sector locations: 109 in the header DIFAT, then chain. --
    var fatSectorLocs = [];
    for (var d = 0; d < 109; d++) {
      var loc = dv.getUint32(76 + d * 4, true);
      if (loc === FREESECT || loc === ENDOFCHAIN) break;
      fatSectorLocs.push(loc);
    }
    var entriesPerDifat = (sectorSize / 4) - 1; // last slot links to next DIFAT
    var difatSid = firstDifatSector;
    var difatGuard = numDifatSectors + 8;
    while (difatSid !== ENDOFCHAIN && difatSid !== FREESECT && difatGuard-- > 0) {
      var dbase = sectorOffset(difatSid);
      if (dbase + sectorSize > bytes.length) break;
      for (var k = 0; k < entriesPerDifat; k++) {
        var fl = dv.getUint32(dbase + k * 4, true);
        if (fl !== FREESECT && fl !== ENDOFCHAIN) fatSectorLocs.push(fl);
      }
      difatSid = dv.getUint32(dbase + entriesPerDifat * 4, true);
    }

    // -- Read the FAT into one flat Uint32Array. --
    var entriesPerSector = sectorSize / 4;
    var fat = new Uint32Array(fatSectorLocs.length * entriesPerSector);
    var fi = 0;
    for (var f = 0; f < fatSectorLocs.length; f++) {
      var foff = sectorOffset(fatSectorLocs[f]);
      for (var e = 0; e < entriesPerSector; e++) {
        fat[fi++] = (foff + e * 4 + 4 <= bytes.length)
          ? dv.getUint32(foff + e * 4, true) : FREESECT;
      }
    }

    // -- Follow a FAT chain, returning its bytes (clamped to sizeLimit). --
    function readChain(startSid, sizeLimit) {
      var chunks = [];
      var sid = startSid;
      var guard = fat.length + 8;
      var collected = 0;
      while (sid !== ENDOFCHAIN && sid !== FREESECT && guard-- > 0) {
        if (sid >= fat.length) break;
        var off = sectorOffset(sid);
        if (off >= bytes.length) break;
        var endOff = Math.min(off + sectorSize, bytes.length);
        chunks.push(bytes.subarray(off, endOff));
        collected += endOff - off;
        if (sizeLimit != null && collected >= sizeLimit) break;
        sid = fat[sid];
      }
      return concat(chunks, sizeLimit);
    }

    // -- Directory. --
    var dirBytes = readChain(firstDirSector, null);
    var dirDv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
    var numEntries = Math.floor(dirBytes.length / 128);
    var root = null;
    var byName = {};
    for (var n = 0; n < numEntries; n++) {
      var base = n * 128;
      var type = dirBytes[base + 66]; // 0 unalloc, 1 storage, 2 stream, 5 root
      if (type !== 1 && type !== 2 && type !== 5) continue;
      var nameLen = dirDv.getUint16(base + 64, true); // bytes incl. terminator
      var name = '';
      if (nameLen > 2) {
        var chars = (nameLen >> 1) - 1;
        for (var c = 0; c < chars; c++) {
          name += String.fromCharCode(dirDv.getUint16(base + c * 2, true));
        }
      }
      var start = dirDv.getUint32(base + 116, true);
      var sizeLow = dirDv.getUint32(base + 120, true);
      var sizeHigh = dirDv.getUint32(base + 124, true); // 0 for v3
      var entry = { name: name, type: type, start: start,
                    size: sizeHigh * 0x100000000 + sizeLow };
      if (type === 5) root = entry;
      else if (type === 2) byName[name] = entry;
    }
    if (!root) return null;

    // -- Mini-FAT + mini stream (small streams live here). --
    var miniStream = readChain(root.start, root.size);
    var miniFatBytes = (numMiniFatSectors > 0 && firstMiniFatSector !== ENDOFCHAIN)
      ? readChain(firstMiniFatSector, null) : new Uint8Array(0);
    var miniFat = new Uint32Array(miniFatBytes.length >> 2);
    var mfDv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
    for (var mi = 0; mi < miniFat.length; mi++) miniFat[mi] = mfDv.getUint32(mi * 4, true);

    function readMiniChain(startSid, sizeLimit) {
      var chunks = [];
      var sid = startSid;
      var guard = miniFat.length + 8;
      var collected = 0;
      while (sid !== ENDOFCHAIN && sid !== FREESECT && guard-- > 0) {
        var off = sid * miniSectorSize;
        if (off >= miniStream.length) break;
        var endOff = Math.min(off + miniSectorSize, miniStream.length);
        chunks.push(miniStream.subarray(off, endOff));
        collected += endOff - off;
        if (sizeLimit != null && collected >= sizeLimit) break;
        sid = (sid < miniFat.length) ? miniFat[sid] : ENDOFCHAIN;
      }
      return concat(chunks, sizeLimit);
    }

    function getStream(entry) {
      // The mini stream itself is always in the FAT; everything else picks a
      // home by size relative to the cutoff.
      return (entry.size >= miniStreamCutoff)
        ? readChain(entry.start, entry.size)
        : readMiniChain(entry.start, entry.size);
    }

    return { byName: byName, getStream: getStream };
  }

  // ------------------------------------------------------------------------
  // Layer 2 — [MS-DOC] Word stream: FIB -> CLX (piece table) -> text
  // ------------------------------------------------------------------------
  function parseWord(wd, cfb) {
    if (wd.length < 0x20) return null;
    var dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);

    // FibBase
    if (dv.getUint16(0, true) !== 0xA5EC) return null;   // wIdent
    var nFib = dv.getUint16(2, true);
    if (nFib < 0x00C1) return null;                       // Word 6/95 or older

    var flags = dv.getUint16(10, true);
    var fEncrypted = (flags & 0x0100) !== 0;  // bit 8
    var fWhichTblStm = (flags & 0x0200) !== 0; // bit 9
    var fObfuscated = (flags & 0x8000) !== 0;  // bit 15 (XOR obfuscation)
    if (fEncrypted || fObfuscated) return null;           // we do not decrypt

    // Walk the variable-length FIB to find fibRgLw (for ccpText) and the
    // FibRgFcLcb blob (for fcClx/lcbClx). Counts are read from the file rather
    // than hard-coded, so the same code works for the Word 97/2000/2002/2003
    // FIB variants (which only ever extend this prefix).
    var pos = 0x20;
    var csw = dv.getUint16(pos, true); pos += 2;          // count of 16-bit
    pos += csw * 2;                                        // skip fibRgW
    var cslw = dv.getUint16(pos, true); pos += 2;          // count of 32-bit
    var fibRgLwStart = pos;
    pos += cslw * 4;                                        // skip fibRgLw
    /* cbRgFcLcb */ dv.getUint16(pos, true); pos += 2;
    var fibRgFcLcbStart = pos;

    // ccpText = fibRgLw[3] (count of chars in the main document).
    if (fibRgLwStart + 16 > wd.length) return null;
    var ccpText = dv.getUint32(fibRgLwStart + 3 * 4, true);
    if (ccpText <= 0) return '';

    // fcClx / lcbClx = pair index 33 of FibRgFcLcb97.
    var clxPair = fibRgFcLcbStart + 33 * 8;
    if (clxPair + 8 > wd.length) return null;
    var fcClx = dv.getUint32(clxPair, true);
    var lcbClx = dv.getUint32(clxPair + 4, true);
    if (lcbClx === 0) return null;                         // no piece table

    // The CLX lives in the table stream chosen by fWhichTblStm.
    var table = cfb.byName[fWhichTblStm ? '1Table' : '0Table'];
    if (!table) table = cfb.byName[fWhichTblStm ? '0Table' : '1Table'];
    if (!table) return null;
    var tableBytes = cfb.getStream(table);
    if (fcClx >= tableBytes.length) return null;
    if (fcClx + lcbClx > tableBytes.length) lcbClx = tableBytes.length - fcClx;

    var pieces = parsePieceTable(tableBytes, fcClx, lcbClx);
    if (!pieces) return null;

    return extractText(wd, pieces, ccpText);
  }

  // Parse the CLX: zero or more Prc records, then a Pcdt holding the PlcPcd
  // (the piece table). See [MS-DOC] structures "Clx", "Pcdt", "PlcPcd", "Pcd",
  // "FcCompressed", and the master algorithm [MS-DOC] 2.4.1 "Retrieving Text".
  function parsePieceTable(table, fcClx, lcbClx) {
    var dv = new DataView(table.buffer, table.byteOffset, table.byteLength);
    var p = fcClx;
    var end = fcClx + lcbClx;
    var plcOff = -1, plcLen = 0;

    while (p < end) {
      var clxt = table[p];
      if (clxt === 0x01) {              // Prc: 1 + 2 + cbGrpprl bytes
        if (p + 3 > end) break;
        var cbGrpprl = dv.getInt16(p + 1, true);
        p += 3 + cbGrpprl;
      } else if (clxt === 0x02) {       // Pcdt: clxt + lcb(4) + PlcPcd(lcb)
        if (p + 5 > end) break;
        plcLen = dv.getUint32(p + 1, true);
        plcOff = p + 5;
        break;
      } else {
        break;
      }
    }
    if (plcOff < 0) return null;
    if (plcOff + plcLen > table.length) plcLen = table.length - plcOff;

    // PlcPcd = (n+1) CPs (4 bytes each) followed by n PCDs (8 bytes each).
    var n = Math.floor((plcLen - 4) / 12);
    if (n < 1) return null;

    var cps = new Array(n + 1);
    for (var i = 0; i <= n; i++) cps[i] = dv.getUint32(plcOff + i * 4, true);
    var pcdBase = plcOff + (n + 1) * 4;

    var pieces = [];
    for (var j = 0; j < n; j++) {
      var fcVal = dv.getUint32(pcdBase + j * 8 + 2, true); // Pcd.fc (FcCompressed)
      var compressed = (fcVal & 0x40000000) !== 0;          // bit 30 = fCompressed
      var fc = fcVal & 0x3FFFFFFF;
      pieces.push({
        cpStart: cps[j],
        cpEnd: cps[j + 1],
        // compressed: 1 cp1252 byte/char at fc/2; else 2 UTF-16LE bytes/char at fc
        offset: compressed ? (fc >>> 1) : fc,
        compressed: compressed
      });
    }
    return pieces;
  }

  // Walk pieces in CP order, decode, strip field codes and control marks,
  // stopping once ccpText main-document characters have been consumed.
  // (Footnote/header/footer ranges follow ccpText in the same piece table;
  // a v2 would continue past ccpText for ccpFtn/ccpHdd and route those into
  // separate buckets.)
  function extractText(wd, pieces, ccpText) {
    var out = [];
    var fieldStack = []; // per open field: true while inside its instruction
    var consumed = 0;

    for (var i = 0; i < pieces.length && consumed < ccpText; i++) {
      var pc = pieces[i];
      var count = pc.cpEnd - pc.cpStart;
      if (count <= 0) continue;
      if (consumed + count > ccpText) count = ccpText - consumed;

      if (pc.compressed) {
        var b0 = pc.offset;
        for (var k = 0; k < count; k++) {
          var b = wd[b0 + k];
          if (b === undefined) break;
          emit(out, fieldStack,
            b < 0x80 ? b : (b <= 0x9F ? FC_COMPRESSED_MAP[b - 0x80] : b));
        }
      } else {
        var u0 = pc.offset;
        for (var m = 0; m < count; m++) {
          var lo = wd[u0 + m * 2];
          if (lo === undefined) break;
          emit(out, fieldStack, lo | ((wd[u0 + m * 2 + 1] || 0) << 8));
        }
      }
      consumed += count;
    }
    return out.join('');
  }

  // Field markers ([MS-DOC] "Special Characters"): 0x13 begin, 0x14 separator,
  // 0x15 end.
  // Text in the instruction region (0x13..0x14) is the field code -> dropped;
  // text in the result region (0x14..0x15) is kept. Nesting is tracked with a
  // stack so a nested field inside an outer instruction is dropped too.
  function emit(out, fieldStack, code) {
    if (code === 0x13) { fieldStack.push(true); return; }
    if (code === 0x14) { if (fieldStack.length) fieldStack[fieldStack.length - 1] = false; return; }
    if (code === 0x15) { if (fieldStack.length) fieldStack.pop(); return; }
    for (var i = 0; i < fieldStack.length; i++) {
      if (fieldStack[i]) return; // inside some field's instruction
    }
    var ch = mapChar(code);
    if (ch !== '') out.push(ch);
  }

  // Map a decoded character to output text. Structural control marks become
  // whitespace; special placeholders are dropped.
  function mapChar(code) {
    switch (code) {
      case 0x07: return '\t';     // cell mark (tables collapse to tabs)
      case 0x09: return '\t';     // tab
      case 0x0A: return '\n';     // line feed
      case 0x0B: return '\n';     // manual line break (Shift+Enter)
      case 0x0C: return '\n';     // page/section break
      case 0x0D: return '\n';     // paragraph mark
      case 0x1E: return '-';      // non-breaking hyphen
      case 0x1F: return '';       // optional (soft) hyphen
      case 0xA0: return ' '; // non-breaking space
      case 0x0001: return '';     // embedded object / picture placeholder
      case 0x0002: return '';     // auto-numbered footnote reference
      case 0x0003: return '';     // short horizontal line (rare)
      case 0x0004: return '';     // reserved
      case 0x0005: return '';     // annotation reference
      case 0x0008: return '';     // drawn object anchor
    }
    if (code < 0x20) return '';   // drop any other control character
    return String.fromCharCode(code);
  }

  // ---- helpers ------------------------------------------------------------
  function toUint8(input) {
    if (input == null) return null;
    if (input instanceof Uint8Array) return input; // also covers Node Buffer
    if (typeof ArrayBuffer !== 'undefined' && input instanceof ArrayBuffer) {
      return new Uint8Array(input);
    }
    if (input.buffer && typeof input.byteLength === 'number') {
      return new Uint8Array(input.buffer, input.byteOffset || 0, input.byteLength);
    }
    return null;
  }

  function concat(chunks, limit) {
    var total = 0, i;
    for (i = 0; i < chunks.length; i++) total += chunks[i].length;
    if (limit != null && limit < total) total = limit;
    var out = new Uint8Array(total);
    var pos = 0;
    for (i = 0; i < chunks.length && pos < total; i++) {
      var ch = chunks[i];
      var take = Math.min(ch.length, total - pos);
      out.set(take === ch.length ? ch : ch.subarray(0, take), pos);
      pos += take;
    }
    return out;
  }

  return docToText;
}));
