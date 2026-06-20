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
 * text and paragraph breaks. No fonts, images, or styles. Tables collapse to
 * tab/newline text. Headers, footers, footnotes (and endnotes, comments,
 * textboxes) are available as separate stories via docToText.sections().
 * Tracked-change *deletions* are dropped (their text is identified via the
 * sprmCFRMarkDel revision mark in the CHPX bin table); tracked *insertions* are
 * kept, so the output reflects the document with all changes accepted.
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
  // Public entry points
  // ------------------------------------------------------------------------

  // Parse the whole document into its text stories, or null on failure.
  function parse(input) {
    try {
      var bytes = toUint8(input);
      if (!bytes || bytes.length < 512) return null;
      var cfb = parseCfb(bytes);
      if (!cfb) return null;
      var wordDocument = cfb.byName['WordDocument'];
      if (!wordDocument) return null;
      return parseWord(cfb.getStream(wordDocument), cfb);
    } catch (e) {
      return null; // any failure -> graceful handoff
    }
  }

  // docToText(input) -> main-body text, or null (unchanged, back-compatible).
  function docToText(input) {
    var doc = parse(input);
    return doc ? doc.body : null;
  }

  // docToText.sections(input) -> { body, footnotes, headers, annotations,
  // endnotes, textboxes, headerTextboxes } (each a string; "" when empty), or
  // null. These are the document's separate text stories, which follow the
  // main body consecutively in the same piece table. `body` === docToText().
  // The object also carries `.html` (see below).
  docToText.sections = parse;

  // docToText.html(input) -> { body, footnotes, ... } where each is styled HTML
  // for that story: text runs wrapped in <span> with the run's character
  // formatting (bold/italic/underline/strike/size/color/font), with \t between
  // table cells and \n at row/paragraph breaks. Or null on failure.
  docToText.html = function (input) { var doc = parse(input); return doc ? doc.html : null; };
  docToText.model = function (input) { var doc = parse(input); return doc ? doc.model : null; };

  // docToText.images(input) -> [{ mime, bytes }] for embedded raster images
  // (PNG/JPEG), best effort. Word stores pictures/OLE images as raw image bytes
  // somewhere in the CFB streams; we carve complete images by signature from the
  // reassembled streams (CFB sector fragmentation already handled), validating
  // each by parsing to its real end marker. WMF/EMF metafiles aren't raster and
  // can't render in-browser, so they're skipped; exact inline placement isn't
  // reconstructed. Returns null on failure (not a CFB).
  docToText.images = function (input) {
    try { var cfb = parseCfb(toUint8(input)); return cfb ? carveImages(cfb) : null; }
    catch (e) { return null; }
  };

  function matchSig(b, i, sig) {
    if (i + sig.length > b.length) return false;
    for (var k = 0; k < sig.length; k++) if (b[i + k] !== sig[k]) return false;
    return true;
  }
  // End of a PNG: walk chunks (length+type+data+crc) to IEND. -1 if malformed.
  function pngEnd(b, start) {
    var p = start + 8;
    while (p + 12 <= b.length) {
      var len = b[p] * 0x1000000 + (b[p + 1] << 16) + (b[p + 2] << 8) + b[p + 3];
      var type = String.fromCharCode(b[p + 4], b[p + 5], b[p + 6], b[p + 7]);
      p += 12 + len;
      if (type === 'IEND') return p <= b.length ? p : -1;
      if (len < 0 || p > b.length) return -1;
    }
    return -1;
  }
  // End of a JPEG: parse markers/segments to EOI (FF D9). -1 if malformed.
  function jpegEnd(b, start) {
    var p = start + 2;
    while (p + 1 < b.length) {
      if (b[p] !== 0xFF) { p++; continue; }
      var m = b[p + 1];
      if (m === 0xD9) return p + 2;                                  // EOI
      if (m === 0x01 || m === 0xFF || (m >= 0xD0 && m <= 0xD8)) { p += 2; continue; }
      if (p + 3 >= b.length) return -1;
      p += 2 + ((b[p + 2] << 8) | b[p + 3]);                         // skip segment
      if (m === 0xDA) {                                              // SOS -> scan entropy data
        while (p + 1 < b.length && !(b[p] === 0xFF && b[p + 1] !== 0x00 && !(b[p + 1] >= 0xD0 && b[p + 1] <= 0xD7))) p++;
      }
    }
    return -1;
  }
  function carveImages(cfb) {
    var SIGS = [
      { mime: 'image/png', sig: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], end: pngEnd },
      { mime: 'image/jpeg', sig: [0xff, 0xd8, 0xff], end: jpegEnd }
    ];
    var images = [], seen = {};
    for (var name in cfb.byName) {
      var b = cfb.getStream(cfb.byName[name]);
      for (var i = 0; i + 3 < b.length; i++) {
        for (var s = 0; s < SIGS.length; s++) {
          if (matchSig(b, i, SIGS[s].sig)) {
            var end = SIGS[s].end(b, i);
            if (end > i + 32) {                          // a real image, not a stray signature
              var key = SIGS[s].mime + ':' + (end - i);
              if (!seen[key]) { seen[key] = 1; images.push({ mime: SIGS[s].mime, bytes: b.subarray(i, end) }); }
              i = end - 1;
            }
            break;
          }
        }
      }
    }
    return images;
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

    // Character counts of each text "story". They sit consecutively in the
    // piece table after the main body, in this order (ccpMcr at index 6 is
    // reserved and not part of the chain). [MS-DOC] FibRgLw97 + the PlcPcd
    // aCP note (last CP = ccpText + sum of these + 1).
    if (fibRgLwStart + 16 > wd.length) return null;
    function rgLw(i) { var o = fibRgLwStart + i * 4; return o + 4 <= wd.length ? dv.getUint32(o, true) : 0; }
    var ccpText = rgLw(3); if (ccpText < 0) ccpText = 0;
    var ccpFtn = rgLw(4), ccpHdd = rgLw(5), ccpAtn = rgLw(7),
        ccpEdn = rgLw(8), ccpTxbx = rgLw(9), ccpHdrTxbx = rgLw(10);

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

    // Character runs (formatting + tracked-deletion flag). Deletions are
    // dropped from every story; the formatting feeds the styled HTML output.
    var chpx = parseChpx(wd, tableBytes, fibRgFcLcbStart, dv);
    var isDeleted = makeIsDeleted(chpx);
    var fonts = parseFonts(tableBytes, fibRgFcLcbStart, dv);
    var styles = parseStsh(tableBytes, fibRgFcLcbStart, dv);
    var papx = parsePapx(wd, tableBytes, fibRgFcLcbStart, dv);
    // resolve(fc): the fully-resolved character props at a WordDocument offset,
    // layered lowest-to-highest priority ([MS-DOC] 2.4.6.2): paragraph style ->
    // character style (sprmCIstd) -> direct run props -> font name.
    var resolve = chpx ? function (fc) {
      var out = {}, k, s;
      if (styles && papx) { var pr = runAt(papx, fc); if (pr && styles[pr.istd]) { s = styles[pr.istd]; for (k in s) out[k] = s[k]; } }
      var r = runAt(chpx, fc), p = r ? r.p : {};
      if (styles && p.istd != null && styles[p.istd]) { s = styles[p.istd]; for (k in s) out[k] = s[k]; }
      for (k in p) if (k !== 'istd') out[k] = p[k];     // direct run props win
      if (fonts && out.ftc != null && fonts[out.ftc]) out.font = fonts[out.ftc];
      // An explicit RGB (sprmCCv) wins; otherwise fall back to the 16-colour
      // palette index (sprmCIco) so both the styled HTML and the model see it.
      if ((out.cv == null || (out.cv & 0xFFFFFF) === 0) && out.ico >= 2 && out.ico <= 16 && ICO_CV[out.ico]) out.cv = ICO_CV[out.ico];
      return out;
    } : null;

    // Story boundaries: body is [0, ccpText); each follows consecutively.
    var bounds = [['body', 0, ccpText]], cp = ccpText;
    function add(name, len) { bounds.push([name, cp, cp + len]); cp += len; }
    add('footnotes', ccpFtn); add('headers', ccpHdd); add('annotations', ccpAtn);
    add('endnotes', ccpEdn); add('textboxes', ccpTxbx); add('headerTextboxes', ccpHdrTxbx);

    var doc = { html: {}, model: {} };
    // Carved images, paired in document order with the body's picture chars so
    // the model (and thus the writer) can round-trip them.
    var modelImages = []; try { modelImages = carveImages(cfb) || []; } catch (e) { }
    var imgCtr = { n: 0 };
    var dataEntry = cfb.byName['Data'], dataStream = null;
    try { dataStream = dataEntry ? cfb.getStream(dataEntry) : null; } catch (e) { }
    function paraAlign(fc) { var r = papx ? runAt(papx, fc) : null; return r ? (r.jc || 0) : 0; }
    var listKind = parseListNfc(tableBytes, fibRgFcLcbStart, dv);
    var footnoteRefCps = parseRefCps(tableBytes, fibRgFcLcbStart, dv, 2); // PlcffndRef #2 -> { bodyCp: footnoteIndex }
    var endnoteRefCps = parseRefCps(tableBytes, fibRgFcLcbStart, dv, 46); // PlcfendRef #46 -> { bodyCp: endnoteIndex }
    var commentRefCps = parseRefCps(tableBytes, fibRgFcLcbStart, dv, 4, 30); // PlcfandRef #4 (ATRD=30) -> { bodyCp: commentIndex }
    var textboxRefCps = parseRefCps(tableBytes, fibRgFcLcbStart, dv, 40, 26); // PlcfspaMom #40 (FSPA=26) -> { bodyCp: shapeIndex }
    function paraList(fc) { var r = papx ? runAt(papx, fc) : null; if (!r || !r.ilfo) return null; return { ilvl: r.ilvl || 0, kind: (listKind && listKind[r.ilfo]) || 'bullet' }; }
    // Paragraph spacing/indentation (twips): left/right/first-line indent, space
    // before/after, and line spacing (LSPD: line + lineMult flag). Only non-zero
    // values are returned, so an unspaced paragraph stays a bare model node.
    function paraPP(fc) {
      var r = papx ? runAt(papx, fc) : null; if (!r) return null;
      var pp = {};
      if (r.indL) pp.indL = r.indL;
      if (r.indR) pp.indR = r.indR;
      if (r.ind1) pp.ind1 = r.ind1;
      if (r.spB) pp.spB = r.spB;
      if (r.spA) pp.spA = r.spA;
      if (r.line) { pp.line = r.line; pp.lineMult = r.lineMult || 0; }
      if (r.keepN) pp.keepNext = 1;       // keep with next paragraph
      if (r.keepL) pp.keepLines = 1;      // keep lines together
      if (r.pgBrk) pp.pageBreak = 1;      // page break before
      return Object.keys(pp).length ? pp : null;
    }
    // The table row's column boundaries (rgdxaCenter twips), from sprmTDefTable on
    // the row-terminator paragraph; null for non-table paragraphs.
    function paraTblw(fc) { var r = papx ? runAt(papx, fc) : null; return (r && r.tblw) ? r.tblw : null; }
    function paraTblShd(fc) { var r = papx ? runAt(papx, fc) : null; return (r && r.tblShd) ? r.tblShd : null; }
    function paraTblMerge(fc) { var r = papx ? runAt(papx, fc) : null; return (r && r.tblMerge) ? r.tblMerge : null; }
    for (var bi = 0; bi < bounds.length; bi++) {
      var nm = bounds[bi][0], a = bounds[bi][1], b = bounds[bi][2];
      doc[nm] = extractRange(wd, pieces, a, b, isDeleted);
      doc.html[nm] = extractRangeStyled(wd, pieces, a, b, isDeleted, resolve);
      doc.model[nm] = extractRangeModel(wd, pieces, a, b, isDeleted, resolve, nm === 'body' ? modelImages : null, imgCtr, paraAlign, nm === 'body' ? dataStream : null, paraList, paraPP, paraTblw, paraTblShd, paraTblMerge, nm === 'body' ? footnoteRefCps : null, nm === 'body' ? endnoteRefCps : null, nm === 'body' ? commentRefCps : null, nm === 'body' ? textboxRefCps : null);
    }
    // Split the header document (PlcfHdd) into the real page header & footer for
    // the first section — skipping the footnote/endnote separator stories (0-5)
    // and preferring odd-page, then first-page, then even-page. doc.model.header /
    // .footer feed the writer (a clean header/footer, not separator noise).
    var hdd = parsePlcfHdd(tableBytes, fibRgFcLcbStart, dv);
    if (hdd && ccpHdd > 0) {
      var hddStart = ccpText + ccpFtn;
      var pick = function (cands) {
        for (var c = 0; c < cands.length; c++) { var k = cands[c]; if (k + 1 < hdd.length && hdd[k + 1] > hdd[k]) return k; }
        return -1;
      };
      var grab = function (k) {
        return extractRangeModel(wd, pieces, hddStart + hdd[k], hddStart + hdd[k + 1], isDeleted, resolve, null, imgCtr, paraAlign, null, paraList, paraPP, paraTblw, paraTblShd, paraTblMerge, null);
      };
      var hK = pick([7, 10, 6]), fK = pick([9, 11, 8]);   // header: odd/first/even; footer: odd/first/even
      if (hK >= 0) doc.model.header = grab(hK);
      if (fK >= 0) doc.model.footer = grab(fK);
    }
    // Page setup: the first section's properties (margins, page size, orientation).
    var sections = parseSections(tableBytes, fibRgFcLcbStart, dv, wd);
    if (sections) {
      var sx = sections.filter(function (s) { return s; });
      if (sx.length) doc.model.sections = sx;   // every section's setup
      if (sections[0]) doc.model.page = sections[0];   // first section (back-compat)
    }
    // Document properties (title/author/...) from the \x05SummaryInformation stream.
    var props = parseSummaryInfo(cfb);
    if (props) doc.model.props = props;
    // Floating-shape (text-box) positions, so the demo can place them on the page.
    var shapes = parseShapes(tableBytes, fibRgFcLcbStart, dv);
    if (shapes) doc.model.shapes = shapes;
    return doc;
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

  // ---- character runs & formatting ([MS-DOC] CHPX bin table) ---------------
  // Every run's character properties live in CHPX FKP "pages" indexed by
  // PlcfBteChpx (FibRgFcLcb97 #12). parseChpx returns the runs in WordDocument
  // byte order, each with the direct character properties we use: the tracked-
  // deletion flag plus bold/italic/strike/underline/size/color/font/char-style.
  function parseChpx(wd, table, fibRgFcLcbStart, fibDv) {
    try {
      var pair = fibRgFcLcbStart + 12 * 8;
      if (pair + 8 > wd.length) return null;
      var fc = fibDv.getUint32(pair, true);
      var lcb = fibDv.getUint32(pair + 4, true);
      if (lcb < 4 || fc < 0 || fc + lcb > table.length) return null;
      var tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var n = Math.floor((lcb - 4) / 8);          // PlcfBteChpx: n+1 FCs + n PNs
      if (n < 1) return null;
      var pnBase = fc + (n + 1) * 4;
      var runs = [];
      for (var i = 0; i < n; i++) {
        var pn = tdv.getUint32(pnBase + i * 4, true) & 0x003FFFFF; // PnFkpChpx.pn
        collectFkpRuns(wd, pn * 512, runs);
      }
      runs.sort(function (x, y) { return x.a - y.a; });
      return runs;
    } catch (e) { return null; }
  }

  // One ChpxFkp page (512 bytes at pn*512): crun at byte 511, rgfc[crun+1],
  // then rgb[crun] (word offsets to each Chpx; 0 = default props).
  function collectFkpRuns(wd, pageOff, runs) {
    if (pageOff < 0 || pageOff + 512 > wd.length) return;
    var dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    var crun = wd[pageOff + 511];
    if (!crun) return;
    var rgbBase = pageOff + 4 * (crun + 1);
    for (var i = 0; i < crun; i++) {
      var a = dv.getUint32(pageOff + i * 4, true);
      var b = dv.getUint32(pageOff + (i + 1) * 4, true);
      if (b <= a) continue;
      var props = {};
      var word = wd[rgbBase + i];
      if (word) {
        var chpxOff = pageOff + word * 2;             // Chpx: cb, then grpprl
        if (chpxOff >= pageOff && chpxOff < pageOff + 512) {
          var cb = wd[chpxOff];
          if (chpxOff + 1 + cb <= pageOff + 512) props = parseChpGrpprl(wd, chpxOff + 1, cb);
        }
      }
      runs.push({ a: a, b: b, p: props });
    }
  }

  // Decode the character sprms we render from a Chpx grpprl. Codes verified
  // empirically against real documents. ToggleOperand props (bold/italic/...)
  // are "on" exactly when bit 0 of the operand is set (covers 0x01 and 0x81 —
  // 0x81 = "invert the off style default" = on; the same rule as deletions).
  function parseChpGrpprl(wd, off, len) {
    var p = {}, q = off, end = off + len;
    while (q + 2 <= end) {
      var sprm = wd[q] | (wd[q + 1] << 8); q += 2;
      switch (sprm) {
        case 0x0800: p.del = (wd[q] & 1) === 1; break;        // sprmCFRMarkDel (deletion)
        case 0x0835: p.b = (wd[q] & 1) === 1; break;          // bold
        case 0x0836: p.i = (wd[q] & 1) === 1; break;          // italic
        case 0x0837: p.strike = (wd[q] & 1) === 1; break;     // strikethrough
        case 0x083C: p.hidden = (wd[q] & 1) === 1; break;     // vanish (hidden text)
        case 0x2A3E: p.u = wd[q]; break;                       // underline kind (0 = none)
        case 0x4A43: p.hps = wd[q] | (wd[q + 1] << 8); break;  // font size, half-points
        case 0x2A42: p.ico = wd[q]; break;                     // color, 16-colour palette index
        case 0x6870: p.cv = (wd[q] | (wd[q + 1] << 8) | (wd[q + 2] << 16)) >>> 0; break; // 24-bit RGB
        case 0x4A4F: p.ftc = wd[q] | (wd[q + 1] << 8); break;  // font index (ftc0) -> SttbfFfn
        case 0x4A30: p.istd = wd[q] | (wd[q + 1] << 8); break; // character style index
        case 0x6A03: p.picLoc = (wd[q] | (wd[q + 1] << 8) | (wd[q + 2] << 16) | (wd[q + 3] << 24)) >>> 0; break; // sprmCPicLocation (Data offset)
        case 0x0855: p.fSpec = wd[q] & 1; break;               // sprmCFSpec (special char, e.g. picture)
      }
      q += sprmOperandLen(sprm, wd, q);
    }
    return p;
  }

  // Operand size from the sprm's spra field (bits 13-15). spra 6 is variable:
  // a 1-byte length precedes the operand (table sprms with 2-byte lengths don't
  // occur in CHPX grpprls).
  function sprmOperandLen(sprm, buf, pos) {
    switch ((sprm >> 13) & 7) {
      case 0: case 1: return 1;
      case 2: case 4: case 5: return 2;
      case 3: return 4;
      case 7: return 3;
      case 6: return 1 + (buf[pos] || 0);
      default: return 1;
    }
  }

  // Binary-search the CHPX run containing a WordDocument byte offset.
  function runAt(runs, fc) {
    var lo = 0, hi = runs.length - 1;
    while (lo <= hi) {
      var mid = (lo + hi) >> 1;
      if (fc < runs[mid].a) hi = mid - 1;
      else if (fc >= runs[mid].b) lo = mid + 1;
      else return runs[mid];
    }
    return null;
  }

  // Tracked-deletion predicate derived from the CHPX runs.
  function makeIsDeleted(runs) {
    if (!runs || !runs.length) return null;
    return function (fc) { var r = runAt(runs, fc); return !!(r && r.p.del); };
  }

  // Font table: SttbfFfn (FibRgFcLcb97 #15) -> array of font names indexed by
  // the ftc used in sprmCRgFtc0. The STTB is "extended" (Unicode); each entry
  // is an FFN whose own cbFfnM1 byte gives its length, with the name (a null-
  // terminated UTF-16 string) at byte 40. [MS-DOC] SttbfFfn / FFN.
  function parseFonts(table, fibRgFcLcbStart, fibDv) {
    try {
      var fc = fibDv.getUint32(fibRgFcLcbStart + 15 * 8, true);
      var lcb = fibDv.getUint32(fibRgFcLcbStart + 15 * 8 + 4, true);
      if (lcb < 6 || fc < 0 || fc + lcb > table.length) return null;
      var dv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var end = fc + lcb, p = fc, cData;
      // Header: optional fExtend (0xFFFF) marker, then cData (count) + cbExtra.
      // FFN names are UTF-16 regardless; each FFN is self-sized via cbFfnM1.
      if (dv.getUint16(p, true) === 0xFFFF) { p += 2; cData = dv.getUint16(p, true); p += 4; }
      else { cData = dv.getUint16(p, true); p += 4; }
      var fonts = [];
      for (var i = 0; i < cData && p < end; i++) {
        var ffnEnd = p + table[p] + 1;                   // table[p] = cbFfnM1
        var name = '';
        for (var q = p + 40; q + 1 < ffnEnd && q + 1 < end; q += 2) {
          var ch = table[q] | (table[q + 1] << 8);
          if (ch === 0) break;
          name += String.fromCharCode(ch);
        }
        fonts.push(name);
        p = ffnEnd;
      }
      return fonts;
    } catch (e) { return null; }
  }

  // Stylesheet: STSH (Stshf at FibRgFcLcb97 #1) -> resolved character props per
  // style index (istd), so formatting carried by a style (a heading's bold, a
  // hyperlink's blue+underline) isn't lost. Each STD has a CHP grpprl and an
  // istdBase parent it inherits from. [MS-DOC] STSH / STSHI / STD / STDF / UPX.
  function parseStsh(table, fibStart, fibDv) {
    try {
      var fc = fibDv.getUint32(fibStart + 1 * 8, true);   // #1 = fcStshf
      var lcb = fibDv.getUint32(fibStart + 1 * 8 + 4, true);
      if (lcb < 6 || fc < 0 || fc + lcb > table.length) return null;
      var dv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var end = fc + lcb;
      var cbStshi = dv.getUint16(fc, true), stshi = fc + 2;
      var cstd = dv.getUint16(stshi, true);               // STSHI.cstd
      var cbBase = dv.getUint16(stshi + 2, true);         // STSHI.cbSTDBaseInFile
      var p = stshi + cbStshi;                            // -> rgStd
      // Pass 1: each style's own CHP grpprl + parent.
      var raw = new Array(cstd);
      for (var i = 0; i < cstd && p + 2 <= end; i++) {
        var cbStd = dv.getUint16(p, true); p += 2;
        var std = p; p += cbStd;
        if (cbStd === 0 || std + cbBase + 2 > end) { raw[i] = null; continue; }
        var word1 = dv.getUint16(std + 2, true);
        var stk = word1 & 0xF;                            // style kind (1=para,2=char), low 4 bits
        var istdBase = (word1 >> 4) & 0x0FFF;             // parent style (0xFFF = none), high 12 bits
        var cupx = dv.getUint16(std + 4, true) & 0xF;     // count of UPX, low 4 bits
        var nameAt = std + cbBase, cch = dv.getUint16(nameAt, true);
        var up = nameAt + 2 + cch * 2 + 2;                // past style name + chTerm
        var chpIdx = stk === 1 ? 1 : 0;                   // para style: CHP is 2nd UPX
        var chp = {};
        for (var u = 0; u < cupx && up + 2 <= std + cbStd; u++) {
          var cbUpx = dv.getUint16(up, true); up += 2;
          if (u === chpIdx) chp = parseChpGrpprl(table, up, cbUpx);
          up += cbUpx; if (up & 1) up++;                  // UPXs are padded to even
        }
        raw[i] = { base: istdBase, chp: chp };
      }
      // Pass 2: resolve inheritance (parent props, then own props).
      var out = new Array(cstd);
      function resolve(i, depth) {
        if (i == null || i < 0 || i >= cstd || !raw[i] || depth > 24) return {};
        if (out[i]) return out[i];
        var r = {}, k, base = raw[i].base !== 0x0FFF ? resolve(raw[i].base, depth + 1) : {};
        for (k in base) r[k] = base[k];
        for (k in raw[i].chp) r[k] = raw[i].chp[k];
        return (out[i] = r);
      }
      for (var j = 0; j < cstd; j++) resolve(j, 0);
      return out;
    } catch (e) { return null; }
  }

  // Paragraph bin table: PlcfBtePapx (FibRgFcLcb97 #13) -> the paragraph style
  // index (istd) for each WordDocument byte range, so a run inherits its
  // paragraph style's character formatting (e.g. a heading's bold). We only
  // need each paragraph's istd. [MS-DOC] PlcfBtePapx / PapxFkp / PapxInFkp.
  function parsePapx(wd, table, fibStart, fibDv) {
    try {
      var fc = fibDv.getUint32(fibStart + 13 * 8, true);
      var lcb = fibDv.getUint32(fibStart + 13 * 8 + 4, true);
      if (lcb < 4 || fc < 0 || fc + lcb > table.length) return null;
      var tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var n = Math.floor((lcb - 4) / 8);
      if (n < 1) return null;
      var pnBase = fc + (n + 1) * 4, runs = [];
      for (var i = 0; i < n; i++) {
        var pn = tdv.getUint32(pnBase + i * 4, true) & 0x003FFFFF;
        collectPapxFkp(wd, pn * 512, runs);
      }
      runs.sort(function (x, y) { return x.a - y.a; });
      return runs;
    } catch (e) { return null; }
  }

  // One PapxFkp page: crun at byte 511, rgfc[crun+1], then rgbx[crun] (BxPap,
  // 13 bytes each; first byte is the word offset to a PapxInFkp, 0 = default).
  function collectPapxFkp(wd, pageOff, runs) {
    if (pageOff < 0 || pageOff + 512 > wd.length) return;
    var dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
    var crun = wd[pageOff + 511];
    if (!crun) return;
    var bxBase = pageOff + 4 * (crun + 1);
    for (var i = 0; i < crun; i++) {
      var a = dv.getUint32(pageOff + i * 4, true);
      var b = dv.getUint32(pageOff + (i + 1) * 4, true);
      if (b <= a) continue;
      var bOff = wd[bxBase + i * 13], istd = 0, jc = 0, ilfo = 0, ilvl = 0;  // BxPap.bOffset
      var indL = 0, indR = 0, ind1 = 0, spB = 0, spA = 0, line = 0, lineMult = 0, tblw = null, keepN = 0, keepL = 0, pgBrk = 0, tblShd = null, tblMerge = null;
      if (bOff) {
        var papx = pageOff + bOff * 2;                   // PapxInFkp
        if (papx >= pageOff && papx < pageOff + 510) {
          var cb = wd[papx];                             // GrpPrlAndIstd starts at:
          var g = cb !== 0 ? papx + 1 : papx + 2;        // cb!=0 -> +1, else +2
          if (g + 2 <= pageOff + 512) istd = wd[g] | (wd[g + 1] << 8);
          // walk the grpprl (after the istd) for alignment, list, spacing/indent
          var grpEnd = cb !== 0 ? papx + 2 * cb : papx + 2 + 2 * (wd[papx + 1] || 0);
          if (grpEnd > pageOff + 512) grpEnd = pageOff + 512;
          for (var gp = g + 2; gp + 2 <= grpEnd;) {
            var sc = wd[gp] | (wd[gp + 1] << 8), ol = sprmOperandLen(sc, wd, gp + 2);
            if (sc === 0x2403 || sc === 0x2461) jc = wd[gp + 2];        // sprmPJc80 / sprmPJc
            else if (sc === 0x460B) ilfo = wd[gp + 2] | (wd[gp + 3] << 8); // sprmPIlfo (list)
            else if (sc === 0x260A) ilvl = wd[gp + 2];                     // sprmPIlvl (level)
            else if (sc === 0x2405) keepL = wd[gp + 2];                    // sprmPFKeep (keep lines together)
            else if (sc === 0x2406) keepN = wd[gp + 2];                    // sprmPFKeepFollow (keep with next)
            else if (sc === 0x2407) pgBrk = wd[gp + 2];                    // sprmPFPageBreakBefore
            else if (sc === 0x840F) indL = dv.getInt16(gp + 2, true);      // sprmPDxaLeft
            else if (sc === 0x840E) indR = dv.getInt16(gp + 2, true);      // sprmPDxaRight
            else if (sc === 0x8411) ind1 = dv.getInt16(gp + 2, true);      // sprmPDxaLeft1 (first line; <0 = hanging)
            else if (sc === 0xA413) spB = dv.getInt16(gp + 2, true);       // sprmPDyaBefore
            else if (sc === 0xA414) spA = dv.getInt16(gp + 2, true);       // sprmPDyaAfter
            else if (sc === 0x6412) { line = dv.getInt16(gp + 2, true); lineMult = wd[gp + 4] | (wd[gp + 5] << 8); } // sprmPDyaLine (LSPD)
            else if (sc === 0xD608) {                                      // sprmTDefTable: capture rgdxaCenter (column boundaries, twips)
              var tcb = wd[gp + 2] | (wd[gp + 3] << 8); ol = 1 + tcb;      // 2-byte cb exception; operand bytes = cb + 1 (Word writes cb = dataLen + 1)
              var itc = wd[gp + 4];                                        // itcMac (number of columns)
              if (itc > 0 && itc < 64 && gp + 5 + (itc + 1) * 2 <= pageOff + 512) {
                tblw = []; for (var tt = 0; tt <= itc; tt++) tblw.push(dv.getInt16(gp + 5 + tt * 2, true));
                // rgTc80 follows rgdxaCenter; each TC80 is 20 bytes and opens with tcgrf (cell-merge flags)
                var tcB = gp + 5 + (itc + 1) * 2, anyM = false, mg = [];
                for (var mc = 0; mc < itc; mc++) {
                  var off = tcB + mc * 20, gf = (off + 1 < pageOff + 512) ? (wd[off] | (wd[off + 1] << 8)) : 0;
                  var hm = (gf & 0x0001) ? 'start' : (gf & 0x0002) ? 'cont' : null;  // fFirstMerged / fMerged
                  var vm = (gf & 0x0040) ? 'restart' : (gf & 0x0020) ? 'cont' : null; // fVertRestart / fVertMerge
                  if (hm || vm) anyM = true;
                  mg.push((hm || vm) ? { h: hm, v: vm } : null);
                }
                if (anyM) tblMerge = mg;
              }
            }
            else if (sc === 0xD612) {                                      // sprmTDefTableShd: per-cell background (Shd, 10 bytes each)
              var scb = wd[gp + 2], sn = Math.floor(scb / 10);             // 1-byte cb; cvFore is the fill (R,G,B,fAuto)
              if (sn > 0 && sn < 64 && gp + 3 + sn * 10 <= pageOff + 512) { tblShd = []; for (var sj = 0; sj < sn; sj++) { var so = gp + 3 + sj * 10; tblShd.push(wd[so + 3] === 0 ? (wd[so] | (wd[so + 1] << 8) | (wd[so + 2] << 16)) : null); } }
            }
            gp += 2 + ol; if (ol <= 0) break;
          }
        }
      }
      runs.push({ a: a, b: b, istd: istd, jc: jc, ilfo: ilfo, ilvl: ilvl, indL: indL, indR: indR, ind1: ind1, spB: spB, spA: spA, line: line, lineMult: lineMult, tblw: tblw, keepN: keepN, keepL: keepL, pgBrk: pgBrk, tblShd: tblShd, tblMerge: tblMerge });
    }
  }

  // PlfLst (FibRgFcLcb #73) + PlfLfo (#74) -> ilfo (1-based) -> 'bullet'|'number'
  // from each list's level-0 number format (nfc: 23 = bullet; 0-22 = decimal/
  // roman/letters). Best-effort: missing/unreadable LVLs default to bullet.
  // A reference PLC (PlcffndRef #2 etc.): N+1 CPs — the first N are the CPs of the
  // reference characters (0x02) in the main story, the last is the doc-end lim —
  // plus N 2-byte FRD records. Returns { bodyCp: refIndex } for the N references.
  function parseRefCps(table, fibStart, fibDv, idx, dataSize) {
    try {
      var cb = dataSize || 2;   // FRD (footnote/endnote) = 2 bytes; ATRD (comments) = 30
      var fc = fibDv.getUint32(fibStart + idx * 8, true), lcb = fibDv.getUint32(fibStart + idx * 8 + 4, true);
      if (lcb < 4 + cb || fc < 0 || fc + lcb > table.length) return null;
      var n = Math.floor((lcb - 4) / (4 + cb));
      if (n < 1) return null;
      var dv = new DataView(table.buffer, table.byteOffset, table.byteLength), map = {};
      for (var i = 0; i < n; i++) map[dv.getUint32(fc + i * 4, true)] = i;
      return map;
    } catch (e) { return null; }
  }

  // PlcfHdd (#11): CPs (relative to the header document) delimiting its stories.
  // Stories 0-5 are footnote/endnote separators; then 6 per section — even/odd/
  // first page header, then even/odd/first page footer. Returns the CP array.
  function parsePlcfHdd(table, fibStart, fibDv) {
    try {
      var fc = fibDv.getUint32(fibStart + 11 * 8, true), lcb = fibDv.getUint32(fibStart + 11 * 8 + 4, true);
      if (lcb < 8 || fc < 0 || fc + lcb > table.length) return null;
      var n = Math.floor(lcb / 4), dv = new DataView(table.buffer, table.byteOffset, table.byteLength), cps = [];
      for (var i = 0; i < n; i++) cps.push(dv.getUint32(fc + i * 4, true));
      return cps;
    } catch (e) { return null; }
  }

  // Document properties from the \x05SummaryInformation stream (an [MS-OLEPS]
  // property set): title/subject/author/keywords/comments. Reads the property-set
  // offset from the header, then walks its (propId, offset) table for the VT_LPSTR
  // (or VT_LPWSTR) string values. Best-effort; returns null if absent/unreadable.
  function parseSummaryInfo(cfb) {
    try {
      var entry = cfb.byName['\x05SummaryInformation'];
      if (!entry) return null;
      var s = cfb.getStream(entry);                              // byName holds the dir entry; read its bytes
      if (!s || s.length < 48) return null;
      var dv = new DataView(s.buffer, s.byteOffset, s.byteLength);
      if (dv.getUint16(0, true) !== 0xFFFE) return null;          // byte-order mark
      var ps = dv.getUint32(44, true);                            // offset to the property set (28-byte header + 16-byte FMTID)
      if (ps + 8 > s.length) return null;
      var num = dv.getUint32(ps + 4, true);
      if (num < 1 || num > 64) return null;
      var names = { 2: 'title', 3: 'subject', 4: 'author', 5: 'keywords', 6: 'comments' }, out = {};
      for (var i = 0; i < num; i++) {
        var ent = ps + 8 + i * 8;
        if (ent + 8 > s.length) break;
        var pid = dv.getUint32(ent, true), off = ps + dv.getUint32(ent + 4, true);
        if (!names[pid] || off + 8 > s.length) continue;
        var type = dv.getUint32(off, true), cch = dv.getUint32(off + 4, true);
        if ((type !== 0x1E && type !== 0x1F) || cch < 1) continue; // VT_LPSTR / VT_LPWSTR
        var str = '';
        if (type === 0x1E) { if (off + 8 + cch > s.length) continue; for (var c = 0; c < cch - 1; c++) str += String.fromCharCode(s[off + 8 + c]); }
        else { if (off + 8 + cch * 2 > s.length) continue; for (var c2 = 0; c2 < cch - 1; c2++) str += String.fromCharCode(dv.getUint16(off + 8 + c2 * 2, true)); }
        str = str.replace(/\0+$/, '');
        if (str) out[names[pid]] = str;
      }
      return Object.keys(out).length ? out : null;
    } catch (e) { return null; }
  }

  // Floating-shape positions from PlcfspaMom (FibRgFcLcb #40): each shape's FSPA
  // bounding box (xaLeft/yaTop/xaRight/yaBottom in twips) plus its anchor CP, so the
  // demo can place a text box where the document actually puts it on the page rather
  // than at its text anchor. PLC layout: N+1 CPs (4 bytes) then N FSPA records (26).
  function parseShapes(table, fibStart, fibDv) {
    try {
      var fc = fibDv.getUint32(fibStart + 40 * 8, true), lcb = fibDv.getUint32(fibStart + 40 * 8 + 4, true);
      if (lcb < 4 + 26 || fc < 0 || fc + lcb > table.length) return null;
      var n = Math.floor((lcb - 4) / 30);
      if (n < 1) return null;
      var dv = new DataView(table.buffer, table.byteOffset, table.byteLength), base = fc + (n + 1) * 4, shapes = [];
      for (var i = 0; i < n; i++) {
        var o = base + i * 26;
        if (o + 26 > table.length) break;
        shapes.push({ cp: dv.getUint32(fc + i * 4, true), xL: dv.getInt32(o + 4, true), yT: dv.getInt32(o + 8, true), xR: dv.getInt32(o + 12, true), yB: dv.getInt32(o + 16, true) });
      }
      return shapes.length ? shapes : null;
    } catch (e) { return null; }
  }

  // A section's page setup from its SEPX (a grpprl of section sprms): margins
  // (sprmSDyaTop 0x9023 / sprmSDyaBottom 0x9024 signed; sprmSDxaLeft 0xB021 /
  // sprmSDxaRight 0xB022), page size (sprmSXaPage 0xB01F / sprmSYaPage 0xB020),
  // orientation (sprmSBOrientation 0x301D), columns (sprmSCcolumns 0x500B). Twips.
  function sepxProps(wdv, wd, fcSepx) {
    if (fcSepx === 0xFFFFFFFF || fcSepx + 2 > wd.length) return null;
    var cb = wdv.getUint16(fcSepx, true), g = fcSepx + 2, end = fcSepx + 2 + cb, p = {};
    while (g + 2 <= end && g + 4 <= wd.length) {
      var sprm = wdv.getUint16(g, true), spra = (sprm >> 13) & 7;
      var opLen = spra <= 1 ? 1 : (spra === 2 || spra === 4 || spra === 5) ? 2 : spra === 3 ? 4 : spra === 7 ? 3 : (1 + (wd[g + 2] || 0));
      if (sprm === 0x9023) p.top = wdv.getInt16(g + 2, true);
      else if (sprm === 0x9024) p.bottom = wdv.getInt16(g + 2, true);
      else if (sprm === 0xB021) p.left = wdv.getUint16(g + 2, true);
      else if (sprm === 0xB022) p.right = wdv.getUint16(g + 2, true);
      else if (sprm === 0xB01F) p.width = wdv.getUint16(g + 2, true);
      else if (sprm === 0xB020) p.height = wdv.getUint16(g + 2, true);
      else if (sprm === 0x301D) p.landscape = wd[g + 2] === 2;
      else if (sprm === 0x500B) p.cols = wdv.getUint16(g + 2, true) + 1;  // sprmSCcolumns (ccolM1)
      g += 2 + opLen; if (opLen <= 0) break;
    }
    return Object.keys(p).length ? p : null;
  }
  // Every section's page setup (PlcfSed #6: (n+1) CPs then n 12-byte SEDs, each
  // SED.fcSepx -> a SEPX in the WordDocument). One entry per section (null for an
  // unreadable one); a single-section document gives a 1-element array.
  function parseSections(table, fibStart, fibDv, wd) {
    try {
      var sedFc = fibDv.getUint32(fibStart + 6 * 8, true), sedLcb = fibDv.getUint32(fibStart + 6 * 8 + 4, true);
      if (sedLcb < 16) return null;
      var n = Math.floor((sedLcb - 4) / 16);
      if (n < 1) return null;
      var tdv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var wdv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength), base = sedFc + (n + 1) * 4, out = [];
      for (var si = 0; si < n; si++) out.push(sepxProps(wdv, wd, tdv.getUint32(base + si * 12 + 2, true)));
      return out;
    } catch (e) { return null; }
  }

  function parseListNfc(table, fibStart, fibDv) {
    try {
      var lstFc = fibDv.getUint32(fibStart + 73 * 8, true), lstLcb = fibDv.getUint32(fibStart + 73 * 8 + 4, true);
      var lfoFc = fibDv.getUint32(fibStart + 74 * 8, true), lfoLcb = fibDv.getUint32(fibStart + 74 * 8 + 4, true);
      if (!lstLcb || !lfoLcb || lstFc + lstLcb > table.length || lfoFc + lfoLcb > table.length) return null;
      var dv = new DataView(table.buffer, table.byteOffset, table.byteLength);
      var cLst = dv.getUint16(lstFc, true), p = lstFc + 2, lsts = [];
      for (var i = 0; i < cLst; i++) { lsts.push({ lsid: dv.getInt32(p, true), simple: table[p + 26] & 1 }); p += 28; }
      var nfcByLsid = {};
      for (i = 0; i < cLst && p + 28 <= lstFc + lstLcb; i++) {
        var nLvl = lsts[i].simple ? 1 : 9;
        for (var lv = 0; lv < nLvl && p + 28 <= lstFc + lstLcb; lv++) {
          if (lv === 0) nfcByLsid[lsts[i].lsid] = table[p + 4];
          var cchOff = p + 28 + table[p + 25] + table[p + 24];
          var cch = cchOff + 2 <= table.length ? dv.getUint16(cchOff, true) : 0;
          p = cchOff + 2 + cch * 2;
        }
      }
      var cLfo = dv.getUint32(lfoFc, true), q = lfoFc + 4, kind = {};
      for (i = 0; i < cLfo && q + 16 <= lfoFc + lfoLcb; i++) { var nfc = nfcByLsid[dv.getInt32(q, true)]; kind[i + 1] = (nfc != null && nfc < 23) ? 'number' : 'bullet'; q += 16; }
      return kind;
    } catch (e) { return null; }
  }

  // Extract the text of one CP range [lo, hi): the body uses [0, ccpText) and
  // each trailing story (footnotes, headers, ...) its own range. Pieces are in
  // CP order (PlcPcd aCP is sorted), so we take each piece's overlap with the
  // range. Decodes, strips field codes/control marks, and drops any chars whose
  // WordDocument offset is in a tracked-deletion range (isDeleted).
  function extractRange(wd, pieces, lo, hi, isDeleted) {
    if (hi <= lo) return '';
    var out = [];
    var fieldStack = []; // per open field: true while inside its instruction
    var state = { cells: 0 }; // run of pending table cell marks (0x07)

    for (var i = 0; i < pieces.length; i++) {
      var pc = pieces[i];
      var a = lo > pc.cpStart ? lo : pc.cpStart;
      var b = hi < pc.cpEnd ? hi : pc.cpEnd;
      if (a >= b) continue;
      var start = a - pc.cpStart; // first char index within this piece
      var count = b - a;

      if (pc.compressed) {
        var base = pc.offset + start;
        for (var k = 0; k < count; k++) {
          if (isDeleted && isDeleted(base + k)) continue; // tracked-change deletion
          var byte = wd[base + k];
          if (byte === undefined) break;
          emit(out, fieldStack, state,
            byte < 0x80 ? byte : (byte <= 0x9F ? FC_COMPRESSED_MAP[byte - 0x80] : byte));
        }
      } else {
        var u = pc.offset + start * 2;
        for (var m = 0; m < count; m++) {
          if (isDeleted && isDeleted(u + m * 2)) continue; // tracked-change deletion
          var loB = wd[u + m * 2];
          if (loB === undefined) break;
          emit(out, fieldStack, state, loB | ((wd[u + m * 2 + 1] || 0) << 8));
        }
      }
    }
    flushCells(out, state); // a row mark may be the very last thing in a story
    return out.join('');
  }

  // ---- styled HTML output ([MS-DOC] character formatting) -----------------
  // Old Word 16-colour palette for sprmCIco (0=auto, 1=black -> default text).
  var ICO_PALETTE = ['', '', 'blue', 'cyan', 'lime', 'magenta', 'red', 'yellow',
    'white', 'navy', 'teal', 'green', 'purple', 'maroon', 'olive', 'gray', 'silver'];
  // The same 16-colour palette as COLORREF ints (0x00BBGGRR) for the paragraph
  // model, so a model run carries an ico colour the same way it carries an sprmCCv
  // one. Indices 0/1 (auto/black) stay 0 = "default text colour, don't store".
  var ICO_CV = [0, 0, 0xFF0000, 0xFFFF00, 0x00FF00, 0xFF00FF, 0x0000FF, 0x00FFFF,
    0xFFFFFF, 0x800000, 0x808000, 0x008000, 0x800080, 0x000080, 0x008080, 0x808080, 0xC0C0C0];

  function escHtml(s) {
    return s.replace(/[&<>]/g, function (c) { return c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'; });
  }

  // Resolved character props -> inline CSS ('' for default formatting).
  function styleCss(p) {
    if (!p) return '';
    var css = '';
    if (p.b) css += 'font-weight:bold;';
    if (p.i) css += 'font-style:italic;';
    var deco = '';
    if (p.u) deco += 'underline ';
    if (p.strike) deco += 'line-through ';
    if (deco) css += 'text-decoration:' + deco.trim() + ';';
    if (p.hps) css += 'font-size:' + (p.hps / 2) + 'pt;';
    if (p.font) css += "font-family:'" + p.font.replace(/['\\<>]/g, '') + "';";
    var col = colorOf(p);
    if (col) css += 'color:' + col + ';';
    return css;
  }
  function colorOf(p) {
    if (p.cv != null && (p.cv & 0xFFFFFF) !== 0) {
      return 'rgb(' + (p.cv & 0xFF) + ',' + ((p.cv >> 8) & 0xFF) + ',' + ((p.cv >> 16) & 0xFF) + ')';
    }
    if (p.ico != null && p.ico > 1 && p.ico < ICO_PALETTE.length) return ICO_PALETTE[p.ico];
    return '';
  }

  // Like extractRange, but wraps styled text runs in <span> (HTML-escaped).
  // Structural marks stay bare — \t between table cells, \n at row/paragraph
  // breaks — so the caller can build <table>/<p>. resolve(fc) returns the
  // resolved character props for the char at WordDocument offset fc.
  function extractRangeStyled(wd, pieces, lo, hi, isDeleted, resolve) {
    if (hi <= lo) return '';
    var out = [], buf = '', curCss = null, fieldStack = [], cells = 0;
    function flushRun() {
      if (!buf) return;
      out.push(curCss ? '<span style="' + curCss + '">' + escHtml(buf) + '</span>' : escHtml(buf));
      buf = '';
    }
    function flushCells() { if (!cells) return; flushRun(); out.push(cells === 1 ? '\t' : '\n'); cells = 0; }
    function feed(code, fc) {
      if (code === 0x13) { flushCells(); fieldStack.push(true); return; }
      if (code === 0x14) { flushCells(); if (fieldStack.length) fieldStack[fieldStack.length - 1] = false; return; }
      if (code === 0x15) { flushCells(); if (fieldStack.length) fieldStack.pop(); return; }
      for (var i = 0; i < fieldStack.length; i++) if (fieldStack[i]) return;
      if (code === 0x07) { cells++; return; }
      flushCells();
      var ch = mapChar(code);
      if (ch === '') return;
      if (ch === '\n' || ch === '\t') { flushRun(); out.push(ch); return; }
      var css = resolve ? styleCss(resolve(fc)) : '';
      if (css !== curCss) { flushRun(); curCss = css; }
      buf += ch;
    }
    for (var i = 0; i < pieces.length; i++) {
      var pc = pieces[i];
      var a = lo > pc.cpStart ? lo : pc.cpStart;
      var b = hi < pc.cpEnd ? hi : pc.cpEnd;
      if (a >= b) continue;
      var start = a - pc.cpStart, count = b - a;
      if (pc.compressed) {
        var base = pc.offset + start;
        for (var k = 0; k < count; k++) {
          var fc = base + k;
          if (isDeleted && isDeleted(fc)) continue;
          var byte = wd[fc]; if (byte === undefined) break;
          feed(byte < 0x80 ? byte : (byte <= 0x9F ? FC_COMPRESSED_MAP[byte - 0x80] : byte), fc);
        }
      } else {
        var u = pc.offset + start * 2;
        for (var m = 0; m < count; m++) {
          var fc2 = u + m * 2;
          if (isDeleted && isDeleted(fc2)) continue;
          var loB = wd[fc2]; if (loB === undefined) break;
          feed(loB | ((wd[fc2 + 1] || 0) << 8), fc2);
        }
      }
    }
    flushCells(); flushRun();
    return out.join('');
  }

  // Structured model for the writer (textToDoc): the same styled runs the HTML
  // output is built from, but preserving the table cell/row marks the text and
  // HTML flatten. Returns an array of paragraphs:
  //   { runs: [{ text, b, i, u, strike, size, font, color }], kind }
  // where kind is 'p' (normal paragraph), 'cell' (table cell, more cells follow
  // in the row) or 'rowEnd' (last cell of a table row). size is in points;
  // color is a COLORREF int (0x00BBGGRR) or null.
  function extractRangeModel(wd, pieces, lo, hi, isDeleted, resolve, images, imgCtr, paraAlign, data, paraList, paraPP, paraTblw, paraTblShd, paraTblMerge, footnoteRefs, endnoteRefs, commentRefs, textboxRefs) {
    var paras = [], runs = [], buf = '', curKey = null, curProps = null, fieldStack = [], cells = 0, lastCellFc = null;
    var instr = '', inInstr = false, curUrl = null; // hyperlink field: instruction text + the URL it yields
    var EMPTY = { b: false, i: false, u: false, strike: false, size: null, font: null, color: null };
    function props(p) {
      var color = null;
      if (p.cv != null && (p.cv & 0xFFFFFF) !== 0) color = p.cv & 0xFFFFFF; // already 0x00BBGGRR
      return { b: !!p.b, i: !!p.i, u: !!p.u, strike: !!p.strike, size: p.hps ? p.hps / 2 : null, font: p.font || null, color: color };
    }
    function key(pp) { return pp.b + '|' + pp.i + '|' + pp.u + '|' + pp.strike + '|' + pp.size + '|' + pp.font + '|' + pp.color; }
    // A HYPERLINK field instruction is `HYPERLINK "addr" [switches]`; the address
    // is the first quoted token (or first bare token for an unquoted URL).
    function parseHyperlink(s) {
      var m = /HYPERLINK\s+(.*)/i.exec(s); if (!m) return null;
      var q = /"([^"]+)"/.exec(m[1]); if (q) return q[1];
      var t = /(\S+)/.exec(m[1]); return t ? t[1] : null;
    }
    function flushRun() { if (buf) { var r = { text: buf, b: curProps.b, i: curProps.i, u: curProps.u, strike: curProps.strike, size: curProps.size, font: curProps.font, color: curProps.color }; if (curUrl) r.url = curUrl; runs.push(r); buf = ''; } }
    function endPara(kind, fc, tblw, tblShd, tblMerge) { flushRun(); var pp = (paraPP && fc != null) ? paraPP(fc) : null; var par = { runs: runs, kind: kind, align: (paraAlign && fc != null) ? paraAlign(fc) : 0, list: (paraList && fc != null) ? paraList(fc) : null }; if (pp) par.pp = pp; if (tblw) par.tblw = tblw; if (tblShd) par.tblShd = tblShd; if (tblMerge) par.tblMerge = tblMerge; paras.push(par); runs = []; curKey = null; curProps = null; }
    // The row's column boundaries (rgdxaCenter) live in the PAPX of the row-terminator
    // mark (the last 0x07 seen), so resolve them there and attach to the rowEnd cell.
    function flushCells() { if (!cells) return; var k = cells === 1 ? 'cell' : 'rowEnd'; var rE = k === 'rowEnd' && lastCellFc != null; var tw = (rE && paraTblw) ? paraTblw(lastCellFc) : null, ts = (rE && paraTblShd) ? paraTblShd(lastCellFc) : null, tm = (rE && paraTblMerge) ? paraTblMerge(lastCellFc) : null; endPara(k, null, tw, ts, tm); cells = 0; }
    function feed(code, fc, cp) {
      // A footnote reference (0x02) whose CP is listed in PlcffndRef: record an
      // anchor run so the writer can re-place it and re-link the footnote text.
      if (code === 0x02 && footnoteRefs && footnoteRefs[cp] != null) { flushRun(); runs.push({ ftnRef: footnoteRefs[cp] }); curKey = null; return; }
      if (code === 0x02 && endnoteRefs && endnoteRefs[cp] != null) { flushRun(); runs.push({ endRef: endnoteRefs[cp] }); curKey = null; return; }
      if (code === 0x05 && commentRefs && commentRefs[cp] != null) { flushRun(); runs.push({ comRef: commentRefs[cp] }); curKey = null; return; }
      if (code === 0x08 && textboxRefs && textboxRefs[cp] != null) { flushRun(); runs.push({ tbxRef: textboxRefs[cp] }); curKey = null; return; }
      // Field marks. For a top-level field we collect the instruction text (so a
      // HYPERLINK URL can be recovered) and tag the result runs with the URL.
      if (code === 0x13) { flushCells(); flushRun(); fieldStack.push(true); if (fieldStack.length === 1) { inInstr = true; instr = ''; curKey = null; } return; }
      if (code === 0x14) { flushCells(); if (fieldStack.length) fieldStack[fieldStack.length - 1] = false; if (fieldStack.length === 1) { inInstr = false; curUrl = parseHyperlink(instr); curKey = null; } return; }
      if (code === 0x15) { flushCells(); flushRun(); if (fieldStack.length) fieldStack.pop(); if (fieldStack.length === 0) { curUrl = null; curKey = null; } return; }
      for (var i = 0; i < fieldStack.length; i++) if (fieldStack[i]) { if (inInstr) { var ic = mapChar(code); if (ic) instr += ic; } return; }
      if (code === 0x07) { cells++; lastCellFc = fc; return; }
      flushCells();
      if (code === 0x0A || code === 0x0B || code === 0x0C || code === 0x0D) { endPara('p', fc); return; }
      // picture placeholder (0x01): emit an image run, paired in order with the
      // carved images, plus the picture's display size from its PICF (dxaGoal/
      // dyaGoal, via sprmCPicLocation) so the writer can re-embed at that size.
      if (code === 0x01 && images && imgCtr && imgCtr.n < images.length) {
        flushRun();
        var img = images[imgCtr.n++], pl = resolve ? resolve(fc).picLoc : null;
        if (data && pl != null && pl + 32 <= data.length) { var pv = new DataView(data.buffer, data.byteOffset, data.byteLength); img.dxa = pv.getInt16(pl + 0x1C, true); img.dya = pv.getInt16(pl + 0x1E, true); }
        runs.push({ image: img });
        return;
      }
      var ch = mapChar(code);
      if (ch === '') return;
      var pp = resolve ? props(resolve(fc)) : EMPTY;
      var kk = key(pp);
      if (kk !== curKey) { flushRun(); curKey = kk; curProps = pp; }
      buf += ch;
    }
    for (var i = 0; i < pieces.length; i++) {
      var pc = pieces[i];
      var a = lo > pc.cpStart ? lo : pc.cpStart, b = hi < pc.cpEnd ? hi : pc.cpEnd;
      if (a >= b) continue;
      var start = a - pc.cpStart, count = b - a;
      if (pc.compressed) {
        var base = pc.offset + start;
        for (var k = 0; k < count; k++) {
          var fc = base + k; if (isDeleted && isDeleted(fc)) continue;
          var byte = wd[fc]; if (byte === undefined) break;
          feed(byte < 0x80 ? byte : (byte <= 0x9F ? FC_COMPRESSED_MAP[byte - 0x80] : byte), fc, a + k);
        }
      } else {
        var u = pc.offset + start * 2;
        for (var m = 0; m < count; m++) {
          var fc2 = u + m * 2; if (isDeleted && isDeleted(fc2)) continue;
          var loB = wd[fc2]; if (loB === undefined) break;
          feed(loB | ((wd[fc2 + 1] || 0) << 8), fc2, a + m);
        }
      }
    }
    flushCells();
    if (buf || runs.length) endPara('p');
    return paras;
  }

  // Field markers ([MS-DOC] "Special Characters"): 0x13 begin, 0x14 separator,
  // 0x15 end. Text in the instruction region (0x13..0x14) is the field code ->
  // dropped; the result region (0x14..0x15) is kept. Nesting is tracked with a
  // stack so a nested field inside an outer instruction is dropped too.
  //
  // Table cell marks (0x07) are buffered as a run: in the binary format every
  // cell ends with a 0x07 and the row terminator adds one more, so a single
  // 0x07 is a column separator (-> tab) while two or more together mark the end
  // of a table row (-> newline). Empty cells can blur this; a fully accurate
  // split would need paragraph properties (sprmPFTtp), but this matches Word for
  // the common case and keeps tables readable instead of one endless line.
  function emit(out, fieldStack, state, code) {
    if (code === 0x13) { flushCells(out, state); fieldStack.push(true); return; }
    if (code === 0x14) { flushCells(out, state); if (fieldStack.length) fieldStack[fieldStack.length - 1] = false; return; }
    if (code === 0x15) { flushCells(out, state); if (fieldStack.length) fieldStack.pop(); return; }
    for (var i = 0; i < fieldStack.length; i++) {
      if (fieldStack[i]) return; // inside some field's instruction
    }
    if (code === 0x07) { state.cells++; return; } // buffer the cell-mark run
    flushCells(out, state);
    var ch = mapChar(code);
    if (ch !== '') out.push(ch);
  }

  // Flush a pending run of cell marks: one -> column tab, two or more -> newline.
  function flushCells(out, state) {
    if (state.cells === 1) out.push('\t');
    else if (state.cells >= 2) out.push('\n');
    state.cells = 0;
  }

  // Map a decoded character to output text. Structural control marks become
  // whitespace; special placeholders are dropped.
  function mapChar(code) {
    switch (code) {
      // 0x07 (cell mark) is handled as a run in emit()/flushCells(), not here.
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
