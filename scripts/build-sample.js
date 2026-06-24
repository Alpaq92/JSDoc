/*
 * build-sample.js — regenerate samples/feature-showcase.doc.
 *
 * A clean-room showcase of the fidelity JSDoc round-trips, written with the bundled
 * writer (src/textToDoc.js) and read back by docToText. It exercises, in one document:
 *   - document properties (title / author / subject)
 *   - character extras: superscript / subscript / highlight (+ bold / italic / underline / colour)
 *   - tab stops with a dot leader and right-aligned numbers (a table of contents)
 *   - paragraph shading + a box border (a callout)
 *   - a table: a column-spanning merged header, shaded status cells, an empty cell, a part-width merge
 *   - bookmarks (named ranges over real text)
 *
 *   node scripts/build-sample.js
 */
var fs = require('fs');
var path = require('path');
var textToDoc = require('../src/textToDoc.js');
var docToText = require('../src/docToText.js');

// COLORREF is 0x00BBGGRR (red is the low byte).
var GREY = 0xC0C0C0, GREEN = 0x00FF00, RED = 0x0000FF, YELLOW = 0x00FFFF, BLUE = 0xFF0000, RULE = 0x808080, CALLOUT = 0xF0F0F0;
var FULL = [0, 3000, 6000, 9000];   // three equal columns
function cell(t, shd, vm) { var c = { runs: t ? [{ text: t }] : [], kind: 'cell' }; if (shd != null) c.shd = shd; if (vm) c.vmerge = vm; return c; }
function rowEnd(t, shd, vm) { var c = { runs: t ? [{ text: t }] : [], kind: 'rowEnd', tblw: FULL }; if (shd != null) c.shd = shd; if (vm) c.vmerge = vm; return c; }
function toc(label, page) { return { runs: [{ text: label + '\t' + page }], kind: 'p', pp: { tabs: [{ pos: 9000, align: 'right', leader: 'dot' }] } }; }
function box(c) { var s = { color: c, width: 1, type: 1 }; return { top: s, left: s, bottom: s, right: s }; }

function model(bookmarks) {
  return {
    props: { title: 'JSDoc feature showcase', author: 'Alpaq92', subject: 'Everything JSDoc round-trips' },
    bookmarks: bookmarks,
    body: [].concat(
      [{ runs: [{ text: 'JSDoc feature showcase', b: true, size: 15 }], kind: 'p' }],
      [{ runs: [{ text: 'One .doc, written by the bundled writer and read back, exercising the formatting JSDoc round-trips.' }], kind: 'p' }],
      // Character extras: superscript / subscript / highlight, with bold / italic / underline / colour.
      [{ runs: [
        { text: 'Inline: E = mc' }, { text: '2', va: 'super' }, { text: ', the formula H' }, { text: '2', va: 'sub' },
        { text: 'O, a ' }, { text: 'highlighted', highlight: YELLOW }, { text: ' term and a ' }, { text: 'flagged', highlight: GREEN },
        { text: ' note — plus ' }, { text: 'bold', b: true }, { text: ', ' }, { text: 'italic', i: true },
        { text: ', ' }, { text: 'underlined', u: true }, { text: ', and ' }, { text: 'blue', color: BLUE }, { text: ' text.' }
      ], kind: 'p' }],
      // More character styles: underline kinds, small caps, all caps, and hidden text.
      [{ runs: [
        { text: 'Also: ' }, { text: 'double', u: true, uStyle: 'double' }, { text: ', ' }, { text: 'dotted', u: true, uStyle: 'dotted' },
        { text: ', and ' }, { text: 'wavy', u: true, uStyle: 'wavy' }, { text: ' underlines; ' }, { text: 'small caps', smallCaps: true },
        { text: '; ' }, { text: 'all caps', caps: true }, { text: '; and ' }, { text: 'hidden', hidden: true }, { text: ' text (shown dimmed).' }
      ], kind: 'p' }],
      // Tab stops: a small table of contents — a dot leader to a right-aligned page number.
      [{ runs: [{ text: 'Contents', b: true }], kind: 'p' }],
      [toc('Introduction', '1'), toc('Methods', '4'), toc('Results', '9')],
      // Paragraph shading + a box border (a callout).
      [{ runs: [{ text: 'Note: this paragraph has a grey background and a box border — sprmPShd plus the four sprmPBrc sides.' }], kind: 'p', pp: { shd: CALLOUT, borders: box(RULE) } }],
      // A status table: merged header, shaded State cells, an empty cell, a part-width
      // merge, and the Owner column vertically merged across the first two rows.
      [{ runs: [{ text: 'Project status — Q2', b: true }], kind: 'rowEnd', shd: GREY, tblw: [0, 9000] }],
      [cell('Phase'), cell('State'), rowEnd('Owner')],
      [cell('Design'), cell('Done', GREEN), rowEnd('Alpaq92', null, 'restart')],   // Owner spans...
      [cell('Build'), cell('Blocked', RED), rowEnd('', null, 'cont')],             // ...these two rows
      [cell('Ship'), cell('Pending', YELLOW), rowEnd('')],                         // empty Owner cell
      [{ runs: [{ text: 'Overall', b: true }], kind: 'cell' }, { runs: [{ text: 'On track', b: true }], kind: 'rowEnd', tblw: [0, 6000, 9000] }],
      [{ runs: [{ text: 'The words “highlighted” and “Methods” above are bookmarks. Built by scripts/build-sample.js.' }], kind: 'p' }]
    )
  };
}

// Two passes: build once to locate where target phrases land in CP space, then rebuild
// with bookmarks spanning those ranges (the text is identical, so the CPs line up).
var body = docToText(textToDoc(model(null)));
function range(name, phrase) { var i = body.indexOf(phrase); return i < 0 ? null : { name: name, start: i, end: i + phrase.length }; }
var bookmarks = [range('highlightedTerm', 'highlighted'), range('methodsEntry', 'Methods')].filter(Boolean);

var out = path.join(__dirname, '..', 'samples', 'feature-showcase.doc');
fs.writeFileSync(out, Buffer.from(textToDoc(model(bookmarks))));
console.log('wrote ' + out + ' (' + fs.statSync(out).size + ' bytes); bookmarks: ' + bookmarks.map(function (b) { return b.name + ' [' + b.start + '–' + b.end + ']'; }).join(', '));
