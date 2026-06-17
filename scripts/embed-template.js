'use strict';
/*
 * embed-template.js — (re)generate the blank-document skeleton embedded in
 * src/textToDoc.js as TEMPLATE_B64.
 *
 * The skeleton is a real, app-saved empty .doc stripped to just the two streams
 * the writer reuses (WordDocument + 1Table — FIB, stylesheet, section table,
 * fonts, list definitions). It has no authored content; textToDoc() injects the
 * body into it. The source isn't bundled in the repo (only the embedded base64
 * is shipped) — pass a blank .doc to regenerate:
 *   node scripts/embed-template.js path/to/blank.doc
 * (textToDoc.buildCfb can strip a full blank .doc to the two streams first.)
 */
var fs = require('fs');
var path = require('path');

var root = path.join(__dirname, '..');
var srcDoc = process.argv[2];
if (!srcDoc) { console.error('usage: node scripts/embed-template.js path/to/blank.doc'); process.exit(1); }
var skeleton = fs.readFileSync(srcDoc);
var b64 = Buffer.from(skeleton).toString('base64');

var file = path.join(root, 'src', 'textToDoc.js');
var src = fs.readFileSync(file, 'utf8');
var re = /var TEMPLATE_B64 = (?:"[^"]*"|'[^']*');/;
if (!re.test(src)) { console.error('TEMPLATE_B64 assignment not found in textToDoc.js'); process.exit(1); }
src = src.replace(re, 'var TEMPLATE_B64 = ' + JSON.stringify(b64) + ';');
fs.writeFileSync(file, src);
console.log('embedded', skeleton.length, 'bytes ->', b64.length, 'base64 chars into src/textToDoc.js');
