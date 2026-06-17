'use strict';
/*
 * embed-template.js — (re)generate the blank-document skeleton embedded in
 * src/textToDoc.js as TEMPLATE_B64.
 *
 * The skeleton is samples/blank-template.doc: a real, app-saved empty .doc
 * stripped to just the two streams the writer reuses (WordDocument + 1Table —
 * which carry the FIB, stylesheet, section table and fonts). It contains no
 * document text; textToDoc() injects the body into it. Run after replacing the
 * skeleton:  node scripts/embed-template.js
 */
var fs = require('fs');
var path = require('path');

var root = path.join(__dirname, '..');
var skeleton = fs.readFileSync(path.join(root, 'samples', 'blank-template.doc'));
var b64 = Buffer.from(skeleton).toString('base64');

var file = path.join(root, 'src', 'textToDoc.js');
var src = fs.readFileSync(file, 'utf8');
var re = /var TEMPLATE_B64 = (?:"[^"]*"|'[^']*');/;
if (!re.test(src)) { console.error('TEMPLATE_B64 assignment not found in textToDoc.js'); process.exit(1); }
src = src.replace(re, 'var TEMPLATE_B64 = ' + JSON.stringify(b64) + ';');
fs.writeFileSync(file, src);
console.log('embedded', skeleton.length, 'bytes ->', b64.length, 'base64 chars into src/textToDoc.js');
