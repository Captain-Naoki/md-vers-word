// Banc de test : Markdown → .docx sans navigateur (DOM simulé par linkedom)
import { parseHTML } from 'linkedom';
import fs from 'node:fs';
const { document, window } = parseHTML('<!doctype html><html><body><article id="p"></article></body></html>');
globalThis.document = document; globalThis.window = window;
const { markdownToHtml, liftDisplayMath } = await import('../src/markdown.js');
const { buildDocx } = await import('../src/docx.js');
const { buildClipboardHtml } = await import('../src/clipboard.js');
const md = fs.readFileSync(process.argv[2], 'utf8');
const root = document.getElementById('p');
root.innerHTML = markdownToHtml(md);
liftDisplayMath(root);
const blob = await buildDocx(root, { name: 'Times New Roman', size: 24 });
fs.writeFileSync(process.argv[3], Buffer.from(await blob.arrayBuffer()));
fs.writeFileSync(process.argv[3].replace(/\.docx$/, '.clip.html'), buildClipboardHtml(root, { font: { css: 'Times New Roman', size: 24 } }));
console.log('ok', root.querySelectorAll('span.math').length, 'equations');
