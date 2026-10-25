// Prépare le HTML que Word comprend lorsqu'on colle (Ctrl+V)
import { esc, mathmlToOmml, ommlToWordHtml } from './omml.js';

const HEAD = '<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
  + 'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" '
  + 'xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><meta name="ProgId" content="Word.Document">'
  + '<meta name="Generator" content="Microsoft Word 15"></head>';

const cleanMathml = (math) => math.outerHTML
  .replace(/ (class|style)="[^"]*"/g, '')
  .replace(/<math(?![^>]*xmlns=)/, '<math xmlns="http://www.w3.org/1998/Math/MathML"');

/**
 * mode = 'omml'   : équations Word natives (format HTML de Word) + MathML en secours
 * mode = 'mathml' : MathML seul
 */
export function buildClipboardHtml(root, { font, mode = 'omml' }) {
  const clone = root.cloneNode(true);
  const snippets = [];
  for (const span of Array.from(clone.querySelectorAll('span.math'))) {
    const display = span.classList.contains('math-display');
    const math = span.querySelector('math');
    let snip;
    if (!math) snip = `<span style="color:#C00000">${esc(span.dataset.tex || span.textContent)}</span>`;
    else if (mode === 'mathml') snip = cleanMathml(math);
    else snip = `<!--[if gte msEquation 12]>${ommlToWordHtml(mathmlToOmml(math, display))}<![endif]--><![if !msEquation]>${cleanMathml(math)}<![endif]>`;
    snippets.push(snip);
    span.replaceWith(document.createTextNode(`${snippets.length - 1}`));
  }
  for (const d of Array.from(clone.querySelectorAll('div.math-block'))) {
    const p = document.createElement('p');
    p.className = 'MsoNormal';
    p.setAttribute('align', 'center');
    p.setAttribute('style', 'text-align:center');
    while (d.firstChild) p.appendChild(d.firstChild);
    d.replaceWith(p);
  }
  for (const p of Array.from(clone.querySelectorAll('p'))) if (!p.className) p.className = 'MsoNormal';
  const style = font && font.css ? ` style="font-family:${font.css};font-size:${font.size / 2}pt"` : '';
  const body = clone.innerHTML.replace(/(\d+)/g, (_, i) => snippets[+i]);
  return `${HEAD}<body lang="FR"><div${style}>${body}</div></body></html>`;
}

/** Texte brut de secours (équations en LaTeX) */
export function buildPlainText(root) {
  const clone = root.cloneNode(true);
  for (const span of Array.from(clone.querySelectorAll('span.math'))) {
    const tex = span.dataset.tex || '';
    span.replaceWith(document.createTextNode(span.classList.contains('math-display') ? `\n${tex}\n` : tex));
  }
  return clone.innerText || clone.textContent;
}

/** Copie riche via l'événement « copy » (aucune réécriture du HTML par le navigateur) */
export function copyRich(html, text) {
  let done = false;
  const handler = (e) => {
    if (html) e.clipboardData.setData('text/html', html);
    e.clipboardData.setData('text/plain', text);
    e.preventDefault();
    done = true;
  };
  document.addEventListener('copy', handler, true);
  try { document.execCommand('copy'); } finally { document.removeEventListener('copy', handler, true); }
  if (!done && navigator.clipboard && window.ClipboardItem) {
    // Secours : API moderne (le navigateur peut alors simplifier le HTML)
    const items = { 'text/plain': new Blob([text], { type: 'text/plain' }) };
    if (html) items['text/html'] = new Blob([html], { type: 'text/html' });
    navigator.clipboard.write([new ClipboardItem(items)]).catch(() => {});
    return true;
  }
  return done;
}

/** Une seule équation, en MathML texte : Word la transforme en équation au collage */
export function copySingleEquation(span) {
  const math = span.querySelector('math');
  if (!math) return false;
  const mml = cleanMathml(math);
  return copyRich(null, mml);
}
