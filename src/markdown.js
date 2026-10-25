// Markdown (format NotebookLM / ChatGPT) + LaTeX → HTML avec équations MathML (Temml)
import { Marked } from 'marked';
import temml from 'temml';

const escAttr = (s) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

export function renderMath(tex, display) {
  let mml;
  try {
    mml = temml.renderToString(tex, { displayMode: display, throwOnError: true, strict: false, trust: false });
  } catch (err) {
    mml = `<span class="math-error" title="${escAttr(String(err.message || err))}">${escAttr(tex)}</span>`;
  }
  const cls = display ? 'math math-display' : 'math math-inline';
  return `<span class="${cls}" data-tex="${escAttr(tex)}">${mml}</span>`;
}

const blockMath = {
  name: 'blockMath',
  level: 'block',
  start(src) { const m = src.match(/^ {0,3}(\$\$|\\\[)/m); return m ? m.index : undefined; },
  tokenizer(src) {
    const m = /^ {0,3}\$\$([\s\S]+?)\$\$[ \t]*(?:\n+|$)/.exec(src)
      || /^ {0,3}\\\[([\s\S]+?)\\\][ \t]*(?:\n+|$)/.exec(src)
      || /^ {0,3}\\begin\{(equation\*?|align\*?|gather\*?|multline\*?)\}([\s\S]+?)\\end\{\1\}[ \t]*(?:\n+|$)/.exec(src);
    if (!m) return undefined;
    let tex = m[2] !== undefined ? m[2] : m[1];
    if (m[2] !== undefined && /^(align|gather|multline)/.test(m[1])) tex = `\\begin{aligned}${tex}\\end{aligned}`;
    return { type: 'blockMath', raw: m[0], text: tex.trim() };
  },
  renderer(t) { return `<div class="math-block">${renderMath(t.text, true)}</div>\n`; },
};

const inlineMath = {
  name: 'inlineMath',
  level: 'inline',
  start(src) { const i = src.search(/(?<!\\)\$|\\\(|\\\[/); return i < 0 ? undefined : i; },
  tokenizer(src) {
    let m;
    if ((m = /^\$\$([\s\S]+?)\$\$/.exec(src))) return { type: 'inlineMath', raw: m[0], text: m[1].trim(), display: true };
    if ((m = /^\\\[([\s\S]+?)\\\]/.exec(src))) return { type: 'inlineMath', raw: m[0], text: m[1].trim(), display: true };
    if ((m = /^\\\(([\s\S]+?)\\\)/.exec(src))) return { type: 'inlineMath', raw: m[0], text: m[1].trim(), display: false };
    // Pas de règle « pas de chiffre après le $ fermant » : NotebookLM colle ses numéros de citation ($i$8.)
    if ((m = /^\$(?![\s$])((?:\\.|[^\\$\n]|\n(?!\n))+?)(?<![\s\\])\$/.exec(src)))
      return { type: 'inlineMath', raw: m[0], text: m[1].trim(), display: false };
    return undefined;
  },
  renderer(t) { return renderMath(t.text, t.display); },
};

const marked = new Marked({ gfm: true, breaks: false });
marked.use({ extensions: [blockMath, inlineMath] });

/** Supprime les marqueurs de citation de NotebookLM : [1], [2, 3], [4-6], ¹ … (hors équations) */
export function stripCitations(md) {
  const parts = md.split(/(\$\$[\s\S]+?\$\$|\$[^$\n]+?\$|\\\[[\s\S]+?\\\]|\\\([\s\S]+?\\\))/);
  return parts.map((p, i) => (i % 2 ? p : p
    .replace(/[  ]?\[(\d+(?:\s*[-–,]\s*\d+)*)\](?!\()/g, '')
  )).join('');
}

export function markdownToHtml(md) {
  // Normalise les retours à la ligne Windows et les espaces insécables collés depuis le web
  return marked.parse(md.replace(/\r\n?/g, '\n'));
}

/**
 * Sort les équations « display » des paragraphes / éléments de liste pour qu'elles
 * forment leur propre bloc (nécessaire pour Word, qui ne met pas d'équation centrée au milieu d'un paragraphe).
 */
export function liftDisplayMath(root) {
  for (const span of Array.from(root.querySelectorAll('p > span.math-display'))) {
    const p = span.parentNode;
    const before = p.cloneNode(false), after = p.cloneNode(false);
    let node = p.firstChild, seen = false;
    while (node) {
      const next = node.nextSibling;
      if (node === span) seen = true;
      else (seen ? after : before).appendChild(node);
      node = next;
    }
    const block = document.createElement('div');
    block.className = 'math-block';
    block.appendChild(span);
    const frag = document.createDocumentFragment();
    if (before.textContent.trim() || before.querySelector('.math')) frag.appendChild(before);
    frag.appendChild(block);
    if (after.textContent.trim() || after.querySelector('.math')) frag.appendChild(after);
    p.replaceWith(frag);
  }
}
