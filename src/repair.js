// Réparation des textes copiés « à la souris » depuis NotebookLM : retours à la ligne perdus
// (« généralePour », « ?La MSE ») et références collées aux mots (« observées78. », « $i$8. »).
// Seules les lignes manifestement aplaties sont modifiées : un Markdown normal reste intact.

const L = 'a-zà-öø-ÿœ';
const U = 'A-ZÀ-ÖØ-ÞŒ';
const MATH = /(\$\$[\s\S]+?\$\$|\\\[[\s\S]+?\\\]|\\\([\s\S]+?\\\)|\$(?![\s$])(?:\\.|[^\\$\n])+?(?<![\s\\])\$)/g;
const PH = '\\d+[DI]'; // équation protégée (D = centrée, I = en ligne)

// Mots où une majuscule au milieu est normale (pas une coupure de ligne perdue)
const CAMEL = new Set(['pytorch', 'tensorflow', 'github', 'gitlab', 'youtube', 'latex', 'deepmind', 'openai', 'notebooklm',
  'javascript', 'typescript', 'powerpoint', 'linkedin', 'whatsapp', 'iphone', 'ipad', 'mcdonald', 'macbook', 'deepseek',
  'postgresql', 'mongodb', 'numpy', 'scipy', 'matplotlib', 'pypi', 'deepl', 'onedrive', 'sharepoint', 'researchgate',
  'scikitlearn', 'keras', 'xgboost', 'lightgbm', 'catboost', 'googlenet', 'resnet', 'unet', 'imagenet', 'leakyrelu']);
// Unités suivies d'un exposant écrit sans ^ (km2, m3…) : ne pas les prendre pour des références
const UNITS = new Set(['km', 'm', 'cm', 'mm', 'dm', 's', 'h', 'kg', 'g', 'ha', 'hpa', 'pa', 'kpa', 'w', 'kw', 'j', 'l', 'ml', 'k', 'mol']);

function protect(text) {
  const maths = [];
  const t = text.replace(MATH, (m) => {
    maths.push(m);
    return `${maths.length - 1}${/^(\$\$|\\\[)/.test(m) ? 'D' : 'I'}`;
  });
  return { t, restore: (s) => s.replace(/(\d+)[DI]/g, (_, i) => maths[+i]) };
}

// « observées78. » → « observées. », « pixel $i$8. » → « pixel $i$. », « CCAM)18. » → « CCAM). »
function stripGluedProtected(t) {
  const re = new RegExp(`([${L}${U}]+|[)\\]»"’]|\\uE001)(\\d{1,3})(?=[.,;:!?)\\s\\uE000${U}]|$)`, 'g');
  return t.replace(re, (m, before) => {
    if (/^\p{L}+$/u.test(before)) {
      if (before.length === 1) return m; // x2, r2
      if (UNITS.has(before.toLowerCase())) return m; // km2, m3
      if (/\p{Lu}$/u.test(before)) return m; // CO2, SSP3, ERA5, CMIP6
    }
    return before;
  });
}

export function stripGluedCitations(md) {
  const { t, restore } = protect(md);
  return restore(stripGluedProtected(t));
}

const GLUED = new RegExp(`[.?!:)][${U}]|[${L}][${U}][${L}]|[^\\s]${PH}`, 'g');
function looksFlattened(t) {
  if (t.length < 200) return false;
  return (t.match(GLUED) || []).length >= 3;
}

function wordAt(s, i) {
  let a = i, b = i + 1;
  while (a > 0 && /\p{L}/u.test(s[a - 1])) a--;
  while (b < s.length && /\p{L}/u.test(s[b])) b++;
  return s.slice(a, b).toLowerCase();
}

const M_COLON = '\u0001', M_SENT = '\u0002', M_ITEM = '\u0003', M_DISP = '\u0004';

function repairLine(t) {
  let s = stripGluedProtected(t);
  // équations centrées sur leur propre ligne
  s = s.replace(new RegExp(`\\s*(\\uE000\\d+D\\uE001)\\s*`, 'g'), `${M_DISP}$1${M_DISP}`);
  // « composante$y_i$ : La valeur… » → élément de liste
  s = s.replace(new RegExp(`(?<=[^\\s${M_DISP}])(\\uE000\\d+I\\uE001)(?=\\s?:)`, 'g'), `${M_ITEM}$1`);
  // « fondamentales :Elle rend… »
  s = s.replace(new RegExp(`:(?=[${U}])`, 'g'), `:${M_COLON}`);
  // « explication) ?La MSE », « façon).Elle », « 2024)Dans » (mais pas les initiales « A.B. »)
  s = s.replace(new RegExp(`(?<!(?:^|[^\\p{L}])\\p{Lu})([.?!)])(?=[${U}])`, 'gu'), `$1${M_SENT}`);
  // « généralePour » (mais pas « PyTorch », « GitHub »…)
  s = s.replace(new RegExp(`([${L}])(?=[${U}][${L}])`, 'g'), (m, c, off, str) => (CAMEL.has(wordAt(str, off)) ? m : c + M_SENT));

  // Assemblage en Markdown
  const parts = s.split(/([\u0001-\u0004])/);
  const blocks = []; // { kind: 'p' | 'item' | 'sub' | 'disp', text }
  let kind = 'p', buf = '';
  const flush = () => {
    const text = buf.trim();
    if (text) blocks.push({ kind: new RegExp(`^\\uE000\\d+D\\uE001$`).test(text) ? 'disp' : kind, text });
    buf = '';
  };
  for (const part of parts) {
    if (part === M_DISP) { flush(); kind = 'p'; }
    else if (part === M_ITEM) { flush(); kind = 'item'; }
    else if (part === M_COLON) { flush(); kind = ['item', 'sub'].includes(blocks[blocks.length - 1]?.kind) ? 'sub' : 'item'; }
    else if (part === M_SENT) { flush(); kind = blocks.length && blocks[blocks.length - 1].kind === 'sub' ? 'sub' : 'p'; }
    else buf += part;
  }
  flush();

  let out = '', prev = null;
  blocks.forEach((b, i) => {
    let text = b.text;
    if (b.kind === 'p') {
      if (i === 0 && /\?$/.test(text) && text.length <= 200) text = `## ${text}`;
      else if (/^[A-H]\.\s.{3,140}$/.test(text) && !/[.:;]$/.test(text)) text = `### ${text}`;
      out += `\n\n${text}`;
    } else if (b.kind === 'item') out += `${prev === 'item' || prev === 'sub' ? '\n' : '\n\n'}- ${text}`;
    else if (b.kind === 'sub') out += prev === 'item' || prev === 'sub' ? `\n  - ${text}` : `\n\n- ${text}`;
    else out += `\n\n${text}\n`;
    prev = b.kind;
  });
  return out.trim();
}

/**
 * Tableaux copiés depuis une page web : le navigateur les donne en texte avec des tabulations
 * entre les cellules. On les convertit en tableau Markdown (au moins 2 lignes consécutives).
 */
export function tabsToTables(md) {
  const lines = md.split('\n');
  const out = [];
  let inCode = false;
  for (let i = 0; i < lines.length; i++) {
    if (/^\s*(```|~~~)/.test(lines[i])) inCode = !inCode;
    if (inCode || !lines[i].includes('\t')) { out.push(lines[i]); continue; }
    let j = i;
    while (j < lines.length && lines[j].includes('\t')) j++;
    if (j - i < 2) { out.push(lines[i]); continue; }
    const rows = lines.slice(i, j).map((l) => l.split('\t').map((c) => c.trim().replace(/(?<!\\)\|/g, '\\|')));
    const n = Math.max(...rows.map((r) => r.length));
    const fmt = (r) => `| ${Array.from({ length: n }, (_, k) => r[k] || ' ').join(' | ')} |`;
    if (out.length && out[out.length - 1].trim()) out.push('');
    out.push(fmt(rows[0]), `|${' --- |'.repeat(n)}`, ...rows.slice(1).map(fmt), '');
    i = j - 1;
  }
  return out.join('\n');
}

/** Répare uniquement les lignes « aplaties » ; le reste du texte est rendu tel quel. */
export function repairFlattened(md) {
  return tabsToTables(md).split('\n').map((line) => {
    if (line.trimStart().startsWith('|')) return line; // tableau
    const { t, restore } = protect(line);
    return looksFlattened(t) ? restore(repairLine(t)) : line;
  }).join('\n');
}
