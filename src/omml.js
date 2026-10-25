// Convertit un arbre MathML (DOM) en OMML, le format natif des équations Word.
// Fonctionne avec n'importe quelle implémentation DOM (navigateur ou xmldom).

const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '∯', '∰', '⋃', '⋂', '⋁', '⋀', '⨁', '⨂', '⨀', '⨄', '⨆']);
const INTEGRALS = new Set(['∫', '∬', '∭', '∮', '∯', '∰']);
const STOP_OPS = new Set(['=', '<', '>', '≤', '≥', '≈', '≠', '≡', '∼', '≃', '≅', '∝', '⇒', '⇔', '⟹', '⟺', '+', '−', '-', '±', '∓', ',', ';']);
const OPEN = new Set(['(', '[', '{', '⟨', '⌊', '⌈', '|', '‖']);
const CLOSE = new Set([')', ']', '}', '⟩', '⌋', '⌉']);
const INVISIBLE = /[⁡-⁤]/g;

// Accents : caractère MathML → caractère combinant attendu par Word
const ACCENTS = {
  'ˆ': '̂', '^': '̂', '̂': '̂',
  'ˇ': '̌', '̌': '̌',
  '˜': '̃', '~': '̃', '̃': '̃',
  '‾': '̅', '¯': '̅', '̅': '̅', '̄': '̄',
  '→': '⃗', '⃗': '⃗', '←': '⃖', '↔': '⃡',
  '˙': '̇', '̇': '̇', '¨': '̈', '̈': '̈',
  '´': '́', '`': '̀', '˘': '̆', '˚': '̊', '∘': '̊',
};

const VARIANTS = {
  normal: { sty: 'p' }, bold: { sty: 'b' }, italic: { sty: 'i' }, 'bold-italic': { sty: 'bi' },
  'double-struck': { scr: 'double-struck', sty: 'p' }, script: { scr: 'script' },
  'bold-script': { scr: 'script', sty: 'b' }, fraktur: { scr: 'fraktur', sty: 'p' },
  'bold-fraktur': { scr: 'fraktur', sty: 'b' }, 'sans-serif': { scr: 'sans-serif', sty: 'p' },
  'bold-sans-serif': { scr: 'sans-serif', sty: 'b' }, monospace: { scr: 'monospace', sty: 'p' },
};

export function esc(s) {
  return String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const name = (n) => (n.localName || n.nodeName || '').toLowerCase().replace(/^.*:/, '');
const kids = (n) => Array.from(n.childNodes || []).filter((c) => c.nodeType === 1);
const attr = (n, a) => (n.getAttribute ? n.getAttribute(a) : null);
const text = (n) => (n.textContent || '');
const hasClass = (n, c) => (' ' + (attr(n, 'class') || '') + ' ').includes(' ' + c + ' ');
const isBold = (n) => /font-weight:\s*(bold|[6-9]00)/.test(attr(n, 'style') || '') || /bold/.test(attr(n, 'mathvariant') || '');

// Lettres « mathématiques » Unicode (𝐀, 𝐱, 𝟐…) → lettres normales + style, comme Word les attend
const ALNUM = [
  [0x1D400, 0x1D419, 65, 'b'], [0x1D41A, 0x1D433, 97, 'b'],
  [0x1D434, 0x1D44D, 65, 'i'], [0x1D44E, 0x1D467, 97, 'i'],
  [0x1D468, 0x1D481, 65, 'bi'], [0x1D482, 0x1D49B, 97, 'bi'],
  [0x1D7CE, 0x1D7D7, 48, 'b'],
];
function demath(t) {
  let style = null, out = '', changed = false;
  for (const ch of t) {
    const cp = ch.codePointAt(0);
    const r = ALNUM.find(([a, z]) => cp >= a && cp <= z);
    if (r) {
      if (style && style !== r[3]) return null;
      style = r[3]; out += String.fromCharCode(r[2] + cp - r[0]); changed = true;
    } else if (cp === 0x210E) { if (style && style !== 'i') return null; style = 'i'; out += 'h'; changed = true; }
    else out += ch;
  }
  return changed ? { text: out, style } : null;
}

function run(t, o = {}) {
  if (t === '') return '';
  let mrPr = '';
  if (o.nor) mrPr = '<m:nor/>';
  else {
    if (o.scr) mrPr += `<m:scr m:val="${o.scr}"/>`;
    if (o.sty) mrPr += `<m:sty m:val="${o.sty}"/>`;
  }
  const wrPr = o.nor
    ? (o.bold || o.italic ? `<w:rPr>${o.bold ? '<w:b/>' : ''}${o.italic ? '<w:i/>' : ''}</w:rPr>` : '')
    : '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
  return `<m:r>${mrPr ? `<m:rPr>${mrPr}</m:rPr>` : ''}${wrPr}<m:t xml:space="preserve">${esc(t)}</m:t></m:r>`;
}

const wrap = (tag, inner) => `<m:${tag}>${inner}</m:${tag}>`;
const e = (inner) => wrap('e', inner);

// Déballe les <mrow> à un seul enfant
function unwrap(n) {
  while (n && ['mrow', 'mstyle', 'mpadded'].includes(name(n)) && kids(n).length === 1) n = kids(n)[0];
  return n;
}

function naryInfo(n) {
  n = unwrap(n);
  if (!n) return null;
  const nm = name(n);
  const k = kids(n);
  const isOp = (b) => { b = unwrap(b); return b && name(b) === 'mo' && NARY.has(text(b).trim()); };
  if (nm === 'mo' && NARY.has(text(n).trim())) return { chr: text(n).trim(), sub: null, sup: null, under: false };
  if (['msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover'].includes(nm) && k.length >= 2 && isOp(k[0])) {
    const chr = text(unwrap(k[0])).trim();
    const under = nm.startsWith('mu') || nm === 'mover';
    let sub = null, sup = null;
    if (nm === 'msub' || nm === 'munder') sub = k[1];
    else if (nm === 'msup' || nm === 'mover') sup = k[1];
    else { sub = k[1]; sup = k[2]; }
    return { chr, sub, sup, under };
  }
  return null;
}

function nary(info, bodyXml) {
  const loc = info.under && !INTEGRALS.has(info.chr) ? 'undOvr' : 'subSup';
  let pr = `<m:chr m:val="${esc(info.chr)}"/><m:limLoc m:val="${loc}"/>`;
  if (!info.sub) pr += '<m:subHide m:val="1"/>';
  if (!info.sup) pr += '<m:supHide m:val="1"/>';
  return `<m:nary><m:naryPr>${pr}</m:naryPr>${wrap('sub', info.sub ? conv(info.sub) : '')}${wrap('sup', info.sup ? conv(info.sup) : '')}${e(bodyXml)}</m:nary>`;
}

// Convertit une liste de frères en gérant les opérateurs n-aires (∑, ∫…) qui absorbent la suite
function convList(list) {
  let out = '';
  for (let i = 0; i < list.length; i++) {
    const info = naryInfo(list[i]);
    if (!info) { out += conv(list[i]); continue; }
    const body = [];
    let depth = 0, j = i + 1;
    for (; j < list.length; j++) {
      const n = unwrap(list[j]);
      const nm = name(n);
      if (body.length === 0 && nm === 'mspace') continue;
      if (nm === 'mspace' && depth === 0 && parseFloat(attr(n, 'width') || '0') >= 0.9) break; // \quad

      if (nm === 'mo') {
        const t = text(n).trim();
        if (OPEN.has(t) && !(t === '|' && depth > 0)) depth++;
        else if (CLOSE.has(t) || (t === '|' && depth > 0)) { if (depth === 0) break; depth--; }
        else if (depth === 0 && body.length > 0 && STOP_OPS.has(t)) break;
      }
      body.push(list[j]);
    }
    out += nary(info, convList(body));
    i = j - 1;
  }
  return out;
}

function isFenceRow(n) {
  const k = kids(n);
  if (k.length < 2) return false;
  const f = k[0], l = k[k.length - 1];
  return name(f) === 'mo' && name(l) === 'mo' && attr(f, 'fence') === 'true' && attr(l, 'fence') === 'true';
}

function fence(n) {
  const k = kids(n);
  const beg = text(k[0]).trim(), end = text(k[k.length - 1]).trim();
  return `<m:d><m:dPr><m:begChr m:val="${esc(beg)}"/><m:endChr m:val="${esc(end)}"/></m:dPr>${e(convList(k.slice(1, -1)))}</m:d>`;
}

function spaceFor(n) {
  const w = attr(n, 'width') || '';
  const v = parseFloat(w);
  if (attr(n, 'linebreak') === 'newline') return '';
  if (!v || v <= 0) return '';
  const em = /em$/.test(w) ? v : /ex$/.test(w) ? v / 2 : /px$/.test(w) ? v / 16 : v;
  if (em >= 1.9) return '  ';
  if (em >= 0.9) return ' ';
  if (em >= 0.25) return ' ';
  return ' ';
}

function table(n) {
  const rows = kids(n).filter((r) => ['mtr', 'mlabeledtr'].includes(name(r)));
  const ncol = Math.max(1, ...rows.map((r) => kids(r).filter((c) => name(c) === 'mtd').length));
  const first = rows[0] ? kids(rows[0]).filter((c) => name(c) === 'mtd') : [];
  const aligns = (attr(n, 'columnalign') || '').split(/\s+/).filter(Boolean);
  let mcs = '';
  for (let c = 0; c < ncol; c++) {
    let jc = aligns[c] || aligns[aligns.length - 1] || 'center';
    const cell = first[c];
    if (cell) {
      if (hasClass(cell, 'tml-left') || attr(cell, 'columnalign') === 'left') jc = 'left';
      if (hasClass(cell, 'tml-right') || attr(cell, 'columnalign') === 'right') jc = 'right';
    }
    mcs += `<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="${jc}"/></m:mcPr></m:mc>`;
  }
  let body = '';
  for (const r of rows) {
    const cells = kids(r).filter((c) => name(c) === 'mtd');
    let mr = '';
    for (let c = 0; c < ncol; c++) mr += e(cells[c] ? convList(kids(cells[c])) : '');
    body += wrap('mr', mr);
  }
  return `<m:m><m:mPr><m:baseJc m:val="center"/><m:plcHide m:val="1"/><m:mcs>${mcs}</m:mcs></m:mPr>${body}</m:m>`;
}

function underOver(n, nm, k) {
  const base = k[0];
  const accChar = (x) => { x = unwrap(x); return x && name(x) === 'mo' ? text(x).trim() : null; };
  if (nm === 'mover') {
    const c = accChar(k[1]);
    if (c === '⏞' || c === '︷') return `<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr>${e(conv(base))}</m:groupChr>`;
    if (c && (c === '‾' || c === '¯' || c === '_') && attr(unwrap(k[1]), 'stretchy') === 'true')
      return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${e(conv(base))}</m:bar>`;
    if (c && ACCENTS[c]) return `<m:acc><m:accPr><m:chr m:val="${ACCENTS[c]}"/></m:accPr>${e(conv(base))}</m:acc>`;
    return `<m:limUpp>${e(conv(base))}${wrap('lim', conv(k[1]))}</m:limUpp>`;
  }
  if (nm === 'munder') {
    const c = accChar(k[1]);
    if (c === '⏟' || c === '︸') return `<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr>${e(conv(base))}</m:groupChr>`;
    if (c === '‾' || c === '_' || c === '¯') return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr>${e(conv(base))}</m:bar>`;
    return `<m:limLow>${e(conv(base))}${wrap('lim', conv(k[1]))}</m:limLow>`;
  }
  // munderover
  return `<m:limLow>${e(`<m:limUpp>${e(conv(base))}${wrap('lim', conv(k[2]))}</m:limUpp>`)}${wrap('lim', conv(k[1]))}</m:limLow>`;
}

function enclose(n) {
  const notes = (attr(n, 'notation') || 'longdiv').split(/\s+/);
  const inner = e(convList(kids(n)));
  if (notes.length === 1 && notes[0] === 'top') return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${inner}</m:bar>`;
  if (notes.length === 1 && notes[0] === 'bottom') return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr>${inner}</m:bar>`;
  const box = notes.some((x) => ['box', 'roundedbox', 'circle', 'longdiv'].includes(x));
  let pr = '';
  if (!box) pr += '<m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/>';
  if (notes.includes('horizontalstrike')) pr += '<m:strikeH m:val="1"/>';
  if (notes.includes('verticalstrike')) pr += '<m:strikeV m:val="1"/>';
  if (notes.includes('updiagonalstrike')) pr += '<m:strikeBLTR m:val="1"/>';
  if (notes.includes('downdiagonalstrike')) pr += '<m:strikeTLBR m:val="1"/>';
  return `<m:borderBox>${pr ? `<m:borderBoxPr>${pr}</m:borderBoxPr>` : ''}${inner}</m:borderBox>`;
}

function conv(n) {
  if (!n) return '';
  if (n.nodeType === 3) { const t = text(n).replace(INVISIBLE, ''); return t.trim() ? run(t) : ''; }
  if (n.nodeType !== 1) return '';
  const nm = name(n);
  const k = kids(n);
  switch (nm) {
    case 'math': case 'mrow': case 'mstyle': case 'mpadded': case 'mtd': case 'merror': case 'mfenced':
      if (nm === 'mrow' && isFenceRow(n)) return fence(n);
      if (nm === 'mfenced') {
        const o = attr(n, 'open') ?? '(', c = attr(n, 'close') ?? ')';
        return `<m:d><m:dPr><m:begChr m:val="${esc(o)}"/><m:endChr m:val="${esc(c)}"/></m:dPr>${k.map((x) => e(conv(x))).join('')}</m:d>`;
      }
      return convList(k);
    case 'semantics': return k.length ? conv(k[0]) : '';
    case 'annotation': case 'annotation-xml': case 'none': case 'mprescripts': return '';
    case 'mi': {
      const t = text(n).replace(INVISIBLE, '');
      const v = attr(n, 'mathvariant');
      if (v && VARIANTS[v]) return run(t, VARIANTS[v]);
      const dm = demath(t);
      if (dm) return run(dm.text, dm.style === 'i' && [...dm.text].length === 1 ? {} : { sty: dm.style });
      return run(t, [...t].length > 1 ? { sty: 'p' } : {});
    }
    case 'mn': {
      const dm = demath(text(n));
      if (dm) return run(dm.text, { sty: dm.style === 'i' ? 'i' : dm.style === 'bi' ? 'bi' : 'b' });
      return run(text(n), attr(n, 'mathvariant') === 'bold' ? { sty: 'b' } : {});
    }
    case 'mo': {
      const t = text(n).replace(INVISIBLE, '');
      return t ? run(t, { sty: 'p' }) : '';
    }
    case 'mtext': {
      const t = text(n).replace(/ /g, ' ');
      const dm = demath(t);
      if (dm) return run(dm.text, { nor: true, bold: dm.style.includes('b'), italic: dm.style.includes('i') });
      return run(t, { nor: true, bold: isBold(n) });
    }
    case 'ms': return run(`"${text(n)}"`, { nor: true });
    case 'mspace': { const s = spaceFor(n); return s ? run(s) : ''; }
    case 'mfrac': {
      const lt = (attr(n, 'linethickness') || '').trim();
      const noBar = /^0(\.0*)?([a-z%]*)$/.test(lt);
      const type = noBar ? 'noBar' : attr(n, 'bevelled') === 'true' ? 'skw' : 'bar';
      return `<m:f><m:fPr><m:type m:val="${type}"/></m:fPr>${wrap('num', conv(k[0]))}${wrap('den', conv(k[1]))}</m:f>`;
    }
    case 'msqrt': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(convList(k))}</m:rad>`;
    case 'mroot': return `<m:rad><m:radPr/>${wrap('deg', conv(k[1]))}${e(conv(k[0]))}</m:rad>`;
    case 'msub': case 'msup': case 'msubsup': case 'munder': case 'mover': case 'munderover': {
      const info = naryInfo(n);
      if (info) return nary(info, '');
      if (nm === 'msub') return `<m:sSub>${e(conv(k[0]))}${wrap('sub', conv(k[1]))}</m:sSub>`;
      if (nm === 'msup') {
        // x^\circ → x° (symbole degré), et pas d'exposant sur une base vide
        const sup = unwrap(k[1]);
        if (sup && ['mo', 'mi'].includes(name(sup)) && ['∘', '°'].includes(text(sup).trim())) return conv(k[0]) + run('°', { sty: 'p' });
        const base = conv(k[0]) || run('​');
        return `<m:sSup>${e(base)}${wrap('sup', conv(k[1]))}</m:sSup>`;
      }
      if (nm === 'msubsup') return `<m:sSubSup>${e(conv(k[0]))}${wrap('sub', conv(k[1]))}${wrap('sup', conv(k[2]))}</m:sSubSup>`;
      return underOver(n, nm, k);
    }
    case 'mmultiscripts': {
      const base = k[0];
      const rest = k.slice(1);
      const p = rest.findIndex((x) => name(x) === 'mprescripts');
      const post = p < 0 ? rest : rest.slice(0, p);
      const pre = p < 0 ? [] : rest.slice(p + 1);
      let core = conv(base);
      if (post.length >= 2) core = `<m:sSubSup>${e(core)}${wrap('sub', conv(post[0]))}${wrap('sup', conv(post[1]))}</m:sSubSup>`;
      if (pre.length >= 2) core = `<m:sPre>${wrap('sub', conv(pre[0]))}${wrap('sup', conv(pre[1]))}${e(core)}</m:sPre>`;
      return core;
    }
    case 'mtable': return table(n);
    case 'menclose': return enclose(n);
    case 'mphantom': return `<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr>${e(convList(k))}</m:phant>`;
    default: return convList(k);
  }
}

/** MathML (élément <math>) → OMML pour un .docx */
export function mathmlToOmml(mathEl, display) {
  const inner = conv(mathEl) || run(' ');
  const om = `<m:oMath>${inner}</m:oMath>`;
  return display ? `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${om}</m:oMathPara>` : om;
}

/** Variante du même OMML au format « HTML Word » (texte directement dans <m:r>) pour le presse-papiers */
export function ommlToWordHtml(omml) {
  return omml.replace(
    /<m:r>(<m:rPr>(.*?)<\/m:rPr>)?(<w:rPr>(.*?)<\/w:rPr>)?<m:t xml:space="preserve">(.*?)<\/m:t><\/m:r>/g,
    (_, _a, mrPr = '', _b, wrPr = '', t) => {
      const nor = mrPr.includes('<m:nor/>');
      const bold = /m:val="b/.test(mrPr) || wrPr.includes('<w:b/>');
      const italic = (!nor && !mrPr.includes('m:sty') && /\p{L}/u.test(t)) || mrPr.includes('m:val="i') || wrPr.includes('<w:i/>');
      const font = nor ? '' : ` style='font-family:"Cambria Math",serif'`;
      let r = `<m:r>${mrPr ? `<m:rPr>${mrPr}</m:rPr>` : ''}${t}</m:r>`;
      if (italic) r = `<i>${r}</i>`;
      if (bold) r = `<b>${r}</b>`;
      return `<span${font}>${r}</span>`;
    },
  ).replace(/<w:rPr>.*?<\/w:rPr>/g, '');
}
