// HTML (aperçu) → document Word .docx avec équations natives (OMML)
import JSZip from 'jszip';
import { esc, mathmlToOmml } from './omml.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const lower = (n) => (n.nodeName || '').toLowerCase();

export function mathXml(span) {
  const display = span.classList.contains('math-display');
  const math = span.querySelector('math');
  if (math) return mathmlToOmml(math, display);
  // Erreur LaTeX : on garde le code source tel quel
  const r = `<w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t xml:space="preserve">${esc(span.dataset.tex || span.textContent)}</w:t></w:r>`;
  return r;
}

class Builder {
  constructor(font) {
    this.font = font;
    this.body = [];
    this.nums = []; // { numId, abstractId }
    this.rels = [];
  }

  newNum(ordered) {
    const numId = this.nums.length + 1;
    this.nums.push({ numId, abstractId: ordered ? 2 : 1 });
    return numId;
  }

  link(href) {
    const id = `rIdL${this.rels.length + 1}`;
    this.rels.push({ id, href });
    return id;
  }

  // Renvoie une liste de segments : { t: 'inline', xml } ou { t: 'display', xml }
  inlines(node, fmt = {}, out = []) {
    for (const c of Array.from(node.childNodes)) {
      if (c.nodeType === 3) {
        let t = c.nodeValue;
        if (!fmt.pre) t = t.replace(/[ \t\r\n]+/g, ' ');
        if (t) out.push({ t: 'inline', xml: this.textRun(t, fmt), text: t, fmt });
        continue;
      }
      if (c.nodeType !== 1) continue;
      const tag = lower(c);
      if (c.classList && c.classList.contains('math')) {
        out.push({ t: c.classList.contains('math-display') ? 'display' : 'inline', xml: mathXml(c), math: true });
        continue;
      }
      if (c.classList && c.classList.contains('math-block')) { this.inlines(c, fmt, out); continue; }
      switch (tag) {
        case 'strong': case 'b': this.inlines(c, { ...fmt, b: true }, out); break;
        case 'em': case 'i': this.inlines(c, { ...fmt, i: true }, out); break;
        case 'del': case 's': this.inlines(c, { ...fmt, strike: true }, out); break;
        case 'sup': this.inlines(c, { ...fmt, va: 'superscript' }, out); break;
        case 'sub': this.inlines(c, { ...fmt, va: 'subscript' }, out); break;
        case 'u': this.inlines(c, { ...fmt, u: true }, out); break;
        case 'code': this.inlines(c, { ...fmt, code: true }, out); break;
        case 'br': out.push({ t: 'inline', xml: '<w:r><w:br/></w:r>' }); break;
        case 'img': if (c.alt) out.push({ t: 'inline', xml: this.textRun(`[${c.alt}]`, fmt) }); break;
        case 'a': {
          const sub = this.inlines(c, { ...fmt, link: true }, []);
          const href = c.getAttribute('href');
          const inner = sub.filter((s) => s.t === 'inline').map((s) => s.xml).join('');
          if (href && /^(https?|mailto):/i.test(href)) out.push({ t: 'inline', xml: `<w:hyperlink r:id="${this.link(href)}" w:history="1">${inner}</w:hyperlink>` });
          else out.push({ t: 'inline', xml: inner });
          break;
        }
        case 'input': if (c.type === 'checkbox') out.push({ t: 'inline', xml: this.textRun(c.checked ? '☒ ' : '☐ ', fmt) }); break;
        case 'ul': case 'ol': case 'p': case 'div': case 'blockquote': case 'pre': case 'table':
          break; // les blocs sont gérés par blocks()
        default: this.inlines(c, fmt, out);
      }
    }
    return out;
  }

  textRun(t, f) {
    let rPr = '';
    if (f.link) rPr += '<w:rStyle w:val="Hyperlink"/>';
    if (f.code) rPr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>';
    if (f.b) rPr += '<w:b/>';
    if (f.i) rPr += '<w:i/>';
    if (f.strike) rPr += '<w:strike/>';
    if (f.u) rPr += '<w:u w:val="single"/>';
    if (f.code) rPr += '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>';
    if (f.va) rPr += `<w:vertAlign w:val="${f.va}"/>`;
    const lines = t.split('\n');
    return lines.map((l, i) => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${esc(l)}</w:t></w:r>`).join('');
  }

  // Écrit des paragraphes à partir de segments, en coupant autour des équations centrées
  paragraphs(segs, pPr = '', contPPr = pPr) {
    let cur = [], first = true;
    const flush = () => {
      // retire les espaces de début / fin
      while (cur.length && cur[0].text !== undefined && !cur[0].text.trim()) cur.shift();
      while (cur.length && cur[cur.length - 1].text !== undefined && !cur[cur.length - 1].text.trim()) cur.pop();
      if (!cur.length) return;
      const a = cur[0], z = cur[cur.length - 1];
      if (a.fmt && !a.fmt.pre) cur[0] = { ...a, xml: this.textRun(a.text.replace(/^\s+/, ''), a.fmt) };
      if (z.fmt && !z.fmt.pre) {
        const zz = cur[cur.length - 1];
        const t = (zz === z ? z.text : zz.text.replace(/^\s+/, '')).replace(/\s+$/, '');
        cur[cur.length - 1] = { ...z, xml: this.textRun(t, z.fmt) };
      }
      this.body.push(`<w:p>${(first ? pPr : contPPr) ? `<w:pPr>${first ? pPr : contPPr}</w:pPr>` : ''}${cur.map((s) => s.xml).join('')}</w:p>`);
      first = false;
      cur = [];
    };
    for (const s of segs) {
      if (s.t === 'display') {
        flush();
        this.body.push(`<w:p><w:pPr><w:pStyle w:val="Equation"/></w:pPr>${s.xml}</w:p>`);
      } else cur.push(s);
    }
    flush();
  }

  blocks(node, ctx = {}) {
    for (const c of Array.from(node.childNodes)) {
      if (c.nodeType === 3) {
        if (c.nodeValue.trim()) this.paragraphs(this.inlines({ childNodes: [c] }, {}), this.basePPr(ctx));
        continue;
      }
      if (c.nodeType !== 1) continue;
      const tag = lower(c);
      if (/^h[1-6]$/.test(tag)) {
        this.paragraphs(this.inlines(c), `<w:pStyle w:val="Heading${tag[1]}"/>`);
      } else if (tag === 'p') {
        this.paragraphs(this.inlines(c), this.basePPr(ctx));
      } else if (c.classList.contains('math-block') || c.classList.contains('math')) {
        this.paragraphs(this.inlines(c.classList.contains('math') ? { childNodes: [c] } : c));
      } else if (tag === 'ul' || tag === 'ol') {
        this.list(c, ctx);
      } else if (tag === 'blockquote') {
        this.blocks(c, { ...ctx, quote: true });
      } else if (tag === 'pre') {
        const lines = c.textContent.replace(/\n$/, '').split('\n');
        for (const l of lines) this.body.push(`<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr>${this.textRun(l || ' ', { code: false })}</w:p>`);
      } else if (tag === 'table') {
        this.table(c);
      } else if (tag === 'hr') {
        this.body.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A0A0A0"/></w:pBdr></w:pPr></w:p>');
      } else {
        this.blocks(c, ctx);
      }
    }
  }

  basePPr(ctx) {
    if (ctx.quote) return '<w:pStyle w:val="Quote"/>';
    if (ctx.level !== undefined) return `<w:pStyle w:val="ListParagraph"/><w:ind w:left="${720 + ctx.level * 360}"/>`;
    return '';
  }

  list(listEl, ctx) {
    const level = ctx.level === undefined ? 0 : ctx.level + 1;
    const numId = this.newNum(lower(listEl) === 'ol');
    const start = parseInt(listEl.getAttribute('start') || '1', 10);
    if (start !== 1) this.nums[numId - 1].start = start;
    for (const li of Array.from(listEl.children)) {
      if (lower(li) !== 'li') continue;
      const numPPr = `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${Math.min(level, 8)}"/><w:numId w:val="${numId}"/></w:numPr>`;
      const contPPr = `<w:pStyle w:val="ListParagraph"/><w:ind w:left="${720 + level * 360}"/>`;
      let first = true;
      // contenu inline direct (listes « serrées ») puis blocs
      const inlineNodes = [], blockNodes = [];
      for (const k of Array.from(li.childNodes)) {
        const t = lower(k);
        if (k.nodeType === 1 && (['p', 'ul', 'ol', 'blockquote', 'pre', 'table', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr'].includes(t) || (k.classList && k.classList.contains('math-block')))) blockNodes.push(k);
        else if (!blockNodes.length) inlineNodes.push(k);
        else blockNodes.push(k);
      }
      const segs = this.inlines({ childNodes: inlineNodes });
      if (segs.some((s) => s.math || (s.text && s.text.trim()) || !s.text)) {
        this.paragraphs(segs, numPPr, contPPr);
        first = false;
      }
      for (const b of blockNodes) {
        const t = lower(b);
        if (t === 'p') { this.paragraphs(this.inlines(b), first ? numPPr : contPPr, contPPr); first = false; }
        else if (t === 'ul' || t === 'ol') this.list(b, { ...ctx, level });
        else if (b.nodeType === 3) { if (b.nodeValue.trim()) { this.paragraphs(this.inlines({ childNodes: [b] }), first ? numPPr : contPPr, contPPr); first = false; } }
        else { this.blocks({ childNodes: [b] }, { ...ctx, level }); }
      }
      if (first) this.body.push(`<w:p><w:pPr>${numPPr}</w:pPr></w:p>`);
    }
  }

  table(tbl) {
    const rows = Array.from(tbl.querySelectorAll('tr'));
    const ncol = Math.max(1, ...rows.map((r) => r.children.length));
    const w = Math.floor(9000 / ncol);
    let x = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${w}"/>`.repeat(ncol)}</w:tblGrid>`;
    for (const r of rows) {
      const head = r.parentNode && lower(r.parentNode) === 'thead';
      x += `<w:tr>${head ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}`;
      const cells = Array.from(r.children);
      for (let i = 0; i < ncol; i++) {
        const cell = cells[i];
        const isTh = cell && lower(cell) === 'th';
        const align = cell && (cell.getAttribute('align') || cell.style.textAlign);
        const jc = align === 'center' ? '<w:jc w:val="center"/>' : align === 'right' ? '<w:jc w:val="right"/>' : '';
        const saved = this.body;
        this.body = [];
        if (cell) this.paragraphs(this.inlines(cell, isTh ? { b: true } : {}), `<w:spacing w:after="0"/>${jc}`);
        const inner = this.body.join('') || '<w:p/>';
        this.body = saved;
        x += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${isTh ? '<w:shd w:val="clear" w:color="auto" w:fill="E7EEF7"/>' : ''}</w:tcPr>${inner}</w:tc>`;
      }
      x += '</w:tr>';
    }
    x += '</w:tbl>';
    this.body.push(x, '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
  }
}

function stylesXml(font) {
  const { name, size } = font; // size en demi-points
  const h = (lvl, sz, color = '1F3864') => `<w:style w:type="paragraph" w:styleId="Heading${lvl}"><w:name w:val="heading ${lvl}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${lvl <= 2 ? 360 : 240}" w:after="120"/><w:jc w:val="left"/><w:outlineLvl w:val="${lvl - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:lang w:val="fr-FR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
${h(1, size + 8)}${h(2, size + 4)}${h(3, size + 2)}${h(4, size, '2F5496')}${h(5, size, '2F5496')}${h(6, size, '2F5496')}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="A5B4C8"/></w:pBdr><w:ind w:left="567" w:right="567"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Equation"><w:name w:val="Équation"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="120" w:after="120"/><w:jc w:val="center"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="19"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
}

function numberingXml(nums) {
  const bullets = ['●', '○', '■', '●', '○', '■', '●', '○', '■'];
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman', 'decimal', 'lowerLetter', 'lowerRoman'];
  let a1 = '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>';
  let a2 = '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>';
  for (let l = 0; l < 9; l++) {
    const ind = `<w:pPr><w:ind w:left="${720 + l * 360}" w:hanging="360"/></w:pPr>`;
    a1 += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[l]}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/><w:sz w:val="16"/></w:rPr></w:lvl>`;
    a2 += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmts[l]}"/><w:lvlText w:val="%${l + 1}."/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
  }
  a1 += '</w:abstractNum>';
  a2 += '</w:abstractNum>';
  const numXml = nums.map((n) => `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="${n.abstractId}"/>${n.abstractId === 2
    ? Array.from({ length: 9 }, (_, l) => `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="${l === 0 && n.start ? n.start : 1}"/></w:lvlOverride>`).join('')
    : ''}</w:num>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${NS}>${a1}${a2}${numXml}</w:numbering>`;
}

/** Construit le .docx (Blob) à partir du conteneur d'aperçu */
export async function buildDocx(root, font) {
  const b = new Builder(font);
  b.blocks(root);
  if (!b.body.length) b.body.push('<w:p/>');
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${b.body.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Texte converti</dc:title><dc:creator>MD vers Word</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>${b.rels.map((r) => `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(r.href)}" TargetMode="External"/>`).join('')}</Relationships>`);
  zip.file('word/document.xml', doc);
  zip.file('word/styles.xml', stylesXml(font));
  zip.file('word/numbering.xml', numberingXml(b.nums));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
