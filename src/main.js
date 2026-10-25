import { markdownToHtml, liftDisplayMath, stripCitations } from './markdown.js';
import { buildDocx } from './docx.js';
import { buildClipboardHtml, buildPlainText, copyRich, copySingleEquation } from './clipboard.js';

const FONTS = {
  tnr: { name: 'Times New Roman', css: "'Times New Roman',serif", size: 24 },
  calibri: { name: 'Calibri', css: 'Calibri,sans-serif', size: 22 },
  arial: { name: 'Arial', css: 'Arial,sans-serif', size: 22 },
  cambria: { name: 'Cambria', css: 'Cambria,serif', size: 24 },
  garamond: { name: 'Garamond', css: 'Garamond,serif', size: 24 },
  none: { name: 'Calibri', css: '', size: 22 },
};

const $ = (id) => document.getElementById(id);
const src = $('src');
const preview = $('preview');
const store = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch { /* stockage indisponible */ } },
};

const SAMPLE = String.raw`## Le calcul étape par étape

Pour chaque simulation de **Modèle Climatique Global (GCM)** :

1. **Calcul de la hausse relative de l'extrême de pluie ($\Delta P_{99.5}$ en %)** : On calcule le pourcentage d'augmentation du *99,5e percentile* des précipitations quotidiennes entre la fin du XXIe siècle (2070–2099, sous scénario fort SSP3-70) et la période historique (1985–2014) :

   $$\Delta P_{99.5}(\%) = \frac{P_{99.5,\text{futur}} - P_{99.5,\text{historique}}}{P_{99.5,\text{historique}}} \times 100$$

2. **Calcul du réchauffement moyen de la température ($\Delta T$ en °C)** : On calcule l'augmentation de la température moyenne de l'air en surface sur la même région entre les deux périodes :

   $$\Delta T(^\circ\text{C}) = T_{\text{futur}} - T_{\text{historique}}$$

   *(Par exemple, dans l'étude, les 5 modèles testés affichent un réchauffement local allant de +1,78 °C à +3,49 °C.)*

$$\text{Pente de dégradation} = \frac{\text{Erreur}_{\text{Futur (+3,0 °C)}} - \text{Erreur}_{\text{Historique (+0 °C)}}}{\Delta T} = \frac{11 - 5}{3{,}0 - 0} = \mathbf{+2{,}0\text{ mm/jour par }^\circ\text{C}}$$

### Autres exemples

| Indicateur | Formule |
|---|---|
| Erreur quadratique | $\mathrm{RMSE} = \sqrt{\frac{1}{n}\sum_{i=1}^{n}(\hat{y}_i - y_i)^2}$ |
| Clausius-Clapeyron | $\frac{d\ln e_s}{dT} \approx 7\,\%\ \text{par}\ ^\circ\text{C}$ |

$$\int_0^{\infty} e^{-x}\,dx = 1 \qquad \begin{pmatrix} a & b \\ c & d \end{pmatrix}$$
`;

let timer = null;
function render() {
  const raw = src.value;
  const md = $('optCite').checked ? stripCitations(raw) : raw;
  preview.innerHTML = md.trim() ? markdownToHtml(md) : '<p class="empty">L’aperçu apparaîtra ici…</p>';
  liftDisplayMath(preview);
  const n = preview.querySelectorAll('span.math').length;
  const errs = preview.querySelectorAll('.math-error').length;
  $('stats').textContent = `${n} équation${n > 1 ? 's' : ''}` + (errs ? ` · ${errs} non reconnue${errs > 1 ? 's' : ''} (en rouge)` : '');
  for (const s of preview.querySelectorAll('span.math')) s.title = 'Cliquer pour copier uniquement cette équation';
  store.set('md2word.src', raw);
}

function applyFont() {
  const f = FONTS[$('optFont').value];
  preview.style.fontFamily = f.css || "'Times New Roman',serif";
  preview.style.fontSize = `${f.size / 2}pt`;
  store.set('md2word.font', $('optFont').value);
}

let toastTimer = null;
function toast(msg, kind = 'ok') {
  const t = $('toast');
  t.textContent = msg;
  t.dataset.kind = kind;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 4200);
}

src.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(render, 180); });
$('optCite').addEventListener('change', () => { store.set('md2word.cite', $('optCite').checked ? '1' : '0'); render(); });
$('optFont').addEventListener('change', applyFont);
$('optMode').addEventListener('change', () => store.set('md2word.mode', $('optMode').value));

$('btnCopy').addEventListener('click', () => {
  if (!src.value.trim()) return toast('Collez d’abord un texte à gauche.', 'warn');
  const html = buildClipboardHtml(preview, { font: FONTS[$('optFont').value], mode: $('optMode').value });
  const ok = copyRich(html, buildPlainText(preview));
  toast(ok ? 'Copié ! Collez maintenant dans Word avec Ctrl+V.' : 'La copie a échoué : utilisez « Télécharger .docx ».', ok ? 'ok' : 'warn');
});

$('btnDocx').addEventListener('click', async () => {
  if (!src.value.trim()) return toast('Collez d’abord un texte à gauche.', 'warn');
  try {
    const blob = await buildDocx(preview, FONTS[$('optFont').value]);
    const a = document.createElement('a');
    const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
    a.href = URL.createObjectURL(blob);
    a.download = `NotebookLM_vers_Word_${stamp}.docx`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    toast('Fichier .docx créé : ouvrez-le, puis copiez-collez où vous voulez dans votre mémoire.');
  } catch (err) {
    console.error(err);
    toast('Erreur pendant la création du .docx : ' + err.message, 'warn');
  }
});

$('btnPaste').addEventListener('click', async () => {
  try {
    const t = await navigator.clipboard.readText();
    if (!t) return toast('Le presse-papiers est vide.', 'warn');
    src.value = t;
    render();
    toast('Texte collé.');
  } catch {
    src.focus();
    toast('Cliquez dans la zone de gauche puis faites Ctrl+V.', 'warn');
  }
});

$('btnClear').addEventListener('click', () => { src.value = ''; render(); src.focus(); });
$('btnSample').addEventListener('click', () => { src.value = SAMPLE; render(); });

preview.addEventListener('click', (e) => {
  const span = e.target.closest('span.math');
  if (!span) return;
  const ok = copySingleEquation(span);
  span.classList.add('flash');
  setTimeout(() => span.classList.remove('flash'), 600);
  toast(ok ? 'Équation copiée : collez-la dans Word (Ctrl+V), elle devient une équation modifiable.' : 'Équation non reconnue.', ok ? 'ok' : 'warn');
});

// Restauration de la session précédente
src.value = store.get('md2word.src') || '';
if (store.get('md2word.cite') === '1') $('optCite').checked = true;
if (store.get('md2word.font') && FONTS[store.get('md2word.font')]) $('optFont').value = store.get('md2word.font');
if (store.get('md2word.mode')) $('optMode').value = store.get('md2word.mode');
applyFont();
render();

// Accès pour les tests automatisés
window.md2word = { render, buildDocx: () => buildDocx(preview, FONTS[$('optFont').value]), buildClipboardHtml: () => buildClipboardHtml(preview, { font: FONTS[$('optFont').value], mode: $('optMode').value }) };
