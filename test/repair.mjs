// Vérifie la réparation des textes aplatis et l'absence d'effet sur un Markdown normal
import fs from 'node:fs';
import { repairFlattened, stripGluedCitations } from '../src/repair.js';
let fail = 0;
const check = (name, ok) => { console.log(ok ? 'OK ' : 'ÉCHEC', name); if (!ok) fail++; };
const normal = fs.readFileSync('test/cas.md', 'utf8');
check('cas.md inchangé par la réparation', repairFlattened(normal) === normal);
const out = repairFlattened(fs.readFileSync('test/aplati.md', 'utf8'));
fs.writeFileSync('test/aplati.repare.md', out);
check('références collées retirées', !/observées78|CCAM\)18|\$i\$8|journée8/.test(out));
check('unités conservées', stripGluedCitations('12 km2 et CO2, SSP3-70, ERA5 et x2.') === '12 km2 et CO2, SSP3-70, ERA5 et x2.');
check('mots camelCase conservés', repairFlattened('x'.repeat(0) + 'PyTorch et GitHub. '.repeat(20)).includes('PyTorch et GitHub'));
// Tableau copié depuis NotebookLM : cellules séparées par des tabulations
const tsv = [
  'Voici les données :',
  'Journée\tMembre 1 (Tirage 1)\tMembre 2 (Tirage 2)\t...\tMembre 10 (Tirage 10)',
  'Jour 1 (01/01/2070)\t0,0 mm\t1,2 mm\t...\t0,0 mm',
  'Jour 2 (02/01/2070)\t12,5 mm\t10,1 mm\t...\t14,0 mm',
  '...\t...\t...\t...\t...',
  'Jour 10 950 (31/12/2099)\t0,5 mm\t$x_1$\t...\t0,2 mm',
  'Texte après le tableau.',
].join('\n');
const tab = repairFlattened(tsv);
check('tableau à tabulations converti', tab.includes('| Journée | Membre 1 (Tirage 1) |') && tab.includes('| --- |') && tab.includes('| Jour 10 950 (31/12/2099) | 0,5 mm | $x_1$ |'));
check('texte autour du tableau conservé', tab.startsWith('Voici les données :\n\n|') && tab.endsWith('\nTexte après le tableau.'));
check('une seule tabulation isolée ignorée', repairFlattened('a\tb\nsuite') === 'a\tb\nsuite');
console.log('\n' + tab + '\n');
console.log('\n' + out);
process.exit(fail ? 1 : 0);
