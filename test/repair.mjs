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
console.log('\n' + out);
process.exit(fail ? 1 : 0);
