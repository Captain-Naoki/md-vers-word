// Construit un fichier HTML unique et autonome : MD-vers-Word.html
import { build } from 'esbuild';
import fs from 'node:fs';

const res = await build({ entryPoints: ['src/main.js'], bundle: true, minify: true, format: 'iife', write: false, target: 'chrome100', legalComments: 'none' });
const js = res.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');

const woff = fs.readFileSync('node_modules/temml/dist/Temml.woff2').toString('base64');
const css = fs.readFileSync('node_modules/temml/dist/Temml-Local.css', 'utf8')
  .replace("url('Temml.woff2')", `url(data:font/woff2;base64,${woff})`);

const html = fs.readFileSync('src/index.html', 'utf8')
  .replace('/*__TEMML_CSS__*/', () => css)
  .replace('/*__APP_JS__*/', () => js);
// Version à télécharger / ouvrir hors ligne, et version publiée sur Vercel (dist/index.html)
fs.writeFileSync('MD-vers-Word.html', html);
fs.mkdirSync('dist', { recursive: true });
fs.writeFileSync('dist/index.html', html);
console.log(`MD-vers-Word.html et dist/index.html : ${(html.length / 1024).toFixed(0)} Ko`);
