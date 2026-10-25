# Markdown → Word

**Collez le texte de NotebookLM, récupérez-le prêt pour Microsoft Word — équations comprises.**

Quand on copie une réponse de NotebookLM (ou de ChatGPT, Gemini, Claude…) et qu'on la colle dans Word, on obtient du code illisible :

```text
$$\Delta P_{99.5}(\%) = \frac{P_{99.5,\text{futur}} - P_{99.5,\text{historique}}}{P_{99.5,\text{historique}}} \times 100$$
```

Cet outil transforme ce texte (Markdown + LaTeX) en contenu Word propre : les **équations deviennent de vraies équations Word** (les mêmes que *Insertion → Équation*), toujours modifiables. Le gras, l'italique, les titres, les listes et les tableaux sont conservés.

🔗 **Version en ligne :** `https://<votre-projet>.vercel.app` *(à remplacer après le déploiement)*

---

## Utilisation

1. Dans NotebookLM, copiez la réponse.
2. Collez-la dans la zone de gauche de l'outil. L'aperçu à droite montre le rendu final.
3. Choisissez une méthode :

| Méthode | Comment | Quand l'utiliser |
|---|---|---|
| **Copier pour Word** | Bouton bleu, puis `Ctrl+V` dans Word | Le plus rapide |
| **Télécharger .docx** | Ouvrez le fichier puis copiez-collez son contenu dans votre document | Méthode la plus sûre, fonctionne à tous les coups |
| **Une équation seule** | Cliquez sur une équation dans l'aperçu, puis `Ctrl+V` dans Word | Pour insérer ou remplacer une formule précise |

**Options**

- **Police** : Times New Roman 12 (par défaut), Calibri, Arial, Cambria, Garamond, ou celle de votre document.
- **Réparer la mise en forme** (activé par défaut) : si le texte a été copié à la souris et que tout est collé sur une seule ligne (« généralePour », « ?La MSE »), l'outil reconstitue les titres, paragraphes, listes et sous-listes, et retire les numéros de référence collés aux mots (« observées78. » → « observées. »). Un texte déjà bien formaté n'est pas modifié.
- **Retirer les références** : supprime les renvois `[1]`, `[2-4]` et les numéros collés à la fin d'un mot, dans tout le texte (désactivé par défaut). Les unités et sigles comme `km2`, `CO2`, `SSP3-70`, `ERA5` sont conservés.
- **Copie → MathML** : méthode de collage alternative si la méthode par défaut ne donne pas un bon résultat dans votre version de Word.

**Confidentialité :** toute la conversion se fait dans votre navigateur. Aucun texte n'est envoyé à un serveur, même sur la version en ligne.

**Hors ligne :** le fichier [`MD-vers-Word.html`](MD-vers-Word.html) contient tout l'outil. Téléchargez-le et ouvrez-le par un double-clic, sans Internet.

## Syntaxe reconnue

- Équations en ligne : `$...$` ou `\(...\)`
- Équations centrées : `$$...$$`, `\[...\]`, `\begin{equation}`, `\begin{align}`
- Markdown GFM : titres, **gras**, *italique*, listes (imbriquées et numérotées), tableaux, citations, code, liens
- LaTeX : fractions, indices et exposants, racines, sommes, intégrales, limites, matrices (`pmatrix`, `bmatrix`…), `cases`, `aligned`, accents (`\hat`, `\bar`, `\vec`…), `\overbrace` / `\underbrace`, `\mathbb`, `\mathcal`, `\mathbf`, `\text`, `\cancel`, la virgule décimale `{,}` et `^\circ` (converti en °)
- Un `$` utilisé comme symbole monétaire (« 5 $ et 10 $ ») n'est pas pris pour une équation.

Une formule LaTeX invalide est affichée en rouge, telle quelle, au lieu de bloquer la conversion.

## Fonctionnement

```text
Markdown + LaTeX ──marked──▶ HTML ──Temml──▶ MathML ──convertisseur maison──▶ OMML (équations Word)
                                                                          ├─▶ presse-papiers (HTML Word)
                                                                          └─▶ fichier .docx (JSZip)
```

| Fichier | Rôle |
|---|---|
| [`src/markdown.js`](src/markdown.js) | Analyse le Markdown et repère les équations (`$`, `$$`, `\(`, `\[`) |
| [`src/repair.js`](src/repair.js) | Reconstitue la mise en forme des textes copiés à la souris et retire les références collées |
| [`src/omml.js`](src/omml.js) | Convertit le MathML en OMML, le format natif des équations Word |
| [`src/docx.js`](src/docx.js) | Génère le fichier `.docx` (styles, listes, tableaux, équations) |
| [`src/clipboard.js`](src/clipboard.js) | Prépare le contenu du presse-papiers au format compris par Word |
| [`src/main.js`](src/main.js) | Interface : aperçu, boutons, options |
| [`src/index.html`](src/index.html) | Page et styles |
| [`build.mjs`](build.mjs) | Assemble le tout en un seul fichier HTML autonome |
| [`test/`](test) | Banc de test : convertit [`test/cas.md`](test/cas.md) en `.docx` sans navigateur |

## Développement

Prérequis : [Node.js](https://nodejs.org) 18 ou plus récent.

```bash
npm install
npm run build   # produit MD-vers-Word.html et dist/index.html
npm test        # teste la réparation (test/aplati.md) puis convertit test/cas.md en test/cas.docx
```

Ouvrez ensuite `MD-vers-Word.html` dans Chrome ou Edge.

## Déploiement sur Vercel

Le fichier [`vercel.json`](vercel.json) est déjà configuré (commande `npm run build`, dossier publié `dist`).

1. Créez un compte sur [vercel.com](https://vercel.com) avec votre compte GitHub.
2. **Add New → Project**, puis importez ce dépôt.
3. Laissez les réglages proposés et cliquez sur **Deploy**.

Chaque `git push` sur la branche principale redéploie automatiquement le site.

## Limites connues

- Le collage direct (« Copier pour Word ») dépend de la version de Word. Si les équations n'apparaissent pas correctement, utilisez **Télécharger .docx**, qui produit toujours de vraies équations Word.
- Fonctionne avec Chrome, Edge et les navigateurs récents. Le rendu de l'aperçu utilise la police Cambria Math, présente sur Windows.

## Licences

Code sous licence [MIT](LICENSE). Bibliothèques utilisées : [marked](https://github.com/markedjs/marked) (MIT), [Temml](https://github.com/ronkok/Temml) (MIT), [JSZip](https://github.com/Stuk/jszip) (MIT).
