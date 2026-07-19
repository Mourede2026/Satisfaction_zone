# Déploiement sur GitHub Pages (avec icônes "écran d'accueil")

## Fichiers à uploader sur GitHub (12 fichiers, à la racine du dépôt)

| Fichier | Rôle |
|---|---|
| `index.html` | Page d'accueil |
| `usagers.html` | Formulaire Usagers |
| `personnels.html` | Formulaire Personnel soignant |
| `cogecs.html` | Formulaire COGECS |
| `admin.html` | Tableau de bord admin |
| `config.js` | **Le seul fichier à modifier** — l'URL de votre API Apps Script |
| `manifest.json` | Fait fonctionner "Ajouter à l'écran d'accueil" sur Android |
| `icon-180.png`, `icon-192.png`, `icon-512.png`, `favicon-32.png` | Icônes de l'application (logo bleu avec une checklist) |

> ⚠️ **Tous les noms de fichiers sont en minuscules.** GitHub Pages est
> sensible à la casse — `Index.html` ne fonctionne pas là où `index.html`
> est attendu, en particulier pour la page d'accueil (`/`).

`Code.gs` **ne va pas sur GitHub** : il reste dans l'éditeur Apps Script (c'est le seul moyen d'écrire dans votre Google Sheet).

## Étape 1 — Apps Script (si pas déjà fait)

1. Remplacez le contenu de `Code.gs` dans l'éditeur Apps Script par la version fournie.
2. **Déployer > Gérer les déploiements > ✏️ > Nouvelle version > Déployer.**
3. Copiez l'URL `.../exec`.

## Étape 2 — `config.js`

Ouvrez `config.js`, remplacez :
```js
const API_URL = "COLLEZ_ICI_VOTRE_URL_APPS_SCRIPT/exec";
```
par votre vraie URL.

## Étape 3 — GitHub

1. Créez un dépôt (Public).
2. **Add file > Upload files**, glissez les 12 fichiers ci-dessus, **Commit changes**.
3. **Settings > Pages** → Source : `Deploy from a branch` → Branch : `main` / `(root)` → **Save**.
4. Votre site est à `https://VOTRE-NOM.github.io/VOTRE-DEPOT/`.

## Étape 4 — Installer sur l'écran d'accueil (mobile)

Une fois le site en ligne, chaque page (`index.html`, `usagers.html`, etc.)
peut être ajoutée à l'écran d'accueil **comme une vraie application**, avec
l'icône bleue et sans barre d'adresse visible :

**Sur Android (Chrome) :**
1. Ouvrez le lien (ex. `.../usagers.html`).
2. Menu ⋮ (en haut à droite) → **Ajouter à l'écran d'accueil** (ou "Installer l'application" si Chrome le propose directement).
3. Confirmez → l'icône apparaît sur l'écran d'accueil.

**Sur iPhone (Safari — obligatoire, ça ne marche pas depuis Chrome iOS) :**
1. Ouvrez le lien dans **Safari**.
2. Bouton **Partager** (le carré avec la flèche vers le haut, en bas de l'écran).
3. **Sur l'écran d'accueil**.
4. Confirmez → l'icône apparaît.

Faites ceci séparément pour chaque lien que vous voulez en accès direct
(ex. les enquêteurs Usagers installent `usagers.html`, ceux du COGECS
installent `cogecs.html`, vous installez `admin.html`) — chacun aura sa
propre icône et son propre nom sur l'écran d'accueil.

## Mettre à jour un fichier plus tard

Sur GitHub : ouvrez le fichier → ✏️ → modifiez → **Commit changes**. Le
site se met à jour en 1-2 minutes. (`Code.gs` doit toujours être redéployé
séparément depuis l'éditeur Apps Script.)

## En cas de souci

- **404 sur la page d'accueil** : vérifiez que le fichier s'appelle bien
  `index.html` (minuscules) dans le dépôt.
- **404 sur un lien précis** (ex. Usagers) : vérifiez que le nom exact du
  fichier sur GitHub correspond au lien cliqué (`usagers.html`, pas
  `Usagers.html` ni `FormUsagers.html`).
- **"API_URL non configurée" ou rien ne se passe à l'envoi** : `config.js`
  n'a pas la bonne URL, ou `Code.gs` n'a pas été redéployé en nouvelle version.
- **Pas d'icône lors de l'ajout à l'écran d'accueil** : vérifiez que
  `manifest.json`, `icon-180.png`, `icon-192.png`, `icon-512.png` et
  `favicon-32.png` sont bien présents dans le dépôt, au même niveau que les
  pages HTML.
