# Déploiement sur GitHub Pages

## Ce qui change par rapport à la version précédente

Avant : tout (pages ET code) vivait dans l'éditeur Apps Script.

Maintenant, il y a **deux parties séparées** :

| Partie | Rôle | Où ça vit |
|---|---|---|
| **API** (`Code.gs`) | Lit/écrit votre Google Sheet | Reste dans Apps Script (obligatoire — c'est le seul moyen d'accéder à Google Sheets) |
| **Pages** (`Index.html`, les 3 formulaires, `AdminDashboard.html`, `config.js`) | Ce que voient les enquêteurs et vous | **GitHub Pages** — une URL fixe, facile à mettre à jour |

Les pages appellent l'API Apps Script en arrière-plan (via `fetch`), au lieu d'être servies par elle. **Votre Google Sheet reste la seule base de données**, rien ne change de ce côté.

## Étape 1 — Mettre à jour l'Apps Script (obligatoire, même si déjà déployé)

Le fichier `Code.gs` a changé de logique (il répond maintenant en JSON plutôt que de servir des pages). Dans l'éditeur Apps Script de votre Google Sheet :

1. Remplacez tout le contenu de `Code.gs` par le nouveau fichier fourni ici.
2. **Vous pouvez supprimer** les fichiers `Index`, `FormUsagers`, `FormPersonnels`, `FormCogecs`, `AdminDashboard` de l'éditeur Apps Script — ils ne servent plus (les pages vivent maintenant sur GitHub). Vous pouvez aussi les laisser, ça ne gêne pas.
3. **Déployer > Gérer les déploiements > ✏️ > Nouvelle version > Déployer.**
4. Copiez l'URL `.../exec` (la même qu'avant si vous ne créez pas un nouveau déploiement, sinon la nouvelle).

## Étape 2 — Configurer `config.js`

Ouvrez `config.js` et remplacez la ligne :

```js
const API_URL = "COLLEZ_ICI_VOTRE_URL_APPS_SCRIPT/exec";
```

par votre vraie URL, par exemple :

```js
const API_URL = "https://script.google.com/macros/s/AKfycbys.../exec";
```

C'est le **seul fichier à modifier** — tous les formulaires et le dashboard le partagent.

## Étape 3 — Créer le dépôt GitHub

1. Sur [github.com](https://github.com), cliquez **New repository**. Nom au choix (ex. `satisfaction-zones-sanitaires`). Laissez-le **Public** (nécessaire pour GitHub Pages gratuit, sauf si vous avez un plan payant).
2. Mettez-y ces 7 fichiers, à la racine du dépôt :
   - `Index.html`
   - `FormUsagers.html`
   - `FormPersonnels.html`
   - `FormCogecs.html`
   - `AdminDashboard.html`
   - `config.js`
   - (n'incluez PAS `Code.gs`, qui reste dans Apps Script uniquement)

   Le plus simple : bouton **Add file > Upload files** sur la page du dépôt, glissez les 6 fichiers, puis **Commit changes**.

## Étape 4 — Activer GitHub Pages

1. Dans le dépôt : **Settings > Pages**.
2. Sous "Build and deployment", **Source** : `Deploy from a branch`.
3. **Branch** : `main` (ou `master`), dossier `/ (root)` → **Save**.
4. Au bout d'une minute ou deux, GitHub affiche l'URL de votre site, du type :
   `https://VOTRE-NOM-UTILISATEUR.github.io/satisfaction-zones-sanitaires/`

## Étape 5 — Partager les liens

- Accueil : `https://VOTRE-NOM.github.io/VOTRE-DEPOT/`
- Usagers : `.../FormUsagers.html`
- Personnel soignant : `.../FormPersonnels.html`
- COGECS : `.../FormCogecs.html`
- Tableau de bord admin : `.../AdminDashboard.html`

## Mettre à jour un fichier plus tard

Sur GitHub, ouvrez le fichier concerné → icône ✏️ (crayon) → modifiez → **Commit changes**. Le site se met à jour automatiquement en une minute ou deux, sans rien redéployer côté Apps Script (sauf si vous modifiez `Code.gs`, qui lui doit toujours être redéployé depuis l'éditeur Apps Script).

## Pourquoi ce découpage était nécessaire

`google.script.run` (utilisé avant) ne fonctionne que si la page HTML est servie *par* Apps Script lui-même — impossible de l'utiliser depuis un site externe comme GitHub Pages. Les fichiers utilisent maintenant `fetch()` vers l'URL Apps Script, qui répond en JSON. C'est pour ça que `Code.gs` doit être redéployé une fois avec sa nouvelle version.

## En cas de souci

- **"API_URL non configurée" ou rien ne se passe à l'envoi** : vérifiez que `config.js` contient bien votre vraie URL `.../exec`, et que vous avez bien redéployé une **nouvelle version** de `Code.gs` (pas juste enregistré).
- **Le formulaire dit "sauvegardé localement" alors que vous avez du réseau** : l'API ne répond pas correctement — vérifiez l'URL dans `config.js`, et que le déploiement Apps Script est bien "Exécuter en tant que : Moi" / "Qui a accès : Tous les utilisateurs".
- Le reste (colonnes, formules, structure du Sheet) n'a pas changé — consultez `README_DEPLOIEMENT.md` d'origine pour ces points.
