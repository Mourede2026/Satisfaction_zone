# Déploiement — Satisfaction Zones Sanitaires du Bénin

## Ce qui change dans cette version

- **L'accueil (`index.html`) n'affiche plus aucun formulaire.** Deux façons d'entrer :
  - **Identifiant + mot de passe** → ouvre directement `admin.html` selon le niveau du compte.
  - **Scanner de QR code intégré** (caméra activée dans la page, pas besoin de l'appli
    caméra du téléphone) → scanner le QR « Enquêteur » ouvre `enqueteur.html`
    (Usagers + COGECS), scanner le QR « Personnel » ouvre `personnels.html`.
- **Une fois l'accès accordé par un scan, plus besoin de rescanner** : il reste valable
  tant que l'appli est réutilisée dans les 6 heures (fenêtre glissante, se prolonge à
  chaque utilisation). Au-delà de 6 h d'inactivité, un nouveau scan est demandé.
- **Gestion des comptes** (onglet 🔐 Comptes & QR de `admin.html`, national/administrateur
  uniquement) : activer/désactiver, modifier (niveau, portée, libellé) ou supprimer
  n'importe lequel des 122 comptes.
- **Enregistrement des questionnaires optimisé** : l'ancienne recherche de "première
  ligne vide" relisait toute la colonne Commune à chaque envoi ; elle est remplacée par
  une lecture directe de la dernière ligne (`getLastRow()`), beaucoup plus rapide,
  surtout à mesure que les feuilles grossissent.
- La capture GPS était déjà bien réglée (position réelle, non mise en cache, jusqu'à
  30 secondes d'attente pour une précision ≤5 m) — je n'y ai pas touché.

## Fichiers du dépôt GitHub Pages (14 fichiers, à la racine)

| Fichier | Rôle |
|---|---|
| `index.html` | Accueil : connexion identifiants + scanner QR caméra |
| `enqueteur.html` | Choix Usagers / COGECS (après scan du QR Enquêteur) |
| `usagers.html`, `cogecs.html` | Formulaires (accès conditionné au scan Enquêteur) |
| `personnels.html` | Formulaire Personnel (accès conditionné au scan Personnel) |
| `admin.html` | Tableau de bord + onglet Comptes & QR |
| `config.js`, `manifest.json`, icônes | Inchangés |

## Fichier Apps Script

`Code.gs` — remplace entièrement votre fichier actuel.

## Fichier à distribuer séparément (jamais sur GitHub)

`Comptes_acces_dashboard.csv` — les 122 identifiants/mots de passe en clair.

## Étape 1 — Apps Script

1. **Extensions > Apps Script**, remplacez tout `Code.gs` par le nouveau fichier.
2. Menu déroulant des fonctions > **setupComptesEtQrTokens** > ▶ **Exécuter** (une fois).
   Crée `Comptes` (122 lignes, avec la colonne **Actif**), `Sessions`, `QRTokens`.
   *(Si vous aviez déjà exécuté une version précédente de cette fonction sans la colonne
   Actif, exécutez plutôt **migrationAjouterColonneActif** pour l'ajouter sans tout
   recréer.)*
3. **Déployer > Gérer les déploiements > ✏️ > Nouvelle version > Déployer.**

Le mot de passe partagé (`ADMIN_PASSWORD`) n'est plus nécessaire pour se connecter au
tableau de bord (la connexion se fait avec un compte personnel), mais reste accepté si
vous l'utilisez ailleurs.

## Étape 2 — GitHub

Uploadez les 8 fichiers HTML/JS du premier tableau (remplacent les fichiers existants).

## Étape 3 — Comptes et QR codes

1. Distribuez `Comptes_acces_dashboard.csv`. Chacun peut changer son mot de passe depuis
   `admin.html`.
2. Connectez-vous une première fois en tant qu'**administrateur** ou **national** pour
   accéder à l'onglet 🔐 Comptes & QR :
   - Les 2 QR (Enquêteur, Personnel) s'affichent et se téléchargent en PNG — à imprimer
     ou afficher à l'écran pour que les enquêteurs/personnels les scannent.
   - **Actualiser** génère un nouveau QR et invalide l'ancien (un scan ultérieur de
     l'ancienne image échoue ; les accès déjà accordés avant l'actualisation restent
     valables jusqu'à leurs 6 h d'inactivité).
   - Le tableau des comptes permet d'activer/désactiver (⛔/✅), modifier (✏️) ou
     supprimer (🗑️) chaque compte.

## Portée de chaque niveau de compte (filtrage réellement côté serveur)

| Niveau | Voit dans le tableau de bord |
|---|---|
| Commune | Uniquement sa commune |
| Zone sanitaire | Toutes les communes de sa zone |
| Département | Toutes les communes de son département |
| National / Administrateur | Tout, sans restriction |

## Comment fonctionne l'accès par QR (technique, pour comprendre le comportement)

1. Le scan (caméra intégrée à `index.html`, ou lien ouvert par l'appareil photo natif du
   téléphone) envoie le rôle et le jeton au serveur pour vérification.
2. Si le jeton correspond au jeton courant, un accès local est enregistré dans le
   téléphone (`localStorage`), horodaté.
3. Chaque utilisation (chargement de page, toutes les 5 minutes en arrière-plan)
   rafraîchit cet horodatage → fenêtre glissante de 6 h, pas de nouveau scan tant que
   l'appli est utilisée.
4. Un lien « 🔄 Scanner un autre code » (page Enquêteur) permet de changer de rôle
   manuellement sur un même appareil.

## ⚠️ Limite assumée

Le contrôle d'accès par QR reste **local à l'appareil** : il ne revérifie pas le serveur
en continu pendant les 6 h (pour fonctionner hors connexion). Une actualisation de QR par
l'administrateur bloque donc les **nouveaux** scans immédiatement, mais ne coupe pas
un accès déjà en cours avant l'expiration de ses 6 h d'inactivité.

## En cas de souci

- **La caméra ne s'active pas** : nécessite HTTPS (GitHub Pages l'est par défaut) et
  l'autorisation caméra du navigateur/téléphone.
- **« Identifiant inconnu » / « Compte désactivé »** : vérifiez l'onglet `Comptes` (colonne
  Actif) ou réactivez le compte depuis l'onglet 🔐 Comptes & QR.
- **QR vide** : vérifiez que la nouvelle version du déploiement Apps Script est active.
- Pour le reste (404, écran d'accueil mobile), rien n'a changé par rapport à avant.
