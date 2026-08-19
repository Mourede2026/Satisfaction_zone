// ============================================================
//  CONFIGURATION — Satisfaction Zones Sanitaires du Bénin
// ============================================================
//  Un seul endroit à modifier : collez ici l'URL de votre
//  application Web Apps Script (Déployer > Gérer les déploiements),
//  qui se termine par /exec
// ============================================================
const API_URL = "https://script.google.com/macros/s/AKfycbz82maFufmFtOk8W0wKeJXp8Krbh_dDK7sOLKDkc1YbOZED7YaLGr9owx5XqogDRv8UDg/exec";

/**
 * Appelle l'API Apps Script en JSON, via XMLHttpRequest (et non fetch).
 *
 * Pourquoi XHR et pas fetch() : les Web Apps Apps Script répondent à toute
 * requête /exec par une redirection HTTP 302 vers
 * script.googleusercontent.com/macros/echo?... . Or la spécification fetch()
 * impose que, sur un 302 reçu en réponse à un POST, le navigateur rejoue la
 * redirection en GET en supprimant le corps de la requête — la requête qui
 * arrive réellement au serveur n'est donc plus un POST accountLogin/etc.
 * mais un GET, qui déclenche doGet() au lieu de doPost() (et dont la réponse
 * contient elle aussi un champ "ok":true, ce qui masque totalement l'échec
 * côté interface : pas de message d'erreur, juste un retour silencieux à
 * l'écran de connexion). XMLHttpRequest ne subit pas cette conversion
 * POST -> GET sur redirection et est la solution standard pour appeler une
 * Web App Apps Script en POST depuis un site externe (GitHub Pages, etc.).
 *
 * On utilise Content-Type: text/plain (et non application/json) pour éviter
 * en plus la requête "preflight" OPTIONS que les Web Apps Apps Script ne
 * gèrent pas.
 */
function callApi(action, data, password) {
  if (!API_URL || API_URL.indexOf("COLLEZ_ICI") !== -1) {
    return Promise.reject(new Error("API_URL non configurée : modifiez config.js"));
  }
  return new Promise(function (resolve, reject) {
    var xhr = new XMLHttpRequest();
    xhr.open("POST", API_URL, true);
    xhr.setRequestHeader("Content-Type", "text/plain;charset=utf-8");
    xhr.onload = function () {
      if (xhr.status < 200 || xhr.status >= 300) {
        reject(new Error("Réponse HTTP " + xhr.status));
        return;
      }
      try {
        resolve(JSON.parse(xhr.responseText));
      } catch (e) {
        reject(new Error("Réponse illisible du serveur (pas du JSON)."));
      }
    };
    xhr.onerror = function () {
      reject(new Error("Échec réseau (connexion au serveur impossible)."));
    };
    xhr.send(JSON.stringify({ action: action, data: data, password: password }));
  });
}
