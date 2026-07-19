// ============================================================
//  CONFIGURATION — Satisfaction Zones Sanitaires du Bénin
// ============================================================
//  Un seul endroit à modifier : collez ici l'URL de votre
//  application Web Apps Script (Déployer > Gérer les déploiements),
//  qui se termine par /exec
// ============================================================
const API_URL = "https://script.google.com/macros/s/AKfycbz1aBzOmKxev5s7Plpj5k_OyxkUSxQEXvfQRYNLvxtcZ_3EIiZiibAdcPqaO6tqaps2Lw/exec";

/**
 * Appelle l'API Apps Script en JSON.
 * On utilise Content-Type: text/plain (et non application/json) pour éviter
 * la requête "preflight" OPTIONS que les applications Web Apps Script ne
 * gèrent pas — c'est la méthode standard pour appeler Apps Script depuis
 * un site externe (GitHub Pages, etc.).
 */
function callApi(action, data, password) {
  if (!API_URL || API_URL.indexOf("COLLEZ_ICI") !== -1) {
    return Promise.reject(new Error("API_URL non configurée : modifiez config.js"));
  }
  return fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: action, data: data, password: password })
  }).then(function (r) {
    if (!r.ok) throw new Error("Réponse HTTP " + r.status);
    return r.json();
  });
}
