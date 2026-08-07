/* ============================================================
 *  MOTEUR GÉNÉRIQUE DE FORMULAIRE — Enquêtes Hôpitaux
 *  Piloté par un schéma JSON (sections -> champs), avec :
 *  cascades de listes (dept -> zone -> fs...), logique de saut
 *  (relevant), file d'attente hors-ligne, GPS (si activé).
 * ============================================================ */
const HopitalEngine = (function(){

  let CFG = null;          // {formKey, submitAction, title, hasGeo, schema}
  let SECTIONS = [];       // schema.sections, avec un index "flat" des champs
  let currentSection = 0;
  let currentAnswers = {};
  let QUEUE_KEY = '';

  // ---- GPS (repris à l'identique du formulaire Usagers zones sanitaires) ----
  let gpsWatchId = null, gpsWaitTimer = null;
  let lastKnownPosition = null, gpsAttempting = false, gpsAttemptFailed = false, gpsSkipped = false;
  const GPS_TARGET_ACCURACY_M = 5;
  const GPS_MAX_WAIT_MS = 30000;

  function stopGpsWatch(){
    if (gpsWatchId !== null){ navigator.geolocation.clearWatch(gpsWatchId); gpsWatchId = null; }
    if (gpsWaitTimer !== null){ clearTimeout(gpsWaitTimer); gpsWaitTimer = null; }
  }
  function onGpsUpdate(pos){
    const acc = (pos.coords && typeof pos.coords.accuracy === 'number') ? pos.coords.accuracy : Infinity;
    const bestAcc = (lastKnownPosition && typeof lastKnownPosition.coords.accuracy === 'number') ? lastKnownPosition.coords.accuracy : Infinity;
    if (!lastKnownPosition || acc < bestAcc) lastKnownPosition = pos;
    if (acc <= GPS_TARGET_ACCURACY_M){ gpsAttempting = false; stopGpsWatch(); }
    renderSection();
  }
  function onGpsError(){
    if (!lastKnownPosition){ gpsAttempting = false; gpsAttemptFailed = true; stopGpsWatch(); renderSection(); }
  }
  function startBackgroundGeolocation(){
    if (!CFG.hasGeo || !navigator.geolocation) return;
    gpsAttempting = true;
    gpsWatchId = navigator.geolocation.watchPosition(onGpsUpdate, onGpsError, { enableHighAccuracy:true, timeout:GPS_MAX_WAIT_MS, maximumAge:0 });
    gpsWaitTimer = setTimeout(() => {
      gpsAttempting = false; stopGpsWatch();
      if (!lastKnownPosition) gpsAttemptFailed = true;
      renderSection();
    }, GPS_MAX_WAIT_MS);
  }
  function retryGps(){
    stopGpsWatch();
    gpsAttempting = true; gpsAttemptFailed = false; gpsSkipped = false;
    renderSection();
    gpsWatchId = navigator.geolocation.watchPosition(onGpsUpdate, onGpsError, { enableHighAccuracy:true, timeout:GPS_MAX_WAIT_MS, maximumAge:0 });
    gpsWaitTimer = setTimeout(() => { gpsAttempting = false; stopGpsWatch(); if (!lastKnownPosition) gpsAttemptFailed = true; renderSection(); }, GPS_MAX_WAIT_MS);
  }
  function skipGps(){ gpsSkipped = true; stopGpsWatch(); renderSection(); }

  // ---- Logique "relevant" (${champ} = 'valeur' / != 'valeur', combinables avec "or") ----
  function isRelevant(field){
    if (!field.relevant) return true;
    const clauses = field.relevant.split(/\s+or\s+/i);
    return clauses.some(clause => {
      const m = clause.match(/\$\{(\w+)\}\s*(=|!=)\s*'([^']*)'/);
      if (!m) return true; // repli sûr : on affiche si l'expression n'est pas reconnue
      const val = currentAnswers[m[1]];
      return m[2] === '=' ? (val === m[3]) : (val !== m[3]);
    });
  }

  function optionsFor(field){
    if (!field.filterField) return field.options;
    const parentVal = currentAnswers[field.filterField];
    if (!parentVal) return [];
    return field.options.filter(o => field.optionFilterVals[o.name] === parentVal);
  }

  function optionLabelFor(fieldName, rawValue){
    if (rawValue === undefined || rawValue === null || rawValue === '') return '';
    for (const sec of SECTIONS){
      const f = sec.fields.find(x => x.name === fieldName);
      if (f && f.options){
        const opt = f.options.find(o => o.name === rawValue);
        if (opt) return opt.label;
      }
    }
    return rawValue; // champ texte/nombre : on affiche la valeur telle quelle
  }
  function interpolate(text){
    if (!text) return text;
    return text.replace(/\$\{(\w+)\}/g, (m, name) => optionLabelFor(name, currentAnswers[name]) || '…');
  }

  // ---- Rendu ----
  function fieldsOfSection(idx){ return SECTIONS[idx].fields.filter(isRelevant); }

  function renderSection(){
    const root = document.getElementById('fieldsRoot');
    const sec = SECTIONS[currentSection];
    let html = `<h2 class="secTitle">${interpolate(sec.title) || ''}</h2>`;

    fieldsOfSection(currentSection).forEach(f => {
      if (f.type === 'note'){ html += `<div class="note">${interpolate(f.label)}</div>`; return; }
      html += `<div class="field" data-name="${f.name}">`;
      html += `<label>${interpolate(f.label)}${f.required ? ' <span class="req">*</span>' : ''}</label>`;
      if (f.hint) html += `<div class="hint">${interpolate(f.hint)}</div>`;

      if (f.type === 'select_one'){
        const opts = optionsFor(f);
        html += '<div class="opts">' + opts.map(o => {
          const checked = currentAnswers[f.name] === o.name ? 'checked' : '';
          return `<label class="opt"><input type="radio" name="${f.name}" value="${o.name}" ${checked} onchange="HopitalEngine.setAnswer('${f.name}', this.value)"> ${o.label}</label>`;
        }).join('') + '</div>';
        if (!opts.length) html += '<div class="hint">Choisissez d\'abord le niveau précédent.</div>';
      } else if (f.type === 'select_multiple'){
        const opts = optionsFor(f);
        const current = (currentAnswers[f.name] || '').split(' ').filter(Boolean);
        html += '<div class="opts">' + opts.map(o => {
          const checked = current.includes(o.name) ? 'checked' : '';
          return `<label class="opt"><input type="checkbox" value="${o.name}" ${checked} onchange="HopitalEngine.toggleMulti('${f.name}', '${o.name}', this.checked)"> ${o.label}</label>`;
        }).join('') + '</div>';
      } else if (f.type === 'integer'){
        html += `<input type="number" inputmode="numeric" value="${currentAnswers[f.name] || ''}" oninput="HopitalEngine.setAnswer('${f.name}', this.value)">`;
      } else if (f.type === 'geopoint'){
        html += renderGpsField();
      } else {
        html += `<textarea rows="3" oninput="HopitalEngine.setAnswer('${f.name}', this.value)">${currentAnswers[f.name] || ''}</textarea>`;
      }
      html += '</div>';
    });

    root.innerHTML = html;
    document.getElementById('progressText').textContent = `Section ${currentSection + 1} / ${SECTIONS.length}`;
    document.getElementById('progressBar').style.width = Math.round(((currentSection + 1) / SECTIONS.length) * 100) + '%';
    document.getElementById('btnPrev').style.display = currentSection === 0 ? 'none' : 'inline-block';
    const isLast = currentSection === SECTIONS.length - 1;
    document.getElementById('btnNext').style.display = isLast ? 'none' : 'inline-block';
    document.getElementById('btnSubmit').style.display = isLast ? 'inline-block' : 'none';
  }

  function renderGpsField(){
    if (gpsSkipped){
      return `<div class="gpsBox">⚠️ Position non enregistrée. <button type="button" onclick="HopitalEngine.retryGps()">Réessayer</button></div>`;
    }
    if (lastKnownPosition){
      const acc = Math.round(lastKnownPosition.coords.accuracy || 0);
      return `<div class="gpsBox gpsOk">📍 Position enregistrée (précision ~${acc} m)` +
        (gpsAttempting ? ' — amélioration en cours…' : '') +
        `<button type="button" onclick="HopitalEngine.retryGps()">Reprendre la position</button></div>`;
    }
    if (gpsAttempting){
      return `<div class="gpsBox">📡 Recherche de la position GPS en cours (jusqu'à 30 s)…</div>`;
    }
    if (gpsAttemptFailed){
      return `<div class="gpsBox gpsErr">❌ Position introuvable.
        <button type="button" onclick="HopitalEngine.retryGps()">Réessayer</button>
        <button type="button" onclick="HopitalEngine.skipGps()">Continuer sans position</button></div>`;
    }
    return `<div class="gpsBox"><button type="button" onclick="HopitalEngine.retryGps()">📍 Activer la localisation</button></div>`;
  }

  function setAnswer(name, value){
    currentAnswers[name] = value;
    // si un champ "parent" de cascade change, on efface les enfants qui en dépendent
    SECTIONS.forEach(sec => sec.fields.forEach(f => {
      if (f.filterField === name) delete currentAnswers[f.name];
    }));
  }
  function toggleMulti(name, optName, checked){
    const cur = new Set((currentAnswers[name] || '').split(' ').filter(Boolean));
    if (checked) cur.add(optName); else cur.delete(optName);
    currentAnswers[name] = Array.from(cur).join(' ');
  }

  function validateSection(idx){
    const missing = fieldsOfSection(idx).filter(f => f.required && f.type !== 'note' && f.type !== 'geopoint' && !currentAnswers[f.name]);
    if (missing.length){
      showToast('Merci de répondre à : ' + missing[0].label);
      return false;
    }
    return true;
  }

  function nextSection(){
    if (!validateSection(currentSection)) return;
    if (currentSection < SECTIONS.length - 1){ currentSection++; renderSection(); window.scrollTo({top:0, behavior:'smooth'}); }
  }
  function prevSection(){
    if (currentSection > 0){ currentSection--; renderSection(); window.scrollTo({top:0, behavior:'smooth'}); }
  }

  // ---- Soumission + file d'attente hors-ligne (repris du formulaire Usagers) ----
  function buildPayload(){
    const data = Object.assign({}, currentAnswers);
    if (CFG.hasGeo){
      data['today'] = new Date().toISOString().slice(0,10);
      if (lastKnownPosition){
        data['_start-geopoint_latitude'] = lastKnownPosition.coords.latitude;
        data['_start-geopoint_longitude'] = lastKnownPosition.coords.longitude;
        data['_start-geopoint_altitude'] = lastKnownPosition.coords.altitude || '';
        data['_start-geopoint_precision'] = lastKnownPosition.coords.accuracy || '';
      }
    }
    return data;
  }

  function submitForm(){
    if (!validateSection(currentSection)) return;
    const btn = document.getElementById('btnSubmit');
    btn.disabled = true;
    const payload = buildPayload();
    sendOrQueue(payload, () => { btn.disabled = false; });
  }

  function sendOrQueue(payload, done){
    if (!navigator.onLine){
      queueSubmission(payload);
      showToast('Hors connexion : réponse sauvegardée localement.');
      resetForm(); done();
      return;
    }
    callApi(CFG.submitAction, payload)
      .then(res => {
        if (res && res.ok){ showToast('✅ Réponse enregistrée avec succès.'); resetForm(); done(); updateQueueBadge(); }
        else { queueSubmission(payload); showToast('Connexion instable : réponse sauvegardée localement.'); resetForm(); done(); }
      })
      .catch(() => { queueSubmission(payload); showToast('Connexion instable : réponse sauvegardée localement.'); resetForm(); done(); });
  }
  function queueSubmission(payload){
    const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
    q.push(payload);
    localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
    updateQueueBadge();
  }
  function updateQueueBadge(){
    const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
    const el = document.getElementById('qCount');
    if (q.length > 0){ el.style.display='inline-block'; el.textContent = q.length; } else { el.style.display='none'; }
    const banner = document.getElementById('status-banner');
    if (!navigator.onLine){ banner.className='offline'; banner.textContent='⚠️ Hors connexion — les réponses seront sauvegardées localement.'; }
    else if (q.length > 0){ banner.className='queued'; banner.textContent = q.length + ' réponse(s) en attente de synchronisation.'; }
    else { banner.className=''; banner.textContent=''; }
  }
  function syncQueue(manual){
    const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
    if (q.length === 0){ if(manual) showToast('Aucune réponse en attente.'); return; }
    if (!navigator.onLine){ if(manual) showToast('Toujours hors connexion.'); return; }
    let remaining = q.slice();
    function next(){
      if (remaining.length === 0){ localStorage.setItem(QUEUE_KEY, JSON.stringify(remaining)); updateQueueBadge(); if(manual) showToast('✅ Synchronisation terminée.'); return; }
      const item = remaining[0];
      callApi(CFG.submitAction, item)
        .then(res => { if (res && res.ok){ remaining.shift(); localStorage.setItem(QUEUE_KEY, JSON.stringify(remaining)); updateQueueBadge(); next(); } else { localStorage.setItem(QUEUE_KEY, JSON.stringify(remaining)); updateQueueBadge(); if(manual) showToast('Synchronisation interrompue (connexion).'); } })
        .catch(() => { localStorage.setItem(QUEUE_KEY, JSON.stringify(remaining)); updateQueueBadge(); if(manual) showToast('Synchronisation interrompue (connexion).'); });
    }
    next();
  }
  function resetForm(){
    currentAnswers = {};
    currentSection = 0;
    lastKnownPosition = null; gpsAttemptFailed = false; gpsSkipped = false;
    renderSection();
    window.scrollTo({top:0, behavior:'smooth'});
  }

  function showToast(msg){
    const t = document.getElementById('toast');
    t.textContent = msg; t.style.display = 'block';
    setTimeout(() => { t.style.display = 'none'; }, 3000);
  }

  // ---- Contrôle d'accès local (grant, cf. index.html / enqueteur.html) ----
  const GRANT_KEY = 'sz_grant';
  const INACTIVITY_LIMIT_MS = 6 * 60 * 60 * 1000;
  function loadGrant(){ try { return JSON.parse(localStorage.getItem(GRANT_KEY)); } catch(e){ return null; } }
  function saveGrant(g){ localStorage.setItem(GRANT_KEY, JSON.stringify(g)); }
  function checkGrant(requiredRole){
    if (!requiredRole) return true; // formulaire sans contrôle d'accès particulier
    const g = loadGrant();
    if (!g || g.role !== requiredRole || (Date.now() - g.lastActivity) > INACTIVITY_LIMIT_MS) return false;
    g.lastActivity = Date.now(); saveGrant(g);
    return true;
  }

  function init(config){
    CFG = config;
    SECTIONS = config.schema.sections;
    QUEUE_KEY = 'queue_' + config.formKey;

    if (config.requiredRole && !checkGrant(config.requiredRole)){
      document.getElementById('fieldsRoot').innerHTML =
        '<div class="note" style="text-align:center"><h2>⚠️ Accès non reconnu</h2>' +
        '<p>Scannez le QR code correspondant depuis l\'accueil pour accéder à ce formulaire.</p>' +
        '<p><a href="hopital.html">← Retour</a></p></div>';
      document.querySelectorAll('.navBar, .status-banner').forEach(el => el.style.display = 'none');
      return;
    }

    document.getElementById('formTitle').textContent = config.title;
    startBackgroundGeolocation();
    renderSection();
    updateQueueBadge();
    syncQueue(false);
    window.addEventListener('online', () => { updateQueueBadge(); syncQueue(false); });
    window.addEventListener('offline', updateQueueBadge);
    setInterval(() => { const g = loadGrant(); if (g){ g.lastActivity = Date.now(); saveGrant(g); } }, 5*60*1000);
  }

  return { init, setAnswer, toggleMulti, nextSection, prevSection, submitForm, syncQueue, retryGps, skipGps };
})();
