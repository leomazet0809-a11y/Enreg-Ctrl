/* =======================================================================
   Travail à plusieurs sur un même classeur (dossier réseau partagé)

   1. Fusion : avant chaque écriture (et au retour dans l'application), le
      classeur est relu ; ce que les autres postes y ont enregistré depuis
      notre dernière synchronisation est fusionné avec nos données
      (fusion à trois versions : ancêtre commun, nos données, le fichier).
      Granularité : un rapport, une famille de gamme, un article (TS,
      matrice), une révision, un document. Si un même élément a été
      modifié des deux côtés, notre version l'emporte (pour un rapport :
      la plus récente).
   2. Verrou : avec l'accès au dossier du classeur, un petit fichier
      « ~verrou-<classeur>.json » signale qu'un poste est en train
      d'écrire. Les autres attendent qu'il disparaisse, ainsi que le
      fichier « ~$<classeur> » créé par Excel tant que le classeur y est
      ouvert, puis enregistrent automatiquement.
   ======================================================================= */

var VERROU_PERIME = 60000;   // un verrou plus ancien est considéré comme abandonné (poste planté)

/* ---------- photo des données partagées ---------- */
function estClePartagee(k){ return /^rc\./.test(k) && k !== 'rc.draft'; }
function instantane(mem){
  var o = {};
  Object.keys(mem || {}).forEach(function(k){ if(estClePartagee(k) && mem[k] !== null && mem[k] !== undefined) o[k] = mem[k]; });
  return o;
}
function memesDonnees(a, b){
  var ka = Object.keys(a).sort(), kb = Object.keys(b).sort();
  if(ka.join('|') !== kb.join('|')) return false;
  return ka.every(function(k){ return sansHorodatage(a[k]) === sansHorodatage(b[k]); });
}

/* ---------- fusion à trois versions ---------- */
function signature(v){ return v === undefined ? '∅' : JSON.stringify(canon(v)); }
function indexer(arr, cle, groupe){
  var ix = {ordre:[], m:Object.create(null)};
  (arr || []).forEach(function(x){
    var k = String(typeof cle === 'function' ? cle(x) : x[cle]);
    if(!(k in ix.m)){ ix.ordre.push(k); ix.m[k] = groupe ? [] : undefined; }
    if(groupe) ix.m[k].push(x); else ix.m[k] = x;
  });
  return ix;
}
function indexerObjet(o){
  var ix = {ordre:Object.keys(o || {}), m:Object.create(null)};
  ix.ordre.forEach(function(k){ ix.m[k] = o[k]; });
  return ix;
}
/* a = ancêtre, o = nos données, t = le fichier ; cpt.n compte les conflits */
function fusion3(a, o, t, cpt, departage){
  var vus = Object.create(null), ordre = [];
  /* ordre du fichier d'abord : sans modification locale, le résultat est identique au fichier */
  t.ordre.concat(o.ordre).forEach(function(k){ if(!vus[k]){ vus[k] = 1; ordre.push(k); } });
  var out = {ordre:[], m:Object.create(null)};
  ordre.forEach(function(k){
    var va = a.m[k], vo = o.m[k], vt = t.m[k];
    var sa = signature(va), so = signature(vo), st = signature(vt), v;
    if(so === st) v = vo;
    else if(so === sa) v = vt;          // seul le fichier a changé
    else if(st === sa) v = vo;          // seuls nous avons changé
    else { cpt.n++; v = departage ? departage(vo, vt) : vo; }
    if(v !== undefined){ out.ordre.push(k); out.m[k] = v; }
  });
  return out;
}
function enTableau(ix, groupe){
  var r = [];
  ix.ordre.forEach(function(k){ if(groupe) r.push.apply(r, ix.m[k]); else r.push(ix.m[k]); });
  return r;
}
function enObjet(ix){ var r = {}; ix.ordre.forEach(function(k){ r[k] = ix.m[k]; }); return r; }

function fusionnerDonnees(anc, ours, theirs){
  var cpt = {n:0};
  function J(m, k, d){ try{ return m && m[k] ? JSON.parse(m[k]) : d; }catch(e){ return d; } }
  function trois(k, d){ return [J(anc, k, d), J(ours, k, d), J(theirs, k, d)]; }
  var res = {};

  var db = trois('rc.db', {});
  function partie(nom, cle, groupe){
    return enTableau(fusion3(indexer(db[0][nom], cle, groupe), indexer(db[1][nom], cle, groupe), indexer(db[2][nom], cle, groupe), cpt), groupe);
  }
  res['rc.db'] = JSON.stringify({gamserie:partie('gamserie', 'famille', true), ts:partie('ts', 'article'), matrice:partie('matrice', 'article')});

  var ve = trois('rc.versions', {gam:{}, ts:{}});
  res['rc.versions'] = JSON.stringify({
    gam:enObjet(fusion3(indexerObjet(ve[0].gam), indexerObjet(ve[1].gam), indexerObjet(ve[2].gam), cpt)),
    ts:enObjet(fusion3(indexerObjet(ve[0].ts), indexerObjet(ve[1].ts), indexerObjet(ve[2].ts), cpt))
  });

  var cleRev = function(x){ return x.type + '|' + x.cle + '|' + x.rev; };
  var hi = trois('rc.revisions', []);
  res['rc.revisions'] = JSON.stringify(enTableau(fusion3(indexer(hi[0], cleRev), indexer(hi[1], cleRev), indexer(hi[2], cleRev), cpt)));

  var dc = trois('rc.docs', []);
  res['rc.docs'] = JSON.stringify(enTableau(fusion3(indexer(dc[0], 'id'), indexer(dc[1], 'id'), indexer(dc[2], 'id'), cpt)));

  /* rapports : un même rapport modifié des deux côtés -> la version la plus récente */
  var rp = trois('rc.reports', []);
  var plusRecent = function(o, t){
    if(o === undefined) return t; if(t === undefined) return o;
    return String(t.majLe || '') > String(o.majLe || '') ? t : o;
  };
  res['rc.reports'] = JSON.stringify(enTableau(fusion3(indexer(rp[0], 'id'), indexer(rp[1], 'id'), indexer(rp[2], 'id'), cpt, plusRecent)));

  /* journal : réunion des évènements des deux côtés */
  var jo = trois('rc.journal', []), vus = Object.create(null), jr = [];
  jo[2].concat(jo[1]).forEach(function(e){
    var k = [e.le, e.type, e.cle, e.rev, e.action, e.detail].join('|');
    if(!vus[k]){ vus[k] = 1; jr.push(e); }
  });
  jr.sort(function(x, y){ return String(x.le) < String(y.le) ? -1 : String(x.le) > String(y.le) ? 1 : 0; });
  res['rc.journal'] = JSON.stringify(jr.slice(-2000));

  /* rapports en cours : celui de ce poste vient de nous, ceux des autres postes du fichier */
  Object.keys(theirs).forEach(function(k){ if(/^rc\.draft\./.test(k) && k !== cleBrouillon()) res[k] = theirs[k]; });
  if(ours[cleBrouillon()] !== undefined) res[cleBrouillon()] = ours[cleBrouillon()];

  /* ne pas créer de clé absente des deux côtés */
  ['rc.db','rc.versions','rc.revisions','rc.docs','rc.reports','rc.journal'].forEach(function(k){
    if(ours[k] === undefined && theirs[k] === undefined) delete res[k];
  });
  return {mem:res, conflits:cpt.n};
}

/* Fusionne le contenu du fichier (lu par lireClasseur) avec nos données.
   Renvoie {recu:true si des données d'un autre poste ou d'Excel sont arrivées, bilan, conflits}. */
function synchroniserAvec(lu){
  var ours = instantane(MEM), fichier = instantane(lu.mem);
  var notreR = R, bilan;
  CHARGEMENT = true;
  var theirs, fus;
  try{
    /* 1. état du fichier, avec les modifications faites dans Excel s'il y en a */
    MEM = Object.assign({}, lu.mem);
    loadAll(); initVersions();
    bilan = integrerFeuillesExcel(lu.wb, lu.neuf);
    theirs = instantane(MEM);
    /* 2. fusion avec nos données */
    fus = fusionnerDonnees(CLASSEUR.ancetre || {}, ours, theirs);
    MEM = fus.mem;
    loadAll();
    R = notreR;
    blobCache = {};
  }finally{
    CHARGEMENT = false;
  }
  CLASSEUR.autres = feuillesUtilisateur(lu.wb);
  CLASSEUR.ancetre = theirs;
  CLASSEUR.modifie = !memesDonnees(instantane(MEM), fichier);
  return {recu:!memesDonnees(instantane(MEM), ours), bilan:bilan, conflits:fus.conflits};
}

/* ---------- accès au dossier du classeur ---------- */
async function contientFichier(d, h){
  try{ return !!(await d.resolve(h)); }catch(e){ return false; }
}
/* Dossier du classeur courant, avec autorisation (demander = true : peut afficher la demande du navigateur) */
async function dossierPret(demander){
  var d = CLASSEUR.dossier;
  if(!d) return null;
  if(!(await autorisation(d, demander))) return null;
  return d;
}
/* Retrouve le dossier mémorisé pour ce classeur (sans fenêtre de choix) */
async function restaurerDossier(h, demander){
  var d = await idbLire('dossier');
  if(!d || typeof d.resolve !== 'function') return false;
  if(!(await autorisation(d, demander))) return false;
  if(!(await contientFichier(d, h))) return false;
  CLASSEUR.dossier = d;
  return true;
}
async function choisirDossier(){
  var h = CLASSEUR.handle;
  if(!h){ toast('Choisissez d\'abord l\'emplacement du classeur (Enregistrer).'); return false; }
  if(await restaurerDossier(h, true)){ toast('Accès au dossier du classeur autorisé.'); render(); return true; }
  if(typeof window.showDirectoryPicker !== 'function'){ toast('Ce navigateur ne permet pas l\'accès au dossier.'); return false; }
  var d;
  try{ d = await window.showDirectoryPicker({id:'enreg-ctrl-dossier', mode:'readwrite', startIn:h}); }
  catch(e){ if(!e || e.name !== 'AbortError') toast('Accès au dossier impossible : ' + (e.message || e.name)); return false; }
  if(!(await contientFichier(d, h))){
    modal('Mauvais dossier', 'Le dossier choisi ne contient pas le classeur <b>' + esc(h.name) + '</b>. Choisissez le dossier où se trouve ce fichier.');
    return false;
  }
  CLASSEUR.dossier = d;
  await idbEcrire('dossier', d);
  await idbEcrire('seul:' + h.name, false);
  toast('Accès au dossier autorisé : le travail à plusieurs est activé.');
  if(VIEW === 'parametres') render();
  return true;
}
/* Proposition unique (par classeur) après ouverture ou création */
async function proposerDossier(){
  var h = CLASSEUR.handle;
  if(!h || CLASSEUR.dossier || typeof window.showDirectoryPicker !== 'function') return;
  if(await idbLire('seul:' + h.name)) return;
  while(document.querySelector('.modal')) await pause(500);   // après le message éventuel déjà affiché
  if(CLASSEUR.handle !== h || CLASSEUR.dossier) return;
  modal('Travail à plusieurs',
    'Si ce classeur est utilisé <b>par plusieurs personnes</b> (dossier réseau partagé), autorisez l\'accès à son dossier : '
    + 'l\'application pourra signaler qu\'un autre utilisateur est en train d\'enregistrer, ou que le classeur est ouvert dans Excel, '
    + 'et attendre qu\'il soit disponible avant d\'enregistrer.'
    + '<p style="margin:10px 0 0">Le navigateur vous demandera de choisir le dossier qui contient <b>' + esc(h.name) + '</b>.</p>',
    function(){}, 'Fermer');
  var ft = document.querySelector('.modal .ft');
  ft.innerHTML = '<button class="btn" data-seul>Je travaille seul</button><button class="btn pri" data-dossier>Autoriser l\'accès au dossier</button>';
  ft.querySelector('[data-seul]').onclick = function(){ document.querySelector('.modal').remove(); idbEcrire('seul:' + h.name, true); };
  ft.querySelector('[data-dossier]').onclick = function(){ document.querySelector('.modal').remove(); choisirDossier(); };
}

/* ---------- verrou et attente ---------- */
function pause(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
function nomVerrou(h){ return '~verrou-' + h.name + '.json'; }
async function lireVerrou(d, nom){
  try{ return JSON.parse(await (await (await d.getFileHandle(nom)).getFile()).text()); }catch(e){ return null; }
}
async function ecrireVerrou(d, nom){
  var w = await (await d.getFileHandle(nom, {create:true})).createWritable();
  await w.write(JSON.stringify({poste:POSTE, depuis:new Date().toISOString()}));
  await w.close();
}
async function supprimerVerrou(d, nom){
  try{ var v = await lireVerrou(d, nom); if(!v || v.poste === POSTE) await d.removeEntry(nom); }catch(e){}
}
async function existe(d, nom){
  try{ await d.getFileHandle(nom); return true; }catch(e){ return false; }
}
/* Fichier « propriétaire » créé par Excel tant que le classeur y est ouvert (nom éventuellement raccourci) */
async function ouvertDansExcel(d, h){
  var noms = ['~$' + h.name, '~$' + h.name.slice(1), '~$' + h.name.slice(2)];
  for(var i=0;i<noms.length;i++) if(await existe(d, noms[i])) return true;
  return false;
}
function verrouActifAutrePoste(v){
  return !!(v && v.poste !== POSTE && (Date.now() - Date.parse(v.depuis || 0)) < VERROU_PERIME);
}

var ATTENTE = {annule:false, forcerExcel:false};
function afficherAttente(motif){
  var el = document.getElementById('attente');
  if(!el){
    el = document.createElement('div');
    el.id = 'attente'; el.className = 'attente'; el.setAttribute('role', 'status');
    document.body.appendChild(el);
  }
  var txt = (motif === 'excel' || motif === 'bloque')
    ? '<b>Classeur ouvert dans Excel</b><span>Votre enregistrement partira dès que le classeur sera fermé.</span>'
    : '<b>Classeur en cours d\'enregistrement par un autre utilisateur</b><span>Votre enregistrement partira dès que le classeur sera disponible.</span>';
  if(el.getAttribute('data-motif') !== motif){
    el.setAttribute('data-motif', motif);
    el.innerHTML = '<i class="roue" aria-hidden="true"></i><div class="att-txt">' + txt + '</div>'
      + (motif === 'excel' ? '<button class="btn sm" data-forcer title="Si personne n\'a le classeur ouvert dans Excel">Enregistrer quand même</button>' : '')
      + '<button class="btn sm" data-annuler>Annuler</button>';
    var f = el.querySelector('[data-forcer]'); if(f) f.onclick = function(){ ATTENTE.forcerExcel = true; };
    el.querySelector('[data-annuler]').onclick = function(){ ATTENTE.annule = true; };
  }
  el.hidden = false;
}
function masquerAttente(){ var el = document.getElementById('attente'); if(el){ el.hidden = true; el.removeAttribute('data-motif'); } }
function erreurAnnulee(){ var e = new Error('Enregistrement annulé'); e.annule = true; return e; }

/* Attend que le classeur soit libre, puis pose notre verrou.
   Renvoie le dossier verrouillé, ou null sans accès au dossier. */
async function prendreVerrou(h){
  var d = await dossierPret(true);
  if(!d) return null;
  var nom = nomVerrou(h);
  for(;;){
    if(ATTENTE.annule) throw erreurAnnulee();
    var autre = verrouActifAutrePoste(await lireVerrou(d, nom));
    var excel = !autre && !ATTENTE.forcerExcel && await ouvertDansExcel(d, h);
    if(!autre && !excel){
      await ecrireVerrou(d, nom);
      await pause(350);                                        // laisse un éventuel concurrent écrire le sien
      var v = await lireVerrou(d, nom);
      if(v && v.poste === POSTE) return d;
      continue;                                                // un autre poste a pris la main : on attend
    }
    afficherAttente(autre ? 'autre' : 'excel');
    await pause(1500);
  }
}
/* Pour une simple lecture : attend la fin de l'écriture d'un autre poste */
async function attendreFinEcriture(h){
  var d = await dossierPret(false);
  if(!d) return;
  var nom = nomVerrou(h);
  while(verrouActifAutrePoste(await lireVerrou(d, nom))){
    if(ATTENTE.annule) throw erreurAnnulee();
    afficherAttente('autre');
    await pause(1500);
  }
}
