/* =======================================================================
   Enregistrement du projet dans un classeur Excel (.xlsx)

   - L'utilisateur choisit l'emplacement du classeur (Chrome / Edge :
     File System Access API). Le lien vers ce fichier est mémorisé pour
     le proposer au prochain lancement.
   - Enregistrer : bouton en haut de page ou Ctrl + S
     (Ctrl + Maj + S : enregistrer sous un autre emplacement).
   - Le classeur contient des onglets lisibles (Rapports, Mesures,
     GAMSERIE, TS, Matrice, Journal), régénérés à chaque enregistrement,
     et un onglet masqué « _Donnees » qui fait foi au rechargement.
   ======================================================================= */

var CLASSEUR = {handle:null, nom:'', enregistreLe:'', modifie:false, ecritureEnCours:false,
                dateFichier:0,   // date de modification du fichier lors du dernier chargement / enregistrement
                autres:[]};      // onglets ajoutés par l'utilisateur, conservés à l'enregistrement
/* Onglets modifiables dans Excel et repris par l'application (ordre d'intégration) */
var FEUILLES_EDITABLES = [
  {base:'matrice',  feuille:'Matrice',  tableau:function(){ return tableauMatrice(); }},
  {base:'gamserie', feuille:'GAMSERIE', tableau:function(){ return tableauGamserie(); }},
  {base:'ts',       feuille:'TS',       tableau:function(){ return tableauTS(); }}
];
var FEUILLES_APPLI = ['Rapports','Mesures','GAMSERIE','TS','Matrice','Journal','Mode d\'emploi', '_Donnees'];
var CHARGEMENT = false;   // pendant le chargement, les écritures ne comptent pas comme modifications
var CLES_DONNEES = ['rc.db','rc.versions','rc.revisions','rc.journal','rc.docs','rc.reports','rc.draft'];
var FEUILLE_DONNEES = '_Donnees';
var TAILLE_PART = 30000;  // une cellule Excel accepte au plus 32 767 caractères
var NOM_DEFAUT = 'Enregistrements-Controles.xlsx';
var FS_OK = typeof window.showSaveFilePicker === 'function' && typeof window.showOpenFilePicker === 'function';
var TYPES_XLSX = [{description:'Classeur Excel', accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}];

var ICONES = {
  save:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M7 3v5h8V3M7 21v-7h10v7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
  excel:'<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="3" fill="currentColor" opacity=".14"/><path d="M8.5 8l7 8M15.5 8l-7 8" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>',
  sun:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><circle cx="12" cy="12" r="4.2" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 2.5v2.2M12 19.3v2.2M2.5 12h2.2M19.3 12h2.2M5.3 5.3l1.6 1.6M17.1 17.1l1.6 1.6M5.3 18.7l1.6-1.6M17.1 6.9l1.6-1.6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
  moon:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
  open:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
  plus:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
  reprendre:'<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path d="M4 12a8 8 0 1 0 2.4-5.7M4 4v4.5h4.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>'
};

/* ---------- état « modifié » ---------- */
function marquerModifie(){
  if(CHARGEMENT || CLASSEUR.modifie) return;
  CLASSEUR.modifie = true;
  majBarreFichier();
}
function estModifie(){ return CLASSEUR.modifie; }
function infoClasseur(){ return {nom:CLASSEUR.nom, enregistreLe:CLASSEUR.enregistreLe}; }
function heureCourante(){ var d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }

function majBarreFichier(){
  var el = document.getElementById('fichier');
  if(el){
    var etat, cls;
    if(CLASSEUR.ecritureEnCours){ etat = 'Enregistrement…'; cls = 'en-cours'; }
    else if(CLASSEUR.modifie){ etat = 'Modifications non enregistrées'; cls = 'mod'; }
    else if(CLASSEUR.nom){ etat = CLASSEUR.enregistreLe ? 'Enregistré à ' + CLASSEUR.enregistreLe : 'À jour'; cls = 'ok'; }
    else { etat = 'Aucun classeur choisi'; cls = 'aucun'; }
    el.className = 'fichier ' + cls;
    el.title = CLASSEUR.nom ? 'Classeur : ' + CLASSEUR.nom : 'Aucun classeur : cliquez sur Enregistrer pour en choisir l\'emplacement';
    el.innerHTML = '<span class="f-ico">' + ICONES.excel + '</span>'
      + '<span class="f-txt"><b>' + esc(CLASSEUR.nom || 'Nouveau projet') + '</b><span><i class="pastille"></i>' + etat + '</span></span>';
  }
  document.title = (CLASSEUR.modifie ? '● ' : '') + (CLASSEUR.nom ? CLASSEUR.nom.replace(/\.xlsx$/i,'') + ' — ' : '') + 'Enregistrements Contrôles';
}

/* ---------- mémorisation du dernier classeur (IndexedDB) ---------- */
function idbOuvrir(){
  return new Promise(function(res, rej){
    var r = indexedDB.open('enreg-ctrl', 1);
    r.onupgradeneeded = function(){ r.result.createObjectStore('kv'); };
    r.onsuccess = function(){ res(r.result); };
    r.onerror = function(){ rej(r.error); };
  });
}
function idbLire(k){
  return idbOuvrir().then(function(db){
    return new Promise(function(res){
      var q = db.transaction('kv').objectStore('kv').get(k);
      q.onsuccess = function(){ res(q.result || null); };
      q.onerror = function(){ res(null); };
    });
  }).catch(function(){ return null; });
}
function idbEcrire(k, v){
  return idbOuvrir().then(function(db){
    return new Promise(function(res){
      var t = db.transaction('kv', 'readwrite');
      t.objectStore('kv').put(v, k);
      t.oncomplete = function(){ res(true); };
      t.onerror = function(){ res(false); };
    });
  }).catch(function(){ return false; });
}

async function autorisation(handle, demander){
  var o = {mode:'readwrite'};
  try{
    if(await handle.queryPermission(o) === 'granted') return true;
    if(demander && await handle.requestPermission(o) === 'granted') return true;
  }catch(e){}
  return false;
}

/* ---------- lecture du classeur ---------- */
function lireClasseur(buf){
  var wb = XLSX.read(buf, {type:'array'});
  var ws = wb.Sheets[FEUILLE_DONNEES];
  if(!ws){
    /* Classeur préparé à la main : on part de bases vides et on reprend ses onglets */
    var aBases = FEUILLES_EDITABLES.some(function(f){ return wb.Sheets[f.feuille]; });
    if(!aBases) throw new Error('Ce classeur ne contient ni les données de l\'application (onglet « ' + FEUILLE_DONNEES + ' ») ni d\'onglet GAMSERIE, TS ou Matrice.');
    return {mem:{'rc.db':JSON.stringify({gamserie:[], ts:[], matrice:[]})}, wb:wb, neuf:true};
  }
  var lignes = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
  var parts = {};
  for(var i=1;i<lignes.length;i++){
    var k = String(lignes[i][0] || '').trim();
    if(CLES_DONNEES.indexOf(k) < 0) continue;
    (parts[k] = parts[k] || [])[+lignes[i][1] || 0] = String(lignes[i][2] === null || lignes[i][2] === undefined ? '' : lignes[i][2]);
  }
  var mem = {};
  for(var c in parts){
    mem[c] = parts[c].join('');
    JSON.parse(mem[c]);   // contrôle d'intégrité : lève une erreur si le contenu a été altéré
  }
  return {mem:mem, wb:wb, neuf:false};
}

/* ---------- reprise des onglets modifiés dans Excel ---------- */
function lignesFeuille(ws){
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:'', blankrows:false});
}
/* Forme comparable d'un tableau : textes rognés, nombres arrondis, vides de fin retirés */
function empreinte(rows){
  return JSON.stringify(rows.map(function(r){
    var a = r.map(function(v){
      if(typeof v === 'number') return String(Math.round(v * 1e9) / 1e9);
      return String(v === null || v === undefined ? '' : v).trim();
    });
    while(a.length && a[a.length-1] === '') a.pop();
    return a;
  }).filter(function(r){ return r.length; }));
}
/* Compare chaque onglet modifiable du classeur à ce que l'application aurait
   écrit ; un onglet différent a été modifié dans Excel et remplace la base.
   tout = true : reprend tous les onglets présents (classeur préparé à la main). */
function integrerFeuillesExcel(wb, tout){
  var bilan = {bases:[], revisions:0, ignores:[], erreurs:[]};
  /* les tableaux de référence sont calculés avant toute intégration (TS dépend de la matrice et de GAMSERIE) */
  var refs = FEUILLES_EDITABLES.map(function(f){ return wb.Sheets[f.feuille] ? empreinte(lignesFeuille(feuille(f.tableau()))) : null; });
  FEUILLES_EDITABLES.forEach(function(f, i){
    var ws = wb.Sheets[f.feuille]; if(!ws) return;
    var rows = lignesFeuille(ws);
    if(!tout && empreinte(rows) === refs[i]) return;
    if(rows.length < 2){ bilan.ignores.push(f.feuille); return; }   // onglet vidé : on ne remplace pas la base par du vide
    try{
      bilan.revisions += importerTableau(f.base, rows, 'remplacer', 'Modifié dans Excel');
      bilan.bases.push(f.feuille);
    }catch(e){ bilan.erreurs.push(f.feuille + ' : ' + (e.message || e)); }
  });
  if(bilan.bases.length){ syncLignes(); saveDraft(); }
  return bilan;
}
function messageBilan(bilan){
  var h = '';
  if(bilan.bases.length){
    h += '<p style="margin:0">Données reprises depuis Excel : <b>' + bilan.bases.map(esc).join(', ') + '</b>.</p>';
    if(bilan.revisions) h += '<p style="margin:10px 0 0"><b>' + bilan.revisions + ' révision' + (bilan.revisions>1?'s':'') + ' en brouillon</b> : '
      + 'les gammes ou valeurs modifiées ne s\'appliquent aux nouveaux rapports qu\'après validation dans <b>Paramètres → GAMSERIE / TS</b> (bouton « Valider les brouillons »).</p>';
    h += '<p style="margin:10px 0 0">Enregistrez (<b>Ctrl + S</b>) pour mettre le classeur à jour.</p>';
  }
  if(bilan.ignores.length) h += '<p style="margin:10px 0 0">Onglet vide ignoré : ' + bilan.ignores.map(esc).join(', ') + '.</p>';
  if(bilan.erreurs.length) h += '<p style="margin:10px 0 0;color:var(--nc)">Non repris : ' + bilan.erreurs.map(esc).join(' · ') + '</p>';
  return h;
}
function feuillesUtilisateur(wb){
  return wb.SheetNames.filter(function(n){ return FEUILLES_APPLI.indexOf(n) < 0; }).map(function(n){ return [n, wb.Sheets[n]]; });
}

/* ---------- construction du classeur ---------- */
function valeurCellule(v){
  if(v === null || v === undefined) return '';
  if(typeof v === 'number') return isFinite(v) ? v : '';
  return v;
}
function feuille(rows){
  rows = rows.map(function(r){ return r.map(valeurCellule); });
  var ws = XLSX.utils.aoa_to_sheet(rows);
  var larg = [];
  rows.forEach(function(r){
    r.forEach(function(v, c){
      var n = String(v).length;
      if(!larg[c] || larg[c] < n) larg[c] = n;
    });
  });
  ws['!cols'] = larg.map(function(n){ return {wch: Math.max(8, Math.min(48, (n || 0) + 2))}; });
  if(rows.length > 1 && rows[0].length) ws['!autofilter'] = {ref: XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:rows.length-1, c:rows[0].length-1}})};
  return ws;
}
function tableauRapports(){
  var rows = [['Date','Article','Lot','Commande','Fournisseur','Contrôleur','Origine','Qté commandée','Qté réceptionnée','Qté échantillon',
    'Conformes','Non conformes','Durée contrôle','Décision','Type NC','Sous-type NC','Gamme','Rév. gamme','Rév. valeurs','Indice de définition','Commentaires','Identifiant']];
  REPORTS.slice().sort(function(a,b){ return String(b.date||'') < String(a.date||'') ? -1 : String(b.date||'') > String(a.date||'') ? 1 : 0; })
    .forEach(function(r){
      var t = totaux2(r);
      rows.push([frDate(r.date), r.article, r.lot, r.commande, r.fournisseur, r.controleur, r.origine||'',
        nombreOuTexte(r.qteCommandee), nombreOuTexte(r.qteReceptionnee), nombreOuTexte(r.qteEchantillon),
        t.conforme, t.nonConforme, r.dureeControle, libDecision(r.decision), r.typeNC||'', r.sousTypeNC||'',
        r.famille||'', r.gamRev===null||r.gamRev===undefined?'':r.gamRev, r.tsRev===null||r.tsRev===undefined?'':r.tsRev,
        r.indiceDef||'', r.commentaires||'', r.id]);
    });
  return rows;
}
function tableauMesures(){
  var rows = [['Date','Article','Lot','Fournisseur','Code','Spécification','Type','Mini.','Nominal','Maxi.','Pièce','Valeur','Résultat','Identifiant rapport']];
  REPORTS.forEach(function(rep){
    var n = nbPieces2(rep), gen = rep.generales || {};
    (rep.lignes||[]).forEach(function(l, i){
      var k = specKind(l);
      if(k === 'generale'){
        var g = String(gen[l.code]||'');
        if(g === '') return;
        rows.push([frDate(rep.date), rep.article, rep.lot, rep.fournisseur, l.code, l.specification, 'Générale', '', '', '', 'lot',
          g === '1' ? 'OK' : 'NOK', g === '1' ? 'Conforme' : 'Non conforme', rep.id]);
        return;
      }
      if(k !== 'individuelle' || !l.testable) return;
      for(var j=0;j<n;j++){
        var brut = (rep.mesures[i]||[])[j];
        if(brut === '' || brut === undefined || brut === null) continue;
        var v = num(brut), hors = !isNaN(v) && (v < l.mini - EPS || v > l.maxi + EPS);
        var val = l.qual ? (String(brut) === '1' ? 'OK' : 'NOK') : (isNaN(v) ? brut : v);
        var res = l.qual ? (String(brut) === '1' ? 'Conforme' : 'Non conforme') : (isNaN(v) ? '' : hors ? 'Hors tolérance' : 'Conforme');
        rows.push([frDate(rep.date), rep.article, rep.lot, rep.fournisseur, l.code, l.specification, l.qual ? 'OK / NOK' : 'Mesure',
          l.qual ? '' : l.mini, l.qual ? '' : l.nom, l.qual ? '' : l.maxi, j+1, val, res, rep.id]);
      }
    });
  });
  return rows;
}
function nombreOuTexte(v){ var n = num(v); return isNaN(n) ? (v || '') : n; }

function construireClasseur(){
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, feuille(tableauRapports()), 'Rapports');
  XLSX.utils.book_append_sheet(wb, feuille(tableauMesures()), 'Mesures');
  XLSX.utils.book_append_sheet(wb, feuille(tableauGamserie()), 'GAMSERIE');
  XLSX.utils.book_append_sheet(wb, feuille(tableauTS()), 'TS');
  XLSX.utils.book_append_sheet(wb, feuille(tableauMatrice()), 'Matrice');
  XLSX.utils.book_append_sheet(wb, feuille(tableauJournal()), 'Journal');
  XLSX.utils.book_append_sheet(wb, feuilleModeEmploi(), 'Mode d\'emploi');
  CLASSEUR.autres.forEach(function(x){ XLSX.utils.book_append_sheet(wb, x[1], x[0]); });

  /* Données de l'application, découpées en parts de 30 000 caractères */
  var aoa = [['Clé','Partie','Contenu (ne pas modifier)']];
  CLES_DONNEES.forEach(function(k){
    var v = get(k);
    if(v === null) return;
    var i = 0, p = 0;
    do{
      var j = Math.min(i + TAILLE_PART, v.length);
      var c = v.charCodeAt(j - 1);
      if(j < v.length && c >= 0xD800 && c <= 0xDBFF) j--;   // ne pas couper une paire de substitution
      aoa.push([k, p++, v.slice(i, j)]);
      i = j;
    } while(i < v.length);
  });
  var wsd = XLSX.utils.aoa_to_sheet(aoa);
  wsd['!cols'] = [{wch:14},{wch:8},{wch:80}];
  XLSX.utils.book_append_sheet(wb, wsd, FEUILLE_DONNEES);

  wb.Workbook = {Sheets: wb.SheetNames.map(function(n){ return {name:n, Hidden: n === FEUILLE_DONNEES ? 1 : 0}; })};
  wb.Props = {Title:'Enregistrements Contrôles', Application:'Enregistrements Contrôles'};
  return XLSX.write(wb, {bookType:'xlsx', type:'array', compression:true});
}

function feuilleModeEmploi(){
  var ws = XLSX.utils.aoa_to_sheet([
    ['Enregistrements Contrôles — mode d\'emploi du classeur'],
    [''],
    ['Onglet', 'Modifiable dans Excel', 'Contenu'],
    ['GAMSERIE', 'Oui', 'Gammes de contrôle : une ligne par famille + code (Famille Méthode, Code, Criticité, Spécification, Moyen, Méthode).'],
    ['TS', 'Oui', 'Valeurs par article : une ligne par article + code (Nominal, Tol -, Tol +, Type spéc. = Générale / Individuelle). Mini et Maxi sont recalculés.'],
    ['Matrice', 'Oui', 'Article, code fournisseur, famille, matière, liasse et références documentaires.'],
    ['Rapports', 'Non', 'Liste des rapports enregistrés (saisie dans l\'application).'],
    ['Mesures', 'Non', 'Toutes les valeurs relevées.'],
    ['Journal', 'Non', 'Historique des révisions.'],
    [''],
    ['Les colonnes Révision, Date de révision, Statut, Spécification (TS), Type saisie, Mini. et Maxi. sont informatives : elles sont recalculées par l\'application.'],
    ['Une gamme ou des valeurs modifiées dans Excel ouvrent une révision en brouillon, à valider dans l\'application avant de servir aux nouveaux rapports.'],
    ['Enregistrez et fermez le classeur dans Excel avant d\'enregistrer depuis l\'application. Vos propres onglets supplémentaires sont conservés (sans mise en forme).'],
    ['Ne modifiez pas l\'onglet masqué « _Donnees ».']
  ]);
  ws['!cols'] = [{wch:16},{wch:22},{wch:120}];
  return ws;
}

/* ---------- modifications faites dans Excel pendant que l'application est ouverte ---------- */
var verifEnCours = false;
async function dateFichier(h){ try{ return (await h.getFile()).lastModified; }catch(e){ return 0; } }
/* Renvoie true si l'on peut poursuivre (pas de modification externe, ou choix fait). */
async function verifierModifExterne(avantEcriture){
  var h = CLASSEUR.handle;
  if(!h || verifEnCours || CLASSEUR.ecritureEnCours) return true;
  var acc = document.getElementById('accueil'); if(acc && !acc.hidden) return true;
  if(!(await autorisation(h, false))) return true;
  verifEnCours = true;
  try{
    var f = await h.getFile();
    if(!CLASSEUR.dateFichier || f.lastModified === CLASSEUR.dateFichier) return true;
    if(!CLASSEUR.modifie && !avantEcriture){
      await chargerDepuisHandle(h, true);   // rien à perdre : rechargement complet
      return false;
    }
    return await choixConflit(h, f);
  }catch(e){ return true; }
  finally{ verifEnCours = false; }
}
function choixConflit(h, f){
  return new Promise(function(res){
    modal('Classeur modifié dans Excel',
      'Le classeur <b>' + esc(h.name) + '</b> a été modifié en dehors de l\'application.'
      + '<ul><li><b>Intégrer les modifications Excel</b> : reprend les onglets GAMSERIE, TS et Matrice du fichier, en gardant vos rapports en cours.</li>'
      + '<li><b>Garder la version de l\'application</b> : les modifications faites dans Excel seront écrasées au prochain enregistrement.</li></ul>',
      function(){}, 'Fermer');
    var ft = document.querySelector('.modal .ft');
    ft.innerHTML = '<button class="btn" data-garder>Garder la version de l\'application</button><button class="btn pri" data-integrer>Intégrer les modifications Excel</button>';
    ft.querySelector('[data-garder]').onclick = function(){
      document.querySelector('.modal').remove();
      CLASSEUR.dateFichier = f.lastModified;
      res(true);
    };
    ft.querySelector('[data-integrer]').onclick = async function(){
      document.querySelector('.modal').remove();
      try{
        var lu = lireClasseur(await f.arrayBuffer());
        CLASSEUR.autres = feuillesUtilisateur(lu.wb);
        var bilan = integrerFeuillesExcel(lu.wb, false);
        CLASSEUR.dateFichier = f.lastModified;
        render();
        if(bilan.bases.length || bilan.erreurs.length) modal('Modifications Excel intégrées', messageBilan(bilan));
        else toast('Aucune modification à reprendre dans GAMSERIE, TS ou Matrice.');
      }catch(e){ modal('Lecture impossible', esc(e.message || e.name)); }
      res(false);
    };
  });
}

/* Bouton « Reprendre les modifications faites dans Excel » */
async function reprendreExcel(){
  var h = CLASSEUR.handle;
  if(!h){ ouvrirClasseur(); return; }
  if(!(await autorisation(h, true))){ toast('Accès au classeur refusé.'); return; }
  var f;
  try{ f = await h.getFile(); }catch(e){ modal('Lecture impossible', esc(e.message || e.name)); return; }
  if(!CLASSEUR.modifie){ await chargerDepuisHandle(h, true); return; }
  await choixConflit(h, f);
}

/* ---------- enregistrement ---------- */
/* Termine l'enregistrement différé du rapport en cours avant d'écrire. */
function avantEnregistrement(){
  if(typeof autoT !== 'undefined') clearTimeout(autoT);
  if(VIEW === 'rapport' && R) enregistrerAuto(true);
}
async function enregistrer(){
  if(CLASSEUR.ecritureEnCours) return;
  avantEnregistrement();
  if(!FS_OK) return telechargerClasseur();
  if(!CLASSEUR.handle) return enregistrerSous();
  if(!(await autorisation(CLASSEUR.handle, true))){ toast('Accès au classeur refusé : choisissez son emplacement.'); return enregistrerSous(); }
  await ecrireClasseur(CLASSEUR.handle);
}
async function enregistrerSous(){
  if(CLASSEUR.ecritureEnCours) return;
  avantEnregistrement();
  if(!FS_OK) return telechargerClasseur();
  var h;
  try{
    h = await window.showSaveFilePicker({suggestedName: CLASSEUR.nom || NOM_DEFAUT, types: TYPES_XLSX, id:'enreg-ctrl'});
  }catch(e){
    if(e && e.name === 'AbortError') return;
    toast('Choix de l\'emplacement impossible : ' + (e.message || e.name)); return;
  }
  CLASSEUR.handle = h; CLASSEUR.nom = h.name; CLASSEUR.dateFichier = 0;
  idbEcrire('classeur', h);
  await ecrireClasseur(h);
}
async function ecrireClasseur(h){
  if(!(await verifierModifExterne(true))) return;
  CLASSEUR.ecritureEnCours = true; majBarreFichier();
  try{
    var data = construireClasseur();
    var w = await h.createWritable();
    await w.write(data);
    await w.close();
    CLASSEUR.modifie = false; CLASSEUR.enregistreLe = heureCourante();
    CLASSEUR.dateFichier = await dateFichier(h);
    toast('Enregistré dans ' + h.name);
  }catch(e){
    modal('Enregistrement impossible',
      'Le classeur <b>' + esc(h.name) + '</b> n\'a pas pu être écrit.'
      + '<ul><li>S\'il est ouvert dans Excel, fermez-le puis enregistrez à nouveau.</li>'
      + '<li>Vérifiez que vous avez le droit d\'écrire dans ce dossier, ou choisissez un autre emplacement (Paramètres → Sauvegarde).</li></ul>'
      + '<p class="hint" style="margin:10px 0 0">Détail : ' + esc(e && (e.message || e.name) || 'erreur inconnue') + '</p>');
  }finally{
    CLASSEUR.ecritureEnCours = false;
    majBarreFichier(); majEtatSauvegarde();
    if(VIEW === 'parametres' && PSET === 'sauvegarde') render();
  }
}
/* Navigateur sans accès direct aux fichiers : téléchargement d'une copie */
function telechargerClasseur(){
  var nom = CLASSEUR.nom || NOM_DEFAUT;
  var blob = new Blob([construireClasseur()], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = nom;
  document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 500);
  CLASSEUR.nom = nom; CLASSEUR.modifie = false; CLASSEUR.enregistreLe = heureCourante();
  majBarreFichier(); majEtatSauvegarde();
  toast('Classeur téléchargé : ' + nom + ' (dossier Téléchargements).');
}

/* ---------- ouverture / création ---------- */
/* lu = résultat de lireClasseur (ou null pour partir des données de démonstration) */
function appliquerDonnees(lu, handle, nom){
  var bilan = null;
  CHARGEMENT = true;
  try{
    MEM = lu ? lu.mem : {};
    CLASSEUR.handle = handle; CLASSEUR.nom = nom || ''; CLASSEUR.enregistreLe = '';
    CLASSEUR.autres = lu ? feuillesUtilisateur(lu.wb) : [];
    initialiserApplication();
    if(lu) bilan = integrerFeuillesExcel(lu.wb, lu.neuf);
  }finally{
    CHARGEMENT = false;
  }
  /* des données reprises d'Excel ne sont pas encore dans l'onglet _Donnees : à enregistrer */
  CLASSEUR.modifie = !!(bilan && bilan.bases.length);
  fermerAccueil();
  render();
  majBarreFichier();
  if(bilan && (bilan.bases.length || bilan.erreurs.length || bilan.ignores.length)) modal('Modifications faites dans Excel', messageBilan(bilan));
  return bilan;
}
function abandonnerModifications(){
  return !CLASSEUR.modifie || confirm('Des modifications ne sont pas enregistrées dans le classeur actuel. Continuer quand même ?');
}
async function chargerDepuisHandle(h, rechargement){
  try{
    var f = await h.getFile();
    var bilan = appliquerDonnees(lireClasseur(await f.arrayBuffer()), h, h.name);
    CLASSEUR.dateFichier = f.lastModified;
    idbEcrire('classeur', h);
    if(!bilan.bases.length) toast(rechargement ? 'Classeur rechargé : modifications faites dans Excel prises en compte.' : 'Classeur ouvert : ' + h.name);
    return true;
  }catch(e){
    modal('Ouverture impossible', 'Le classeur <b>' + esc(h.name) + '</b> n\'a pas pu être lu.'
      + '<p style="margin:10px 0 0">' + esc(e && (e.message || e.name) || 'erreur inconnue') + '</p>');
    return false;
  }
}
async function ouvrirClasseur(){
  if(!abandonnerModifications()) return;
  if(!FS_OK){
    pickFileBinaire(function(buf, nom){
      try{ appliquerDonnees(lireClasseur(buf), null, nom); toast('Classeur ouvert : ' + nom); }
      catch(e){ modal('Ouverture impossible', esc(e.message || e.name)); }
    });
    return;
  }
  var hs;
  try{ hs = await window.showOpenFilePicker({types: TYPES_XLSX, id:'enreg-ctrl', multiple:false}); }
  catch(e){ if(!e || e.name !== 'AbortError') toast('Ouverture impossible : ' + (e.message || e.name)); return; }
  await chargerDepuisHandle(hs[0]);
}
async function nouveauClasseur(){
  if(!abandonnerModifications()) return;
  appliquerDonnees(null, null, '');
  CLASSEUR.modifie = true; majBarreFichier();
  if(FS_OK) await enregistrerSous();
}
function pickFileBinaire(cb){
  var el = document.getElementById('filepick');
  el.value = ''; el.accept = '.xlsx'; el.multiple = false;
  el.onchange = function(){
    var f = el.files[0]; if(!f) return;
    var fr = new FileReader();
    fr.onload = function(){ cb(fr.result, f.name); };
    fr.readAsArrayBuffer(f);
  };
  el.click();
}

/* ---------- écran d'accueil ---------- */
function fermerAccueil(){ var a = document.getElementById('accueil'); if(a) a.hidden = true; }
async function afficherAccueil(){
  var a = document.getElementById('accueil');
  var dernier = FS_OK ? await idbLire('classeur') : null;
  var h = '<div class="accueil-box" role="dialog" aria-modal="true" aria-labelledby="acc-titre">'
    + '<div class="acc-logo">' + LOGO + '</div>'
    + '<h1 id="acc-titre">Enregistrements Contrôles</h1>'
    + '<p class="acc-intro">Rapports de contrôle d\'entrée, gammes, valeurs et révisions. '
    + 'Le projet est enregistré dans un <b>classeur Excel</b> dont vous choisissez l\'emplacement.</p>'
    + '<div class="acc-actions">';
  if(dernier) h += '<button class="acc-btn pri" id="acc-reprendre">' + ICONES.reprendre + '<span><b>Reprendre « ' + esc(dernier.name) + ' »</b><small>Dernier classeur utilisé sur ce poste</small></span></button>';
  h += '<button class="acc-btn' + (dernier ? '' : ' pri') + '" id="acc-ouvrir">' + ICONES.open + '<span><b>Ouvrir un classeur existant…</b><small>Fichier .xlsx créé par l\'application</small></span></button>'
    + '<button class="acc-btn" id="acc-nouveau">' + ICONES.plus + '<span><b>Créer un nouveau classeur…</b><small>Choisissez le dossier et le nom du fichier</small></span></button>'
    + '</div>'
    + '<button class="acc-lien" id="acc-demo">Découvrir l\'application sans classeur (données de démonstration)</button>';
  if(!FS_OK) h += '<p class="acc-alerte">Ce navigateur ne permet pas d\'écrire directement dans un fichier : chaque enregistrement téléchargera une copie du classeur. Utilisez <b>Google Chrome</b> ou <b>Microsoft Edge</b> pour enregistrer à l\'emplacement de votre choix.</p>';
  h += '</div>';
  a.innerHTML = h;
  a.hidden = false;
  var b;
  if(b = document.getElementById('acc-reprendre')) b.onclick = async function(){
    if(await autorisation(dernier, true)) await chargerDepuisHandle(dernier);
    else toast('Accès refusé au classeur ' + dernier.name + '.');
  };
  document.getElementById('acc-ouvrir').onclick = ouvrirClasseur;
  document.getElementById('acc-nouveau').onclick = nouveauClasseur;
  document.getElementById('acc-demo').onclick = function(){
    appliquerDonnees(null, null, '');
    toast('Mode découverte : cliquez sur Enregistrer pour créer votre classeur.');
  };
  var first = a.querySelector('.acc-btn'); if(first) first.focus();
}

/* ---------- thème clair / sombre ---------- */
function themeCourant(){ return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
function appliquerTheme(t){
  document.documentElement.setAttribute('data-theme', t);
  try{ localStorage.setItem('ec.theme', t); }catch(e){}
  var b = document.getElementById('b-theme');
  if(b){
    b.innerHTML = t === 'dark' ? ICONES.sun : ICONES.moon;
    b.title = t === 'dark' ? 'Passer en mode clair' : 'Passer en mode sombre';
    b.setAttribute('aria-label', b.title);
  }
}

var LOGO = '<svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true"><rect x="2" y="2" width="28" height="28" rx="8" fill="var(--accent)"/>'
  + '<path d="M10 9.5h9M10 14h12M10 18.5h6" stroke="var(--accent-ink)" stroke-width="2" stroke-linecap="round"/>'
  + '<path d="M17.5 22l2.3 2.3 4.7-4.8" fill="none" stroke="var(--accent-ink)" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

/* ---------- branchements globaux ---------- */
(function(){
  document.getElementById('brand-logo').innerHTML = LOGO;
  var bs = document.getElementById('b-enregistrer');
  bs.innerHTML = ICONES.save + '<span>Enregistrer</span>';
  bs.onclick = function(){ enregistrer(); };
  document.getElementById('b-theme').onclick = function(){ appliquerTheme(themeCourant() === 'dark' ? 'light' : 'dark'); };
  appliquerTheme(themeCourant());

  window.addEventListener('keydown', function(e){
    if((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === 's' || e.key === 'S')){
      e.preventDefault();
      var acc = document.getElementById('accueil');
      if(acc && !acc.hidden) return;
      if(e.shiftKey) enregistrerSous(); else enregistrer();
    }
  }, true);
  window.addEventListener('focus', function(){ verifierModifExterne(false); });
  document.addEventListener('visibilitychange', function(){ if(document.visibilityState === 'visible') verifierModifExterne(false); });
  window.addEventListener('beforeunload', function(e){
    if(CLASSEUR.modifie){ e.preventDefault(); e.returnValue = ''; }
  });

  majBarreFichier();
  afficherAccueil();
})();
