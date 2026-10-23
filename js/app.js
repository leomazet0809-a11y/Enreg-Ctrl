/* =======================================================================
   1. Stockage en mémoire de travail
   Les données vivent ici pendant la session ; elles sont écrites dans le
   classeur Excel choisi par l'utilisateur (bouton Enregistrer ou Ctrl+S),
   voir js/classeur.js. Toute écriture qui change une valeur signale des
   modifications non enregistrées.
   ======================================================================= */
var LS = null;            // plus de mémoire navigateur : le classeur fait foi
var MEM = {};
function put(k,v){
  var avant = MEM.hasOwnProperty(k) ? MEM[k] : null;
  MEM[k] = v;
  if(typeof marquerModifie === 'function' && sansHorodatage(avant) !== sansHorodatage(v)) marquerModifie();
}
function get(k){ return MEM.hasOwnProperty(k) ? MEM[k] : null; }
/* L'horodatage « majLe » seul ne constitue pas une modification. */
function sansHorodatage(s){ return s === null ? null : String(s).replace(/"majLe":(null|"[^"]*")/g, ''); }

var SEED = {"gamserie":[{"famille":"Liaison CBMI","code":"A","criticite":"Critique","specification":"Certificat fournisseur (CCPU)","moyen":"N/A","methode":"Cf référentiel NORME"},{"famille":"Liaison CBMI","code":"B","criticite":"Majeure","specification":"Conditionnement/étiquetage","moyen":"N/A","methode":"Cf référentiel CONDI"},{"famille":"Liaison CBMI","code":"C","criticite":"Majeure","specification":"État visuel","moyen":"N/A","methode":"Cf référentiel DEF"},{"famille":"Liaison CBMI","code":"D","criticite":"Majeure","specification":"Propreté","moyen":"N/A","methode":"Cf référentiel DEF"},{"famille":"Liaison CBMI","code":"E","criticite":"Majeure","specification":"Hauteur","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"F","criticite":"Majeure","specification":"Longueur coté tige","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"G","criticite":"Majeure","specification":"Epaisseur","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"H","criticite":"Majeure","specification":"Longueur coté boitier","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"I","criticite":"Majeure","specification":"Largeur","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"J","criticite":"Majeure","specification":"Largeur croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"K","criticite":"Majeure","specification":"Largeur croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"L","criticite":"Majeure","specification":"Epaisseur croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"M","criticite":"Majeure","specification":"Epaisseur croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"N","criticite":"Majeure","specification":"Position X croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"O","criticite":"Majeure","specification":"Position Y croix","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"P","criticite":"Majeure","specification":"Largeur méplat ou Ø trou passage vis (DN10)","moyen":"Pied à coulisse ou pige","methode":"N/A"},{"famille":"Liaison CBMI","code":"Q","criticite":"Majeure","specification":"Ø passage tige","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"R","criticite":"Majeure","specification":"Position X méplat ou trou","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"Liaison CBMI","code":"S","criticite":"Majeure","specification":"Position Y méplat ou trou","moyen":"Pied à coulisse","methode":"N/A"},{"famille":"familleA","code":"A","criticite":"Majeure","specification":"spécification 1","moyen":"","methode":""},{"famille":"familleA","code":"B","criticite":"Majeure","specification":"spécification 2","moyen":"","methode":""},{"famille":"familleA","code":"C","criticite":"Majeure","specification":"spécification 3","moyen":"","methode":""},{"famille":"familleA","code":"D","criticite":"Majeure","specification":"spécification 4","moyen":"","methode":""},{"famille":"familleA","code":"E","criticite":"Majeure","specification":"spécification 5","moyen":"","methode":""},{"famille":"familleA","code":"F","criticite":"Majeure","specification":"spécification 6","moyen":"","methode":""}],"ts":[{"article":"041038_CBMI","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[13,-0.1,0.1,"Individuelle"],"F":[28,-0.1,0.1,"Individuelle"],"G":[1,-0.1,0.1,"Individuelle"],"H":[18.6,-0.1,0.1,"Individuelle"],"I":[24,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[8,-0.1,0.1,"Individuelle"],"O":[12,-0.1,0.1,"Individuelle"],"P":[10,-0.1,0.1,"Individuelle"],"Q":[5.2,0,0.1,"Individuelle"],"R":[16,-0.1,0.1,"Individuelle"],"S":[12,-0.1,0.1,"Individuelle"]}},{"article":"401538_CBMI","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[14,-0.1,0.1,"Individuelle"],"F":[25,-0.1,0.1,"Individuelle"],"G":[1,-0.1,0.1,"Individuelle"],"H":[25,-0.1,0.1,"Individuelle"],"I":[17,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[8,-0.1,0.1,"Individuelle"],"O":[8.5,-0.1,0.1,"Individuelle"],"P":[6,0,0.1,"Individuelle"],"Q":[10.2,0,0.2,"Individuelle"],"R":[8,-0.1,0.1,"Individuelle"],"S":[8.5,-0.1,0.1,"Individuelle"]}},{"article":"402038_CBMI","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[14,-0.1,0.1,"Individuelle"],"F":[27,-0.1,0.1,"Individuelle"],"G":[1,-0.1,0.1,"Individuelle"],"H":[27,-0.1,0.1,"Individuelle"],"I":[24,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[10,-0.1,0.1,"Individuelle"],"O":[12,-0.1,0.1,"Individuelle"],"P":[8,0,0.1,"Individuelle"],"Q":[12.2,0,0.2,"Individuelle"],"R":[10,-0.1,0.1,"Individuelle"],"S":[12,-0.1,0.1,"Individuelle"]}},{"article":"043238_CBMI","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[17,-0.1,0.1,"Individuelle"],"F":[33.5,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":[33.5,-0.1,0.1,"Individuelle"],"I":[28,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[11.5,-0.1,0.1,"Individuelle"],"O":[14,-0.1,0.1,"Individuelle"],"P":[11,0,0.1,"Individuelle"],"Q":[16.2,0,0.2,"Individuelle"],"R":[11.5,-0.1,0.1,"Individuelle"],"S":[14,-0.1,0.1,"Individuelle"]}},{"article":"045038_CBMI","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[20,-0.4,0.4,"Individuelle"],"F":[40,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":[40,-0.1,0.1,"Individuelle"],"I":[34,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[13.5,-0.1,0.1,"Individuelle"],"O":[17,-0.1,0.1,"Individuelle"],"P":[12,0,0.1,"Individuelle"],"Q":[18.2,0,0.2,"Individuelle"],"R":[13.5,-0.1,0.1,"Individuelle"],"S":[17,-0.1,0.1,"Individuelle"]}},{"article":"041038_CBMIF05","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[73.73,-0.1,0.1,"Individuelle"],"F":[31,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":["-",-0.1,0.1,"Individuelle"],"I":[22,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[10,-0.1,0.1,"Individuelle"],"O":[11,-0.1,0.1,"Individuelle"],"P":[12,-0.1,0.1,"Individuelle"],"Q":[5.2,0,0.1,"Individuelle"],"R":[19,-0.1,0.1,"Individuelle"],"S":[11,-0.1,0.1,"Individuelle"]}},{"article":"041538_CBMIF05","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[73.5,-0.1,0.1,"Individuelle"],"F":[24,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.05,"Individuelle"],"H":[24,-0.1,0.1,"Individuelle"],"I":[22,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[9,-0.1,0.1,"Individuelle"],"O":[11,-0.1,0.1,"Individuelle"],"P":[6,0,0.1,"Individuelle"],"Q":[10.2,0,0.2,"Individuelle"],"R":[9,-0.1,0.1,"Individuelle"],"S":[11,-0.1,0.1,"Individuelle"]}},{"article":"042038_CBMIF05","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[71,-0.1,0.1,"Individuelle"],"F":[27,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":[27,-0.1,0.1,"Individuelle"],"I":[22,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[10,-0.1,0.1,"Individuelle"],"O":[11,-0.1,0.1,"Individuelle"],"P":[8,0,0.1,"Individuelle"],"Q":[12.2,0,0.2,"Individuelle"],"R":[10,-0.1,0.1,"Individuelle"],"S":[11,-0.1,0.1,"Individuelle"]}},{"article":"043238_CBMIF05","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[74,-0.1,0.1,"Individuelle"],"F":[34,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":[34,-0.1,0.1,"Individuelle"],"I":[28,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[11.5,-0.1,0.1,"Individuelle"],"O":[14,-0.1,0.1,"Individuelle"],"P":[11,0,0.1,"Individuelle"],"Q":[16.2,0,0.2,"Individuelle"],"R":[11.5,-0.1,0.1,"Individuelle"],"S":[14,-0.1,0.1,"Individuelle"]}},{"article":"045038_CBMIF05","vals":{"A":[1,0,0,"Générale"],"B":[1,0,0,"Générale"],"C":[1,0,0,"Générale"],"D":[1,0,0,"Générale"],"E":[75,-0.1,0.1,"Individuelle"],"F":[40,-0.1,0.1,"Individuelle"],"G":[1.5,-0.1,0.1,"Individuelle"],"H":[38,-0.1,0.1,"Individuelle"],"I":[34,-0.1,0.1,"Individuelle"],"J":[12.2,0,0.1,"Individuelle"],"K":[12.2,0,0.1,"Individuelle"],"L":[6,0,0.1,"Individuelle"],"M":[6,0,0.1,"Individuelle"],"N":[13.5,-0.1,0.1,"Individuelle"],"O":[17,-0.1,0.1,"Individuelle"],"P":[12,0,0.1,"Individuelle"],"Q":[18.2,0,0.2,"Individuelle"],"R":[5.5,-0.1,0.1,"Individuelle"],"S":[17,-0.1,0.1,"Individuelle"]}},{"article":"articleA","vals":{"A":[5,-1,1,"Individuelle"],"B":[10,-1,1,"Individuelle"],"C":[15,-1,1,"Individuelle"],"D":[20,-1,1,"Individuelle"],"E":[25,-1,1,"Individuelle"],"F":[30,-1,1,"Individuelle"]}}],"matrice":[{"article":"articleA","fournisseur":"","famille":"familleA","matiere":"","liasse":"","norme":"","planMethode":"","planControle":"","planSituation":"","conditionnement":"","tauxEchantillon":"","defautheque":"","gammeSerie":"","tableauSpec":""},{"article":"041038_CBMI","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"041038_CBMIF05","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"041538_CBMIF05","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"042038_CBMIF05","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"043238_CBMI","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"043238_CBMIF05","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"045038_CBMI","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"045038_CBMIF05","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"401538_CBMI","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"},{"article":"402038_CBMI","fournisseur":"4001017","famille":"Liaison CBMI","matiere":"1.4307","liasse":"OUI","norme":"N/A","planMethode":"PM-0081","planControle":"PC-0081","planSituation":"PS-0035","conditionnement":"CONDI-0045","tauxEchantillon":"ECHANT-100","defautheque":"DEF-0051","gammeSerie":"GAMSERIE-0073","tableauSpec":"TS-0081"}]};
var DB = {gamserie:[],ts:[],matrice:[]};
var REPORTS = [];
var DOCS = [];             // documents PDF stockés (méthodes de contrôle)
var VERS = {gam:{}, ts:{}};// révision courante par famille (gam) et par article (ts)
var HIST = [];             // révisions figées, avec leur photo des données
var JOURNAL = [];          // journal des évènements de version
var VTYPE = 'gam';         // écran Révisions : type affiché
var VCLE = '';             // écran Révisions : clé affichée
var R = null;              // rapport courant
var VIEW = 'rapport';
var PSET = 'gamserie';     // onglet paramètres actif
var TSVUE = '';            // tableau TS : '' | 'trous'
var FGAM = {}, FTS = {};   // filtres saisis dans les en-têtes de colonnes
var GARCH = false, TARCH = false;  // affichage des révisions archivées

function loadAll(){
  var d = get('rc.db');
  if(d){ try{ DB = JSON.parse(d); }catch(e){ DB = clone(SEED); } } else { DB = clone(SEED); saveDB(); }
  if(!DB.gamserie) DB.gamserie=[]; if(!DB.ts) DB.ts=[]; if(!DB.matrice) DB.matrice=[];
  var vs = get('rc.versions'); VERS = vs ? JSON.parse(vs) : {gam:{}, ts:{}};
  if(!VERS.gam) VERS.gam = {}; if(!VERS.ts) VERS.ts = {};
  var hs = get('rc.revisions'); HIST = hs ? JSON.parse(hs) : [];
  var jl = get('rc.journal'); JOURNAL = jl ? JSON.parse(jl) : [];
  var dc = get('rc.docs'); DOCS = dc ? JSON.parse(dc) : [];
  var r = get('rc.reports'); REPORTS = r ? JSON.parse(r) : [];
  /* rapport en cours : propre à chaque poste (plusieurs PC partagent le classeur) */
  var dr = get(cleBrouillon());
  if(dr === null && get('rc.draft') !== null){ dr = get('rc.draft'); }
  delete MEM['rc.draft'];
  R = dr ? JSON.parse(dr) : newReport();
}
function saveDB(){ put('rc.db', JSON.stringify(DB)); }
function saveReports(){ put('rc.reports', JSON.stringify(REPORTS)); }
function saveDocs(){ put('rc.docs', JSON.stringify(DOCS)); return true; }
function saveDraft(){ R.majLe = new Date().toISOString(); put(cleBrouillon(), JSON.stringify(R)); }
/* Identifiant de ce poste (navigateur), mémorisé localement */
var POSTE = (function(){
  var p = null;
  try{ p = localStorage.getItem('ec.poste'); }catch(e){}
  if(!p){
    p = 'P' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    try{ localStorage.setItem('ec.poste', p); }catch(e){}
  }
  return p;
})();
function cleBrouillon(){ return 'rc.draft.' + POSTE; }
function clone(o){ return JSON.parse(JSON.stringify(o)); }

/* =======================================================================
   2. Utilitaires numériques  (accepte 12,5 et 12.5 comme dans le classeur)
   ======================================================================= */
var EPS = 1e-9;
function num(v){
  if(typeof v === 'number') return isFinite(v) ? v : NaN;
  if(v === null || v === undefined) return NaN;
  var s = String(v).replace(/\s/g,'').replace(/\u00a0/g,'').trim();
  if(!s) return NaN;
  if(s.indexOf(',')>-1 && s.indexOf('.')>-1) s = s.replace(/\./g,'').replace(',','.');
  else s = s.replace(',','.');
  var n = Number(s);
  return isFinite(n) ? n : NaN;
}
function fmt(n){
  if(typeof n !== 'number' || !isFinite(n)) return String(n === undefined || n === null ? '' : n);
  var r = Math.round(n*1000)/1000;
  return String(r).replace('.',',');
}
function esc(s){ return String(s===undefined||s===null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function uid(){ return 'R'+Date.now().toString(36)+Math.random().toString(36).slice(2,6); }
function todayISO(){ var d=new Date(); return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate()); }
function pad(n){ return (n<10?'0':'')+n; }
function frDate(iso){ if(!iso) return ''; var p=String(iso).split('-'); return p.length===3 ? p[2]+'/'+p[1]+'/'+p[0] : iso; }

/* =======================================================================
   3. Moteur : article -> famille -> spécifications -> valeurs attendues
   ======================================================================= */
/* Colonnes « références documentaires » de la matrice article / fournisseur.
   k = champ stocké sur la ligne de matrice ; lib = libellé écran ;
   csv = en-tête du fichier ; norm = en-tête normalisé (repérage à l'import). */
var REFCOLS = [
  {k:'norme',           lib:'Norme',                    csv:'Ref.Norme',                    norm:'refnorme'},
  {k:'planMethode',     lib:'Plan méthode',             csv:'Ref.Plan methode',             norm:'refplanmethode'},
  {k:'planControle',    lib:'Plan contrôle',            csv:'Ref.Plan controle',            norm:'refplancontrole'},
  {k:'planSituation',   lib:'Plan de situation',        csv:'Ref.Plan de Situation',        norm:'refplandesituation'},
  {k:'conditionnement', lib:'Conditionnement',          csv:'Ref.Conditionnement',          norm:'refconditionnement'},
  {k:'tauxEchantillon', lib:'Taux échantillon',         csv:'Ref.Taux echantillon',         norm:'reftauxechantillon'},
  {k:'defautheque',     lib:'Défauthèque',              csv:'Ref.Defautheque',              norm:'refdefautheque'}
];
/* Colonnes complètes de la matrice (identité + références). */
function matriceColonnes(){
  return [
    {k:'article',     lib:'Article'},
    {k:'fournisseur', lib:'Code fournisseur'},
    {k:'famille',     lib:'Famille'},
    {k:'matiere',     lib:'Matière'},
    {k:'liasse',      lib:'Liasse'}
  ].concat(REFCOLS.map(function(c){ return {k:c.k, lib:c.lib}; }));
}
function familleOf(article){
  for(var i=0;i<DB.matrice.length;i++) if(DB.matrice[i].article===article) return DB.matrice[i].famille;
  return '';
}
function matriceOf(article){
  for(var i=0;i<DB.matrice.length;i++) if(DB.matrice[i].article===article) return DB.matrice[i];
  return null;
}
/* Document PDF stocké dont le nom (sans extension) correspond à une référence. */
function docByRef(ref){
  var r = String(ref||'').trim().toLowerCase().replace(/\.pdf$/,'');
  if(!r || r === 'n/a') return null;
  for(var i=0;i<DOCS.length;i++){
    if(DOCS[i].nom.toLowerCase().replace(/\.pdf$/,'') === r) return DOCS[i];
  }
  return null;
}
/* Références documentaires d'un article : libellé, référence, et document stocké s'il existe. */
function refsArticle(article){
  var m = matriceOf(article); if(!m) return [];
  var out = [];
  REFCOLS.forEach(function(c){
    var ref = String(m[c.k]||'').trim();
    if(!ref || ref.toUpperCase() === 'N/A') return;
    out.push({lib:c.lib, ref:ref, doc:docByRef(ref)});
  });
  return out;
}
function tsOf(article){
  for(var i=0;i<DB.ts.length;i++) if(DB.ts[i].article===article) return DB.ts[i];
  return null;
}
function articles(){
  var s = {}, out = [];
  DB.matrice.forEach(function(m){ s[m.article]=1; });
  DB.ts.forEach(function(t){ s[t.article]=1; });
  for(var a in s) out.push(a);
  return out.sort();
}
function familles(){
  var s = {}, out = [];
  DB.gamserie.forEach(function(g){ s[g.famille]=1; });
  DB.matrice.forEach(function(m){ if(m.famille) s[m.famille]=1; });
  for(var f in s) out.push(f);
  return out.sort();
}

/* Construit les lignes du rapport : GAMSERIE (spécifications de la famille)
   + TS (nominal / tol - / tol +)  ->  Mini = nom+tol-, Maxi = nom+tol+     */
function buildLignes(article){
  var fam = familleOf(article);
  var ts = tsOf(article);
  var out = [];
  DB.gamserie.forEach(function(g){
    if(g.famille !== fam) return;
    var v = (ts && ts.vals && ts.vals[g.code]) ? ts.vals[g.code] : null;
    var brut = v ? v[0] : null;
    var nom = num(brut), tm = num(v?v[1]:null), tp = num(v?v[2]:null);
    if(isNaN(tm)) tm = 0;
    if(isNaN(tp)) tp = 0;
    var na = !v || String(brut===null||brut===undefined?'':brut).trim()==='-' || isNaN(nom);
    var l = {
      code:g.code, specification:g.specification, criticite:g.criticite||'',
      outil:g.moyen||'', methode:g.methode||'', na:na,
      typeSpec: v ? String(v[3]===undefined||v[3]===null?'':v[3]).trim() : ''  // Générale / Individuelle / (vide)
    };
    if(na){ l.mini=null; l.maxi=null; l.nom=null; l.tolM=null; l.tolP=null; l.testable=false; l.qual=false; }
    else{
      l.nom=nom; l.tolM=tm; l.tolP=tp; l.mini=nom+tm; l.maxi=nom+tp; l.testable=true;
      l.qual = (nom===1 && tm===0 && tp===0);   // point qualitatif : 1 = OK, 0 = NOK
    }
    out.push(l);
  });
  return out;
}
/* Portée d'une spécification dans le rapport, d'après la colonne TYPE SPEC du TS :
   'generale' = relevé une seule fois (OK/NOK) au-dessus du tableau ;
   'individuelle' = relevé pièce par pièce dans le tableau ;
   '' = non déroulée dans le rapport. */
function specKind(l){
  var t = String(l && l.typeSpec || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
  if(t === 'generale') return 'generale';
  if(t === 'individuelle') return 'individuelle';
  return '';
}
function aIndividuelles(){ return R.lignes.some(function(l){ return specKind(l) === 'individuelle' && l.testable; }); }

/* =======================================================================
   Statut du rapport, gel (lecture seule) et non-conformités
   ======================================================================= */
var ORIGINES = ['Contrôle d\'entrée', 'Produit « déjà entré dans l\'usine »'];
var STATUTS  = ['', 'Attente décision méthodes', 'Accepté', 'Refusé'];
/* Flux curatif des NC (PO12 / PO70) : type -> sous-types */
var TYPES_NC = [
  {t:'Type 1 — Réception documentaire', sous:['CCPU manquant','CCPU papier abîmé']},
  {t:'Type 2 — Réception', sous:['Quantité manquante','Livraison de la mauvaise référence','Mauvais conditionnement','Conditionnement abîmé','Mauvaise identification (carton, étiquette, etc.)','Identification non lisible']},
  {t:'Type 3 — Contrôle', sous:['Plan non respecté (dimension, géométrie, état de surface, marquage produit)','Aspect visuel (chocs, rayures, criques, etc.)','CCPU erroné (manque information, mauvaise information)']}
];
function libDecision(s){ return s || 'Sans décision'; }
/* '' = modifiable · 'attente' = figé sauf le statut · 'fige' = figé, réouverture méthodiste */
function figeNiveau(){
  if(R.decision === 'Accepté' || R.decision === 'Refusé') return 'fige';
  if(R.decision === 'Attente décision méthodes') return 'attente';
  return '';
}
function rapportRO(){ return figeNiveau() !== ''; }          // champs de saisie non modifiables
function statutVerrouille(){ return figeNiveau() === 'fige'; } // le statut lui-même est verrouillé
/* Nombre de cases de mesure encore vides (individuelles + générales). */
function casesVides(){
  var n = nbPieces(), v = 0;
  R.lignes.forEach(function(l,i){
    if(specKind(l) === 'individuelle' && l.testable){
      for(var j=0;j<n;j++){ var x = (R.mesures[i]||[])[j]; if(x === '' || x === undefined || x === null) v++; }
    }
  });
  R.lignes.forEach(function(l){
    if(specKind(l) === 'generale'){ if(String((R.generales||{})[l.code]||'') === '') v++; }
  });
  return v;
}

/* =======================================================================
   Cartes de contrôle type Minitab : valeurs individuelles (I) + étendues
   mobiles (MR). Rendu en SVG en ligne (aucune librairie externe).
   ======================================================================= */
function statsIMR(vals){
  var n = vals.length, mean = 0, i;
  for(i=0;i<n;i++) mean += vals[i]; mean = n ? mean/n : 0;
  var mr = [], s = 0;
  for(i=1;i<n;i++){ var d = Math.abs(vals[i]-vals[i-1]); mr.push(d); s += d; }
  var mrbar = mr.length ? s/mr.length : 0;
  return {n:n, mean:mean, mr:mr, mrbar:mrbar,
    iUCL:mean+2.66*mrbar, iLCL:mean-2.66*mrbar,
    mrUCL:3.267*mrbar, mrLCL:0};
}
/* Un graphique : points reliés, ligne centrale + limites, lignes spéc. option. */
function svgChart(pts, cl, ucl, lcl, extra, titre){
  var W = 760, H = 188, ml = 74, mr = 92, mt = 24, mb = 22;
  var iw = W-ml-mr, ih = H-mt-mb;
  var ys = [cl, ucl, lcl];
  pts.forEach(function(p){ ys.push(p); });
  (extra||[]).forEach(function(e){ ys.push(e.y); });
  ys = ys.filter(function(v){ return typeof v==='number' && isFinite(v); });
  var lo = Math.min.apply(null, ys), hi = Math.max.apply(null, ys);
  if(lo === hi){ lo -= 1; hi += 1; }
  var pad = (hi-lo)*0.12; lo -= pad; hi += pad;
  function X(i){ return ml + (pts.length<=1 ? iw/2 : iw*i/(pts.length-1)); }
  function Y(v){ return mt + ih*(hi-v)/(hi-lo); }
  var s = '<svg viewBox="0 0 '+W+' '+H+'" class="ctlchart" preserveAspectRatio="xMidYMid meet">';
  s += '<rect class="cc-plot" x="'+ml+'" y="'+mt+'" width="'+iw+'" height="'+ih+'"/>';
  /* cls : cc-spec (limites de spécification), cc-nom, cc-lim (limites de contrôle), cc-moy */
  function hline(v, cls, dash, lab, side){
    if(typeof v!=='number' || !isFinite(v)) return '';
    var y = Y(v);
    var lx = side==='L' ? (ml-4) : (ml+iw+4), anc = side==='L' ? 'end' : 'start';
    return '<line class="'+cls+'" x1="'+ml+'" y1="'+y.toFixed(1)+'" x2="'+(ml+iw)+'" y2="'+y.toFixed(1)+'"'+(dash?' stroke-dasharray="'+dash+'"':'')+' stroke-width="1"/>'
      + '<text class="'+cls+'" x="'+lx+'" y="'+(y+3).toFixed(1)+'" font-size="10" text-anchor="'+anc+'">'+esc(lab)+'</text>';
  }
  (extra||[]).forEach(function(e){ s += hline(e.y, e.cls||'cc-spec', '4 3', e.label, 'L'); });
  s += hline(ucl, 'cc-lim', '5 4', 'LSC '+fmt(ucl), 'R');
  s += hline(lcl, 'cc-lim', '5 4', 'LIC '+fmt(lcl), 'R');
  s += hline(cl, 'cc-moy', '', 'x̄ '+fmt(cl), 'R');
  s += '<text class="cc-titre" x="'+ml+'" y="14" font-size="11" font-weight="600">'+esc(titre)+'</text>';
  s += '<text class="cc-axe" x="'+(ml-6)+'" y="'+(Y(hi)+8).toFixed(1)+'" font-size="9" text-anchor="end">'+fmt(hi)+'</text>';
  s += '<text class="cc-axe" x="'+(ml-6)+'" y="'+(Y(lo)).toFixed(1)+'" font-size="9" text-anchor="end">'+fmt(lo)+'</text>';
  var d = '';
  pts.forEach(function(p,i){ d += (i?'L':'M')+X(i).toFixed(1)+' '+Y(p).toFixed(1)+' '; });
  if(pts.length>1) s += '<path class="cc-trace" d="'+d+'" fill="none" stroke-width="1.4"/>';
  pts.forEach(function(p,i){
    var out = (typeof ucl==='number' && p>ucl+1e-9) || (typeof lcl==='number' && p<lcl-1e-9);
    s += '<circle class="'+(out?'cc-pt-hors':'cc-pt')+'" cx="'+X(i).toFixed(1)+'" cy="'+Y(p).toFixed(1)+'" r="'+(out?3.6:2.8)+'"/>';
  });
  s += '</svg>';
  return s;
}
/* Bloc I-MR complet pour une série de valeurs (avec bornes spéc. facultatives). */
function blocIMR(vals, spec, titre){
  if(!vals || !vals.length) return '<div class="doc-empty">Aucune mesure numérique à tracer.</div>';
  var st = statsIMR(vals);
  var extra = [];
  if(spec){
    if(typeof spec.maxi==='number' && isFinite(spec.maxi)) extra.push({y:spec.maxi, cls:'cc-spec', label:'Maxi '+fmt(spec.maxi)});
    if(typeof spec.mini==='number' && isFinite(spec.mini)) extra.push({y:spec.mini, cls:'cc-spec', label:'Mini '+fmt(spec.mini)});
    if(typeof spec.nom==='number' && isFinite(spec.nom))  extra.push({y:spec.nom,  cls:'cc-nom', label:'Nom '+fmt(spec.nom)});
  }
  var h = '<div class="chartwrap">';
  h += svgChart(vals, st.mean, st.iUCL, st.iLCL, extra, (titre||'')+' — valeurs individuelles (I)');
  h += svgChart(st.mr, st.mrbar, st.mrUCL, st.mrLCL, null, 'Étendues mobiles (MR)');
  h += '<div class="chartinfo">n = '+st.n+' · moyenne = <b>'+fmt(st.mean)+'</b> · MR moyen = <b>'+fmt(st.mrbar)+'</b> · LSC(I) = '+fmt(st.iUCL)+' · LIC(I) = '+fmt(st.iLCL)+'</div>';
  h += '</div>';
  return h;
}
/* Série de valeurs mesurées d'un rapport pour une côte (code) ou toutes. */
function serieRapportCote(rep, cote){
  var out = [], n = nbPieces2(rep);
  (rep.lignes||[]).forEach(function(l,i){
    if(specKind(l) !== 'individuelle' || !l.testable || l.qual) return;   // valeurs numériques uniquement
    if(cote && l.code !== cote) return;
    for(var j=0;j<n;j++){
      var v = num((rep.mesures[i]||[])[j]);
      if(!isNaN(v)) out.push(v);
    }
  });
  return out;
}
/* Contenu (SVG) de la carte de contrôle du rapport courant, selon la côte choisie. */
function chartRapportInner(){
  var cote = R.chartCote || '', spec = null, titre = 'Toutes côtes';
  if(cote && cote !== '__all'){
    var lg = null; R.lignes.forEach(function(l){ if(l.code === cote) lg = l; });
    if(lg){ spec = {mini:lg.mini, maxi:lg.maxi, nom:lg.nom}; titre = 'Côte ' + cote; }
  }
  return blocIMR(serieRapportCote(R, cote === '__all' ? '' : cote), spec, titre);
}
/* Fenêtre superposée (modale) présentant les deux cartes de contrôle. */
function ouvrirCarteControle(){
  var old = document.querySelector('.cartemodal'); if(old) old.remove();
  var cotes = [];
  R.lignes.forEach(function(l){ if(specKind(l)==='individuelle' && l.testable && !l.qual) cotes.push(l.code); });
  if(R.chartCote !== '__all' && cotes.indexOf(R.chartCote) < 0) R.chartCote = cotes[0] || '__all';
  var m = document.createElement('div');
  m.className = 'modal cartemodal';
  m.innerHTML = '<div class="box cartebox" role="dialog" aria-modal="true" aria-label="Carte de contrôle">'
    + '<div class="cmh"><b>Carte de contrôle — suivi des déviances</b>'
    + '<select id="f-chartcote">'
    + '<option value="__all"'+(R.chartCote==='__all'?' selected':'')+'>Toutes les mesures</option>'
    + cotes.map(function(c){ return '<option value="'+esc(c)+'"'+(R.chartCote===c?' selected':'')+'>Côte '+esc(c)+'</option>'; }).join('')
    + '</select>'
    + '<button class="close" id="cm-close" title="Fermer et revenir au rapport" aria-label="Fermer">'
    + '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M5 5 L19 19 M19 5 L5 19" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" fill="none"/></svg>'
    + '</button></div>'
    + '<div class="cmb" id="chart-imr">'+chartRapportInner()+'</div></div>';
  document.body.appendChild(m);
  function close(){ m.remove(); document.removeEventListener('keydown', k); }
  function k(e){ if(e.key === 'Escape') close(); }
  document.addEventListener('keydown', k);
  m.addEventListener('click', function(e){ if(e.target === m) close(); });
  m.querySelector('#cm-close').onclick = close;
  m.querySelector('#f-chartcote').onchange = function(){
    R.chartCote = this.value; saveDraft();
    var box = m.querySelector('#chart-imr'); if(box) box.innerHTML = chartRapportInner();
  };
}

function newReport(){
  return {
    id:uid(), creeLe:new Date().toISOString(), majLe:null, statut:'brouillon',
    article:'', lot:'', commande:'', fournisseur:'', controleur:'', date:todayISO(),
    qteCommandee:'', qteReceptionnee:'', qteEchantillon:15, dureeControle:'', decision:'',
    origine:'', typeNC:'', sousTypeNC:'',
    commentaires:'',
    famille:'', gamRev:null, gamStatut:'', gamDateEffet:'', tsRev:null, tsStatut:'', tsDateEffet:'', indiceDef:'',
    lignes:[], mesures:[], generales:{}, moyens:{}, chartCote:''
  };
}

/* Recharge les spécifications depuis les bases en conservant les mesures saisies */
function syncLignes(){
  if(!R.generales) R.generales = {};
  var anciennes = R.lignes || [], anciennesMes = R.mesures || [];
  var index = {};
  anciennes.forEach(function(l,i){ index[l.code] = anciennesMes[i] || []; });
  R.lignes = R.article ? buildLignes(R.article) : [];
  if(R.article){
    var fam = familleOf(R.article), vg = version('gam', fam), vt = version('ts', R.article);
    R.famille = fam;
    R.gamRev = vg ? vg.rev : null;
    R.gamStatut = vg ? vg.statut : 'sans révision';
    R.gamDateEffet = vg ? vg.dateEffet : '';
    R.tsRev = vt ? vt.rev : null;
    R.tsStatut = vt ? vt.statut : 'sans révision';
    R.tsDateEffet = vt ? vt.dateEffet : '';
    R.indiceDef = vt ? (vt.indiceDef || '') : '';
  }
  var n = nbPieces();
  R.mesures = R.lignes.map(function(l){
    var src = index[l.code] || [];
    var row = [];
    for(var j=0;j<n;j++) row.push(src[j] !== undefined ? src[j] : '');
    return row;
  });
}
function nbPieces(){
  var n = num(R.qteEchantillon);
  if(isNaN(n) || n < 1) n = 15;
  return Math.min(Math.floor(n), 60);
}
function resizeMesures(){
  var n = nbPieces();
  R.mesures = R.lignes.map(function(l,i){
    var row = (R.mesures[i] || []).slice(0,n);
    while(row.length < n) row.push('');
    return row;
  });
}

/* --- calculs de conformité : règle identique au classeur -----------------
   Conforme    = pièce avec au moins une mesure ET aucune mesure hors tol.
   Non conforme= pièce avec au moins une mesure hors tolérance
   Les lignes non testables (« - ») sont exclues du calcul.               */
function pieceStatut(j){
  var saisies = 0, hors = 0;
  for(var i=0;i<R.lignes.length;i++){
    var l = R.lignes[i];
    if(!l.testable || specKind(l) !== 'individuelle') continue;  // seules les spéc. individuelles se relèvent pièce par pièce
    var v = num(R.mesures[i] ? R.mesures[i][j] : '');
    if(isNaN(v)) continue;
    saisies++;
    if(v < l.mini - EPS || v > l.maxi + EPS) hors++;
  }
  if(hors > 0) return 'nc';
  if(saisies > 0) return 'ok';
  return '';
}
function totaux(){
  var t = {conforme:0, nonConforme:0, encours:0, n:nbPieces()};
  if(!aIndividuelles()) return t;   // aucun point individuel à relever : pas de décompte par pièce
  for(var j=0;j<t.n;j++){
    var s = pieceStatut(j);
    if(s === 'ok') t.conforme++; else if(s === 'nc') t.nonConforme++; else t.encours++;
  }
  return t;
}
/* Décompte des spécifications générales : OK / NOK / non renseignées. */
function totauxGen(){
  var g = {ok:0, nc:0, reste:0, total:0};
  R.lignes.forEach(function(l){
    if(specKind(l) !== 'generale') return;
    g.total++;
    var v = String((R.generales||{})[l.code]||'');
    if(v === '1') g.ok++; else if(v === '0') g.nc++; else g.reste++;
  });
  return g;
}
function mesureStatut(i,j){
  var l = R.lignes[i];
  if(!l || !l.testable) return '';
  var v = num(R.mesures[i][j]);
  if(isNaN(v)) return '';
  return (v < l.mini - EPS || v > l.maxi + EPS) ? 'out' : 'in';
}
/* Nombre de spécifications individuelles hors tolérance (ou NOK) sur une pièce */
function ncParPiece(j){
  var c = 0;
  R.lignes.forEach(function(l,i){ if(specKind(l) === 'individuelle' && mesureStatut(i,j) === 'out') c++; });
  return c;
}
/* Saisie validée : le séparateur décimal est toujours la virgule */
function virgule(v){ return String(v === null || v === undefined ? '' : v).replace(/\./g, ','); }
/* Valeur mesurée affichée avec virgule et au moins deux décimales (13 -> 13,00 ; 13,5 -> 13,50).
   Les décimales supplémentaires saisies sont conservées (13,025 reste 13,025). */
function formatMesure(v){
  var s = virgule(v).trim();
  if(/^[-+]?,\d+$/.test(s)) s = s.replace(',', '0,');
  var m = /^([-+]?\d+)(?:,(\d*))?$/.exec(s);
  if(!m) return s;                                  // texte non numérique : laissé tel quel
  var dec = m[2] || '';
  while(dec.length < 2) dec += '0';
  return m[1] + ',' + dec;
}
function ncParLigne(i){
  var c = 0, n = nbPieces();
  for(var j=0;j<n;j++) if(mesureStatut(i,j) === 'out') c++;
  return c;
}

/* =======================================================================
   3 ter. Révisions — indice 0, 1, 2, … par famille (GAMSERIE) et par article (TS)
   Règle : une révision applicable est figée. Toute modification ouvre la
   révision suivante à l'état brouillon, qui doit être validée pour servir.
   ======================================================================= */
function saveVersions(){ put('rc.versions', JSON.stringify(VERS)); }
function saveHist(){ put('rc.revisions', JSON.stringify(HIST)); }
function saveJournal(){ put('rc.journal', JSON.stringify(JOURNAL)); }

function journaliser(type, cle, rev, action, detail){
  JOURNAL.push({le:new Date().toISOString(), type:type, cle:cle, rev:rev, action:action, detail:detail||''});
  if(JOURNAL.length > 2000) JOURNAL = JOURNAL.slice(-2000);
  saveJournal();
}
function photoGam(fam){
  return DB.gamserie.filter(function(g){ return g.famille === fam; }).map(function(g){
    return {code:g.code, criticite:g.criticite, specification:g.specification, moyen:g.moyen, methode:g.methode};
  });
}
function photoTS(article){
  var t = tsOf(article);
  return t ? clone(t.vals) : {};
}
function photo(type, cle){ return type === 'gam' ? photoGam(cle) : photoTS(cle); }
function version(type, cle){ return VERS[type][cle] || null; }
function libRev(v){ return v ? ('rév. ' + v.rev + (v.statut === 'brouillon' ? ' (brouillon)' : '')) : 'sans révision'; }

/* Toutes les familles et tous les articles existants sont initialisés en révision 0 */
function initVersions(){
  var n = 0, d = todayISO();
  familles().forEach(function(f){
    if(VERS.gam[f]) return;
    VERS.gam[f] = {rev:0, statut:'applicable', dateEffet:d, valideePar:'Reprise initiale', motif:'Version initiale', creeLe:new Date().toISOString()};
    figer('gam', f, 'Version initiale'); n++;
  });
  DB.ts.forEach(function(t){
    if(VERS.ts[t.article]) return;
    VERS.ts[t.article] = {rev:0, statut:'applicable', dateEffet:d, valideePar:'Reprise initiale', motif:'Version initiale', indiceDef:'', creeLe:new Date().toISOString()};
    figer('ts', t.article, 'Version initiale'); n++;
  });
  if(n){ saveVersions(); saveHist(); }
  return n;
}
/* Enregistre la photo des données dans l'historique des révisions */
function figer(type, cle, motif){
  var v = VERS[type][cle]; if(!v) return;
  var e = {type:type, cle:cle, rev:v.rev, dateEffet:v.dateEffet, valideePar:v.valideePar, motif:motif || v.motif || '',
           indiceDef:v.indiceDef || '', figeeLe:new Date().toISOString(), photo:photo(type, cle)};
  var f = -1;
  for(var i=0;i<HIST.length;i++) if(HIST[i].type===type && HIST[i].cle===cle && HIST[i].rev===v.rev){ f=i; break; }
  if(f>-1) HIST[f] = e; else HIST.push(e);
  saveHist();
  journaliser(type, cle, v.rev, 'Révision validée', e.motif);
}
function revisionsDe(type, cle){
  return HIST.filter(function(h){ return h.type===type && h.cle===cle; }).sort(function(a,b){ return b.rev - a.rev; });
}
function photoRev(type, cle, rev){
  var l = revisionsDe(type, cle).filter(function(h){ return h.rev === rev; });
  return l.length ? l[0].photo : null;
}
/* Une révision applicable a-t-elle été modifiée depuis son figeage ? */
function ecartPhoto(type, cle){
  var v = version(type, cle); if(!v) return false;
  var ref = photoRev(type, cle, v.rev);
  return ref ? JSON.stringify(canon(ref)) !== JSON.stringify(canon(photo(type, cle))) : false;
}
/* Forme normalisée pour comparer deux photos : clés triées, vides = '',
   cases vides de fin de tableau ignorées (ex. TYPE SPEC absent). */
function canon(o){
  if(o === undefined || o === null) return '';
  if(Array.isArray(o)){
    var a = o.map(canon);
    while(a.length && a[a.length-1] === '') a.pop();
    return a;
  }
  if(typeof o === 'object'){
    var r = {};
    Object.keys(o).sort().forEach(function(k){ r[k] = canon(o[k]); });
    return r;
  }
  return o;
}
function nouvelleRevision(type, cle){
  var v = version(type, cle);
  if(!v){
    VERS[type][cle] = {rev:0, statut:'brouillon', dateEffet:'', valideePar:'', motif:'', indiceDef:'', creeLe:new Date().toISOString()};
  } else if(v.statut === 'applicable'){
    VERS[type][cle] = {rev:v.rev + 1, statut:'brouillon', dateEffet:'', valideePar:'', motif:'',
                       indiceDef:v.indiceDef || '', creeLe:new Date().toISOString()};
  }
  saveVersions();
  journaliser(type, cle, VERS[type][cle].rev, 'Brouillon ouvert', '');
  return VERS[type][cle];
}
function validerRevision(type, cle){
  var v = version(type, cle);
  if(!v || v.statut !== 'brouillon'){ toast('Aucun brouillon à valider.'); return; }
  var extra = (type === 'ts')
    ? '<div style="margin-top:10px"><label>Indice de définition (plan, révision client)<br><input id="v-def" value="'+esc(v.indiceDef||'')+'" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="ex. plan rév. D"></label></div>'
    : '';
  modal('Valider la révision ' + v.rev + ' — ' + cle,
    '<label>Validée par<br><input id="v-par" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="Nom du valideur"></label>'
    + '<div style="margin-top:10px"><label>Date d\'effet<br><input type="date" id="v-date" value="'+todayISO()+'" style="width:100%;padding:6px;border:1px solid var(--rule)"></label></div>'
    + '<div style="margin-top:10px"><label>Motif de la révision<br><input id="v-motif" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="ex. mise à jour plan client"></label></div>'
    + extra
    + '<p style="margin:12px 0 0">Une fois validée, cette révision est figée : toute modification ultérieure ouvrira la révision ' + (v.rev+1) + '.</p>',
    function(){}, 'Annuler');
  var ft = document.querySelector('.modal .ft');
  ft.innerHTML = '<button class="btn" data-x2>Annuler</button><button class="btn pri" data-ok2>Valider la révision '+v.rev+'</button>';
  ft.querySelector('[data-x2]').onclick = function(){ document.querySelector('.modal').remove(); };
  ft.querySelector('[data-ok2]').onclick = function(){
    var par = document.getElementById('v-par').value.trim();
    var mot = document.getElementById('v-motif').value.trim();
    var dt = document.getElementById('v-date').value;
    if(!par || !mot){ toast('Valideur et motif sont obligatoires.'); return; }
    var anc = v.rev - 1;
    v.valideePar = par; v.motif = mot; v.dateEffet = dt; v.statut = 'applicable';
    var def = document.getElementById('v-def');
    if(def) v.indiceDef = def.value.trim();
    saveVersions(); figer(type, cle, mot);
    if(anc >= 0) journaliser(type, cle, anc, 'Révision périmée', 'remplacée par la rév. ' + v.rev);
    document.querySelector('.modal').remove();
    render();
    toast(cle + ' — révision ' + v.rev + ' applicable depuis le ' + frDate(dt));
  };
}
function abandonnerBrouillon(type, cle){
  var v = version(type, cle);
  if(!v || v.statut !== 'brouillon') return;
  var prec = v.rev - 1, ref = prec >= 0 ? photoRev(type, cle, prec) : null;
  if(!confirm('Abandonner le brouillon de la révision ' + v.rev + ' ? ' + (ref ? 'Les données reviendront à la révision ' + prec + '.' : 'Les données saisies seront conservées telles quelles.'))) return;
  if(ref){
    if(type === 'gam'){
      DB.gamserie = DB.gamserie.filter(function(g){ return g.famille !== cle; })
        .concat(ref.map(function(l){ var o = clone(l); o.famille = cle; return o; }));
    } else {
      var t = tsOf(cle); if(t) t.vals = clone(ref);
    }
    saveDB();
    VERS[type][cle] = {rev:prec, statut:'applicable', dateEffet:revisionsDe(type,cle)[0] ? null : null, creeLe:new Date().toISOString()};
    var h = revisionsDe(type, cle).filter(function(x){ return x.rev === prec; })[0];
    if(h){ VERS[type][cle].dateEffet = h.dateEffet; VERS[type][cle].valideePar = h.valideePar; VERS[type][cle].motif = h.motif; VERS[type][cle].indiceDef = h.indiceDef; }
    HIST = HIST.filter(function(x){ return !(x.type===type && x.cle===cle && x.rev===v.rev); });
    saveHist();
  } else {
    VERS[type][cle].statut = 'applicable';
  }
  saveVersions(); journaliser(type, cle, v.rev, 'Brouillon abandonné', '');
  syncLignes(); saveDraft(); render();
  toast('Brouillon abandonné.');
}
function modifiable(type, cle){
  var v = version(type, cle);
  return !!v && v.statut === 'brouillon';
}
/* Après un import CSV : ouvre un brouillon partout où les données ont changé */
function revisionsApresImport(type, source){
  var cles = type === 'gam' ? familles() : DB.ts.map(function(t){ return t.article; });
  var n = 0;
  cles.forEach(function(c){
    var v = version(type, c);
    if(!v){
      VERS[type][c] = {rev:0, statut:'brouillon', dateEffet:'', valideePar:'', motif:source, indiceDef:'', creeLe:new Date().toISOString()};
      journaliser(type, c, 0, 'Créée par import', source); n++; return;
    }
    if(v.statut === 'brouillon') return;
    if(ecartPhoto(type, c)){
      nouvelleRevision(type, c);
      VERS[type][c].motif = source;
      journaliser(type, c, VERS[type][c].rev, 'Modifiée par import', source); n++;
    }
  });
  saveVersions();
  return n;
}
/* Comparaison de deux révisions */
function comparer(type, cle, ra, rb){
  var a = photoRev(type, cle, ra) || (version(type,cle) && version(type,cle).rev === ra ? photo(type,cle) : null);
  var b = photoRev(type, cle, rb) || (version(type,cle) && version(type,cle).rev === rb ? photo(type,cle) : null);
  var out = [];
  if(!a || !b) return out;
  if(type === 'gam'){
    var ma = {}, mb = {};
    a.forEach(function(l){ ma[l.code] = l; }); b.forEach(function(l){ mb[l.code] = l; });
    var codes = {}; a.forEach(function(l){ codes[l.code]=1; }); b.forEach(function(l){ codes[l.code]=1; });
    Object.keys(codes).sort().forEach(function(c){
      if(!ma[c]) out.push({code:c, type:'ajout', detail:'ajoutée : ' + (mb[c].specification||'')});
      else if(!mb[c]) out.push({code:c, type:'retrait', detail:'supprimée : ' + (ma[c].specification||'')});
      else {
        ['specification','criticite','moyen','methode'].forEach(function(f){
          if(String(ma[c][f]||'') !== String(mb[c][f]||''))
            out.push({code:c, type:'modif', detail:libChamp(f) + ' : « ' + (ma[c][f]||'—') + ' » → « ' + (mb[c][f]||'—') + ' »'});
        });
      }
    });
  } else {
    var codes2 = {};
    for(var k in a) codes2[k]=1; for(var k2 in b) codes2[k2]=1;
    Object.keys(codes2).sort().forEach(function(c){
      var va = a[c], vb = b[c];
      if(!va) out.push({code:c, type:'ajout', detail:'valeurs ajoutées : ' + trio(vb)});
      else if(!vb) out.push({code:c, type:'retrait', detail:'valeurs supprimées : ' + trio(va)});
      else if(trio(va) !== trio(vb)) out.push({code:c, type:'modif', detail:trio(va) + ' → ' + trio(vb)});
    });
  }
  return out;
}
function trio(v){
  if(!v) return '—';
  var ts = String(v[3]===undefined||v[3]===null?'':v[3]).trim();
  return 'nom ' + (v[0]===''||v[0]===null||v[0]===undefined?'—':fmt(num(v[0])||v[0]))
    + ' / tol ' + (v[1]===''?'—':fmt(num(v[1]))) + ' ' + (v[2]===''?'—':fmt(num(v[2])))
    + (ts ? ' / type ' + ts : '');
}
function libChamp(f){
  return {specification:'Spécification', criticite:'Criticité', moyen:'Moyen de contrôle', methode:'Méthode de contrôle'}[f] || f;
}


function csvEscape(v){
  var s = (v===null||v===undefined) ? '' : (typeof v==='number' ? fmt(v) : String(v));
  if(/[";\r\n]/.test(s)) s = '"' + s.replace(/"/g,'""') + '"';
  return s;
}
function toCSV(rows){
  return rows.map(function(r){ return r.map(csvEscape).join(';'); }).join('\r\n');
}
function download(name, text, mime){
  var blob = new Blob(['\ufeff' + text], {type:(mime||'text/csv')+';charset=utf-8'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
function parseCSV(text){
  text = text.replace(/^\ufeff/,'');
  var first = text.split(/\r?\n/)[0] || '';
  var sep = ';';
  var cs = (first.match(/;/g)||[]).length, cc = (first.match(/,/g)||[]).length, ct = (first.match(/\t/g)||[]).length;
  if(ct > cs && ct > cc) sep = '\t'; else if(cc > cs) sep = ',';
  var rows = [], row = [], cur = '', q = false, i = 0;
  while(i < text.length){
    var ch = text[i];
    if(q){
      if(ch === '"'){ if(text[i+1] === '"'){ cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else {
      if(ch === '"') q = true;
      else if(ch === sep){ row.push(cur); cur = ''; }
      else if(ch === '\n'){ row.push(cur); rows.push(row); row = []; cur = ''; }
      else if(ch === '\r'){ /* ignore */ }
      else cur += ch;
    }
    i++;
  }
  if(cur !== '' || row.length){ row.push(cur); rows.push(row); }
  return rows.filter(function(r){ return r.join('').trim() !== ''; });
}

function exportGamserie(){ download('GAMSERIE.csv', toCSV(tableauGamserie())); }
/* Tableaux à plat réutilisés par les exports CSV et par le classeur Excel */
function tableauGamserie(){
  var rows = [['Révision','Date de révision','Statut','Famille Méthode','Code','Criticité','Spécification','Moyen de contrôle','Méthode de contrôle']];
  lignesGam(false).forEach(function(l){
    rows.push([l.rev===null?'':l.rev, l.dateRev?frDate(l.dateRev):'', statutLib(l.statut),
      l.famille, l.code, l.criticite, l.specification, l.moyen, l.methode]);
  });
  return rows;
}
function exportMatrice(){ download('MATRICE_article_fournisseur.csv', toCSV(tableauMatrice())); }
function tableauMatrice(){
  var head = ['Article','Code fournisseur','Famille','Matière','Liasse documentaire'];
  REFCOLS.forEach(function(c){ head.push(c.csv); });
  var rows = [head];
  DB.matrice.forEach(function(m){
    var r = [m.article||'', m.fournisseur||'', m.famille||'', m.matiere||'', m.liasse||''];
    REFCOLS.forEach(function(c){ r.push(m[c.k]||''); });
    rows.push(r);
  });
  return rows;
}
function codesUtilises(){
  var s = {};
  DB.gamserie.forEach(function(g){ if(g.code) s[String(g.code).trim()] = 1; });
  DB.ts.forEach(function(t){ for(var c in t.vals) s[c] = 1; });
  var out = []; for(var c in s) out.push(c);
  return out.sort();
}
/* Export à plat : une ligne par couple article x point de contrôle */
function exportTS(){ download('TS.csv', toCSV(tableauTS())); }
function tableauTS(){
  var rows = [['Révision','Date de révision','Statut','Indice de définition','Article','Famille','Code',
               'Spécification','Type spéc.','Type saisie','Nominal','Tol -','Tol +','Mini.','Maxi.']];
  lignesTS().forEach(function(l){
    rows.push([l.rev===null?'':l.rev, l.dateRev?frDate(l.dateRev):'', l.statut, l.indiceDef,
      l.article, l.famille, l.code, l.specification, l.typeSpec||'', l.type,
      l.brut[0], l.brut[1], l.brut[2],
      l.mini===null?'':l.mini, l.maxi===null?'':l.maxi]);
  });
  return rows;
}
/* Export au format large du prototype (Article ; A ; A min ; A max ; …) */
function exportTSLarge(){
  var codes = codesUtilises();
  var head = ['Article'];
  codes.forEach(function(c){ head.push(c, c+' min', c+' max'); });
  head.push('Révision','Statut','Indice de définition','Date d\'effet');
  var rows = [head];
  DB.ts.forEach(function(t){
    var r = [t.article];
    codes.forEach(function(c){
      var v = t.vals[c];
      r.push(v?v[0]:'', v?v[1]:'', v?v[2]:'');
    });
    var ver = version('ts', t.article);
    r.push(ver?ver.rev:'', ver?ver.statut:'', ver?(ver.indiceDef||''):'', ver&&ver.dateEffet?frDate(ver.dateEffet):'');
    rows.push(r);
  });
  download('TS_format-large.csv', toCSV(rows));
}

function importCSV(base, mode){
  pickFile(function(text, name){
    var rows = parseCSV(text);
    if(rows.length < 2){ toast('Fichier vide ou illisible : ' + name); return; }
    try{
      var nrev = importerTableau(base, rows, mode, 'Import ' + name);
      syncLignes(); saveDraft(); render();
      toast(nomBase(base) + ' importée (' + (mode === 'remplacer' ? 'remplacement' : 'mise à jour') + ') — ' + name
        + (nrev ? ' · ' + nrev + ' révision' + (nrev>1?'s':'') + ' à valider' : ''));
    }catch(e){ toast('Import impossible : ' + e.message); }
  });
}
/* Intègre un tableau (1re ligne = en-têtes) dans une base. Renvoie le nombre
   de révisions ouvertes en brouillon. Utilisé par l'import CSV et par la
   reprise des onglets GAMSERIE / TS / Matrice modifiés dans Excel. */
function importerTableau(base, rows, mode, source){
    var head = rows[0].map(function(h){ return String(h).trim(); });
    {
      if(base === 'gamserie'){
        var nrm = head.map(function(x){
          return String(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[.\s]/g,'');
        });
        var cl = function(){
          for(var z=0;z<arguments.length;z++){ var k = nrm.indexOf(arguments[z]); if(k > -1) return k; }
          return -1;
        };
        var iF = cl('famillemethode','famille'), iC = cl('code'), iCr = cl('criticite'),
            iS = cl('specification'), iM = cl('moyendecontrole','moyen'), iMe = cl('methodedecontrole','methode');
        if(iF < 0 || iC < 0){ iF = 0; iC = 1; iCr = 2; iS = 3; iM = 4; iMe = 5; }  // ancien format positionnel
        var g = [];
        for(var i=1;i<rows.length;i++){
          var r = rows[i];
          if(!t(r[iF]) && !t(r[iC])) continue;
          g.push({famille:t(r[iF]),code:t(r[iC]),criticite:iCr>-1?t(r[iCr]):'',specification:iS>-1?t(r[iS]):'',
                  moyen:iM>-1?t(r[iM]):'',methode:iMe>-1?t(r[iMe]):''});
        }
        DB.gamserie = (mode === 'remplacer') ? g : fusionGam(DB.gamserie, g);
      }
      else if(base === 'matrice'){
        var nmm = head.map(function(x){
          return String(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[.\s]/g,'');
        });
        var cm = function(){
          for(var z=0;z<arguments.length;z++){ var k = nmm.indexOf(arguments[z]); if(k > -1) return k; }
          return -1;
        };
        var iA = cm('article','reference','ref'), iFo = cm('codefournisseur','fournisseur'),
            iFa = cm('famille','famillemethode'), iMa = cm('matiere'), iLi = cm('liassedocumentaire','liasse');
        if(iA < 0) iA = 0;
        var refIdx = {};
        REFCOLS.forEach(function(c){ refIdx[c.k] = cm(c.norm); });
        var m = [];
        for(var i2=1;i2<rows.length;i2++){
          var r2 = rows[i2];
          if(!String(r2[iA]||'').trim()) continue;
          var row = {article:t(r2[iA]), fournisseur:iFo>-1?t(r2[iFo]):'', famille:iFa>-1?t(r2[iFa]):'',
                     matiere:iMa>-1?t(r2[iMa]):'', liasse:iLi>-1?t(r2[iLi]):''};
          REFCOLS.forEach(function(c){ var k = refIdx[c.k]; row[c.k] = k>-1 ? t(r2[k]) : ''; });
          m.push(row);
        }
        DB.matrice = (mode === 'remplacer') ? m : fusionCle(DB.matrice, m, 'article');
      }
      else if(base === 'ts'){
        var norm = head.map(function(x){
          return String(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[.\s]/g,'');
        });
        var col = function(){
          for(var z=0;z<arguments.length;z++){ var k = norm.indexOf(arguments[z]); if(k > -1) return k; }
          return -1;
        };
        var cArt = col('article','reference','ref'), cCode = col('code'), cNom = col('nominal','nom','valeur');
        var ts = [];
        if(cArt > -1 && cCode > -1 && cNom > -1){
          /* --- format à plat : une ligne par couple article x code --- */
          var cTm = col('tol-','tolmoins','tolerance-','tolmin'), cTp = col('tol+','tolplus','tolerance+','tolmax');
          var cTy = col('typespec','typespecification'), hasTy = cTy > -1;
          var cFam = col('famille','famillemethode');
          var map = {}, ordre = [];
          for(var i3=1;i3<rows.length;i3++){
            var r3 = rows[i3];
            var art = t(r3[cArt]), code = t(r3[cCode]).toUpperCase();
            if(!art || !code) continue;
            /* en remplacement, une ligne sans aucune valeur = point non renseigné */
            if(mode === 'remplacer' && t(r3[cNom]) === '' && (cTm < 0 || t(r3[cTm]) === '') && (cTp < 0 || t(r3[cTp]) === '') && (!hasTy || t(r3[cTy]) === '')) continue;
            if(!map[art]){ map[art] = {article:art, vals:{}}; ordre.push(art); }
            map[art].vals[code] = hasTy
              ? [conv(t(r3[cNom])), conv(cTm>-1?t(r3[cTm]):''), conv(cTp>-1?t(r3[cTp]):''), t(r3[cTy])]
              : [conv(t(r3[cNom])), conv(cTm>-1?t(r3[cTm]):''), conv(cTp>-1?t(r3[cTp]):'')];
            if(cFam > -1){
              var fm = t(r3[cFam]);
              if(fm && !familleOf(art)) DB.matrice.push({article:art, famille:fm});
            }
          }
          ordre.forEach(function(a){ ts.push(map[a]); });
          if(mode !== 'remplacer'){
            ts.forEach(function(n){
              var ex = tsOf(n.article);
              if(!ex){ DB.ts.push(n); return; }
              for(var c in n.vals){
                if(!hasTy && ex.vals[c] && ex.vals[c][3] !== undefined && ex.vals[c][3] !== '')
                  n.vals[c][3] = ex.vals[c][3];   // conserve le TYPE SPEC existant si absent du fichier
                ex.vals[c] = n.vals[c];
              }
            });
            ts = DB.ts;
          }
        } else {
          /* --- format large : en-tête = code seul, puis 2 colonnes tol -, tol + --- */
          var idx = [];
          head.forEach(function(h,k){ if(/^[A-Z]{1,3}$/.test(h)) idx.push([h,k]); });
          if(!idx.length){
            throw new Error('en-tête non reconnu : colonnes Article / Code / Nominal attendues, ou colonnes de code (A, B, C…).');
          }
          for(var i4=1;i4<rows.length;i4++){
            var r4 = rows[i4];
            if(!String(r4[0]||'').trim()) continue;
            var vals = {};
            idx.forEach(function(pp){
              var c = pp[0], k = pp[1];
              var a = t(r4[k]), b = t(r4[k+1]), cc = t(r4[k+2]);
              if(a === '' && b === '' && cc === '') return;
              vals[c] = [conv(a), conv(b), conv(cc)];
            });
            ts.push({article:t(r4[0]), vals:vals});
          }
        }
        DB.ts = (mode === 'remplacer') ? ts : fusionCle(DB.ts, ts, 'article');
      }
      saveDB();
      var nrev = 0;
      if(base === 'gamserie') nrev = revisionsApresImport('gam', source);
      else if(base === 'ts') nrev = revisionsApresImport('ts', source);
      return nrev;
    }
  function t(v){ return String(v===undefined||v===null?'':v).trim(); }
  function conv(v){ var n = num(v); return isNaN(n) ? v : n; }
}
function fusionGam(base, ajouts){
  var out = base.slice();
  ajouts.forEach(function(a){
    var f = -1;
    for(var i=0;i<out.length;i++) if(out[i].famille===a.famille && out[i].code===a.code){ f=i; break; }
    if(f>-1) out[f] = a; else out.push(a);
  });
  return out;
}
function fusionCle(base, ajouts, cle){
  var out = base.slice();
  ajouts.forEach(function(a){
    var f = -1;
    for(var i=0;i<out.length;i++) if(out[i][cle] === a[cle]){ f=i; break; }
    if(f>-1) out[f] = a; else out.push(a);
  });
  return out;
}
function nomBase(b){ return b==='gamserie'?'GAMSERIE':(b==='ts'?'TS':'Matrice'); }

function pickFile(cb, accept, multiple){
  var el = document.getElementById('filepick');
  el.value = '';
  el.accept = accept || '.csv,.txt,.json';
  el.multiple = !!multiple;
  el.onchange = function(){
    var fs = Array.prototype.slice.call(el.files);
    fs.forEach(function(f){
      var fr = new FileReader();
      fr.onload = function(){ cb(String(fr.result), f.name, f.size, f); };
      if(accept && accept.indexOf('pdf') > -1) fr.readAsDataURL(f); else fr.readAsText(f, 'UTF-8');
    });
  };
  el.click();
}

/* =======================================================================
   3 bis. Documents de méthode de contrôle (PDF stockés dans l'application)
   ======================================================================= */
function docId(){ return 'D' + Date.now().toString(36) + Math.random().toString(36).slice(2,5); }
function docSize(){ var s = 0; DOCS.forEach(function(d){ s += (d.data || '').length; }); return s; }
function ko(n){ return n >= 1048576 ? (Math.round(n/104857.6)/10) + ' Mo' : Math.round(n/1024) + ' Ko'; }

function importPDF(){
  pickFile(function(dataUrl, name, size){
    if(dataUrl.indexOf('application/pdf') < 0 && !/\.pdf$/i.test(name)){ toast('Seuls les fichiers PDF sont acceptés.'); return; }
    var av = DOCS.slice();
    var f = -1;
    for(var i=0;i<DOCS.length;i++) if(DOCS[i].nom.toLowerCase() === name.toLowerCase()){ f = i; break; }
    var d = {id: f>-1 ? DOCS[f].id : docId(), nom:name, taille:size, ajouteLe:new Date().toISOString(), data:dataUrl};
    if(f > -1) DOCS[f] = d; else DOCS.push(d);
    if(!saveDocs()){
      DOCS = av; saveDocs();
      modal('Stockage saturé',
        'Le document <b>' + esc(name) + '</b> (' + ko(size) + ') dépasse la capacité de la mémoire locale du navigateur, environ 5 Mo tous documents confondus.'
        + '<ul><li>Compressez le PDF ou n\'y gardez que les pages utiles.</li><li>Ou supprimez un document déjà stocké.</li></ul>');
      return;
    }
    if(blobCache[d.id]){ URL.revokeObjectURL(blobCache[d.id]); delete blobCache[d.id]; }
    render();
    toast((f>-1 ? 'Document remplacé : ' : 'Document ajouté : ') + name);
  }, 'application/pdf,.pdf', true);
}
function supprimerDoc(id){
  var d = docById(id);
  if(!d || !confirm('Supprimer le document ' + d.nom + ' ? Les gammes qui le citent perdront le lien.')) return;
  if(blobCache[id]){ URL.revokeObjectURL(blobCache[id]); delete blobCache[id]; }
  DOCS = DOCS.filter(function(x){ return x.id !== id; });
  saveDocs(); render(); toast('Document supprimé.');
}
function docById(id){ for(var i=0;i<DOCS.length;i++) if(DOCS[i].id === id) return DOCS[i]; return null; }

var blobCache = {};
function blobURL(d){
  if(blobCache[d.id]) return blobCache[d.id];
  var b64 = String(d.data).split(',')[1] || '';
  var bin = atob(b64), arr = new Uint8Array(bin.length);
  for(var i=0;i<bin.length;i++) arr[i] = bin.charCodeAt(i);
  blobCache[d.id] = URL.createObjectURL(new Blob([arr], {type:'application/pdf'}));
  return blobCache[d.id];
}
var docOuvert = null;
function ouvrirDoc(id){
  var d = docById(id);
  if(!d){ toast('Document introuvable.'); return; }
  var url = blobURL(d);
  docOuvert = d;
  document.body.classList.add('viewer');
  document.getElementById('vp-nom').textContent = d.nom;
  document.getElementById('vp-frame').src = url;
  document.getElementById('vp-info').innerHTML = ko(d.taille || 0) + ' — importé le ' + frDate(String(d.ajouteLe).slice(0,10))
    + '. Si l\'aperçu reste vide, votre navigateur n\'affiche pas les PDF : utilisez « Ouvrir hors de l\'appli ».';
}
function fermerDoc(){
  document.body.classList.remove('viewer');
  document.getElementById('vp-frame').src = 'about:blank';
  docOuvert = null;
}
function ouvrirDocExterne(){
  if(!docOuvert) return;
  window.open(blobURL(docOuvert), '_blank');
}

/* ---------- modale d'information ---------- */
function modal(titre, htmlCorps, onOk, labelOk){
  var old = document.querySelector('.modal'); if(old) old.remove();
  var m = document.createElement('div');
  m.className = 'modal';
  m.innerHTML = '<div class="box" role="dialog" aria-modal="true"><h3>' + esc(titre) + '</h3>'
    + '<div class="bd">' + htmlCorps + '</div>'
    + '<div class="ft">' + (onOk ? '<button class="btn" data-x>Annuler</button>' : '')
    + '<button class="btn pri" data-ok>' + esc(labelOk || (onOk ? 'Confirmer' : 'J\'ai compris')) + '</button></div></div>';
  document.body.appendChild(m);
  function close(){ m.remove(); document.removeEventListener('keydown', k); }
  function k(e){ if(e.key === 'Escape') close(); }
  document.addEventListener('keydown', k);
  m.addEventListener('click', function(e){ if(e.target === m) close(); });
  m.querySelector('[data-ok]').onclick = function(){ close(); if(onOk) onOk(); };
  var c = m.querySelector('[data-x]'); if(c) c.onclick = close;
  m.querySelector('[data-ok]').focus();
}

/* Sauvegarde complète (bases + rapports) au format JSON */
function exportJSON(){
  download('sauvegarde-controle-' + todayISO() + '.json',
    JSON.stringify({version:3, exportLe:new Date().toISOString(), db:DB, rapports:REPORTS, documents:DOCS,
      versions:VERS, revisions:HIST, journal:JOURNAL}, null, 1), 'application/json');
}
function importJSON(){
  pickFile(function(text){
    try{
      var o = JSON.parse(text);
      if(!o.db) throw new Error('fichier non reconnu');
      if(!confirm('Restaurer cette sauvegarde ? Les bases, les documents et les rapports enregistrés seront remplacés.')) return;
      DB = o.db; REPORTS = o.rapports || []; DOCS = o.documents || [];
      VERS = o.versions || {gam:{},ts:{}}; HIST = o.revisions || []; JOURNAL = o.journal || [];
      if(!VERS.gam) VERS.gam = {}; if(!VERS.ts) VERS.ts = {};
      blobCache = {};
      saveDB(); saveReports(); saveDocs(); saveVersions(); saveHist(); saveJournal();
      initVersions(); syncLignes(); saveDraft(); render();
      toast('Sauvegarde restaurée.');
    }catch(e){ toast('Restauration impossible : ' + e.message); }
  });
}

/* Export du rapport courant en CSV */
function exportRapport(rep){
  var t = totaux2(rep), n = nbPieces2(rep);
  var rows = [
    ['Article',rep.article,'Lot',rep.lot,'Commande',rep.commande,'Fournisseur',rep.fournisseur,'Contrôleur',rep.controleur,'Date',frDate(rep.date)],
    ['Quantité commandée',rep.qteCommandee,'Quantité réceptionnée',rep.qteReceptionnee,'Quantité échantillon',rep.qteEchantillon,
     'Conforme',t.conforme,'Non conforme',t.nonConforme,'Durée contrôle',rep.dureeControle,'Décision',rep.decision],
    ['Commentaires',rep.commentaires||''],
    ['Gamme',rep.famille||'','Révision gamme',rep.gamRev===null||rep.gamRev===undefined?'':rep.gamRev,
     'Révision valeurs',rep.tsRev===null||rep.tsRev===undefined?'':rep.tsRev,'Indice de définition',rep.indiceDef||''],
    []
  ];
  var gen = rep.generales || {};
  var aGen = rep.lignes.some(function(l){ return specKind(l) === 'generale'; });
  if(aGen){
    rows.push(['Spécifications générales']);
    rows.push(['Code','Spécification','Méthode','Résultat']);
    rep.lignes.forEach(function(l){
      if(specKind(l) !== 'generale') return;
      var v = String(gen[l.code]||'');
      rows.push([l.code, l.specification, l.methode||'', v==='1'?'OK':v==='0'?'NOK':'']);
    });
    rows.push([]);
  }
  var head = ['Code','Spécification','Mini.','Tol -','Nom.','Tol +','Maxi.','Outil','Méthode'];
  for(var j=0;j<n;j++) head.push(String(j+1));
  rows.push(head);
  rep.lignes.forEach(function(l,i){
    if(specKind(l) !== 'individuelle') return;   // seules les spéc. individuelles ont un relevé par pièce
    var r = [l.code, l.specification,
      l.na?'-':l.mini, l.na?'-':l.tolM, l.na?'-':l.nom, l.na?'-':l.tolP, l.na?'-':l.maxi, l.outil, l.methode||''];
    for(var j2=0;j2<n;j2++) r.push((rep.mesures[i]||[])[j2] || '');
    rows.push(r);
  });
  download('rapport_' + (rep.article||'sans-article') + '_' + (rep.lot||rep.id) + '.csv', toCSV(rows));
}
function nbPieces2(rep){ var n = num(rep.qteEchantillon); if(isNaN(n)||n<1) n=15; return Math.min(Math.floor(n),60); }
function totaux2(rep){ var s = R; R = rep; var t = totaux(); R = s; return t; }

/* =======================================================================
   5. Rendu
   ======================================================================= */
function render(){
  document.querySelectorAll('.nav button').forEach(function(b){ b.setAttribute('aria-current', b.dataset.view === VIEW ? 'true' : 'false'); });
  majBarreFichier();
  var m = document.getElementById('main');
  if(VIEW === 'rapport') m.innerHTML = viewRapport();
  else if(VIEW === 'historique') m.innerHTML = viewHistorique();
  else if(VIEW === 'analyse') m.innerHTML = viewAnalyse();
  else m.innerHTML = viewParametres();
  bind();
}

/* ---------- vue rapport ---------- */
function viewRapport(){
  var t = totaux(), n = nbPieces();
  var arts = articles();
  var fam = R.article ? familleOf(R.article) : '';
  var ro = rapportRO();                 // rapport figé (attente / accepté / refusé) : saisie bloquée
  var rdo = ro ? ' readonly' : '';       // pour les champs texte
  var dis = ro ? ' disabled' : '';       // pour les boutons / listes / mesures
  var h = '';
  h += '<div class="sheet '+(R.decision==='Accepté'?'dec-a':R.decision==='Refusé'?'dec-r':'')+(ro?' fige':'')+'">';
  h += '<div class="rbar"><b>Rapport de contrôle</b>'
    + '<span class="etat" id="etat-save">' + etatSauvegarde() + '</span>'
    + '<button class="close noprint" id="b-fermer" title="Fermer et revenir aux rapports enregistrés" aria-label="Fermer le rapport">'
    + '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M5 5 L19 19 M19 5 L5 19" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" fill="none"/></svg>'
    + '</button></div>';
  if(ro){
    h += '<div class="figebar noprint"><span class="lock">🔒</span> Rapport figé — statut <b>'+esc(libDecision(R.decision))+'</b>. '
      + (statutVerrouille()
          ? 'Modifiable uniquement après réouverture par un méthodiste.'
          : 'Modifiez le statut (Accepté / Refusé / Sans décision) pour le débloquer.')
      + (statutVerrouille() ? '<button class="btn sm" id="b-rouvrir" style="margin-left:auto">Rouvrir (méthodiste)</button>' : '') + '</div>';
  }
  h += '<div class="hdr">';
  h += champ('Article', '<input list="dl-art" id="f-article" value="'+esc(R.article)+'" placeholder="Saisir ou choisir…" autocomplete="off"'+rdo+'>');
  h += champ('Lot', '<input id="f-lot" value="'+esc(R.lot)+'"'+rdo+'>');
  h += champ('Commande', '<input id="f-commande" value="'+esc(R.commande)+'"'+rdo+'>');
  h += champ('Fournisseur', '<input id="f-fournisseur" value="'+esc(R.fournisseur)+'"'+rdo+'>');
  h += champ('Contrôleur', '<input id="f-controleur" value="'+esc(R.controleur)+'"'+rdo+'>');
  h += champ('Date', '<input type="date" id="f-date" value="'+esc(R.date)+'"'+rdo+'>');
  h += '</div>';
  h += '<div class="hdr">';
  h += champ('Quantité commandée', '<input id="f-qteCommandee" inputmode="decimal" value="'+esc(R.qteCommandee)+'"'+rdo+'>', 'num');
  h += champ('Quantité réceptionnée', '<input id="f-qteReceptionnee" inputmode="decimal" value="'+esc(R.qteReceptionnee)+'"'+rdo+'>', 'num');
  h += champ('Quantité échantillon', '<input id="f-qteEchantillon" inputmode="decimal" value="'+esc(R.qteEchantillon)+'"'+rdo+'>', 'num');
  h += champ('Conforme', '<div class="res ok"><span id="tot-ok">'+t.conforme+'</span><small>/ '+n+' pièces</small></div>');
  h += champ('Non conforme', '<div class="res nc"><span id="tot-nc">'+t.nonConforme+'</span><small>/ '+n+' pièces</small></div>');
  h += champ('Durée contrôle', '<input id="f-dureeControle" value="'+esc(R.dureeControle)+'" placeholder="ex. 0:45"'+rdo+'>');
  h += '</div>';
  h += '<div class="hdr"><div class="cell" style="grid-column:1 / -1"><label>Origine du contrôle</label>'
    + '<select id="f-origine" style="max-width:380px;border:1px solid var(--rule);background:var(--field);padding:5px 7px;border-radius:2px"'+dis+'>'
    + '<option value="">—</option>'
    + ORIGINES.map(function(o){ return '<option value="'+esc(o)+'"'+(R.origine===o?' selected':'')+'>'+esc(o)+'</option>'; }).join('')
    + '</select></div></div>';
  h += '<div class="cmtrow">'
    + '<div class="cmt"><label for="f-commentaires">Commentaires</label>'
    + '<textarea id="f-commentaires" rows="2" placeholder="Observations, réserves, éléments à signaler…">'+esc(R.commentaires||'')+'</textarea></div>'
    + docPaneRapport()
    + '</div>';

  if(!R.article){
    h += '<div class="empty"><b>Aucun article renseigné</b>Choisissez un article : les spécifications de sa famille et les valeurs attendues se déclinent automatiquement.'
      + (arts.length ? '' : '<br><br>Aucune base chargée. Rendez-vous dans <b style="display:inline">Paramètres</b> pour importer GAMSERIE, TS et Matrice.') + '</div>';
  } else if(!R.lignes.length){
    h += '<div class="empty"><b>Pas de gamme pour cet article</b>'
      + (fam ? 'Famille « '+esc(fam)+' » : aucune spécification dans GAMSERIE.' : 'Article absent de la Matrice : aucune famille associée.')
      + '</div>';
  } else {
    h += bandeauRevision();
    var gens = R.lignes.filter(function(l){ return specKind(l) === 'generale'; });
    var hasIndiv = R.lignes.some(function(l){ return specKind(l) === 'individuelle'; });

    /* --- spécifications générales : relevé unique OK / NOK, au-dessus du tableau --- */
    if(gens.length){
      var g = totauxGen();
      h += '<div class="genwrap"><div class="genh">Spécifications générales '
        + '<span class="hint" style="text-transform:none;letter-spacing:0">— un seul relevé OK / NOK pour le lot'
        + (g.nc?' · <b style="color:var(--nc)">'+g.nc+' NOK</b>':'') + (g.reste?' · '+g.reste+' à renseigner':'') + '</span></div>';
      R.lignes.forEach(function(l){
        if(specKind(l) !== 'generale') return;
        var v = String((R.generales||{})[l.code]||'');
        var cls = v==='1'?'v1':v==='0'?'v0':'';
        h += '<div class="genrow">'
          + '<span class="gencode">'+esc(l.code)+'</span>'
          + '<span class="genspec"><span class="crit '+esc(String(l.criticite).toLowerCase())+'" title="'+esc(l.criticite)+'"></span>'+esc(l.specification)+'</span>'
          + '<span class="genmeth">'+esc(l.methode||'')+'</span>'
          + '<input class="genmoy" data-moy="'+esc(l.code)+'" value="'+esc((R.moyens||{})[l.code]||'')+'" placeholder="N° moyen" title="Numéro de moyen de contrôle"'+rdo+'>'
          + '<button class="q genq '+cls+'" data-gen="'+esc(l.code)+'" title="Basculer OK / NOK"'+dis+'>'+(cls==='v1'?'OK':cls==='v0'?'NOK':'–')+'</button>'
          + '</div>';
      });
      h += '</div>';
    }

    /* --- spécifications individuelles : relevé pièce par pièce --- */
    if(hasIndiv){
      h += '<div class="scroll releve"><table class="ctl"><thead><tr>'
        + '<th class="c-code sticky1">Code</th><th class="c-spec sticky2">Spécification'
        + (fam ? ' <span style="font-weight:400;text-transform:none;letter-spacing:0;color:var(--ink3)">— '+esc(fam)+'</span>' : '')
        + '</th><th class="c-tol">Mini.</th><th class="c-tol">Tol -</th><th class="c-tol">Nom.</th><th class="c-tol">Tol +</th><th class="c-tol">Maxi.</th><th class="c-out">Outil</th><th class="c-meth">Méthode</th><th class="c-moy">N° moyen</th>';
      for(var j=0;j<n;j++){ var ncp = ncParPiece(j); h += '<th class="c-p" data-p="'+j+'">'+(j+1)+(ncp?'<span class="nccount">'+ncp+'</span>':'')+'</th>'; }
      h += '</tr></thead><tbody>';
      R.lignes.forEach(function(l,i){
        if(specKind(l) !== 'individuelle') return;
        var nc = ncParLigne(i);
        h += '<tr class="'+(l.na?'na':'')+'">';
        h += '<td class="code sticky1">'+esc(l.code)+'</td>';
        h += '<td class="spec sticky2"><span class="crit '+esc(String(l.criticite).toLowerCase())+'" title="'+esc(l.criticite)+'"></span>'+esc(l.specification)
           + (nc?'<span class="nccount">'+nc+'</span>':'')+'</td>';
        if(l.na){
          h += '<td class="t">-</td><td class="t">-</td><td class="t nom">-</td><td class="t">-</td><td class="t">-</td>';
        } else {
          h += '<td class="t">'+fmt(l.mini)+'</td><td class="t">'+fmt(l.tolM)+'</td><td class="t nom">'+fmt(l.nom)+'</td><td class="t">'+fmt(l.tolP)+'</td><td class="t">'+fmt(l.maxi)+'</td>';
        }
        h += '<td class="outil">'+esc(l.outil)+'</td>';
        h += '<td class="meth">'+esc(l.methode)+'</td>';
        h += '<td class="moy"><input class="cmoy" data-moy="'+esc(l.code)+'" value="'+esc((R.moyens||{})[l.code]||'')+'" title="Numéro de moyen de contrôle"'+rdo+'></td>';
        for(var k=0;k<n;k++){
          if(!l.testable){ h += '<td class="mes off" data-r="'+i+'" data-c="'+k+'"></td>'; continue; }
          var st = mesureStatut(i,k), v = R.mesures[i][k];
          if(l.qual){
            var cls = v==='1'||v===1 ? 'v1' : (v==='0'||v===0 ? 'v0' : '');
            h += '<td class="mes '+st+'" data-r="'+i+'" data-c="'+k+'"><button class="q '+cls+'" data-r="'+i+'" data-c="'+k+'"'+dis+'>'+(cls==='v1'?'OK':cls==='v0'?'NOK':'–')+'</button></td>';
          } else {
            h += '<td class="mes '+st+'" data-r="'+i+'" data-c="'+k+'"><input inputmode="decimal" data-r="'+i+'" data-c="'+k+'" value="'+esc(formatMesure(v))+'"'+dis+'></td>';
          }
        }
        h += '</tr>';
      });
      h += '</tbody><tfoot><tr><td class="lbl" colspan="10">Verdict par pièce</td>';
      for(var p=0;p<n;p++){
        var s = pieceStatut(p);
        h += '<td class="'+(s==='ok'?'pok':s==='nc'?'pnc':'')+'" data-p="'+p+'">'+(s==='ok'?'C':s==='nc'?'NC':'·')+'</td>';
      }
      h += '</tr></tfoot></table></div>';

      /* --- carte de contrôle type Minitab (I-MR) : ouverte dans une fenêtre superposée --- */
      var cotes = [];
      R.lignes.forEach(function(l){ if(specKind(l)==='individuelle' && l.testable && !l.qual) cotes.push(l.code); });
      if(cotes.length){
        if(R.chartCote !== '__all' && cotes.indexOf(R.chartCote) < 0) R.chartCote = cotes[0];
        h += '<div class="bar noprint" style="border-top:0"><button class="btn pri" id="b-carte">📈 Carte de contrôle — suivi des déviances</button>'
          + '<span class="hint">Valeurs individuelles (I) + étendues mobiles (MR), en fenêtre superposée.</span></div>';
      }
    }

    if(!gens.length && !hasIndiv){
      h += '<div class="empty"><b>Aucun point à relever</b>Renseignez la colonne <b style="display:inline">TYPE SPEC</b> (Générale ou Individuelle) dans <b style="display:inline">Paramètres → TS</b> : les points « Générale » se relèvent une fois au-dessus du tableau, les points « Individuelle » pièce par pièce.</div>';
    }
  }

  h += '<div class="decprint">'+(R.commentaires?'Commentaires : '+esc(R.commentaires)+'<br>':'')+'Gamme <b>'+esc(R.famille||'—')+'</b> rév. <b>'+(R.gamRev===null||R.gamRev===undefined?'—':R.gamRev)+'</b>'
     + ' · Valeurs rév. <b>'+(R.tsRev===null||R.tsRev===undefined?'—':R.tsRev)+'</b>'+(R.indiceDef?' ('+esc(R.indiceDef)+')':'')+'<br>'
     + 'Décision : <b>'+esc(R.decision||'—')+'</b> · Conforme : <b>'+t.conforme+'</b> · Non conforme : <b>'+t.nonConforme+'</b>'
     + ' · Contrôleur : <b>'+esc(R.controleur)+'</b> · Date : <b>'+frDate(R.date)+'</b></div>';
  var stCls = R.decision==='Accepté'?'dec-a':R.decision==='Refusé'?'dec-r':R.decision==='Attente décision méthodes'?'dec-att':'';
  var stampCls = R.decision==='Accepté'?'a':R.decision==='Refusé'?'r':R.decision==='Attente décision méthodes'?'att':'';
  h += '<div class="bar noprint">';
  h += '<label class="barlab">Statut</label>';
  h += '<select id="f-statut" class="'+stCls+'"'+(statutVerrouille()?' disabled':'')+' style="border:1px solid var(--rule);padding:5px 9px;border-radius:2px">'
    + STATUTS.map(function(s){ return '<option value="'+esc(s)+'"'+(R.decision===s?' selected':'')+'>'+esc(libDecision(s))+'</option>'; }).join('')
    + '</select>';
  if(R.decision === 'Refusé'){
    var subs = []; TYPES_NC.forEach(function(x){ if(x.t === R.typeNC) subs = x.sous; });
    h += '<label class="barlab">Type NC</label>'
      + '<select id="f-typenc" style="border:1px solid var(--rule);padding:5px 9px;border-radius:2px;max-width:230px">'
      + '<option value="">— type —</option>'
      + TYPES_NC.map(function(x){ return '<option value="'+esc(x.t)+'"'+(R.typeNC===x.t?' selected':'')+'>'+esc(x.t)+'</option>'; }).join('')
      + '</select>'
      + '<select id="f-soustypenc" style="border:1px solid var(--rule);padding:5px 9px;border-radius:2px;max-width:300px"'+(subs.length?'':' disabled')+'>'
      + '<option value="">— sous-type —</option>'
      + subs.map(function(s){ return '<option value="'+esc(s)+'"'+(R.sousTypeNC===s?' selected':'')+'>'+esc(s)+'</option>'; }).join('')
      + '</select>';
  }
  h += '<span class="hint" id="hint-pieces">'+(t.encours?t.encours+' pièce'+(t.encours>1?'s':'')+' non commencée'+(t.encours>1?'s':''):'Toutes les pièces sont renseignées')+'</span>';
  if(R.decision) h += '<span class="stamp '+stampCls+'">'+esc(libDecision(R.decision))+'</span>';
  h += '<span class="sp"></span>';
  h += '<button class="btn dgr" id="b-eqf" title="Déclarer un écart qualité fournisseur dans X3">Déclarer un EQf X3</button>';
  if(R.decision === 'Accepté') h += '<button class="btn" id="b-mvtA" title="Changement de stock X3 : statut Q → A">Mouvement X3 statut A</button>';
  else if(R.decision === 'Refusé') h += '<button class="btn" id="b-mvtR" title="Changement de stock X3 : statut Q → R">Mouvement X3 statut R</button>';
  h += '<button class="btn" id="b-vider"'+dis+'>Vider les mesures</button>';
  h += '<button class="btn" id="b-nouveau">Nouveau rapport</button>';
  h += '<button class="btn" id="b-csv">Exporter en CSV</button>';
  h += '<button class="btn" id="b-print">Imprimer / PDF</button>';
  h += '</div></div>';
  h += '<datalist id="dl-art">' + arts.map(function(a){ return '<option value="'+esc(a)+'">'; }).join('') + '</datalist>';
  return h;
}
/* Bandeau : révisions appliquées au rapport + écart avec la révision en vigueur */
function bandeauRevision(){
  var vg = version('gam', R.famille), vt = version('ts', R.article);
  var h = '<div class="rev">'
    + '<span>Gamme <b>' + esc(R.famille || '—') + '</b> — <b>rév. ' + (R.gamRev === null || R.gamRev === undefined ? '—' : R.gamRev) + '</b>'
    + (R.gamDateEffet ? ' <span class="hint">applicable depuis le ' + frDate(R.gamDateEffet) + '</span>' : '') + '</span>'
    + '<span>Valeurs <b>' + esc(R.article) + '</b> — <b>rév. ' + (R.tsRev === null || R.tsRev === undefined ? '—' : R.tsRev) + '</b>'
    + (R.indiceDef ? ' <span class="hint">' + esc(R.indiceDef) + '</span>' : '') + '</span>';
  h += '</div>';

  var alertes = [];
  if(R.gamStatut === 'brouillon') alertes.push('La gamme « ' + esc(R.famille) + ' » est en <b>brouillon de révision ' + R.gamRev + '</b>, non validée.');
  if(R.tsStatut === 'brouillon') alertes.push('Les valeurs de <b>' + esc(R.article) + '</b> sont en <b>brouillon de révision ' + R.tsRev + '</b>, non validées.');
  if(R.gamStatut === 'sans révision' || R.tsStatut === 'sans révision') alertes.push('Données sans révision enregistrée : validez-les dans Paramètres → Révisions.');
  var chgGam = vg && R.gamRev !== null && (vg.rev !== R.gamRev || (R.gamStatut === 'brouillon' && vg.statut === 'applicable'));
  var chgTs = vt && R.tsRev !== null && (vt.rev !== R.tsRev || (R.tsStatut === 'brouillon' && vt.statut === 'applicable'));
  var perime = chgGam || chgTs;
  if(perime){
    alertes.push('Ce rapport a été établi selon la gamme <b>rév. ' + R.gamRev + (R.gamStatut==='brouillon'?' (brouillon)':'')
      + '</b> et les valeurs <b>rév. ' + R.tsRev + (R.tsStatut==='brouillon'?' (brouillon)':'') + '</b>. '
      + 'En vigueur aujourd\'hui : gamme <b>rév. ' + (vg?vg.rev:'—') + '</b>, valeurs <b>rév. ' + (vt?vt.rev:'—') + '</b>.');
  }
  if(alertes.length){
    h += '<div class="alerte noprint"><div>' + alertes.join('<br>') + '</div>'
      + (perime ? '<button class="btn sm" id="b-rebasculer">Rebasculer sur les révisions en vigueur</button>' : '') + '</div>';
  }
  return h;
}
function rebasculer(){
  var avant = {g:R.gamRev, t:R.tsRev};
  syncLignes(); resizeMesures();
  R.rebascules = R.rebascules || [];
  R.rebascules.push({le:new Date().toISOString(), de:'gamme rév. '+avant.g+' / valeurs rév. '+avant.t,
                     vers:'gamme rév. '+R.gamRev+' / valeurs rév. '+R.tsRev});
  enregistrerAuto(true); render();
  toast('Rapport rebasculé sur la gamme rév. ' + R.gamRev + ' et les valeurs rév. ' + R.tsRev + '. Vérifiez les mesures déjà saisies.');
}

function champ(label, inner, cls){
  return '<div class="cell '+(cls||'')+'"><label>'+esc(label)+'</label>'+inner+'</div>';
}
/* Panneau « Documents applicables » : à droite des commentaires, il liste les
   références documentaires de l'article (matrice) et ouvre le PDF stocké s'il existe. */
function docPaneRapport(){
  var refs = R.article ? refsArticle(R.article) : [];
  var dispo = refs.filter(function(r){ return r.doc; }).length;
  var h = '<div class="docpane"><label>Documents applicables'
    + (refs.length ? ' <span class="hint" style="text-transform:none;letter-spacing:0">'+dispo+'/'+refs.length+' disponible'+(dispo>1?'s':'')+'</span>' : '')
    + '</label>';
  if(!R.article){
    h += '<div class="doc-empty">Renseignez un article pour voir ses documents.</div>';
  } else if(!refs.length){
    h += '<div class="doc-empty">Aucun document recensé dans la matrice pour cet article.</div>';
  } else {
    h += '<div class="doclist">';
    refs.forEach(function(r){
      h += '<div class="docitem"><span class="doc-lib">'+esc(r.lib)+'</span>'
        + (r.doc
            ? '<a class="doc" data-doc="'+esc(r.doc.id)+'" title="Ouvrir '+esc(r.doc.nom)+'" tabindex="0">'+esc(r.ref)+'</a>'
            : '<span class="doc-miss">'+esc(r.ref)+' <em>· PDF non importé</em></span>')
        + '</div>';
    });
    h += '</div>';
  }
  return h + '</div>';
}

/* ---------- vue historique ---------- */
function viewHistorique(){
  var list = REPORTS.slice().sort(function(a,b){
    var da = String(a.date||''), db = String(b.date||'');
    if(da !== db) return da < db ? 1 : -1;                 // colonne DATE, du plus récent au plus ancien
    return (b.creeLe||'') < (a.creeLe||'') ? -1 : 1;        // à date égale : dernier créé en premier
  });
  /* Menu principal : accès aux rapports, aux analyses et aux paramètres */
  var h = '<div class="sheet"><div class="menuhead"><h2 class="sec">Rapports enregistrés</h2>'
     + '<div class="menuact">'
     + '<button class="btn" id="h-analyse">'+ICO_MENU.analyse+'<span>Analyses des dérives</span></button>'
     + '<button class="btn" id="h-parametres">'+ICO_MENU.reglages+'<span>Paramètres</span></button>'
     + '<button class="btn pri" id="h-nouveau">'+ICO_MENU.plus+'<span>Nouveau rapport</span></button>'
     + '</div></div><div class="pad">';
  if(brouillonEnCours()){
    h += '<div class="encours"><span><b>Rapport en cours, non enregistré</b> — '
      + esc([R.article || 'sans article', R.lot ? 'lot ' + R.lot : ''].filter(Boolean).join(' · ')) + '</span>'
      + '<span class="sp"></span><button class="btn sm" id="h-abandon">Abandonner</button><button class="btn sm pri" id="h-reprendre">Reprendre</button></div>';
  }
  h += '<div class="toolbar"><input type="search" id="q" placeholder="Rechercher un article, un lot, un fournisseur…" value="'+esc(window.__q||'')+'">'
     + '<span class="hint" id="q-count">'+list.length+' rapport'+(list.length>1?'s':'')+'</span>'
     + '<span class="hint" style="margin-left:auto">Cliquez sur un rapport pour l\'ouvrir.</span></div>';
  if(!list.length){
    h += '<div class="empty"><b>Aucun rapport enregistré</b>Cliquez sur « Nouveau rapport » pour commencer un contrôle.</div>';
  } else {
    h += '<div class="rowlist" data-list="rapports"><div class="r head"><div>Article</div><div>Lot / Commande</div><div>Date</div><div>Échant. / rév.</div><div>Résultat</div><div>Décision</div><div style="text-align:right">Actions</div></div>';
    list.forEach(function(r){
      var t = totaux2(r);
      h += '<div class="r" data-open="'+r.id+'" tabindex="0" title="Ouvrir ce rapport" data-s="'+esc([r.article,r.lot,r.commande,r.fournisseur,r.controleur,r.decision].join(' ').toLowerCase())+'">'
        + '<div><b>'+esc(r.article)+'</b><br><span class="hint">'+esc(r.fournisseur)+'</span></div>'
        + '<div>'+esc(r.lot)+'<br><span class="hint">'+esc(r.commande)+'</span></div>'
        + '<div>'+frDate(r.date)+'<br><span class="hint">'+esc(r.controleur)+'</span></div>'
        + '<div>'+nbPieces2(r)+'<br><span class="hint">rév. '+(r.gamRev===null||r.gamRev===undefined?'—':r.gamRev)+' / '+(r.tsRev===null||r.tsRev===undefined?'—':r.tsRev)+'</span></div>'
        + '<div><span style="color:var(--ok);font-weight:700">'+t.conforme+' C</span> / <span style="color:var(--nc);font-weight:700">'+t.nonConforme+' NC</span></div>'
        + '<div><span class="tag '+(r.decision==='Accepté'?'a':r.decision==='Refusé'?'r':r.decision==='Attente décision méthodes'?'att':'d')+'">'+esc(libDecision(r.decision))+'</span></div>'
        + '<div class="act2">'
          + '<button class="btn sm" data-exp="'+r.id+'">CSV</button>'
          + '<button class="btn sm dgr" data-del="'+r.id+'">Supprimer</button>'
        + '</div></div>';
    });
    h += '</div>';
  }
  h += '</div></div>';
  return h;
}

/* =======================================================================
   Vue « Analyses des dérives » : extraction des relevés de tous les
   rapports, filtres, tableau et cartes de contrôle type Minitab.
   ======================================================================= */
var FANA = {article:'', lot:'', fournisseur:'', code:'', date:''};
/* Un relevé = une valeur numérique mesurée (article × lot × code × pièce). */
function relevesTous(){
  var out = [];
  REPORTS.forEach(function(rep){
    var n = nbPieces2(rep);
    (rep.lignes||[]).forEach(function(l,i){
      if(specKind(l) !== 'individuelle' || !l.testable || l.qual) return;
      for(var j=0;j<n;j++){
        var v = num((rep.mesures[i]||[])[j]);
        if(isNaN(v)) continue;
        out.push({article:rep.article||'', lot:rep.lot||'', fournisseur:rep.fournisseur||'', date:rep.date||'',
          code:l.code, spec:l.specification||'', mini:l.mini, tolM:l.tolM, nom:l.nom, tolP:l.tolP, maxi:l.maxi,
          valeur:v, piece:j+1, gamRev:rep.gamRev, tsRev:rep.tsRev});
      }
    });
  });
  return out;
}
function analyseFiltrees(){
  var recs = relevesTous(), f = FANA;
  function ct(v,q){ return !q || String(v||'').toLowerCase().indexOf(String(q).toLowerCase()) > -1; }
  return recs.filter(function(r){
    return ct(r.article,f.article) && ct(r.lot,f.lot) && ct(r.fournisseur,f.fournisseur)
      && (!f.code || r.code === f.code) && (!f.date || r.date === f.date);
  });
}
/* Sortie (carte de contrôle + tableau) recalculée à chaque changement de filtre. */
function analyseOut(){
  var f = FANA, filt = analyseFiltrees();
  var vals = filt.map(function(r){ return r.valeur; });
  var spec = (f.code && filt.length) ? {mini:filt[0].mini, maxi:filt[0].maxi, nom:filt[0].nom} : null;
  var h = '<div class="chartbox" style="border-top:0;padding:0 0 12px;background:transparent">'
    + blocIMR(vals, spec, f.code ? ('Côte '+f.code) : 'Relevés superposés') + '</div>';
  if(!filt.length){
    h += '<div class="empty"><b>Aucun relevé</b>' + (relevesTous().length ? 'Aucun relevé ne correspond aux filtres.' : 'Enregistrez des rapports avec des mesures pour alimenter l\'analyse.') + '</div>';
    return h;
  }
  var cap = 600, list = filt.slice(0, cap);
  h += '<div class="scroll tall"><table class="flat" data-tbl="ana"><thead><tr>'
    + '<th>Article</th><th>Lot</th><th>Fournisseur</th><th>Date</th><th>Code</th><th>Spécification</th>'
    + '<th class="w-val">Mini.</th><th class="w-val">Tol -</th><th class="w-val">Nominal</th><th class="w-val">Tol +</th><th class="w-val">Maxi.</th>'
    + '<th class="w-val">Valeur mesurée</th><th class="w-rev">Ind. GAMSERIE</th><th class="w-rev">Ind. TS</th></tr></thead><tbody>';
  list.forEach(function(r){
    var hors = (typeof r.mini==='number' && r.valeur < r.mini-EPS) || (typeof r.maxi==='number' && r.valeur > r.maxi+EPS);
    h += '<tr'+(hors?' class="trou"':'')+'>'
      + '<td>'+esc(r.article)+'</td><td>'+esc(r.lot)+'</td><td>'+esc(r.fournisseur)+'</td><td class="dt">'+frDate(r.date)+'</td>'
      + '<td class="cd">'+esc(r.code)+'</td><td class="sp">'+esc(r.spec)+'</td>'
      + '<td class="bn">'+(r.mini===null||r.mini===undefined?'—':fmt(r.mini))+'</td>'
      + '<td class="bn">'+(r.tolM===null||r.tolM===undefined?'—':fmt(r.tolM))+'</td>'
      + '<td class="bn">'+(r.nom===null||r.nom===undefined?'—':fmt(r.nom))+'</td>'
      + '<td class="bn">'+(r.tolP===null||r.tolP===undefined?'—':fmt(r.tolP))+'</td>'
      + '<td class="bn">'+(r.maxi===null||r.maxi===undefined?'—':fmt(r.maxi))+'</td>'
      + '<td class="bn" style="font-weight:700'+(hors?';color:var(--nc)':'')+'">'+formatMesure(Math.round(r.valeur*1e6)/1e6)+'</td>'
      + '<td class="rv"><b>'+(r.gamRev===null||r.gamRev===undefined?'—':r.gamRev)+'</b></td>'
      + '<td class="rv"><b>'+(r.tsRev===null||r.tsRev===undefined?'—':r.tsRev)+'</b></td></tr>';
  });
  h += '</tbody></table></div>';
  if(filt.length > cap) h += '<p class="hint" style="margin-top:10px">Affichage limité à '+cap+' relevés : affinez les filtres.</p>';
  return h;
}
function viewAnalyse(){
  var recs = relevesTous(), f = FANA;
  var arts = {}, codes = {};
  recs.forEach(function(r){ if(r.article) arts[r.article]=1; if(r.code) codes[r.code]=1; });

  var h = '<div class="sheet"><div class="rbar"><b>Analyses des dérives</b>'
    + '<span class="etat">' + recs.length + ' relevé' + (recs.length>1?'s':'') + ' dans ' + REPORTS.length + ' rapport' + (REPORTS.length>1?'s':'') + '</span>'
    + '<button class="close noprint" id="a-fermer" title="Fermer et revenir aux rapports enregistrés" aria-label="Fermer">'
    + '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M5 5 L19 19 M19 5 L5 19" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" fill="none"/></svg>'
    + '</button></div><div class="pad">';
  h += '<div class="toolbar">'
    + '<input type="text" id="a-article" list="dl-aart" placeholder="Article" value="'+esc(f.article)+'" style="min-width:150px">'
    + '<input type="text" id="a-lot" placeholder="Lot" value="'+esc(f.lot)+'" style="min-width:110px">'
    + '<input type="text" id="a-fournisseur" placeholder="Fournisseur" value="'+esc(f.fournisseur)+'" style="min-width:130px">'
    + '<select id="a-code"><option value="">Tous points de contrôle</option>'
    + Object.keys(codes).sort().map(function(c){ return '<option value="'+esc(c)+'"'+(f.code===c?' selected':'')+'>Point '+esc(c)+'</option>'; }).join('') + '</select>'
    + '<label class="hint" style="display:flex;align-items:center;gap:5px">Date <input type="date" id="a-date" value="'+esc(f.date)+'"></label>'
    + '<button class="btn sm" id="a-reset">Réinitialiser</button>'
    + '<span class="sp" style="margin-left:auto"></span>'
    + '<span class="hint" id="a-count"></span>'
    + '</div>';
  h += '<datalist id="dl-aart">'+Object.keys(arts).sort().map(function(a){ return '<option value="'+esc(a)+'">'; }).join('')+'</datalist>';
  h += '<div id="ana-out">'+analyseOut()+'</div>';
  h += '</div></div>';
  return h;
}
function majAnalyse(){
  var out = document.getElementById('ana-out'); if(out) out.innerHTML = analyseOut();
  var filt = analyseFiltrees(), c = document.getElementById('a-count');
  if(c) c.textContent = filt.length + ' relevé' + (filt.length>1?'s':'') + ' filtré' + (filt.length>1?'s':'');
  setupColResize();   // ré-active le redimensionnement sur le tableau reconstruit
}
function bindAnalyse(){
  var b;
  if(b = document.getElementById('a-fermer')) b.onclick = function(){ VIEW = 'historique'; render(); };
  if(b = document.getElementById('a-reset')) b.onclick = function(){ FANA = {article:'', lot:'', fournisseur:'', code:'', date:''}; render(); };
  function lier(id, key){
    var el = document.getElementById(id); if(!el) return;
    el.oninput = el.onchange = function(){ FANA[key] = el.value; majAnalyse(); };
  }
  lier('a-article','article'); lier('a-lot','lot'); lier('a-fournisseur','fournisseur');
  lier('a-code','code'); lier('a-date','date');
  majAnalyse();   // renseigne le compteur initial
}

/* ---------- vue paramètres ---------- */
function viewParametres(){
  var h = '<div class="sheet">'
    + '<div class="rbar"><b>Paramètres</b>'
    + '<button class="close noprint" id="p-fermer" title="Fermer les paramètres et revenir aux rapports enregistrés" aria-label="Fermer les paramètres">'
    + '<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M5 5 L19 19 M19 5 L5 19" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" fill="none"/></svg>'
    + '</button></div>'
    + '<div class="tabs">'
    + '<button data-pset="gamserie" aria-current="'+(PSET==='gamserie')+'">GAMSERIE — gammes de contrôle</button>'
    + '<button data-pset="ts" aria-current="'+(PSET==='ts')+'">TS — valeurs par article</button>'
    + '<button data-pset="matrice" aria-current="'+(PSET==='matrice')+'">Matrice — article / fournisseur</button>'
    + '<button data-pset="revisions" aria-current="'+(PSET==='revisions')+'">Révisions</button>'
    + '<button data-pset="documents" aria-current="'+(PSET==='documents')+'">Documents ('+DOCS.length+')</button>'
    + '<button data-pset="sauvegarde" aria-current="'+(PSET==='sauvegarde')+'">Sauvegarde</button>'
    + '</div><div class="pad">';
  if(PSET === 'gamserie') h += psetGamserie();
  else if(PSET === 'ts') h += psetTS();
  else if(PSET === 'matrice') h += psetMatrice();
  else if(PSET === 'documents') h += psetDocuments();
  else if(PSET === 'revisions') h += psetRevisions();
  else h += psetSauvegarde();
  h += '</div></div>';
  return h;
}

function barreVersion(type, cle){
  if(!cle) return '';
  var v = version(type, cle);
  var st = !v ? '<span class="st non">sans révision</span>'
    : v.statut === 'brouillon' ? '<span class="st bro">brouillon rév. ' + v.rev + '</span>'
    : '<span class="st app">rév. ' + v.rev + ' applicable</span>';
  var meta = '';
  if(v && v.statut === 'applicable'){
    meta = '<span class="meta">depuis le ' + frDate(v.dateEffet) + (v.valideePar ? ' · validée par ' + esc(v.valideePar) : '')
      + (v.motif ? ' · ' + esc(v.motif) : '') + (v.indiceDef ? ' · ' + esc(v.indiceDef) : '') + '</span>';
  } else if(v){
    meta = '<span class="meta">modification en cours, non applicable tant qu\'elle n\'est pas validée</span>';
  }
  var h = '<div class="vbar">' + st + meta + '<span class="sp2"></span>';
  if(!v || v.statut === 'applicable'){
    h += '<button class="btn sm" data-newrev="'+type+'" data-cle="'+esc(cle)+'">Créer la révision ' + (v ? v.rev + 1 : 0) + '</button>';
  } else {
    h += '<button class="btn sm pri" data-valrev="'+type+'" data-cle="'+esc(cle)+'">Valider la révision ' + v.rev + '</button>'
       + '<button class="btn sm dgr" data-abrev="'+type+'" data-cle="'+esc(cle)+'">Abandonner</button>';
  }
  h += '<button class="btn sm" data-hist="'+type+'" data-cle="'+esc(cle)+'">Historique</button></div>';
  return h;
}

/* =======================================================================
   Tableaux à plat GAMSERIE et TS : révision, date, statut, filtres d'en-tête
   ======================================================================= */
function statutLib(st){ return st === 'brouillon' ? 'brouillon' : st === 'archive' ? 'archive' : 'valide'; }
function tagStatut(st){ return '<span class="tag ' + (st==='brouillon'?'d':st==='archive'?'ar':'a') + '">' + statutLib(st) + '</span>'; }

/* --- GAMSERIE : une ligne par spécification, groupée par famille --- */
function lignesGam(archives){
  var out = [];
  familles().forEach(function(fam){
    var v = version('gam', fam);
    DB.gamserie.forEach(function(g, i){
      if(g.famille !== fam) return;
      out.push({famille:fam, code:g.code, criticite:g.criticite||'', specification:g.specification||'',
        moyen:g.moyen||'', methode:g.methode||'', idx:i,
        rev:v?v.rev:null, dateRev:v?v.dateEffet:'', statut:v?(v.statut==='brouillon'?'brouillon':'valide'):'valide',
        edit: v ? v.statut === 'brouillon' : false});
    });
    if(archives){
      revisionsDe('gam', fam).forEach(function(h){
        if(v && h.rev === v.rev) return;
        (h.photo||[]).forEach(function(l){
          out.push({famille:fam, code:l.code, criticite:l.criticite||'', specification:l.specification||'',
            moyen:l.moyen||'', methode:l.methode||'', idx:-1,
            rev:h.rev, dateRev:h.dateEffet, statut:'archive', edit:false});
        });
      });
    }
  });
  return out;
}

function psetGamserie(){
  var lignes = lignesGam(GARCH);
  var brouillons = {}, valides = {};
  lignes.forEach(function(l){ if(l.statut==='brouillon') brouillons[l.famille]=1; else if(l.statut==='valide') valides[l.famille]=1; });
  var nb = Object.keys(brouillons).length;

  var h = '<div class="toolbar">'
    + '<button class="btn sm" id="g-newfam">Nouvelle famille</button>'
    + '<button class="btn sm" id="g-add">Ajouter une spécification</button>'
    + '<label class="chk"><input type="checkbox" id="g-arch"'+(GARCH?' checked':'')+'> Afficher les révisions archivées</label>'
    + '<span class="sp" style="margin-left:auto"></span>'
    + '<button class="btn sm" id="g-exp">Exporter en CSV</button>'
    + '<button class="btn sm" id="g-imp-r">Importer (remplacer)</button>'
    + '<button class="btn sm" id="g-imp-m">Importer (mettre à jour)</button>'
    + '</div>';
  h += '<div class="toolbar">'
    + '<span class="hint" id="gam-count">' + lignes.length + ' ligne' + (lignes.length>1?'s':'') + '</span>'
    + (Object.keys(valides).length ? '<button class="btn sm" id="g-newrev">Ouvrir une révision (familles affichées)</button>' : '')
    + (nb ? '<button class="btn sm pri" id="g-valrev">Valider les ' + nb + ' brouillon' + (nb>1?'s':'') + '</button>' : '')
    + '<span class="hint">Une famille en révision valide est figée : ouvrez une révision pour la modifier.</span>'
    + '</div>';

  if(!lignes.length) return h + '<div class="empty"><b>Aucune spécification</b>Créez une famille ou importez GAMSERIE.csv.</div>';

  h += '<div class="scroll tall"><table class="flat" data-tbl="gam"><thead><tr>'
    + thf('gam','rev','Rév.','w-rev') + thf('gam','date','Date rév.','w-date') + thf('gam','statut','Statut','w-st')
    + thf('gam','famille','Famille','w-fam') + thf('gam','code','Code','w-code') + thf('gam','criticite','Criticité','w-crit')
    + thf('gam','spec','Spécification','') + thf('gam','moyen','Moyen de contrôle','w-moy') + thf('gam','methode','Méthode de contrôle','w-moy')
    + '<th class="w-act"></th></tr></thead><tbody>';
  var prec = '';
  lignes.forEach(function(l){
    var cle = l.famille + '|' + l.rev + '|' + l.statut;
    var debut = cle !== prec; prec = cle;
    h += '<tr class="'+(debut?'grp':'')+(l.statut==='archive'?' arch':'')+'">';
    h += cellRev('gam', l, debut)
      + td('statut', statutLib(l.statut), tagStatut(l.statut), 'st')
      + td('famille', l.famille, esc(l.famille), 'fam' + (debut?'':' sub'))
      + td('code', l.code, esc(l.code), 'cd')
      + (l.edit
          ? '<td data-k="criticite" data-v="'+esc(l.criticite)+'" class="pz"><select data-b="gamserie" data-i="'+l.idx+'" data-f="criticite">'
            + ['','Critique','Majeure','Mineure'].map(function(c){ return '<option'+(String(l.criticite)===c?' selected':'')+'>'+c+'</option>'; }).join('')
            + '</select></td>'
          : td('criticite', l.criticite, esc(l.criticite||'—'), 'ty'))
      + tdTxt('spec', l, 'specification', 'sp')
      + tdTxt('moyen', l, 'moyen', 'fam')
      + tdMeth(l)
      + '<td class="act3">' + (l.edit ? '<button class="btn sm dgr" data-delrow="gamserie" data-i="'+l.idx+'">Suppr.</button>' : '') + '</td>'
      + '</tr>';
  });
  h += '</tbody></table></div>';
  return h;
}
function tdTxt(k, l, champ, cls){
  var v = l[champ] || '';
  if(!l.edit) return td(k, v, esc(v), cls);
  return '<td data-k="'+k+'" data-v="'+esc(v)+'" class="pz"><input data-b="gamserie" data-i="'+l.idx+'" data-f="'+champ+'" data-cell="'+k+'" value="'+esc(v)+'"></td>';
}
/* « Méthode de contrôle » : simple texte libre — les documents ne sont plus reliés
   depuis la GAMSERIE mais via la matrice article / fournisseur. */
function tdMeth(l){ return tdTxt('methode', l, 'methode', 'fam'); }
function td(k, v, html, cls){
  return '<td data-k="'+k+'" data-v="'+esc(v===null||v===undefined?'':v)+'"'+(cls?' class="'+cls+'"':'')+'>'+html+'</td>';
}
/* Colonnes Rév. et Date rév. : détaillées sur la 1re ligne du groupe, rappelées ensuite */
function cellRev(type, l, debut){
  var cle = type === 'gam' ? l.famille : l.article;
  var rv = l.rev === null || l.rev === undefined ? '—' : l.rev;
  var h;
  if(debut && l.statut !== 'archive'){
    h = '<td data-k="rev" data-v="'+rv+'" class="rv"><b>'+rv+'</b>'
      + '<button class="btn sm" data-vcle="'+esc(cle)+'" data-vtype="'+type+'" title="Gérer les révisions">…</button></td>';
  } else {
    h = '<td data-k="rev" data-v="'+rv+'" class="rv sub">'+rv+'</td>';
  }
  return h + '<td data-k="date" data-v="'+(l.dateRev?frDate(l.dateRev):'')+'" class="dt'+(debut?'':' sub')+'">'
    + (l.dateRev ? frDate(l.dateRev) : '—') + '</td>';
}
/* En-tête avec champ de filtre à droite du libellé */
function thf(tbl, k, libelle, cls){
  var F = tbl === 'gam' ? FGAM : FTS;
  return '<th'+(cls?' class="'+cls+'"':'')+'><div class="thw"><span>'+esc(libelle)+'</span>'
    + '<input class="fh" data-tbl="'+tbl+'" data-f="'+k+'" value="'+esc(F[k]||'')+'" title="Filtrer sur '+esc(libelle)+'"></div></th>';
}

/* --- TS : une ligne par couple article x point de contrôle --- */
function lignesTS(archives){
  var out = [];
  articles().forEach(function(art){
    var fam = familleOf(art), t = tsOf(art), vals = t ? t.vals : {};
    var v = version('ts', art);
    var codes = DB.gamserie.filter(function(g){ return g.famille === fam; })
      .map(function(g){ return {code:g.code, specification:g.specification}; });
    var vus = {}; codes.forEach(function(c){ vus[c.code] = 1; });
    for(var c in vals) if(!vus[c]) codes.push({code:c, specification:'(code absent de la gamme)'});
    codes.forEach(function(g){
      out.push(ligneTS(art, fam, g, vals[g.code], v ? v.rev : null, v ? v.dateEffet : '',
        v ? (v.statut === 'brouillon' ? 'brouillon' : 'valide') : 'valide',
        v ? v.statut === 'brouillon' : false, v ? (v.indiceDef||'') : ''));
    });
    if(archives){
      revisionsDe('ts', art).forEach(function(hh){
        if(v && hh.rev === v.rev) return;
        var ph = hh.photo || {};
        codes.forEach(function(g){
          if(!ph[g.code]) return;
          out.push(ligneTS(art, fam, g, ph[g.code], hh.rev, hh.dateEffet, 'archive', false, hh.indiceDef||''));
        });
      });
    }
  });
  return out;
}
function ligneTS(art, fam, g, val, rev, dateRev, statut, edit, indiceDef){
  val = val || ['','',''];
  var nom = num(val[0]), tm = num(val[1]), tp = num(val[2]);
  var brut = String(val[0]===null||val[0]===undefined?'':val[0]).trim();
  var vide = brut === '', na = brut === '-' || (isNaN(nom) && !vide);
  return {
    article:art, famille:fam||'', code:g.code, specification:g.specification||'',
    brut:val, vide:vide, na:na,
    mini:(na||vide||isNaN(nom))?null:nom+(isNaN(tm)?0:tm),
    maxi:(na||vide||isNaN(nom))?null:nom+(isNaN(tp)?0:tp),
    type: na ? 'non applicable' : vide ? 'à renseigner' : (nom===1 && tm===0 && tp===0) ? 'Booléen' : 'Valeur',
    typeSpec: String(val[3]===undefined||val[3]===null?'':val[3]).trim(),
    rev:rev, dateRev:dateRev, statut:statut, edit:edit, indiceDef:indiceDef
  };
}

function psetTS(){
  var lignes = lignesTS(TARCH);
  if(TSVUE === 'trous') lignes = lignes.filter(function(l){ return l.vide; });
  var brouillons = {}, valides = {};
  lignes.forEach(function(l){ if(l.statut==='brouillon') brouillons[l.article]=1; else if(l.statut==='valide') valides[l.article]=1; });
  var nb = Object.keys(brouillons).length;

  var h = '<div class="toolbar">'
    + '<select id="ts-vue"><option value="">Toutes les lignes</option>'
    + '<option value="trous"'+(TSVUE==='trous'?' selected':'')+'>Valeurs manquantes seulement</option></select>'
    + '<button class="btn sm" id="ts-add">Ajouter un article</button>'
    + '<label class="chk"><input type="checkbox" id="ts-arch"'+(TARCH?' checked':'')+'> Afficher les révisions archivées</label>'
    + '<span class="sp" style="margin-left:auto"></span>'
    + '<button class="btn sm" id="ts-exp">Exporter (tableau à plat)</button>'
    + '<button class="btn sm" id="ts-expl">Export format large</button>'
    + '<button class="btn sm" id="ts-imp-r">Importer (remplacer)</button>'
    + '<button class="btn sm" id="ts-imp-m">Importer (mettre à jour)</button>'
    + '</div>';
  h += '<div class="toolbar">'
    + '<span class="hint" id="ts-count">' + lignes.length + ' ligne' + (lignes.length>1?'s':'') + '</span>'
    + (Object.keys(valides).length ? '<button class="btn sm" id="ts-newrev">Ouvrir une révision (articles affichés)</button>' : '')
    + (nb ? '<button class="btn sm pri" id="ts-valrev">Valider les ' + nb + ' brouillon' + (nb>1?'s':'') + '</button>' : '')
    + '<span class="hint">Type saisie : nominal « - » = non applicable · nominal 1 avec tolérances 0 et 0 = booléen (OK / NOK) · sinon valeur à mesurer</span>'
    + '</div>';

  if(!lignes.length) return h + '<div class="empty"><b>Aucune ligne</b>Ajoutez un article, changez de filtre ou importez un fichier TS.</div>';

  var cap = 800, tronque = lignes.length > cap;
  if(tronque) lignes = lignes.slice(0, cap);

  h += '<div class="scroll tall"><table class="flat" data-tbl="ts"><thead><tr>'
    + thf('ts','rev','Rév.','w-rev') + thf('ts','date','Date rév.','w-date') + thf('ts','statut','Statut','w-st')
    + thf('ts','article','Article','w-art') + thf('ts','famille','Famille','w-fam') + thf('ts','code','Code','w-code')
    + thf('ts','spec','Spécification','') + thf('ts','typespec','Type spéc.','w-type') + thf('ts','typesaisie','Type saisie','w-type')
    + thf('ts','nom','Nominal','w-val') + thf('ts','tolm','Tol -','w-val')
    + thf('ts','tolp','Tol +','w-val') + thf('ts','mini','Mini.','w-val') + thf('ts','maxi','Maxi.','w-val')
    + '</tr></thead><tbody>';
  var prec = '';
  lignes.forEach(function(l){
    var cle = l.article + '|' + l.rev + '|' + l.statut;
    var debut = cle !== prec; prec = cle;
    h += '<tr class="'+(debut?'grp':'')+(l.vide?' trou':'')+(l.na?' napp':'')+(l.statut==='archive'?' arch':'')+'">'
      + cellRev('ts', l, debut)
      + td('statut', statutLib(l.statut), tagStatut(l.statut), 'st')
      + td('article', l.article, (debut ? '<b>'+esc(l.article)+'</b>' + (l.indiceDef ? '<br><span class="hint">'+esc(l.indiceDef)+'</span>' : '') : esc(l.article)), 'art' + (debut?'':' sub'))
      + td('famille', l.famille, esc(l.famille||'—'), 'fam')
      + td('code', l.code, esc(l.code), 'cd')
      + td('spec', l.specification, esc(l.specification), 'sp')
      + tdTypeSpec(l)
      + td('typesaisie', l.type, l.type, 'ty')
      + tdVal(l, 0, 'nom') + tdVal(l, 1, 'tolm') + tdVal(l, 2, 'tolp')
      + td('mini', l.mini===null?'':fmt(l.mini), l.mini===null?'—':fmt(l.mini), 'bn')
      + td('maxi', l.maxi===null?'':fmt(l.maxi), l.maxi===null?'—':fmt(l.maxi), 'bn')
      + '</tr>';
  });
  h += '</tbody></table></div>';
  if(tronque) h += '<p class="hint" style="margin-top:10px">Affichage limité à '+cap+' lignes : affinez les filtres pour voir le reste.</p>';
  return h;
}
function tdVal(l, k, colonne){
  var v = l.brut[k];
  v = (v===null||v===undefined) ? '' : v;
  if(!l.edit) return td(colonne, v, esc(v===''?'—':v), 'bn');
  return '<td data-k="'+colonne+'" data-v="'+esc(v)+'" class="mono"><input data-tsa="'+esc(l.article)+'" data-ts="'+esc(l.code)+'" data-k="'+k+'" data-cell="'+colonne+'" inputmode="decimal" value="'+esc(v)+'"></td>';
}
/* Colonne TYPE SPEC : Générale / Individuelle / (vide) — stockée en 4e position du TS (index 3) */
function tdTypeSpec(l){
  var v = l.typeSpec || '';
  if(!l.edit) return td('typespec', v, esc(v||'—'), 'ty');
  var opts = ['','Générale','Individuelle'].map(function(o){
    return '<option value="'+esc(o)+'"'+(v===o?' selected':'')+'>'+(o||'—')+'</option>';
  }).join('');
  return '<td data-k="typespec" data-v="'+esc(v)+'" class="pz"><select data-tsa="'+esc(l.article)+'" data-ts="'+esc(l.code)+'" data-k="3" data-cell="typespec">'+opts+'</select></td>';
}

function majBornes(el, val){
  var tr = el.parentNode ? el.parentNode.parentNode : null;
  if(!tr) return;
  var nom = num(val[0]), tm = num(val[1]), tp = num(val[2]);
  var brut = String(val[0]===null||val[0]===undefined?'':val[0]).trim();
  var ko = brut === '' || brut === '-' || isNaN(nom);
  var mini = tr.querySelector('td[data-k="mini"]'), maxi = tr.querySelector('td[data-k="maxi"]'), ty = tr.querySelector('td[data-k="typesaisie"]');
  if(mini){ mini.textContent = ko ? '—' : fmt(nom + (isNaN(tm)?0:tm)); mini.setAttribute('data-v', ko?'':fmt(nom + (isNaN(tm)?0:tm))); }
  if(maxi){ maxi.textContent = ko ? '—' : fmt(nom + (isNaN(tp)?0:tp)); maxi.setAttribute('data-v', ko?'':fmt(nom + (isNaN(tp)?0:tp))); }
  var t = brut === '-' ? 'non applicable' : brut === '' ? 'à renseigner' : (nom===1 && tm===0 && tp===0) ? 'Booléen' : 'Valeur';
  if(ty){ ty.textContent = t; ty.setAttribute('data-v', t); }
  tr.className = tr.className.replace(/ ?trou| ?napp/g,'') + (brut === '' ? ' trou' : brut === '-' ? ' napp' : '');
}

/* --- filtres d'en-tête : masquage des lignes en direct, sans re-rendu --- */
function applyFilters(tbl){
  var F = tbl === 'gam' ? FGAM : FTS;
  var t = document.querySelector('table.flat[data-tbl="'+tbl+'"]');
  if(!t) return;
  var rows = t.querySelectorAll('tbody tr'), n = 0;
  for(var i=0;i<rows.length;i++){
    var ok = true;
    for(var k in F){
      var q = String(F[k]||'').trim().toLowerCase();
      if(!q) continue;
      var c = rows[i].querySelector('td[data-k="'+k+'"]');
      var v = c ? String(c.getAttribute('data-v')||'').toLowerCase() : '';
      if(v.indexOf(q) < 0){ ok = false; break; }
    }
    rows[i].style.display = ok ? '' : 'none';
    if(ok) n++;
  }
  var c2 = document.getElementById(tbl === 'gam' ? 'gam-count' : 'ts-count');
  if(c2) c2.textContent = n + ' ligne' + (n>1?'s':'') + (n < rows.length ? ' sur ' + rows.length : '');
}
/* Clés (familles ou articles) visibles après filtrage, hors archives */
function clesAffichees(tbl){
  var out = {}, t = document.querySelector('table.flat[data-tbl="'+tbl+'"]');
  if(!t) return [];
  var rows = t.querySelectorAll('tbody tr');
  for(var i=0;i<rows.length;i++){
    if(rows[i].style.display === 'none' || rows[i].className.indexOf('arch') > -1) continue;
    var c = rows[i].querySelector('td[data-k="'+(tbl==='gam'?'famille':'article')+'"]');
    if(c) out[c.getAttribute('data-v')] = 1;
  }
  return Object.keys(out);
}

/* --- actions groupées de révision --- */
function nouvellesRevisions(type, cles){
  cles = cles.filter(function(c){ var v = version(type, c); return !v || v.statut === 'applicable'; });
  if(!cles.length){ toast('Aucun élément à réviser dans la sélection.'); return; }
  modal('Ouvrir une révision',
    'Une nouvelle révision en brouillon va être ouverte pour <b>' + cles.length + (type==='gam'?' famille':' article') + (cles.length>1?'s':'') + '</b>.'
    + '<p style="margin:10px 0 0">Les données deviendront modifiables ; elles ne s\'appliqueront aux nouveaux rapports qu\'après validation.</p>',
    function(){
      cles.forEach(function(c){ nouvelleRevision(type, c); });
      syncLignes(); saveDraft(); render();
      toast(cles.length + ' révision' + (cles.length>1?'s':'') + ' ouverte' + (cles.length>1?'s':'') + ' en brouillon.');
    }, 'Ouvrir les révisions');
}
function validerBrouillons(type, cles){
  cles = cles.filter(function(c){ var v = version(type, c); return v && v.statut === 'brouillon'; });
  if(!cles.length){ toast('Aucun brouillon dans la sélection.'); return; }
  modal('Valider ' + cles.length + ' brouillon' + (cles.length>1?'s':''),
    '<label>Validée par<br><input id="v-par" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="Nom du valideur"></label>'
    + '<div style="margin-top:10px"><label>Date d\'effet<br><input type="date" id="v-date" value="'+todayISO()+'" style="width:100%;padding:6px;border:1px solid var(--rule)"></label></div>'
    + '<div style="margin-top:10px"><label>Motif de la révision<br><input id="v-motif" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="ex. mise à jour plan client"></label></div>'
    + (type === 'ts' ? '<div style="margin-top:10px"><label>Indice de définition (vide = conserver celui de chaque article)<br><input id="v-def" style="width:100%;padding:6px;border:1px solid var(--rule)" placeholder="ex. plan rév. D"></label></div>' : '')
    + '<p style="margin:12px 0 0">Concernés : ' + esc(cles.slice(0,6).join(', ')) + (cles.length>6 ? ' … (+' + (cles.length-6) + ')' : '') + '.</p>',
    function(){}, 'Annuler');
  var ft = document.querySelector('.modal .ft');
  ft.innerHTML = '<button class="btn" data-x2>Annuler</button><button class="btn pri" data-ok2>Valider</button>';
  ft.querySelector('[data-x2]').onclick = function(){ document.querySelector('.modal').remove(); };
  ft.querySelector('[data-ok2]').onclick = function(){
    var par = document.getElementById('v-par').value.trim();
    var mot = document.getElementById('v-motif').value.trim();
    var dt = document.getElementById('v-date').value;
    var defEl = document.getElementById('v-def'), def = defEl ? defEl.value.trim() : '';
    if(!par || !mot){ toast('Valideur et motif sont obligatoires.'); return; }
    cles.forEach(function(c){
      var v = version(type, c);
      if(!v || v.statut !== 'brouillon') return;
      v.valideePar = par; v.motif = mot; v.dateEffet = dt; v.statut = 'applicable';
      if(def) v.indiceDef = def;
      figer(type, c, mot);
      if(v.rev > 0) journaliser(type, c, v.rev - 1, 'Révision archivée', 'remplacée par la rév. ' + v.rev);
    });
    saveVersions();
    document.querySelector('.modal').remove();
    syncLignes(); saveDraft(); render();
    toast(cles.length + ' élément' + (cles.length>1?'s':'') + ' validé' + (cles.length>1?'s':'') + ' au ' + frDate(dt));
  };
}
/* Menu de révision d'une famille ou d'un article */
function menuRevision(type, cle){
  var v = version(type, cle);
  var h = '<p style="margin:0">' + (type==='gam'?'Famille':'Article') + ' <b>' + esc(cle) + '</b> — ' + (v ? libRev(v) : 'sans révision')
    + (v && v.statut === 'applicable' && v.dateEffet ? ', valide depuis le ' + frDate(v.dateEffet) : '') + '.</p>';
  modal('Révisions de ' + cle, h, function(){}, 'Fermer');
  var ft = document.querySelector('.modal .ft');
  var boutons = '<button class="btn" data-x2>Fermer</button>';
  if(!v || v.statut === 'applicable') boutons += '<button class="btn pri" data-new2>Ouvrir la révision ' + (v ? v.rev + 1 : 0) + '</button>';
  else boutons += '<button class="btn dgr" data-ab2>Abandonner</button><button class="btn pri" data-val2>Valider la révision ' + v.rev + '</button>';
  boutons += '<button class="btn" data-h2>Historique</button>';
  ft.innerHTML = boutons;
  function close(){ var m = document.querySelector('.modal'); if(m) m.remove(); }
  ft.querySelector('[data-x2]').onclick = close;
  var e;
  if(e = ft.querySelector('[data-new2]')) e.onclick = function(){ close(); nouvelleRevision(type, cle); syncLignes(); saveDraft(); render(); toast('Révision ' + version(type,cle).rev + ' ouverte en brouillon.'); };
  if(e = ft.querySelector('[data-val2]')) e.onclick = function(){ close(); validerRevision(type, cle); };
  if(e = ft.querySelector('[data-ab2]')) e.onclick = function(){ close(); abandonnerBrouillon(type, cle); };
  if(e = ft.querySelector('[data-h2]')) e.onclick = function(){ close(); VTYPE = type; VCLE = cle; PSET = 'revisions'; window.__cmp = null; render(); };
}

function psetMatrice(){
  var cols = matriceColonnes();
  var h = '<div class="toolbar">'
    + '<button class="btn sm" id="m-add">Ajouter un article</button>'
    + '<span class="hint">'+DB.matrice.length+' article'+(DB.matrice.length>1?'s':'')+'</span>'
    + '<span class="sp" style="margin-left:auto"></span>'
    + '<button class="btn sm" id="m-exp">Exporter en CSV</button>'
    + '<button class="btn sm" id="m-imp-r">Importer (remplacer)</button>'
    + '<button class="btn sm" id="m-imp-m">Importer (mettre à jour)</button>'
    + '</div>'
    + '<p class="hint" style="margin:0 0 10px;max-width:90ch">Chaque article est relié à son fournisseur, sa famille de contrôle et ses références documentaires. '
    + 'Un PDF importé dans <b>Documents</b> et nommé d\'après une référence (ex. <b>PM-0081.pdf</b>) s\'affiche automatiquement en haut du rapport des articles qui le citent.</p>';
  if(!DB.matrice.length) return h + '<div class="empty"><b>Matrice vide</b>Importez la matrice article / fournisseur.</div>';
  h += '<div class="scroll tall"><table class="flat" data-tbl="mat"><thead><tr>';
  cols.forEach(function(c){ h += '<th'+(c.k==='article'?' class="w-art"':'')+'>'+esc(c.lib)+'</th>'; });
  h += '<th class="w-act"></th></tr></thead><tbody>';
  DB.matrice.forEach(function(m,i){
    h += '<tr>';
    cols.forEach(function(c){
      h += '<td class="pz"><input data-b="matrice" data-i="'+i+'" data-f="'+c.k+'" value="'+esc(m[c.k]||'')+'"'
        + (c.k==='article'?' style="font-weight:600"':'') + '></td>';
    });
    h += '<td class="act3"><button class="btn sm dgr" data-delrow="matrice" data-i="'+i+'">Suppr.</button></td></tr>';
  });
  h += '</tbody></table></div>';
  return h;
}
function tdInput(base, i, f, v, edit){
  return '<td><input data-b="'+base+'" data-i="'+i+'" data-f="'+f+'" value="'+esc(v)+'"'
    + (edit === false ? ' readonly class="ro"' : '') + '></td>';
}

function psetRevisions(){
  var cles = VTYPE === 'gam' ? familles() : DB.ts.map(function(t){ return t.article; }).sort();
  if(VCLE && cles.indexOf(VCLE) < 0) VCLE = '';
  if(!VCLE && cles.length) VCLE = cles[0];
  var h = '<div class="toolbar">'
    + '<select id="v-type"><option value="gam"'+(VTYPE==='gam'?' selected':'')+'>Gammes de contrôle (GAMSERIE) — par famille</option>'
    + '<option value="ts"'+(VTYPE==='ts'?' selected':'')+'>Valeurs (TS) — par article</option></select>'
    + '<select id="v-cle">' + cles.map(function(c){ return '<option value="'+esc(c)+'"'+(VCLE===c?' selected':'')+'>'+esc(c)+'</option>'; }).join('') + '</select>'
    + '<span class="sp" style="margin-left:auto"></span>'
    + '<button class="btn sm" id="v-jrn">Exporter le journal (CSV)</button>'
    + '</div>';
  if(!VCLE) return h + '<div class="empty"><b>Rien à versionner</b>Importez d\'abord GAMSERIE et TS.</div>';

  h += barreVersion(VTYPE, VCLE);

  var revs = revisionsDe(VTYPE, VCLE);
  var v = version(VTYPE, VCLE);
  var courant = v && v.statut === 'brouillon' ? [{rev:v.rev, brouillon:true, dateEffet:'', valideePar:'', motif:v.motif||'', indiceDef:v.indiceDef||''}] : [];
  var lignes = courant.concat(revs);
  h += '<table class="grid" data-tbl="rev"><thead><tr><th style="width:80px">Révision</th><th style="width:110px">Date d\'effet</th>'
     + '<th style="width:150px">Validée par</th><th>Motif</th>' + (VTYPE==='ts' ? '<th style="width:150px">Indice de définition</th>' : '')
     + '<th style="width:130px">Rapports liés</th><th style="width:110px"></th></tr></thead><tbody>';
  lignes.forEach(function(r){
    var lies = REPORTS.filter(function(x){
      return VTYPE === 'gam' ? (x.famille === VCLE && x.gamRev === r.rev) : (x.article === VCLE && x.tsRev === r.rev);
    }).length;
    var etat = r.brouillon ? '<span class="tag d">brouillon</span>' : (v && v.rev === r.rev ? '<span class="tag a">applicable</span>' : '<span class="tag d">périmée</span>');
    h += '<tr><td style="padding:6px 9px"><b style="font-family:var(--mono);font-size:14px">rév. '+r.rev+'</b> '+etat+'</td>'
      + '<td style="padding:6px 9px">'+(r.dateEffet?frDate(r.dateEffet):'—')+'</td>'
      + '<td style="padding:6px 9px">'+esc(r.valideePar||'—')+'</td>'
      + '<td style="padding:6px 9px">'+esc(r.motif||'—')+'</td>'
      + (VTYPE==='ts' ? '<td style="padding:6px 9px">'+esc(r.indiceDef||'—')+'</td>' : '')
      + '<td style="padding:6px 9px">'+(lies?lies+' rapport'+(lies>1?'s':''):'<span class="hint">aucun</span>')+'</td>'
      + '<td class="act">'+(r.rev>0?'<button class="btn sm" data-cmp="'+r.rev+'">Comparer à '+(r.rev-1)+'</button>':'')+'</td></tr>';
  });
  h += '</tbody></table>';

  if(window.__cmp !== undefined && window.__cmp !== null){
    var b = window.__cmp, a = b - 1;
    var diffs = comparer(VTYPE, VCLE, a, b);
    h += '<h3 class="soustitre">'
      + 'Révision ' + a + ' → révision ' + b + ' · ' + esc(VCLE) + '</h3>';
    if(!diffs.length) h += '<p class="hint">Aucun écart entre ces deux révisions.</p>';
    else{
      h += '<div class="diff"><div class="d head"><div>Code</div><div>Nature</div><div>Écart</div></div>';
      diffs.forEach(function(d){
        h += '<div class="d"><div style="font-family:var(--display);font-weight:700;font-size:15px">'+esc(d.code)+'</div>'
          + '<div><span class="tg '+d.type+'">'+(d.type==='ajout'?'ajout':d.type==='retrait'?'retrait':'modif.')+'</span></div>'
          + '<div>'+esc(d.detail)+'</div></div>';
      });
      h += '</div>';
    }
  }

  var jrn = JOURNAL.filter(function(j){ return j.type === VTYPE && j.cle === VCLE; }).slice(-12).reverse();
  if(jrn.length){
    h += '<h3 class="soustitre">Journal</h3><div class="rowlist">';
    jrn.forEach(function(j){
      h += '<div style="display:grid;grid-template-columns:150px 160px 90px 1fr;border-bottom:1px solid var(--rule2);font-size:12px">'
        + '<div style="padding:6px 9px;border-right:1px solid var(--rule2);color:var(--ink2)">'+frDate(String(j.le).slice(0,10))+' '+String(j.le).slice(11,16)+'</div>'
        + '<div style="padding:6px 9px;border-right:1px solid var(--rule2)">'+esc(j.action)+'</div>'
        + '<div style="padding:6px 9px;border-right:1px solid var(--rule2);font-family:var(--mono)">rév. '+j.rev+'</div>'
        + '<div style="padding:6px 9px;color:var(--ink2)">'+esc(j.detail)+'</div></div>';
    });
    h += '</div>';
  }
  return h;
}
function exportJournal(){ download('journal-revisions-' + todayISO() + '.csv', toCSV(tableauJournal())); }
function tableauJournal(){
  var rows = [['Date','Type','Clé','Révision','Évènement','Détail']];
  JOURNAL.forEach(function(j){
    rows.push([frDate(String(j.le).slice(0,10)) + ' ' + String(j.le).slice(11,16),
      j.type === 'gam' ? 'GAMSERIE' : 'TS', j.cle, j.rev, j.action, j.detail]);
  });
  return rows;
}

function psetDocuments(){
  var total = docSize(), pc = Math.min(100, Math.round(total / 31457280 * 100));
  var h = '<div class="toolbar"><button class="btn sm" id="d-add">Importer des PDF</button>'
    + '<span class="hint">' + DOCS.length + ' document' + (DOCS.length>1?'s':'') + ' — ' + ko(total) + ' dans le classeur (30 Mo conseillés au maximum)</span></div>';
  h += '<div class="bar-use"><i style="width:'+pc+'%"></i></div>';
  h += '<p style="max-width:82ch;color:var(--ink2);line-height:1.55;margin:0 0 14px">'
    + 'Les PDF sont stockés dans le classeur Excel du projet. <b>Nommez chaque PDF d\'après sa référence</b> dans la matrice article / fournisseur '
    + '(ex. <b>PM-0081.pdf</b> pour la référence <b>PM-0081</b>) : il s\'affiche alors automatiquement en haut du rapport de chaque article qui le cite.</p>';
  if(!DOCS.length) return h + '<div class="empty"><b>Aucun document</b>Importez les référentiels, plans ou instructions de contrôle au format PDF.</div>';
  h += '<div class="rowlist" data-list="docs"><div class="docrow head"><div>Document</div><div>Taille</div><div>Articles liés</div><div style="padding:7px 10px;text-align:right">Actions</div></div>';
  DOCS.forEach(function(d){
    var liees = DB.matrice.filter(function(mm){
      return REFCOLS.some(function(c){ var x = docByRef(mm[c.k]); return x && x.id === d.id; });
    }).length;
    h += '<div class="docrow"><div><b>'+esc(d.nom)+'</b></div><div>'+ko(d.taille||0)+'</div>'
      + '<div>' + (liees ? liees + ' article' + (liees>1?'s':'') : '<span class="hint">aucun</span>') + '</div>'
      + '<div style="padding:5px 10px;display:flex;gap:5px;justify-content:flex-end;overflow:visible">'
        + '<button class="btn sm" data-vdoc="'+esc(d.id)+'">Afficher</button>'
        + '<button class="btn sm dgr" data-xdoc="'+esc(d.id)+'">Supprimer</button>'
      + '</div></div>';
  });
  h += '</div>';
  return h;
}

function psetSauvegarde(){
  var info = infoClasseur();
  return '<div class="kv"><span>GAMSERIE : <b>'+DB.gamserie.length+'</b> lignes</span><span>TS : <b>'+DB.ts.length+'</b> articles</span>'
    + '<span>Matrice : <b>'+DB.matrice.length+'</b> articles</span><span>Documents : <b>'+DOCS.length+'</b> ('+ko(docSize())+')</span>'
    + '<span>Rapports enregistrés : <b>'+REPORTS.length+'</b></span>'
    + '<span>Révisions figées : <b>'+HIST.length+'</b></span><span>Journal : <b>'+JOURNAL.length+'</b> évènements</span></div>'
    + '<div class="carte-classeur"><div class="cc-ico">'+ICONES.excel+'</div><div>'
    + '<div class="cc-nom">' + (info.nom ? esc(info.nom) : 'Aucun classeur choisi') + '</div>'
    + '<div class="hint">' + (info.nom
        ? (info.enregistreLe ? 'Dernier enregistrement à ' + info.enregistreLe : 'Pas encore enregistré pendant cette session')
          + (estModifie() ? ' · <b style="color:var(--amber)">modifications non enregistrées</b>' : '')
        : 'Choisissez l\'emplacement du classeur pour pouvoir enregistrer.') + '</div></div></div>'
    + (info.handle ? '<div class="partage-etat ' + (info.dossier ? 'on' : '') + '"><span><b>Travail à plusieurs</b> — '
        + (info.dossier
            ? 'activé : l\'application signale qu\'un autre utilisateur enregistre ou que le classeur est ouvert dans Excel, et enregistre dès qu\'il est disponible.'
            : 'non activé : autorisez l\'accès au dossier du classeur si plusieurs personnes l\'utilisent.')
        + '</span>' + (info.dossier ? '' : '<button class="btn sm pri" id="s-dossier">Autoriser l\'accès au dossier…</button>') + '</div>' : '')
    + '<p style="max-width:78ch;color:var(--ink2);line-height:1.6">Le projet complet (bases, rapports, révisions, documents) est enregistré dans ce classeur Excel (.xlsx). '
    + 'Cliquez sur <b>Enregistrer</b> en haut de page ou appuyez sur <b>Ctrl + S</b>. '
    + 'l\'onglet masqué « _Donnees » contient les données de l\'application : ne le modifiez pas.</p>'
    + '<p style="max-width:78ch;color:var(--ink2);line-height:1.6"><b>Modifier les données dans Excel</b> : les onglets <b>GAMSERIE</b>, <b>TS</b> et <b>Matrice</b> '
    + 'peuvent être complétés ou corrigés directement dans Excel (ajout de lignes, modification de valeurs). Enregistrez et fermez le classeur dans Excel : '
    + 'à l\'ouverture du classeur, ou dès que vous revenez dans l\'application, les modifications sont reprises. Une gamme ou des valeurs modifiées ouvrent une '
    + 'révision en brouillon, à valider ici avant de servir aux nouveaux rapports. Les onglets Rapports, Mesures et Journal sont en lecture seule.</p>'
    + '<div class="toolbar"><button class="btn pri" id="s-saveas">Choisir l\'emplacement du classeur…</button>'
    + (info.nom ? '<button class="btn" id="s-reload">Relire le classeur (Excel, autres utilisateurs)</button>' : '')
    + '<button class="btn" id="s-open">Ouvrir un autre classeur…</button>'
    + '<button class="btn" id="s-new">Nouveau classeur…</button></div>'
    + '<h3 class="soustitre">Copie de secours</h3>'
    + '<div class="toolbar"><button class="btn" id="s-exp">Exporter une copie (JSON)</button>'
    + '<button class="btn" id="s-imp">Restaurer une copie (JSON)</button>'
    + '<button class="btn dgr" id="s-reset">Réinitialiser avec les données du prototype</button></div>';
}

/* =======================================================================
   6. Interactions
   ======================================================================= */
function bind(){
  document.querySelectorAll('.nav button').forEach(function(b){
    b.onclick = function(){ VIEW = b.dataset.view; render(); };
  });

  if(VIEW === 'rapport') bindRapport();
  if(VIEW === 'historique') bindHistorique();
  if(VIEW === 'analyse') bindAnalyse();
  if(VIEW === 'parametres') bindParametres();

  setupColResize();
}

/* =======================================================================
   Redimensionnement des colonnes (toutes les tables + listes en grille).
   Largeur mémorisée en session (COLW) et réappliquée à chaque rendu ;
   largeur minimale dynamique = contenu le plus large de la colonne.
   ======================================================================= */
var COLW = {};   // COLW[clé] = { n:<nb colonnes>, w:[largeurs px] }
var _rzCv, _rzCx;
function _measureText(txt, font){
  if(!_rzCv){ _rzCv = document.createElement('canvas'); _rzCx = _rzCv.getContext('2d'); }
  _rzCx.font = font; return _rzCx.measureText(String(txt)).width;
}
function _cellText(cell){
  var f = cell.querySelector('input,select,textarea');
  if(f){
    if(f.tagName === 'SELECT'){ var o = f.options[f.selectedIndex]; return o ? o.text : ''; }
    return f.value || f.placeholder || '';
  }
  return cell.textContent || '';
}
/* Largeur minimale d'une colonne = largeur du contenu le plus large (plafonnée). */
function colMin(cells, floor){
  var min = floor || 34, font = null;
  for(var i=0;i<cells.length;i++){
    var cell = cells[i]; if(!cell) continue;
    if(font === null){ var cs = getComputedStyle(cell); font = cs.fontWeight+' '+cs.fontSize+' '+cs.fontFamily; }
    var txt = String(_cellText(cell)).replace(/\s+/g,' ').trim();
    if(txt) min = Math.max(min, Math.ceil(_measureText(txt, font)) + 24);
  }
  return Math.min(min, 460);
}
function _startDrag(getW, setW, min){
  return function(e){
    e.preventDefault(); e.stopPropagation();
    var startX = (e.touches?e.touches[0].clientX:e.clientX), startW = getW();
    var handle = e.currentTarget; if(handle.classList) handle.classList.add('on');
    document.body.classList.add('rz-active');
    function mv(ev){
      var x = ev.touches?ev.touches[0].clientX:ev.clientX;
      setW(Math.max(min, Math.round(startW + (x - startX))));
    }
    function up(){
      document.removeEventListener('mousemove',mv); document.removeEventListener('mouseup',up);
      document.removeEventListener('touchmove',mv); document.removeEventListener('touchend',up);
      document.body.classList.remove('rz-active'); if(handle.classList) handle.classList.remove('on');
    }
    document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
    document.addEventListener('touchmove',mv,{passive:false}); document.addEventListener('touchend',up);
  };
}
function _addHandle(host, getW, setW, minFn){
  if(host.querySelector('.rzh')) return;
  var hd = document.createElement('span'); hd.className = 'rzh'; host.appendChild(hd);
  function begin(e){ _startDrag(getW, setW, minFn())(e); }
  hd.addEventListener('mousedown', begin);
  hd.addEventListener('touchstart', begin, {passive:false});
}
function setupColResize(){
  /* --- tables <table> (flat + grid) — largeur fixe + poignées --- */
  document.querySelectorAll('table.flat[data-tbl], table.grid[data-tbl]').forEach(function(tbl){
    var head = tbl.tHead && tbl.tHead.rows[0]; if(!head) return;
    var key = 'T:' + tbl.getAttribute('data-tbl'), ths = head.cells, n = ths.length;
    var nat = [];
    for(var c=0;c<n;c++) nat.push(ths[c].offsetWidth);   // largeurs naturelles (layout auto) avant de figer
    tbl.classList.add('rzt');
    var store = (COLW[key] && COLW[key].n === n) ? COLW[key].w : null;
    var total = 0;
    function syncW(){ var s=0; for(var i=0;i<n;i++) s += ths[i].offsetWidth; tbl.style.width = s+'px'; }
    for(var c=0;c<n;c++){
      (function(c){
        var th = ths[c];
        var w = (store && store[c]) ? store[c] : nat[c];
        th.style.width = w+'px'; total += w;
        _addHandle(th,
          function(){ return th.offsetWidth; },
          function(w){ th.style.width = w+'px'; COLW[key] = COLW[key] || {n:n,w:[]}; COLW[key].n = n; COLW[key].w[c] = w; syncW(); },
          function(){ var cc=[]; for(var r=0;r<tbl.rows.length;r++){ var x=tbl.rows[r].cells[c]; if(x) cc.push(x); } return colMin(cc,34); });
      })(c);
    }
    tbl.style.minWidth = '0'; tbl.style.width = total + 'px';   // largeur figée = somme des colonnes
  });
  /* --- listes en grille CSS (rapports enregistrés, documents) --- */
  document.querySelectorAll('.rowlist[data-list]').forEach(function(rl){
    var isDoc = rl.getAttribute('data-list') === 'docs';
    var rowSel = isDoc ? '.docrow' : '.r', varName = isDoc ? '--dcols' : '--cols';
    var head = rl.querySelector(rowSel + '.head'); if(!head) return;
    var cells = Array.prototype.slice.call(head.children), n = cells.length;
    var key = 'L:' + rl.getAttribute('data-list');
    var rows = Array.prototype.slice.call(rl.querySelectorAll(rowSel));
    var store = (COLW[key] && COLW[key].n === n) ? COLW[key].w : null;
    var widths = [];
    for(var c=0;c<n;c++) widths.push((store && store[c]) ? store[c] : Math.round(cells[c].getBoundingClientRect().width));
    function apply(){ rl.style.setProperty(varName, widths.map(function(w){ return w+'px'; }).join(' ')); }
    apply();
    cells.forEach(function(cell, c){
      cell.setAttribute('data-rzc','1');
      if(c === n-1) return;   // pas de poignée après la dernière colonne
      _addHandle(cell,
        function(){ return widths[c]; },
        function(w){ widths[c] = w; apply(); COLW[key] = {n:n, w:widths.slice()}; },
        function(){ return colMin(rows.map(function(r){ return r.children[c]; }), 40); });
    });
  });
}

function bindRapport(){
  ['lot','commande','fournisseur','controleur','date','qteCommandee','qteReceptionnee','dureeControle','commentaires'].forEach(function(f){
    var el = document.getElementById('f-'+f);
    if(el) el.oninput = function(){
      R[f] = el.value;
      if(el.classList) el.classList.remove('miss');
      enregistrerAuto();
    };
  });
  var art = document.getElementById('f-article');
  if(art) art.oninput = art.onchange = function(){
    if(R.article === art.value.trim()) return;
    R.article = art.value.trim(); syncLignes(); enregistrerAuto(true); render();
  };
  var qe = document.getElementById('f-qteEchantillon');
  if(qe) qe.onchange = function(){ R.qteEchantillon = qe.value; resizeMesures(); enregistrerAuto(true); render(); };

  /* Origine du contrôle */
  var org = document.getElementById('f-origine');
  if(org) org.onchange = function(){ R.origine = org.value; enregistrerAuto(true); };

  /* Statut du rapport : sortie d'« Attente décision méthodes » protégée par mot de passe,
     et contrôle des cases vides avant Accepté / Refusé */
  var st = document.getElementById('f-statut');
  if(st) st.onchange = function(){
    var nv = st.value;
    if(R.decision === 'Attente décision méthodes' && nv !== R.decision){
      var mdp = prompt('Sortie du statut « Attente décision méthodes » réservée aux méthodistes.\nMot de passe :', '');
      if(mdp === null){ st.value = R.decision; return; }
      if(mdp !== '1'){ toast('Mot de passe incorrect.'); st.value = R.decision; return; }
    }
    if(nv === 'Accepté' || nv === 'Refusé'){
      var vd = casesVides();
      if(vd > 0){ toast('Impossible : ' + vd + ' mesure' + (vd>1?'s':'') + ' non renseignée' + (vd>1?'s':'') + '. Complétez toutes les cases avant d\'accepter ou refuser.'); st.value = R.decision; return; }
    }
    R.decision = nv;
    if(nv !== 'Refusé'){ R.typeNC = ''; R.sousTypeNC = ''; }
    enregistrerAuto(true); render();
  };
  /* Type / sous-type de NC (statut Refusé) */
  var tnc = document.getElementById('f-typenc');
  if(tnc) tnc.onchange = function(){ R.typeNC = tnc.value; R.sousTypeNC = ''; enregistrerAuto(true); render(); };
  var snc = document.getElementById('f-soustypenc');
  if(snc) snc.onchange = function(){ R.sousTypeNC = snc.value; enregistrerAuto(true); };
  /* Réouverture d'un rapport figé (méthodiste, mot de passe « 1 ») */
  var rv = document.getElementById('b-rouvrir');
  if(rv) rv.onclick = function(){
    var mdp = prompt('Réouverture réservée aux méthodistes.\nMot de passe :', '');
    if(mdp === null) return;
    if(mdp !== '1'){ toast('Mot de passe incorrect.'); return; }
    R.decision = ''; enregistrerAuto(true); render(); toast('Rapport rouvert (statut : sans décision).');
  };
  /* Déclaration d'un EQf dans X3 (message + acquittement) */
  var eqf = document.getElementById('b-eqf');
  if(eqf) eqf.onclick = function(){ modal('EQf X3', '<p style="margin:0">Génère un EQf dans X3</p>', null, 'Acquitter'); };
  /* Mouvements de stock X3 selon la décision */
  var mvA = document.getElementById('b-mvtA');
  if(mvA) mvA.onclick = function(){ modal('Mouvement X3 — statut A', '<p style="margin:0">Génère un changement de stock X3 du statut Q au statut A</p>', null, 'Acquitter'); };
  var mvR = document.getElementById('b-mvtR');
  if(mvR) mvR.onclick = function(){ modal('Mouvement X3 — statut R', '<p style="margin:0">Génère un changement de stock X3 du statut Q au statut R</p>', null, 'Acquitter'); };
  /* Numéros de moyen de contrôle (par point) */
  document.querySelectorAll('input.cmoy, input.genmoy').forEach(function(el){
    el.oninput = function(){ R.moyens = R.moyens || {}; R.moyens[el.dataset.moy] = el.value; enregistrerAuto(); };
  });
  /* Ouverture de la carte de contrôle en fenêtre superposée */
  var bc = document.getElementById('b-carte');
  if(bc) bc.onclick = ouvrirCarteControle;

  document.querySelectorAll('td.mes input').forEach(function(inp){
    inp.oninput = function(){
      R.mesures[+inp.dataset.r][+inp.dataset.c] = inp.value;
      enregistrerAuto(); refresh();
    };
    /* à la validation de la cellule (Entrée, Tab, clic ailleurs) : points -> virgules */
    inp.onchange = function(){
      var v = formatMesure(inp.value);
      if(v === inp.value) return;
      inp.value = v;
      R.mesures[+inp.dataset.r][+inp.dataset.c] = v;
      enregistrerAuto();
    };
    inp.onkeydown = navKey;
  });
  document.querySelectorAll('button.q:not(.genq)').forEach(function(b){
    b.onclick = function(){
      var i = +b.dataset.r, j = +b.dataset.c, v = String(R.mesures[i][j]);
      R.mesures[i][j] = (v === '1') ? '0' : (v === '0') ? '' : '1';
      enregistrerAuto(true); render();
      var next = document.querySelector('button.q[data-r="'+i+'"][data-c="'+j+'"]');
      if(next) next.focus();
    };
    b.onkeydown = function(e){
      var k = e.key.toLowerCase();
      if(k === 'o' || k === '1'){ R.mesures[+b.dataset.r][+b.dataset.c] = '1'; e.preventDefault(); enregistrerAuto(true); render(); focusCell(+b.dataset.r+1, +b.dataset.c); }
      else if(k === 'n' || k === '0'){ R.mesures[+b.dataset.r][+b.dataset.c] = '0'; e.preventDefault(); enregistrerAuto(true); render(); focusCell(+b.dataset.r+1, +b.dataset.c); }
      else navKey(e);
    };
  });

  /* Spécifications générales : un seul OK / NOK, cycle – → OK → NOK → – */
  document.querySelectorAll('button.genq').forEach(function(b){
    if(!R.generales) R.generales = {};
    b.onclick = function(){
      var c = b.dataset.gen, v = String(R.generales[c]||'');
      R.generales[c] = (v === '1') ? '0' : (v === '0') ? '' : '1';
      enregistrerAuto(true); render();
      var next = document.querySelector('button.genq[data-gen="'+c+'"]');
      if(next) next.focus();
    };
    b.onkeydown = function(e){
      var k = e.key.toLowerCase();
      if(k === 'o' || k === '1'){ R.generales[b.dataset.gen] = '1'; e.preventDefault(); enregistrerAuto(true); render(); }
      else if(k === 'n' || k === '0'){ R.generales[b.dataset.gen] = '0'; e.preventDefault(); enregistrerAuto(true); render(); }
    };
  });

  document.querySelectorAll('a.doc').forEach(function(a){
    a.onclick = function(){ ouvrirDoc(a.dataset.doc); };
    a.onkeydown = function(e){ if(e.key === 'Enter' || e.key === ' '){ e.preventDefault(); ouvrirDoc(a.dataset.doc); } };
  });

  var b;
  if(b = document.getElementById('b-rebasculer')) b.onclick = rebasculer;
  if(b = document.getElementById('b-fermer')) b.onclick = fermerRapport;
  if(b = document.getElementById('b-nouveau')) b.onclick = function(){
    if(!confirm('Ouvrir un nouveau rapport vierge ? Le rapport en cours sera perdu s\'il n\'est pas enregistré.')) return;
    R = newReport(); saveDraft(); render();
  };
  if(b = document.getElementById('b-vider')) b.onclick = function(){
    if(!confirm('Effacer toutes les mesures de ce rapport ?')) return;
    R.mesures = R.lignes.map(function(){ var r=[]; for(var j=0;j<nbPieces();j++) r.push(''); return r; });
    enregistrerAuto(true); render();
  };
  if(b = document.getElementById('b-print')) b.onclick = function(){ window.print(); };
  if(b = document.getElementById('b-csv')) b.onclick = function(){ exportRapport(R); };
}

function navKey(e){
  var i = +e.target.dataset.r, j = +e.target.dataset.c, k = e.key;
  if(k === 'Enter' || k === 'ArrowDown'){ e.preventDefault(); focusCell(i+1, j); }
  else if(k === 'ArrowUp'){ e.preventDefault(); focusCell(i-1, j); }
  else if(k === 'ArrowRight' && (e.target.selectionStart === undefined || e.target.selectionStart === String(e.target.value).length)){ focusCell(i, j+1); }
  else if(k === 'ArrowLeft' && (e.target.selectionStart === undefined || e.target.selectionStart === 0)){ focusCell(i, j-1); }
}
function focusCell(i, j){
  if(i < 0 || j < 0) return;
  var el = document.querySelector('td.mes[data-r="'+i+'"][data-c="'+j+'"] input, td.mes[data-r="'+i+'"][data-c="'+j+'"] button');
  if(el){ el.focus(); if(el.select) el.select(); }
  else if(i < R.lignes.length) focusCell(i+1, j);      // saute les lignes non applicables
}

/* mise à jour en direct des couleurs, des verdicts et des totaux */
function refresh(){
  var n = nbPieces();
  R.lignes.forEach(function(l,i){
    if(specKind(l) !== 'individuelle') return;
    var nc = 0, tr = null;
    for(var j=0;j<n;j++){
      var td = document.querySelector('td.mes[data-r="'+i+'"][data-c="'+j+'"]');
      if(!td) continue;
      if(!tr) tr = td.parentNode;
      var st = mesureStatut(i,j);
      td.className = 'mes ' + (l.testable ? st : 'off');
      if(st === 'out') nc++;
    }
    if(tr){
      var badge = tr.querySelector('.nccount');
      if(nc && !badge){ var s = document.createElement('span'); s.className='nccount'; s.textContent=nc; tr.querySelector('td.spec').appendChild(s); }
      else if(nc && badge) badge.textContent = nc;
      else if(!nc && badge) badge.remove();
    }
  });
  for(var p=0;p<n;p++){
    var th = document.querySelector('th.c-p[data-p="'+p+'"]');
    if(th){
      var ncp = ncParPiece(p), bp = th.querySelector('.nccount');
      if(ncp && !bp){ bp = document.createElement('span'); bp.className = 'nccount'; th.appendChild(bp); }
      if(ncp) bp.textContent = ncp; else if(bp) bp.remove();
    }
    var td2 = document.querySelector('tfoot td[data-p="'+p+'"]');
    if(!td2) continue;
    var s2 = pieceStatut(p);
    td2.className = s2==='ok' ? 'pok' : s2==='nc' ? 'pnc' : '';
    td2.textContent = s2==='ok' ? 'C' : s2==='nc' ? 'NC' : '·';
  }
  var t = totaux();
  var eo = document.getElementById('tot-ok'), en = document.getElementById('tot-nc');
  if(eo) eo.textContent = t.conforme;
  if(en) en.textContent = t.nonConforme;
  var hint = document.getElementById('hint-pieces');
  if(hint) hint.textContent = t.encours ? t.encours+' pièce'+(t.encours>1?'s':'')+' non commencée'+(t.encours>1?'s':'') : 'Toutes les pièces sont renseignées';
  var chartBox = document.getElementById('chart-imr');
  if(chartBox) chartBox.innerHTML = chartRapportInner();   // suivi temps réel des déviances
}

var CHAMPS_OBLIGATOIRES = [
  {f:'article', id:'f-article', nom:'Article', ok:function(v){ return String(v||'').trim() !== ''; }},
  {f:'lot', id:'f-lot', nom:'Lot', ok:function(v){ return String(v||'').trim() !== ''; }},
  {f:'qteEchantillon', id:'f-qteEchantillon', nom:'Quantité échantillon', ok:function(v){ var n = num(v); return !isNaN(n) && n >= 1; }}
];
function champsManquants(){
  return CHAMPS_OBLIGATOIRES.filter(function(c){ return !c.ok(R[c.f]); });
}
/* Enregistrement automatique : à chaque saisie, dès que les champs
   obligatoires (Article, Lot, Quantité échantillon) sont renseignés. */
var autoT = null;
function enregistrerAuto(immediat){
  if(champsManquants().length){
    R.statut = 'brouillon';
    saveDraft(); majEtatSauvegarde();
    return false;
  }
  var t = totaux();
  R.statut = 'enregistré';
  R.conforme = t.conforme; R.nonConforme = t.nonConforme;
  R.majLe = new Date().toISOString();
  saveDraft();
  clearTimeout(autoT);
  var ecrire = function(){
    var f = -1;
    for(var i=0;i<REPORTS.length;i++) if(REPORTS[i].id === R.id){ f = i; break; }
    if(f > -1) REPORTS[f] = clone(R); else REPORTS.push(clone(R));
    saveReports();
    majEtatSauvegarde();
  };
  if(immediat) ecrire(); else autoT = setTimeout(ecrire, 400);
  majEtatSauvegarde();
  return true;
}
function etatSauvegarde(){
  var m = champsManquants();
  if(m.length) return '<span class="ko">Non enregistré — ' + m.map(function(c){ return esc(c.nom); }).join(', ') + ' à renseigner</span>';
  if(estModifie()) return '<span class="okk">Rapport complet — à enregistrer dans le classeur (Ctrl+S)</span>';
  return '<span class="okk">Enregistré dans le classeur</span>';
}
function majEtatSauvegarde(){
  var e = document.getElementById('etat-save');
  if(e) e.innerHTML = etatSauvegarde();
}
/* Fermeture du rapport : retour aux rapports enregistrés */
function fermerRapport(){
  var manque = champsManquants();
  if(manque.length && (R.article || R.lot || R.lignes.length)){
    modal('Rapport non enregistré',
      'Ce rapport ne peut pas rejoindre les rapports enregistrés tant qu\'il manque :'
      + '<ul>' + manque.map(function(c){ return '<li><b>' + esc(c.nom) + '</b></li>'; }).join('') + '</ul>'
      + '<p style="margin:12px 0 0">Si vous fermez maintenant, il reste en brouillon : vous le retrouverez avec le bouton « Reprendre » de l\'écran Rapports enregistrés.</p>',
      function(){ VIEW = 'historique'; render(); }, 'Fermer quand même');
    var ft = document.querySelector('.modal .ft');
    ft.innerHTML = '<button class="btn" data-x3>Compléter</button><button class="btn pri" data-ok3>Fermer quand même</button>';
    ft.querySelector('[data-x3]').onclick = function(){
      document.querySelector('.modal').remove();
      manque.forEach(function(c){ var el = document.getElementById(c.id); if(el) el.classList.add('miss'); });
      var first = document.getElementById(manque[0].id);
      if(first){ first.focus(); if(first.select) first.select(); }
    };
    ft.querySelector('[data-ok3]').onclick = function(){ document.querySelector('.modal').remove(); VIEW = 'historique'; render(); };
    return;
  }
  enregistrerAuto(true);
  VIEW = 'historique'; render();
}

function bindHistorique(){
  var q = document.getElementById('q');
  if(q) q.oninput = function(){ window.__q = q.value; filtreRapports(); };
  filtreRapports();
  var bn = document.getElementById('h-nouveau');
  if(bn) bn.onclick = function(){
    if(brouillonEnCours() && !confirm('Un rapport en cours n\'est pas enregistré (article, lot ou quantité échantillon manquant). L\'abandonner pour en commencer un nouveau ?')) return;
    R = newReport(); saveDraft(); VIEW = 'rapport'; render();
  };
  var b;
  if(b = document.getElementById('h-analyse')) b.onclick = function(){ VIEW = 'analyse'; render(); };
  if(b = document.getElementById('h-parametres')) b.onclick = function(){ VIEW = 'parametres'; render(); };
  if(b = document.getElementById('h-reprendre')) b.onclick = function(){ VIEW = 'rapport'; render(); };
  if(b = document.getElementById('h-abandon')) b.onclick = function(){
    if(!confirm('Abandonner le rapport en cours ? Les saisies seront perdues.')) return;
    R = newReport(); saveDraft(); render();
  };
  /* un clic sur la ligne (hors boutons d'action) ouvre le rapport */
  document.querySelectorAll('.r[data-open]').forEach(function(row){
    function ouvrir(){
      if(brouillonEnCours() && R.id !== row.dataset.open && !confirm('Un rapport en cours n\'est pas enregistré. L\'abandonner pour ouvrir ce rapport ?')) return;
      R = clone(byId(row.dataset.open)); saveDraft(); VIEW = 'rapport'; render();
    }
    row.onclick = function(e){ if(!e.target.closest('button')) ouvrir(); };
    row.onkeydown = function(e){ if((e.key === 'Enter' || e.key === ' ') && e.target === row){ e.preventDefault(); ouvrir(); } };
  });
  document.querySelectorAll('[data-exp]').forEach(function(b){ b.onclick = function(){ exportRapport(byId(b.dataset.exp)); }; });
  document.querySelectorAll('[data-del]').forEach(function(b){
    b.onclick = function(){
      var r = byId(b.dataset.del);
      if(!confirm('Supprimer définitivement le rapport ' + (r.article||'') + (r.lot?' / lot '+r.lot:'') + ' ?')) return;
      REPORTS = REPORTS.filter(function(x){ return x.id !== b.dataset.del; });
      saveReports(); render(); toast('Rapport supprimé.');
    };
  });
}
/* Filtrage des rapports enregistrés sans re-rendu : le curseur reste en place */
function filtreRapports(){
  var q = String(window.__q || '').trim().toLowerCase();
  var rows = document.querySelectorAll('.rowlist .r[data-s]'), n = 0;
  for(var i=0;i<rows.length;i++){
    var ok = !q || String(rows[i].getAttribute('data-s')||'').indexOf(q) > -1;
    rows[i].style.display = ok ? '' : 'none';
    if(ok) n++;
  }
  var c = document.getElementById('q-count');
  if(c) c.textContent = n + ' rapport' + (n>1?'s':'') + (q ? ' sur ' + rows.length : '');
}
/* Rapport en cours de saisie qui n'a pas encore rejoint les rapports enregistrés */
function brouillonEnCours(){
  return !!(R && (R.article || R.lot) && !byId(R.id));
}
var ICO_MENU = {
  analyse:'<svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><path d="M4 19h16M6 15l4-5 3 3 5-7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  reglages:'<svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><path d="M4 7h10M18 7h2M4 17h4M12 17h8" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="16" cy="7" r="2" fill="none" stroke="currentColor" stroke-width="1.8"/><circle cx="10" cy="17" r="2" fill="none" stroke="currentColor" stroke-width="1.8"/></svg>',
  plus:'<svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>'
};
function byId(id){ for(var i=0;i<REPORTS.length;i++) if(REPORTS[i].id === id) return REPORTS[i]; return null; }

function bindParametres(){
  var pf = document.getElementById('p-fermer');
  if(pf) pf.onclick = function(){ VIEW = 'historique'; render(); };
  document.querySelectorAll('[data-pset]').forEach(function(b){ b.onclick = function(){ PSET = b.dataset.pset; render(); }; });

  document.querySelectorAll('[data-b]').forEach(function(el){
    el.oninput = el.onchange = function(){
      DB[el.dataset.b][+el.dataset.i][el.dataset.f] = el.value;
      if(el.dataset.cell && el.parentNode) el.parentNode.setAttribute('data-v', el.value);
      saveDB(); syncLignes(); saveDraft();
    };
  });
  document.querySelectorAll('[data-delrow]').forEach(function(b){
    b.onclick = function(){
      if(!confirm('Supprimer cette ligne ?')) return;
      DB[b.dataset.delrow].splice(+b.dataset.i, 1);
      saveDB(); syncLignes(); saveDraft(); render();
    };
  });
  document.querySelectorAll('[data-tsa]').forEach(function(el){
    el.oninput = el.onchange = function(){
      var art = el.dataset.tsa, c = el.dataset.ts, k = +el.dataset.k;
      var t = tsOf(art);
      if(!t){ t = {article:art, vals:{}}; DB.ts.push(t); }
      if(!t.vals[c]) t.vals[c] = ['','',''];
      var n = num(el.value);
      t.vals[c][k] = (el.value.trim() === '') ? '' : (isNaN(n) ? el.value : n);
      saveDB(); syncLignes(); saveDraft();
      if(el.parentNode) el.parentNode.setAttribute('data-v', el.value);
      majBornes(el, t.vals[c]);
    };
  });

  var b;

  /* filtres d'en-tête : masquage en direct, sans re-rendu (le focus est conservé) */
  document.querySelectorAll('input.fh').forEach(function(el){
    el.oninput = function(){
      var F = el.dataset.tbl === 'gam' ? FGAM : FTS;
      F[el.dataset.f] = el.value;
      applyFilters(el.dataset.tbl);
    };
  });
  applyFilters('gam'); applyFilters('ts');

  document.querySelectorAll('[data-vcle]').forEach(function(x){
    x.onclick = function(){ menuRevision(x.dataset.vtype, x.dataset.vcle); };
  });

  if(b = document.getElementById('g-arch')) b.onchange = function(){ GARCH = b.checked; render(); };
  if(b = document.getElementById('g-add')) b.onclick = function(){
    var brouillons = familles().filter(function(f){ return modifiable('gam', f); });
    if(!brouillons.length){ toast('Ouvrez d\'abord une révision sur la famille à compléter.'); return; }
    var f = brouillons.length === 1 ? brouillons[0] : prompt('Famille à compléter (' + brouillons.join(', ') + ') :', brouillons[0]);
    if(!f || !modifiable('gam', f.trim())){ toast('Famille non modifiable : sa révision doit être en brouillon.'); return; }
    DB.gamserie.push({famille:f.trim(), code:'', criticite:'Majeure', specification:'', moyen:'', methode:''});
    saveDB(); render();
  };
  if(b = document.getElementById('g-newrev')) b.onclick = function(){ nouvellesRevisions('gam', clesAffichees('gam')); };
  if(b = document.getElementById('g-valrev')) b.onclick = function(){ validerBrouillons('gam', clesAffichees('gam')); };
  if(b = document.getElementById('g-exp')) b.onclick = exportGamserie;
  if(b = document.getElementById('g-imp-r')) b.onclick = function(){ importCSV('gamserie','remplacer'); };
  if(b = document.getElementById('g-imp-m')) b.onclick = function(){ importCSV('gamserie','maj'); };

  if(b = document.getElementById('ts-vue')) b.onchange = function(){ TSVUE = b.value; render(); };
  if(b = document.getElementById('ts-arch')) b.onchange = function(){ TARCH = b.checked; render(); };
  if(b = document.getElementById('ts-add')) b.onclick = function(){
    var a = prompt('Référence du nouvel article :', '');
    if(!a) return; a = a.trim();
    if(!a) return;
    if(!tsOf(a)) DB.ts.push({article:a, vals:{}});
    if(!familleOf(a)){
      var f = prompt('Famille de contrôle pour ' + a + ' :', familles()[0] || '');
      if(f) DB.matrice.push({article:a, famille:f.trim()});
    }
    VERS.ts[a] = {rev:0, statut:'brouillon', dateEffet:'', valideePar:'', motif:'Création de l\'article', indiceDef:'', creeLe:new Date().toISOString()};
    saveDB(); saveVersions(); journaliser('ts', a, 0, 'Article créé', '');
    FTS.article = a; render();
    toast(a + ' créé en révision 0 (brouillon).');
  };
  if(b = document.getElementById('ts-newrev')) b.onclick = function(){ nouvellesRevisions('ts', clesAffichees('ts')); };
  if(b = document.getElementById('ts-valrev')) b.onclick = function(){ validerBrouillons('ts', clesAffichees('ts')); };
  if(b = document.getElementById('ts-exp')) b.onclick = exportTS;
  if(b = document.getElementById('ts-expl')) b.onclick = exportTSLarge;
  if(b = document.getElementById('ts-imp-r')) b.onclick = function(){ importCSV('ts','remplacer'); };
  if(b = document.getElementById('ts-imp-m')) b.onclick = function(){ importCSV('ts','maj'); };

  if(b = document.getElementById('m-add')) b.onclick = function(){ DB.matrice.push({article:'', famille:''}); saveDB(); render(); };
  if(b = document.getElementById('m-exp')) b.onclick = exportMatrice;
  if(b = document.getElementById('m-imp-r')) b.onclick = function(){ importCSV('matrice','remplacer'); };
  if(b = document.getElementById('m-imp-m')) b.onclick = function(){ importCSV('matrice','maj'); };

  document.querySelectorAll('[data-newrev]').forEach(function(x){
    x.onclick = function(){
      var ty = x.dataset.newrev, c = x.dataset.cle, v = version(ty, c);
      modal('Ouvrir une nouvelle révision',
        (v ? 'La <b>révision ' + v.rev + '</b> de « ' + esc(c) + ' » est applicable, donc figée : elle a pu servir à des rapports déjà enregistrés.<p style="margin:10px 0 0">'
           : 'Aucune révision enregistrée pour « ' + esc(c) +' ».<p style="margin:10px 0 0">')
        + 'La <b>révision ' + (v ? v.rev + 1 : 0) + '</b> s\'ouvre en brouillon : vous pourrez la modifier librement, elle ne s\'appliquera aux nouveaux rapports qu\'une fois validée.</p>',
        function(){ nouvelleRevision(ty, c); syncLignes(); saveDraft(); render(); toast('Révision ' + version(ty,c).rev + ' ouverte en brouillon.'); },
        'Ouvrir la révision ' + (v ? v.rev + 1 : 0));
    };
  });
  document.querySelectorAll('[data-valrev]').forEach(function(x){ x.onclick = function(){ validerRevision(x.dataset.valrev, x.dataset.cle); }; });
  document.querySelectorAll('[data-abrev]').forEach(function(x){ x.onclick = function(){ abandonnerBrouillon(x.dataset.abrev, x.dataset.cle); }; });
  document.querySelectorAll('[data-hist]').forEach(function(x){
    x.onclick = function(){ VTYPE = x.dataset.hist; VCLE = x.dataset.cle; PSET = 'revisions'; window.__cmp = null; render(); };
  });
  document.querySelectorAll('[data-cmp]').forEach(function(x){ x.onclick = function(){ window.__cmp = +x.dataset.cmp; render(); }; });
  if(b = document.getElementById('v-type')) b.onchange = function(){ VTYPE = b.value; VCLE = ''; window.__cmp = null; render(); };
  if(b = document.getElementById('v-cle')) b.onchange = function(){ VCLE = b.value; window.__cmp = null; render(); };
  if(b = document.getElementById('v-jrn')) b.onclick = exportJournal;
  if(b = document.getElementById('g-newfam')) b.onclick = function(){
    var f = prompt('Nom de la nouvelle famille de contrôle :', '');
    if(!f) return; f = f.trim();
    if(familles().indexOf(f) > -1){ toast('Cette famille existe déjà.'); return; }
    DB.gamserie.push({famille:f, code:'A', criticite:'Majeure', specification:'', moyen:'', methode:''});
    VERS.gam[f] = {rev:0, statut:'brouillon', dateEffet:'', valideePar:'', motif:'Création de la famille', creeLe:new Date().toISOString()};
    saveDB(); saveVersions(); journaliser('gam', f, 0, 'Famille créée', '');
    FGAM.famille = f; render();
  };

  if(b = document.getElementById('d-add')) b.onclick = importPDF;
  document.querySelectorAll('[data-vdoc]').forEach(function(x){ x.onclick = function(){ ouvrirDoc(x.dataset.vdoc); }; });
  document.querySelectorAll('[data-xdoc]').forEach(function(x){ x.onclick = function(){ supprimerDoc(x.dataset.xdoc); }; });

  if(b = document.getElementById('s-saveas')) b.onclick = function(){ enregistrerSous(); };
  if(b = document.getElementById('s-open')) b.onclick = function(){ ouvrirClasseur(); };
  if(b = document.getElementById('s-reload')) b.onclick = function(){ reprendreExcel(); };
  if(b = document.getElementById('s-dossier')) b.onclick = function(){ choisirDossier(); };
  if(b = document.getElementById('s-new')) b.onclick = function(){ nouveauClasseur(); };
  if(b = document.getElementById('s-exp')) b.onclick = exportJSON;
  if(b = document.getElementById('s-imp')) b.onclick = importJSON;
  if(b = document.getElementById('s-reset')) b.onclick = function(){
    if(!confirm('Remplacer les 3 bases par les données du prototype Excel ? Les rapports enregistrés sont conservés.')) return;
    DB = clone(SEED); VERS = {gam:{}, ts:{}}; HIST = []; 
    saveDB(); saveVersions(); saveHist();
    journaliser('gam', '(toutes)', 0, 'Réinitialisation', 'bases remises à la version du prototype');
    initVersions(); syncLignes(); saveDraft(); render(); toast('Bases réinitialisées en révision 0.');
  };
}

var toastT = null;
function toast(msg){
  var el = document.querySelector('.toast');
  if(!el){ el = document.createElement('div'); el.className = 'toast'; document.body.appendChild(el); }
  el.textContent = msg;
  clearTimeout(toastT);
  toastT = setTimeout(function(){ el.remove(); }, 3800);
}

/* ---------- démarrage ----------
   Appelé par js/classeur.js après chargement des données (MEM) depuis le
   classeur Excel, ou avec MEM vide pour partir des données du prototype. */
function initialiserApplication(){
  VIEW = 'historique'; blobCache = {}; fermerDoc();
  loadAll();
  initVersions();          // toute famille et tout article non versionné démarre en révision 0
  if(R.article && !R.lignes.length) syncLignes();
  render();
}
document.getElementById('vp-close').onclick = fermerDoc;
document.getElementById('vp-ext').onclick = ouvrirDocExterne;
document.addEventListener('keydown', function(e){ if(e.key === 'Escape' && docOuvert && !document.querySelector('.modal')) fermerDoc(); });
