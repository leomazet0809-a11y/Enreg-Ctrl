/* =======================================================================
   Info-bulles : sur chaque bouton, une bulle précise l'action au bout
   d'une seconde de survol (ou de focus clavier). Le texte vient de
   l'attribut title (repris pour éviter la bulle native du navigateur),
   sinon du tableau ci-dessous, sinon du libellé du bouton.
   ======================================================================= */
(function(){
  var DELAI = 1000;

  /* Textes par identifiant de bouton */
  var PAR_ID = {
    'acc-ouvrir':  'Ouvrir un classeur Excel déjà créé par l\'application',
    'acc-nouveau': 'Créer un nouveau classeur Excel à l\'emplacement de votre choix',
    'acc-demo':    'Essayer l\'application avec des données d\'exemple, sans classeur',
    'acc-reprendre': 'Rouvrir le dernier classeur utilisé sur ce poste',
    'b-fermer':    'Fermer le rapport et revenir aux rapports enregistrés',
    'pp-fermer':   'Fermer les paramètres',
    'p-fermer':    'Fermer et revenir aux rapports enregistrés',
    'a-fermer':    'Fermer les analyses et revenir aux rapports enregistrés',
    'r-actualiser':'Recharger toutes les données du classeur Excel dans l\'application, sans rien enregistrer dans le classeur',
    'h-nouveau':   'Commencer un nouveau rapport de contrôle',
    'h-analyse':   'Analyses des dérives : cartes de contrôle et tendances des mesures',
    'h-referentiel':'Référentiel : gammes, valeurs TS, matrice, cohérence, révisions et documents',
    'h-parametres':'Paramètres : sauvegarde du classeur et affichage (fenêtre en haut à droite)',
    'h-reprendre': 'Reprendre la saisie du rapport en cours',
    'h-abandon':   'Abandonner le rapport en cours (les saisies sont perdues)',
    'h-incoh':     'Voir les incohérences du paramétrage',
    'b-carte':     'Afficher la carte de contrôle (I-MR) de la cote choisie',
    'b-vider':     'Effacer toutes les mesures saisies dans ce rapport',
    'b-csv':       'Exporter ce rapport dans un fichier CSV',
    'b-print':     'Imprimer le rapport ou l\'enregistrer en PDF',
    'b-rebasculer':'Appliquer à ce rapport la révision de gamme et de valeurs en vigueur',
    'a-reset':     'Effacer tous les filtres de l\'analyse',
    'g-newfam':    'Créer une nouvelle famille de contrôle (gamme)',
    'g-add':       'Ajouter un point de contrôle à une famille en révision brouillon',
    'g-exp':       'Exporter les gammes (GAMSERIE) en CSV',
    'g-imp-r':     'Importer un fichier GAMSERIE qui remplace toutes les gammes',
    'g-imp-m':     'Importer un fichier GAMSERIE qui complète ou modifie les gammes existantes',
    'g-newrev':    'Ouvrir une révision brouillon sur les familles affichées',
    'g-valrev':    'Valider les révisions brouillon des familles affichées',
    'ts-add':      'Créer un article : ses points de contrôle sont repris de la gamme de sa famille',
    'ts-exp':      'Exporter les valeurs TS en CSV, une ligne par article et par code',
    'ts-expl':     'Exporter les valeurs TS en CSV, une ligne par article',
    'ts-imp-r':    'Importer un fichier TS qui remplace toutes les valeurs',
    'ts-imp-m':    'Importer un fichier TS qui complète ou modifie les valeurs existantes',
    'ts-newrev':   'Ouvrir une révision brouillon sur les articles affichés',
    'ts-valrev':   'Valider les révisions brouillon des articles affichés',
    'm-add':       'Créer un article : ses points de contrôle sont repris de la gamme de sa famille',
    'm-exp':       'Exporter la matrice article / fournisseur en CSV',
    'm-imp-r':     'Importer une matrice qui remplace la matrice actuelle',
    'm-imp-m':     'Importer une matrice qui complète ou modifie la matrice actuelle',
    'v-jrn':       'Exporter le journal des révisions en CSV',
    'd-add':       'Importer des documents PDF (plans, méthodes…) dans le classeur',
    's-saveas':    'Choisir le dossier et le nom du classeur Excel',
    's-open':      'Ouvrir un autre classeur Excel',
    's-new':       'Créer un nouveau classeur Excel vide',
    's-reload':    'Relire le classeur : modifications faites dans Excel ou par les autres utilisateurs',
    's-dossier':   'Autoriser l\'accès au dossier du classeur (travail à plusieurs)',
    's-exp':       'Télécharger une copie de sauvegarde au format JSON',
    's-imp':       'Restaurer les données depuis une copie de sauvegarde JSON',
    's-reset':     'Remplacer toutes les données par celles du prototype',
    't-moins':     'Réduire la taille du texte',
    't-plus':      'Agrandir la taille du texte',
    't-reset':     'Revenir à la taille de texte normale (100 %)',
    'vp-ext':      'Ouvrir le document dans le lecteur PDF du poste',
    'vp-close':    'Fermer le document'
  };
  /* Textes par attribut data-… (boutons répétés dans les listes) */
  var PAR_ATTR = [
    ['data-exp',    'Exporter ce rapport dans un fichier CSV'],
    ['data-del',    'Supprimer définitivement ce rapport'],
    ['data-delrow', 'Supprimer cette ligne'],
    ['data-newrev', 'Ouvrir une nouvelle révision en brouillon'],
    ['data-valrev', 'Valider la révision brouillon : elle s\'appliquera aux nouveaux rapports'],
    ['data-abrev',  'Abandonner la révision brouillon'],
    ['data-hist',   'Voir l\'historique des révisions'],
    ['data-cmp',    'Comparer avec cette révision'],
    ['data-pset',   null],
    ['data-incfam', 'Rattacher cet article à une famille de contrôle'],
    ['data-incts',  'Ouvrir les valeurs TS de cet article'],
    ['data-incdel', 'Retirer de TS les valeurs sur des codes absents de la gamme'],
    ['data-annuler','Revenir au rapport sans le fermer'],
    ['data-fermer', 'Fermer sans enregistrer : le contenu reste dans l\'application, à enregistrer plus tard'],
    ['data-enreg',  'Enregistrer dans le classeur Excel puis fermer'],
    ['data-doc',    'Ouvrir ce document']
  ];
  var PAR_ONGLET = {
    gamserie:'Gammes de contrôle : points, moyens et méthodes par famille',
    ts:'Valeurs par article : nominal, tolérances et type de spécification',
    matrice:'Matrice article / fournisseur, famille et références documentaires',
    coherence:'Incohérences entre la Matrice, les gammes et les valeurs TS',
    revisions:'Historique et comparaison des révisions',
    documents:'Documents PDF importés dans le classeur',
    sauvegarde:'Classeur Excel, travail à plusieurs et copies de sauvegarde',
    affichage:'Taille du texte de l\'application'
  };

  function texte(el){
    if(el.title){ el.setAttribute('data-tip', el.title); el.removeAttribute('title'); }  // pas de bulle native en double
    if(el.id && PAR_ID[el.id]) return PAR_ID[el.id];
    var t = el.getAttribute('data-tip');
    if(t) return t;
    if(el.hasAttribute('data-pset')) return PAR_ONGLET[el.getAttribute('data-pset')] || el.textContent.trim();
    if(el.hasAttribute('data-ponglet')) return PAR_ONGLET[el.getAttribute('data-ponglet')] || el.textContent.trim();
    for(var i=0;i<PAR_ATTR.length;i++) if(PAR_ATTR[i][1] && el.hasAttribute(PAR_ATTR[i][0])) return PAR_ATTR[i][1];
    return el.getAttribute('aria-label') || el.textContent.replace(/\s+/g,' ').trim();
  }

  var bulle = null, cible = null, minuterie = null;
  function creer(){
    bulle = document.createElement('div');
    bulle.className = 'infobulle'; bulle.setAttribute('role','tooltip'); bulle.id = 'infobulle';
    document.documentElement.appendChild(bulle);   // hors du <body> : non affectée par la taille du texte
  }
  function masquer(){
    clearTimeout(minuterie); minuterie = null;
    if(bulle) bulle.classList.remove('vis');
    if(cible){ cible.removeAttribute('aria-describedby'); cible = null; }
  }
  function afficher(el){
    var t = texte(el); if(!t) return;
    if(!bulle) creer();
    bulle.textContent = t;
    bulle.style.left = '0px'; bulle.style.top = '0px';
    bulle.classList.add('vis');
    var r = el.getBoundingClientRect(), b = bulle.getBoundingClientRect();
    var W = document.documentElement.clientWidth, H = document.documentElement.clientHeight;
    var x = Math.max(8, Math.min(W - b.width - 8, r.left + r.width / 2 - b.width / 2));
    var y = r.bottom + 8;
    if(y + b.height > H - 8) y = r.top - b.height - 8;      // pas de place dessous : au-dessus
    bulle.style.left = Math.round(x) + 'px'; bulle.style.top = Math.round(Math.max(8, y)) + 'px';
    el.setAttribute('aria-describedby', 'infobulle');
  }
  function armer(el){
    if(el === cible) return;
    masquer(); cible = el;
    minuterie = setTimeout(function(){ if(cible === el && document.contains(el)) afficher(el); }, DELAI);
  }
  function bouton(n){ return n && n.closest ? n.closest('button, .acc-btn, [role="button"]') : null; }

  document.addEventListener('mouseover', function(e){ var el = bouton(e.target); if(el) armer(el); else if(cible) masquer(); });
  document.addEventListener('mouseout', function(e){
    var el = bouton(e.target); if(!el || el !== cible) return;
    if(e.relatedTarget && el.contains(e.relatedTarget)) return;
    masquer();
  });
  document.addEventListener('focusin', function(e){
    var el = bouton(e.target);
    if(el && el.matches(':focus-visible')) armer(el);
  });
  document.addEventListener('focusout', function(e){ if(bouton(e.target) === cible) masquer(); });
  ['mousedown','keydown','wheel','scroll'].forEach(function(ev){ window.addEventListener(ev, masquer, true); });
  window.addEventListener('blur', masquer);
})();
