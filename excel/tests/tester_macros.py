#!/usr/bin/env python3
# Exécute les macros du classeur dans LibreOffice (mode VBA) et vérifie les résultats.
# Usage : python3 tests/tester_macros.py Enregistrements-Controle.xlsm
# Chaque instruction VBA est précédée d'un repère (DbgL) pour localiser une erreur éventuelle.
import os, random, shutil, subprocess, sys, tempfile, time
import uno
from com.sun.star.beans import PropertyValue

SRC = sys.argv[1]
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

work = tempfile.mkdtemp(prefix='ec-test-')
fichier = os.path.join(work, 'test.xlsm'); shutil.copy(SRC, fichier)
port = str(2900 + random.randint(0, 500)); profil = tempfile.mkdtemp(prefix='lo-p-')
proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '-env:UserInstallation=file://' + profil,
                         '--accept=socket,host=localhost,port=%s;urp;' % port], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for _ in range(120):
    try:
        local = uno.getComponentContext()
        ctx = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local).resolve(
            'uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext' % port); break
    except Exception: time.sleep(0.5)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(fichier), '_blank', 0, (pv('Hidden', False), pv('MacroExecutionMode', 4)))
sp = doc.getScriptProvider()
import re as _re
_vb = doc.BasicLibraries.getByName('VBAProject')
for _m in _vb.ElementNames:
    _src = _vb.getByName(_m).splitlines(); _out = []; _inproc = False; _prev = ''
    for _i, _l in enumerate(_src):
        _t = _l.strip()
        if _re.match(r'(Public |Private )?(Sub|Function) ', _t): _inproc = True
        if _re.match(r'End (Sub|Function)', _t): _inproc = False
        if (_inproc and _t and not _t.startswith("'") and not _re.match(r'(Public |Private )?(Sub|Function|Dim|Case|Static) ', _t)
                and not _prev.rstrip().endswith(' _') and not _prev.strip().startswith('Select Case') and not _re.match(r'^\w+:$', _t)):
            _out.append('DbgL = "%s:%d"' % (_m, _i + 1))
        _out.append(_l); _prev = _l
    _j = '\n'.join(_out)
    _vb.replaceByName(_m, _j)
doc.BasicLibraries.getByName('VBAProject').insertByName('modDbg', '''Option VBASupport 1
Public DbgL As String
Public DbgX As String
Sub SelMenu()
    Worksheets("Menu").Activate
    Worksheets("Menu").Range("C9").Select
End Sub
Function LireX() As String
    LireX = DbgX
End Function
Function Appel(ByVal nom As String, Optional ByVal a1 As Variant) As String
    On Error GoTo e
    If IsMissing(a1) Then
        Application.Run nom
    Else
        Application.Run nom, a1
    End If
    Appel = ""
    Exit Function
e:
    Appel = "ERREUR " & nom & " @" & DbgL & " [" & DbgX & "] : " & Err.Number & " " & Err.Description
End Function
''')

def macro(nom, *args):
    if not nom.startswith('modCommun.Test'):
        r = sp.getScript('vnd.sun.star.script:VBAProject.modDbg.Appel?language=Basic&location=document').invoke((nom.split('.')[1],) + tuple(args), (), ())[0]
        if r: print('     !!', r)
        return r
    s = sp.getScript('vnd.sun.star.script:VBAProject.%s?language=Basic&location=document' % nom)
    r = s.invoke(tuple(args), (), ())
    return r[0]

def sh(n): return doc.Sheets.getByName(n)
def cell(n, a): return sh(n).getCellRangeByName(a)
def val(n, a):
    c = cell(n, a); return c.getString() if c.getType().value == 'TEXT' or c.getFormula().startswith('=') and c.getString() and not c.getValue() else (c.getValue() if c.getType().value != 'EMPTY' else '')
def s(n, a): return cell(n, a).getString()
def setv(n, a, v):
    c = cell(n, a)
    if isinstance(v, (int, float)): c.setValue(v)
    else: c.setString(v)
def msgs():
    m = macro('modCommun.TestMessages'); return m.strip().replace('\n', ' / ')
def demarrer(rep=''): macro('modCommun.TestDemarrer', rep)
ECHECS = []
def ok(cond, txt):
    if not cond: ECHECS.append(txt)
    print(('OK   ' if cond else 'ÉCHEC') + ' ' + txt)
    return cond

def lignes_rapport():
    out = []
    for r in range(18, 60):
        if not s('Rapport', 'A%d' % r) and not s('Rapport', 'B%d' % r): break
        out.append((r, s('Rapport', 'B%d' % r), s('Rapport', 'G%d' % r), s('Rapport', 'H%d' % r), s('Rapport', 'I%d' % r), s('Rapport', 'J%d' % r)))
    return out

try:
    demarrer()
    macro('modMenu.Auto_Open')
    nrev = sum(1 for r in range(4, 200) if s('Révisions', 'A%d' % r))
    ok(nrev > 0, 'Auto_Open : %d révision(s) initiale(s) créée(s)' % nrev)
    nh = sum(1 for r in range(4, 2000) if s('Historique', 'A%d' % r))
    ok(nh > 0, 'Historique : %d ligne(s) de photo' % nh)
    print('     boutons Menu :', sh('Menu').DrawPage.Count, '· Rapport :', sh('Rapport').DrawPage.Count)

    # ---- nouveau rapport + préparation
    demarrer()
    macro('modRapport.NouveauRapport')
    ok(s('Rapport', 'O12') == 'Nouveau' and s('Rapport', 'L6') == '15', 'Nouveau rapport : état « Nouveau », échantillon 15 par défaut (%s / %s)' % (s('Rapport', 'O12'), s('Rapport', 'L6')))
    setv('Rapport', 'C4', '041038_CBMI'); setv('Rapport', 'F4', 'LOT-1'); setv('Rapport', 'L6', 3)
    demarrer(); macro('modRapport.PreparerRapport'); m = msgs()
    L = lignes_rapport()
    ok(len(L) == 19, 'Préparer : %d lignes (attendu 19) %s' % (len(L), m))
    print('     ', L[0], L[4], L[6])
    ok(s('Rapport', 'C12') == 'Liaison CBMI' and 'rév. 0' in s('Rapport', 'F12'), 'Famille / révision : %s / %s / %s' % (s('Rapport', 'C12'), s('Rapport', 'F12'), s('Rapport', 'I12')))
    ok(s('Rapport', 'L4') == '4001017', 'Fournisseur repris de la Matrice : %s' % s('Rapport', 'L4'))
    ok(s('Rapport', 'R4') != '', 'Documents applicables : %s %s / %s %s' % (s('Rapport', 'Q4'), s('Rapport', 'R4'), s('Rapport', 'Q5'), s('Rapport', 'R5')))
    # ---- saisie
    for r, code, ty, *_ in L:
        if ty == 'Générale': setv('Rapport', 'L%d' % r, 'OK')
    r_e = [x for x in L if x[1] == 'E'][0][0]
    r_f = [x for x in L if x[1] == 'F'][0][0]
    setv('Rapport', 'M%d' % r_e, 13.02); setv('Rapport', 'N%d' % r_e, 13.3); setv('Rapport', 'O%d' % r_e, '12,95')   # 13,3 hors tolérance ; 12,95 en texte avec virgule
    setv('Rapport', 'M%d' % r_f, '28.01')
    doc.calculateAll()
    ok(s('Rapport', 'L%d' % r_e) == '1', 'Badge NC du point E : %s (attendu 1)' % s('Rapport', 'L%d' % r_e))
    ok(s('Rapport', 'N15') == '1' and s('Rapport', 'M15') == '0', 'NC par pièce : P1=%s P2=%s' % (s('Rapport', 'M15'), s('Rapport', 'N15')))
    ok(s('Rapport', 'N16') == 'NC' and s('Rapport', 'M16') == 'En cours', 'Statut pièce : P1=%s P2=%s P4=%r' % (s('Rapport', 'M16'), s('Rapport', 'N16'), s('Rapport', 'P16')))
    demarrer(); macro('modRapport.PreparerRapport'); msgs()
    ok(abs(cell('Rapport', 'N%d' % r_e).getValue() - 13.3) < 1e-9 and s('Rapport', 'L18') == 'OK', 'Re-préparer garde les valeurs saisies')
    demarrer(); macro('modRapport.EnregistrerRapport'); m = msgs()
    ok('Rapport enregistré' in m, 'Enregistrer : %s' % m)
    ok(cell('Rapport', 'O%d' % r_e).getType().value == 'VALUE' and abs(cell('Rapport', 'O%d' % r_e).getValue() - 12.95) < 1e-9, 'Texte « 12,95 » converti en nombre')
    ok(abs(cell('Rapport', 'M%d' % r_f).getValue() - 28.01) < 1e-9, 'Texte « 28.01 » (point) converti en nombre')
    rid = s('Rapport', 'L12')
    ok(rid.startswith('R') and s('Rapports', 'A4') == rid, 'Ligne Rapports : %s · décision %s · C=%s NC=%s' % (s('Rapports', 'A4'), s('Rapports', 'M4'), s('Rapports', 'T4'), s('Rapports', 'U4')))
    nmes = sum(1 for r in range(4, 500) if s('Mesures', 'A%d' % r))
    ok(nmes == 19 + 4, 'Mesures : %d ligne(s) (19 lignes de structure/générales + 4 relevés)' % nmes)
    nc = [(s('Mesures', 'C%d' % r), s('Mesures', 'O%d' % r), s('Mesures', 'P%d' % r), s('Mesures', 'Q%d' % r)) for r in range(4, 12)]
    print('     ', nc)
    ok(s('Menu', 'C9') == '041038_CBMI' and s('Menu', 'L9') == rid, 'Liste du Menu : %s %s %s' % (s('Menu', 'B9'), s('Menu', 'C9'), s('Menu', 'D9')))
    # ---- réouverture
    demarrer(); macro('modRapport.NouveauRapport')
    ok(s('Rapport', 'C4') == '' and s('Rapport', 'B18') == '', 'Nouveau rapport : formulaire vidé')
    demarrer(); macro('modRapport.ChargerRapport', rid)
    L2 = lignes_rapport()
    ok(len(L2) == 19 and s('Rapport', 'F4') == 'LOT-1', 'Rapport rouvert : %d lignes, lot %s' % (len(L2), s('Rapport', 'F4')))
    r_e2 = [x for x in L2 if x[1] == 'E'][0][0]
    print('      E:', [s('Rapport', c + str(r_e2)) for c in 'BGHIJKLMNO'], 'A:', s('Rapport','L18'))
    ok(abs(cell('Rapport', 'N%d' % r_e2).getValue() - 13.3) < 1e-9 and s('Rapport', 'L%d' % [x for x in L2 if x[2] == 'Générale'][0][0]) == 'OK', 'Mesures et résultats généraux restaurés')
    # ---- décision « Attente », puis sortie protégée par mot de passe
    setv('Rapport', 'F8', 'Attente décision méthodes')
    demarrer(); macro('modRapport.EnregistrerRapport'); msgs()
    setv('Rapport', 'F8', 'Accepté')
    demarrer('mauvais'); macro('modRapport.EnregistrerRapport'); m = msgs()
    ok('incorrect' in m and s('Rapports', 'M4') == 'Attente décision méthodes', 'Sortie d\'« Attente » refusée sans le bon mot de passe : %s' % m)
    setv('Rapport', 'F8', 'Accepté')
    demarrer('1'); macro('modRapport.EnregistrerRapport'); m = msgs()
    ok(s('Rapport', 'O12') == 'Figé' and s('Rapports', 'M4') == 'Accepté', 'Accepté avec mot de passe : rapport figé (%s)' % m)
    demarrer(); macro('modRapport.PreparerRapport'); m = msgs()
    ok('figé' in m, 'Rapport figé : préparation refusée')
    demarrer('1'); macro('modRapport.RouvrirRapport'); m = msgs()
    ok(s('Rapport', 'O12') == 'Rouvert', 'Rouvrir (méthodiste) : %s' % m)
    # ---- nouvel article
    demarrer('NOUV-42|1|F099|041038_CBMIF05'); macro('modReferentiel.CreerArticle'); m = msgs()
    nts = [r for r in range(4, 400) if s('TS', 'B%d' % r) == 'NOUV-42']
    ok(len(nts) == 19, 'Nouvel article : %d lignes TS créées · %s' % (len(nts), m[:120]))
    doc.calculateAll()
    ok(s('TS', 'C%d' % nts[0]) == 'Liaison CBMI' and s('TS', 'D%d' % nts[0]) == 'A', 'Famille / code lus dans la gamme : %s / %s' % (s('TS', 'C%d' % nts[0]), s('TS', 'D%d' % nts[0])))
    ok(s('TS', 'F%d' % nts[4]) != '', 'Valeurs copiées du modèle : %s' % s('TS', 'F%d' % nts[4]))
    ok(s('Matrice', 'A15') == 'NOUV-42' or any(s('Matrice', 'A%d' % r) == 'NOUV-42' for r in range(4, 40)), 'Article ajouté à la Matrice')
    # ---- cohérence
    setv('Matrice', 'A40', 'SANS-FAM')
    demarrer(); macro('modReferentiel.VerifierCoherence')
    coh = [(s('Cohérence', 'A%d' % r), s('Cohérence', 'B%d' % r)) for r in range(4, 30) if s('Cohérence', 'A%d' % r)]
    ok(('Article sans famille', 'SANS-FAM') in coh, 'Cohérence : %s' % coh[:6])
    # ---- analyses
    demarrer(); macro('modMenu.ActualiserAnalyses')
    ok(s('Analyses', 'A14') == '041038_CBMI', 'Analyses : %s %s n=%s moy=%s NC=%s' % (s('Analyses', 'A14'), s('Analyses', 'B14'), s('Analyses', 'G14'), s('Analyses', 'H14'), s('Analyses', 'L14')))

    # ---- révisions
    doc.CurrentController.setActiveSheet(sh('TS'))
    rts = [r for r in range(4, 400) if s('TS', 'B%d' % r) == '041038_CBMI' and s('TS', 'F%d' % r)][0]
    avant = s('TS', 'F%d' % rts)
    demarrer('041038_CBMI|oui'); macro('modRevisions.OuvrirRevision'); m = msgs()
    revs = lambda: [(s('Révisions', 'B%d' % r), s('Révisions', 'C%d' % r), s('Révisions', 'D%d' % r)) for r in range(4, 60) if s('Révisions', 'B%d' % r) == '041038_CBMI']
    ok(('041038_CBMI', '1', 'Brouillon') in revs(), 'Révision 1 ouverte : %s · %s' % (revs(), m[:80]))
    setv('TS', 'F%d' % rts, 99.5)
    demarrer('041038_CBMI|oui'); macro('modRevisions.AbandonnerRevision'); m = msgs()
    ok(s('TS', 'F%d' % rts) == avant and ('041038_CBMI', '1', 'Brouillon') not in revs(), 'Abandon : valeur restaurée (%s) · %s' % (s('TS', 'F%d' % rts), revs()))
    rts = [r for r in range(4, 400) if s('TS', 'B%d' % r) == '041038_CBMI' and s('TS', 'F%d' % r)][0]
    demarrer('041038_CBMI|oui'); macro('modRevisions.OuvrirRevision'); msgs()
    setv('TS', 'F%d' % rts, 99.5)
    demarrer('041038_CBMI|Testeur|Nouvelle cote|10/11/2026'); macro('modRevisions.ValiderRevision'); m = msgs()
    ok(('041038_CBMI', '1', 'Applicable') in revs() and ('041038_CBMI', '0', 'Archivée') in revs(), 'Validation : %s · %s' % (revs(), m[:80]))
    # modification hors révision détectée
    r2 = [r for r in range(4, 400) if s('TS', 'B%d' % r) == '041038_CBMIF05' and s('TS', 'F%d' % r)][0]
    setv('TS', 'F%d' % r2, 1234)
    demarrer(); macro('modReferentiel.VerifierCoherence')
    coh = [(s('Cohérence', 'A%d' % r), s('Cohérence', 'B%d' % r)) for r in range(4, 40) if s('Cohérence', 'A%d' % r)]
    ok(any('révision' in a.lower() and '041038_CBMIF05' in b for a, b in coh), 'Modification sans révision signalée : %s' % [c for c in coh if 'vision' in c[0]])
    # nouveau rapport après révision : rév. 1
    demarrer(); macro('modRapport.NouveauRapport'); setv('Rapport', 'C4', '041038_CBMI'); setv('Rapport', 'F4', 'LOT-2')
    demarrer(); macro('modRapport.PreparerRapport'); msgs()
    ok('rév. 1' in s('Rapport', 'I12'), 'Nouveau rapport : révision TS %s' % s('Rapport', 'I12'))
    # révision de gamme
    doc.CurrentController.setActiveSheet(sh('GAMSERIE'))
    demarrer('Liaison CBMI|oui'); macro('modRevisions.OuvrirRevision'); m = msgs()
    demarrer('Liaison CBMI|Testeur|Ajout point|'); macro('modRevisions.ValiderRevision'); m = msgs()
    ok('validée' in m, 'Révision de gamme : %s' % m[:90])
    # suppression d'un rapport
    nrap = sum(1 for r in range(4, 50) if s('Rapports', 'A%d' % r))
    print('      rapports :', nrap, '· journal :', sum(1 for r in range(4, 300) if s('Journal', 'A%d' % r)), 'ligne(s)')
    macro('modDbg.SelMenu')
    demarrer('oui'); macro('modMenu.SupprimerRapportSelectionne'); m = msgs()
    ok(s('Rapports', 'A4') == '' and not any(s('Mesures', 'A%d' % r) for r in range(4, 40)) and 'Aucun' in s('Menu', 'A9') + s('Menu', 'B9'), 'Suppression du rapport : %s' % m[:80])
    # ---- lignes TS manquantes
    demarrer(); macro('modReferentiel.PreparerLignesTSManquantes'); print('     lignes manquantes :', msgs())
finally:
    try:
        doc.close(True)
    except Exception as e:
        print('fermeture', e)
    try: desktop.terminate()
    except Exception: pass
    proc.wait(timeout=30)
print('\n%d échec(s)' % len(ECHECS))
sys.exit(1 if ECHECS else 0)
