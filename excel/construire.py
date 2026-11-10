#!/usr/bin/env python3
"""
Construit le classeur « Enregistrements-Controle.xlsm » (application complète en VBA).

1. LibreOffice (sans écran) compile les modules VBA de excel/vba/*.bas dans un .xlsm minimal ;
2. openpyxl ajoute les feuilles, la mise en page, les formules, les mises en forme
   conditionnelles, les listes déroulantes et les données d'exemple, en gardant les macros.

Utilisation :  python3 excel/construire.py
"""
import glob, json, os, random, re, subprocess, sys, tempfile, time

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)
SORTIE = os.path.join(ICI, 'Enregistrements-Controle.xlsm')

# ---------------------------------------------------------------- 1. macros
def compiler_vba(chemin):
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.script import ModuleInfo
    from com.sun.star.script.ModuleType import NORMAL

    def pv(n, v):
        p = PropertyValue(); p.Name = n; p.Value = v; return p

    port = str(2100 + random.randint(0, 800))
    profil = tempfile.mkdtemp(prefix='lo-profil-')
    proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '-env:UserInstallation=file://' + profil,
                             '--accept=socket,host=localhost,port=%s;urp;' % port],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for _ in range(120):
        try:
            local = uno.getComponentContext()
            res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
            ctx = res.resolve('uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext' % port)
            break
        except Exception:
            time.sleep(0.5)
    if ctx is None:
        sys.exit('LibreOffice ne répond pas')
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desktop.loadComponentFromURL('private:factory/scalc', '_blank', 0, (pv('Hidden', True),))
    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    libs.ProjectName = 'VBAProject'
    if not libs.hasByName('VBAProject'):
        libs.createLibrary('VBAProject')
    lib = libs.getByName('VBAProject')
    for f in sorted(glob.glob(os.path.join(ICI, 'vba', '*.bas'))):
        nom = os.path.splitext(os.path.basename(f))[0]
        code = open(f, encoding='utf-8').read()
        code = re.sub(r'^Attribute VB_Name = .*\n', '', code, flags=re.M)
        for c in code:
            if ord(c) > 255:
                sys.exit('%s : caractère hors page de codes Windows occidentale : %r' % (f, c))
        mi = ModuleInfo(); mi.ModuleType = NORMAL
        lib.insertModuleInfo(nom, mi)
        lib.insertByName(nom, 'Option VBASupport 1\n' + code)
    doc.storeToURL(uno.systemPathToFileUrl(chemin), (pv('FilterName', 'Calc MS Excel 2007 VBA XML'),))
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    try:
        proc.wait(timeout=30)
    except Exception:
        proc.kill()

# ---------------------------------------------------------------- 2. classeur
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L

BLEU, ENCRE, GRIS, GRIS2, FOND = '2563EB', '111827', '6B7280', 'E5E7EB', 'F3F4F6'
ROUGE, ROUGEF, VERT, VERTF, AMBRE, AMBREF = 'B91C1C', 'FEE2E2', '15803D', 'DCFCE7', 'A15C07', 'FEF3C7'
TITRE = Font(name='Calibri', size=16, bold=True, color=ENCRE)
GRAS = Font(name='Calibri', size=11, bold=True, color=ENCRE)
ETIQ = Font(name='Calibri', size=10, color=GRIS)
ENTETE_FILL = PatternFill('solid', fgColor=FOND)
SAISIE_FILL = PatternFill('solid', fgColor='FFFFFF')
CALC_FONT = Font(name='Calibri', size=11, color=GRIS)
fin = Side(style='thin', color=GRIS2)
CADRE = Border(left=fin, right=fin, top=fin, bottom=fin)
BAS = Border(bottom=Side(style='thin', color='9CA3AF'))

R_L0, R_LMAX, R_PMAX = 18, 200, 100
R_LFIN = R_L0 + R_LMAX - 1
C_P1 = 13
C_PFIN = C_P1 + R_PMAX - 1


def remplir(ws, rng, **kw):
    for row in ws[rng]:
        for c in row:
            for k, v in kw.items():
                setattr(c, k, v)


def feuille_tableau(wb, nom, titre, entetes, largeurs, aide=''):
    ws = wb.create_sheet(nom)
    ws['A1'] = titre
    ws['A1'].font = TITRE
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 32
    if aide:
        ws.cell(1, max(4, len(entetes) - 2)).value = aide
        ws.cell(1, max(4, len(entetes) - 2)).font = ETIQ
    for i, h in enumerate(entetes, 1):
        c = ws.cell(3, i, h)
        c.font = GRAS
        c.fill = ENTETE_FILL
        c.border = CADRE
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[3].height = 30
    for i, w in enumerate(largeurs, 1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = 'A3:%s3' % L(len(entetes))
    ws.sheet_view.showGridLines = True
    return ws


def nom_defini(wb, nom, ref):
    wb.defined_names[nom] = DefinedName(nom, attr_text=ref)


def liste(ws, cellules, formule, titre='Valeur', message='Choisissez une valeur dans la liste.'):
    dv = DataValidation(type='list', formula1=formule.lstrip('='), allow_blank=True)
    dv.error = message
    dv.errorTitle = titre
    dv.errorStyle = 'warning'
    ws.add_data_validation(dv)
    dv.add(cellules)


def construire(base):
    wb = load_workbook(base, keep_vba=True)
    for ws in list(wb.worksheets):
        wb.remove(ws)

    # ------------------------------------------------ données d'exemple (celles de l'application web)
    app = open(os.path.join(RACINE, 'js', 'app.js'), encoding='utf-8').read()
    seed = json.loads(re.search(r'var SEED = (\{.*?\});\n', app).group(1))
    rnd = random.Random(20261110)
    alpha = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
    ids, id_de = set(), {}
    for g in seed['gamserie']:
        while True:
            i = ''.join(rnd.choice(alpha) for _ in range(6))
            if i not in ids:
                break
        ids.add(i)
        g['id'] = i
        id_de[(g['famille'], g['code'])] = i
    fam_de = {m['article']: m['famille'] for m in seed['matrice']}

    # ------------------------------------------------ Menu
    ws = wb.create_sheet('Menu')
    ws['A1'] = "Enregistrements Contrôle d'entrée"
    ws['A1'].font = Font(name='Calibri', size=20, bold=True, color=ENCRE)
    ws.row_dimensions[1].height = 34
    for r in (3, 5):
        ws.row_dimensions[r].height = 32
    ws['B7'].font = ETIQ
    ws['H7'].font = Font(name='Calibri', size=10, color=AMBRE, bold=True)
    entetes = ['', 'Date', 'Article', 'Lot', 'Commande', 'Fournisseur', 'Contrôleur', 'Échantillon', 'Conformes', 'Non conformes', 'Décision', 'ID']
    larg = [2, 12, 20, 16, 16, 16, 16, 12, 11, 13, 26, 18]
    for i, (h, w) in enumerate(zip(entetes, larg), 1):
        c = ws.cell(8, i, h)
        if h:
            c.font = GRAS; c.fill = ENTETE_FILL; c.border = BAS
        ws.column_dimensions[L(i)].width = w
    ws.column_dimensions['L'].hidden = True
    for r in range(9, 2009):
        ws.cell(r, 2).number_format = 'DD/MM/YYYY'
    ws.freeze_panes = 'A9'
    ws.conditional_formatting.add('K9:K2008', FormulaRule(formula=['$K9="Accepté"'], font=Font(color=VERT, bold=True)))
    ws.conditional_formatting.add('K9:K2008', FormulaRule(formula=['$K9="Refusé"'], font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add('K9:K2008', FormulaRule(formula=['$K9="Attente décision méthodes"'], font=Font(color=AMBRE, bold=True)))
    ws.conditional_formatting.add('J9:J2008', FormulaRule(formula=['AND(ISNUMBER($J9),$J9>0)'], font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add('I9:I2008', FormulaRule(formula=['AND(ISNUMBER($I9),$I9>0)'], font=Font(color=VERT, bold=True)))
    ws.sheet_view.showGridLines = False

    # ------------------------------------------------ Rapport
    ws = wb.create_sheet('Rapport')
    ws['A1'] = 'Rapport de contrôle'
    ws['A1'].font = TITRE
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 32
    largeurs = {1: 9, 2: 8, 3: 34, 4: 10, 5: 18, 6: 18, 7: 11, 8: 9, 9: 9, 10: 9, 11: 10, 12: 10}
    for c in range(1, C_PFIN + 1):
        ws.column_dimensions[L(c)].width = largeurs.get(c, 8)
    ws.column_dimensions['A'].hidden = True
    champs = [  # (étiquette, cellule de saisie)
        ('Article', 'C4'), ('Lot', 'F4'), ('Commande', 'I4'), ('Fournisseur', 'L4'), ('Contrôleur', 'O4'),
        ('Date', 'C6'), ('Qté commandée', 'F6'), ('Qté réceptionnée', 'I6'), ('Qté échantillon', 'L6'), ('Durée contrôle', 'O6'),
        ('Origine', 'C8'), ('Décision', 'F8'), ('Type NC', 'I8'), ('Sous-type NC', 'L8'),
        ('Commentaires', 'C10'),
    ]
    for etiq, cel in champs:
        col = ws[cel].column
        lab = ws.cell(ws[cel].row, col - 1, etiq)
        lab.font = ETIQ
        lab.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
        ws[cel].fill = SAISIE_FILL
        ws[cel].border = CADRE
        ws[cel].font = Font(name='Calibri', size=11, color=ENCRE, bold=cel in ('C4', 'F4'))
        ws[cel].alignment = Alignment(vertical='center')
    for rng in ('C4:D4', 'C6:D6', 'C8:D8', 'F8:G8', 'I8:J8', 'L8:O8', 'C10:O10'):
        ws.merge_cells(rng)
    ws['C6'].number_format = 'DD/MM/YYYY'
    for r in (4, 6, 8, 10):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[10].height = 36
    ws['C10'].alignment = Alignment(wrap_text=True, vertical='top')
    infos = [('Famille', 'C12'), ('Rév. gamme', 'F12'), ('Rév. valeurs TS', 'I12'), ('N° rapport', 'L12'), ('État', 'O12'),
             ('Pièces conformes', 'C13'), ('Pièces non conformes', 'F13'), ('Généraux NOK', 'I13'), ('Pièces à terminer', 'L13')]
    for etiq, cel in infos:
        lab = ws.cell(ws[cel].row, ws[cel].column - 1, etiq)
        lab.font = ETIQ
        lab.alignment = Alignment(horizontal='right')
        ws[cel].font = Font(name='Calibri', size=11, bold=True, color=ENCRE)
    ws['C13'] = '=COUNTIF(M16:%s16,"C")' % L(C_PFIN)
    ws['F13'] = '=COUNTIF(M16:%s16,"NC")' % L(C_PFIN)
    ws['I13'] = '=COUNTIFS(G%d:G%d,"Générale",L%d:L%d,"NOK")' % (R_L0, R_LFIN, R_L0, R_LFIN)
    ws['L13'] = '=IF(L6="","",MAX(0,L6-C13-F13))'
    ws['C13'].font = Font(name='Calibri', size=13, bold=True, color=VERT)
    ws['F13'].font = Font(name='Calibri', size=13, bold=True, color=ROUGE)
    ws['I13'].font = Font(name='Calibri', size=13, bold=True, color=ROUGE)
    ws['Q3'] = 'Documents applicables'
    ws['Q3'].font = GRAS
    ws.column_dimensions['Q'].width = 16
    ws.column_dimensions['R'].width = 16
    # tableau
    ent = ['ID', 'Code', 'Spécification', 'Criticité', 'Moyen de contrôle', 'Méthode', 'Type', 'Mini', 'Nominal', 'Maxi', 'N° moyen', 'Résultat']
    for i, h in enumerate(ent, 1):
        c = ws.cell(17, i, h)
        c.font = GRAS; c.fill = ENTETE_FILL; c.border = CADRE
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.cell(15, 12, 'NC / pièce').font = ETIQ
    ws.cell(16, 12, 'Statut pièce').font = ETIQ
    ws.cell(15, 12).alignment = ws.cell(16, 12).alignment = Alignment(horizontal='right')
    g, h, j = '$G$%d:$G$%d' % (R_L0, R_LFIN), '$H$%d:$H$%d' % (R_L0, R_LFIN), '$J$%d:$J$%d' % (R_L0, R_LFIN)
    for c in range(C_P1, C_PFIN + 1):
        col = L(c)
        m = '%s$%d:%s$%d' % (col, R_L0, col, R_LFIN)
        hors = 'COLUMN()-12>$L$6'
        ws.cell(17, c, '=IF(%s,"","P"&(COLUMN()-12))' % hors)
        ws.cell(15, c, '=IF(%s,"",SUMPRODUCT(--(%s="Individuelle"),--ISNUMBER(%s),--((%s<%s)+(%s>%s)>0))+COUNTIF(%s,"NOK"))' % (hors, g, m, m, h, m, j, m))
        nb = 'COUNTIF(%s,"Individuelle")' % g
        rempl = 'COUNTIFS(%s,"Individuelle",%s,"<>")' % (g, m)
        ws.cell(16, c, '=IF(%s,"",IF(%s=0,"",IF(%s15>0,"NC",IF(%s=%s,"C",IF(%s>0,"En cours","")))))' % (hors, nb, col, rempl, nb, rempl))
        for r in (15, 16, 17):
            ws.cell(r, c).alignment = Alignment(horizontal='center')
        ws.cell(17, c).font = GRAS
        ws.cell(17, c).fill = ENTETE_FILL
        ws.cell(17, c).border = CADRE
    for r in range(R_L0, R_LFIN + 1):
        for c in (8, 9, 10):
            ws.cell(r, c).number_format = '0.00'
        for c in range(C_P1, C_PFIN + 1):
            ws.cell(r, c).number_format = '0.00'
            ws.cell(r, c).alignment = Alignment(horizontal='center')
        ws.cell(r, 12).alignment = Alignment(horizontal='center')
        ws.cell(r, 3).alignment = Alignment(wrap_text=False)
    zone = 'M%d:%s%d' % (R_L0, L(C_PFIN), R_LFIN)
    lignes = 'A%d:L%d' % (R_L0, R_LFIN)
    ws.conditional_formatting.add(zone, FormulaRule(formula=['OR(COLUMN(M%d)-12>$L$6,$G%d<>"Individuelle")' % (R_L0, R_L0)], fill=PatternFill('solid', fgColor=FOND), stopIfTrue=True))
    ws.conditional_formatting.add(zone, FormulaRule(formula=['AND(ISNUMBER(M%d),ISNUMBER($H%d),OR(M%d<$H%d,M%d>$J%d))' % ((R_L0,) * 6)], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add(zone, FormulaRule(formula=['M%d="NOK"' % R_L0], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add(zone, FormulaRule(formula=['M%d="OK"' % R_L0], font=Font(color=VERT, bold=True)))
    ws.conditional_formatting.add(zone, FormulaRule(formula=['AND(ISNUMBER(M%d),ISNUMBER($H%d))' % (R_L0, R_L0)], font=Font(color=VERT)))
    ws.conditional_formatting.add(lignes, FormulaRule(formula=['$G%d="N/A"' % R_L0], font=Font(color='9CA3AF', italic=True)))
    resz = 'L%d:L%d' % (R_L0, R_LFIN)
    ws.conditional_formatting.add(resz, FormulaRule(formula=['AND($G%d="Individuelle",N(L%d)>0)' % (R_L0, R_L0)], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add(resz, FormulaRule(formula=['L%d="NOK"' % R_L0], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add(resz, FormulaRule(formula=['L%d="OK"' % R_L0], fill=PatternFill('solid', fgColor=VERTF), font=Font(color=VERT, bold=True)))
    ws.conditional_formatting.add(resz, FormulaRule(formula=['$G%d="Générale"' % R_L0], fill=PatternFill('solid', fgColor='EFF6FF')))
    ws.conditional_formatting.add('M15:%s15' % L(C_PFIN), FormulaRule(formula=['AND(ISNUMBER(M15),M15>0)'], fill=PatternFill('solid', fgColor=ROUGE), font=Font(color='FFFFFF', bold=True)))
    ws.conditional_formatting.add('M16:%s16' % L(C_PFIN), FormulaRule(formula=['M16="NC"'], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add('M16:%s16' % L(C_PFIN), FormulaRule(formula=['M16="C"'], fill=PatternFill('solid', fgColor=VERTF), font=Font(color=VERT, bold=True)))
    ws.conditional_formatting.add('M16:%s16' % L(C_PFIN), FormulaRule(formula=['M16="En cours"'], font=Font(color=AMBRE)))
    ws.conditional_formatting.add('O12', FormulaRule(formula=['$O$12="Figé"'], fill=PatternFill('solid', fgColor=AMBREF), font=Font(color=AMBRE, bold=True)))
    ws.conditional_formatting.add('F8', FormulaRule(formula=['$F$8="Accepté"'], fill=PatternFill('solid', fgColor=VERTF), font=Font(color=VERT, bold=True)))
    ws.conditional_formatting.add('F8', FormulaRule(formula=['$F$8="Refusé"'], fill=PatternFill('solid', fgColor=ROUGEF), font=Font(color=ROUGE, bold=True)))
    ws.conditional_formatting.add('F8', FormulaRule(formula=['$F$8="Attente décision méthodes"'], fill=PatternFill('solid', fgColor=AMBREF), font=Font(color=AMBRE, bold=True)))
    dv = DataValidation(type='custom', formula1='OR(M%d="",ISNUMBER(M%d),M%d="OK",M%d="NOK")' % ((R_L0,) * 4), allow_blank=True)
    dv.errorTitle = 'Valeur mesurée'
    dv.error = 'Saisissez un nombre avec une virgule (par exemple 13,02), ou OK / NOK pour un point qualitatif.'
    ws.add_data_validation(dv)
    dv.add(zone)
    liste(ws, resz, '=ListeOKNOK', 'Résultat', 'Choisissez OK ou NOK.')
    liste(ws, 'C4', '=Matrice!$A$4:$A$3000', 'Article', "Choisissez un article de la Matrice (ou créez-le avec « Nouvel article »).")
    liste(ws, 'C8', '=ListeOrigines')
    liste(ws, 'F8', '=ListeDecisions')
    liste(ws, 'I8', '=ListeTypesNC')
    liste(ws, 'L8', '=ListeSousTypesNC')
    dq = DataValidation(type='whole', operator='between', formula1='1', formula2=str(R_PMAX), allow_blank=True)
    dq.error = 'Quantité échantillon : nombre entier de 1 à %d.' % R_PMAX
    ws.add_data_validation(dq)
    dq.add('L6')
    ws.freeze_panes = 'D1'
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ------------------------------------------------ GAMSERIE, TS, Matrice, Documents
    wg = feuille_tableau(wb, 'GAMSERIE', 'GAMSERIE - gammes de contrôle',
                         ['ID', 'Famille', 'Code', 'Criticité', 'Spécification', 'Moyen de contrôle', 'Méthode de contrôle'],
                         [10, 22, 8, 11, 44, 26, 26], "ID : identifiant du point, ne pas modifier (attribué automatiquement).")
    for k, gl in enumerate(seed['gamserie'], 4):
        for c, v in enumerate([gl['id'], gl['famille'], gl['code'], gl.get('criticite', ''), gl.get('specification', ''), gl.get('moyen', ''), gl.get('methode', '')], 1):
            wg.cell(k, c, v)
    for r in range(4, 2004):
        wg.cell(r, 1).font = CALC_FONT
    liste(wg, 'D4:D2003', '=ListeCriticites')

    wt = feuille_tableau(wb, 'TS', 'TS - valeurs par article',
                         ['ID point', 'Article', 'Famille', 'Code', 'Type spéc.', 'Nominal', 'Tol -', 'Tol +'],
                         [10, 20, 20, 8, 13, 11, 9, 9], "Famille et Code sont lus dans la gamme (colonnes grises). Nominal « - » = non applicable.")
    k = 4
    for t in seed['ts']:
        fam = fam_de.get(t['article'], '')
        for code, v in t['vals'].items():
            i = id_de.get((fam, code))
            if not i:
                continue
            wt.cell(k, 1, i); wt.cell(k, 2, t['article'])
            wt.cell(k, 5, v[3] if len(v) > 3 else '')
            for c, x in zip((6, 7, 8), v[:3]):
                wt.cell(k, c, x)
            k += 1
    for r in range(4, 3004):
        wt.cell(r, 3, '=IF($A%d="","",IFERROR(INDEX(GAMSERIE!$B:$B,MATCH($A%d,GAMSERIE!$A:$A,0)),"(ID inconnu)"))' % (r, r))
        wt.cell(r, 4, '=IF($A%d="","",IFERROR(INDEX(GAMSERIE!$C:$C,MATCH($A%d,GAMSERIE!$A:$A,0)),""))' % (r, r))
        for c in (1, 3, 4):
            wt.cell(r, c).font = CALC_FONT
    liste(wt, 'E4:E3003', '=ListeTypesSpec', 'Type spéc.', 'Choisissez Générale ou Individuelle.')
    wt.conditional_formatting.add('F4:H3003', FormulaRule(formula=['AND($B4<>"",$F4="")'], fill=PatternFill('solid', fgColor=ROUGEF)))

    wm = feuille_tableau(wb, 'Matrice', 'Matrice - article / fournisseur',
                         ['Article', 'Code fournisseur', 'Famille', 'Matière', 'Liasse', 'Réf. Norme', 'Réf. Plan méthode', 'Réf. Plan contrôle',
                          'Réf. Plan de situation', 'Réf. Conditionnement', 'Réf. Taux échantillon', 'Réf. Défauthèque'],
                         [20, 15, 20, 12, 9, 14, 14, 14, 14, 16, 16, 14])
    cles = ['article', 'fournisseur', 'famille', 'matiere', 'liasse', 'norme', 'planMethode', 'planControle', 'planSituation', 'conditionnement', 'tauxEchantillon', 'defautheque']
    for k, m in enumerate(seed['matrice'], 4):
        for c, cle in enumerate(cles, 1):
            wm.cell(k, c, m.get(cle, ''))

    wd = feuille_tableau(wb, 'Documents', 'Documents PDF', ['Référence', 'Fichier', 'Ajouté le'], [22, 80, 18],
                         "Référence = nom du fichier sans « .pdf », citée dans la Matrice.")

    # ------------------------------------------------ Analyses, Cohérence
    wa = wb.create_sheet('Analyses')
    wa['A1'] = 'Analyses des dérives'
    wa['A1'].font = TITRE
    wa.row_dimensions[2].height = 32
    for i, (etiq, aide) in enumerate([('Article', 'contient'), ('Lot', 'contient'), ('Fournisseur', 'contient'),
                                      ('Point (code ou spécification)', ''), ('Du (date)', 'jj/mm/aaaa'), ('Au (date)', 'jj/mm/aaaa')], 4):
        wa.cell(i, 2, etiq).font = ETIQ
        wa.cell(i, 3).border = CADRE
        wa.cell(i, 4, aide).font = ETIQ
    wa['C8'].number_format = wa['C9'].number_format = 'DD/MM/YYYY'
    wa['B10'] = 'Renseignez les filtres (facultatif), puis cliquez sur « Actualiser l\'analyse ».'
    wa['B10'].font = ETIQ
    for c, w in zip('ABCDEFGHIJKLM', [12, 22, 30, 12, 10, 10, 10, 11, 10, 10, 11, 13, 11]):
        wa.column_dimensions[c].width = w
    wa.sheet_view.showGridLines = False

    wc = feuille_tableau(wb, 'Cohérence', 'Cohérence du paramétrage', ['Type', 'Article', 'Famille', 'Détail'], [28, 20, 20, 110])

    # ------------------------------------------------ Révisions, Historique, Journal, Rapports, Mesures
    wr = feuille_tableau(wb, 'Révisions', 'Révisions des gammes (par famille) et des valeurs TS (par article)',
                         ['Type', 'Famille / article', 'Révision', 'Statut', "Date d'effet", 'Validée par', 'Motif', 'Ouverte le'],
                         [10, 22, 10, 12, 13, 18, 40, 18])
    wh = feuille_tableau(wb, 'Historique', 'Historique - photo de chaque révision validée',
                         ['Type', 'Famille / article', 'Révision', 'ID point', 'Code', 'Criticité', 'Spécification', 'Moyen', 'Méthode', 'Type spéc.', 'Nominal', 'Tol -', 'Tol +'],
                         [10, 22, 10, 10, 8, 11, 40, 22, 22, 13, 10, 9, 9])
    wj = feuille_tableau(wb, 'Journal', 'Journal des révisions', ['Date', 'Type', 'Famille / article', 'Révision', 'Action', 'Détail'], [18, 10, 22, 10, 22, 60])
    wp = feuille_tableau(wb, 'Rapports', 'Rapports enregistrés (données)',
                         ['ID', 'Date', 'Article', 'Lot', 'Commande', 'Fournisseur', 'Contrôleur', 'Qté commandée', 'Qté réceptionnée', 'Qté échantillon',
                          'Durée', 'Origine', 'Décision', 'Type NC', 'Sous-type NC', 'Commentaires', 'Famille', 'Rév. gamme', 'Rév. TS', 'Conformes',
                          'Non conformes', 'Créé le', 'Modifié le'],
                         [18, 11, 18, 14, 14, 14, 14, 10, 10, 10, 10, 20, 22, 20, 24, 30, 18, 12, 12, 10, 10, 17, 17],
                         "Rempli par « Enregistrer le rapport » : ne pas modifier à la main.")
    wme = feuille_tableau(wb, 'Mesures', 'Mesures (une ligne par relevé)',
                          ['ID rapport', 'ID point', 'Code', 'Spécification', 'Criticité', 'Moyen', 'Méthode', 'Type spéc.', 'Mini', 'Nominal', 'Maxi',
                           'Tol -', 'Tol +', 'N° moyen', 'Pièce (0 = générale)', 'Valeur', 'Résultat', 'Article', 'Lot', 'Fournisseur', 'Date'],
                          [18, 10, 8, 32, 11, 18, 18, 12, 9, 9, 9, 8, 8, 10, 10, 10, 9, 18, 14, 14, 11],
                          "Rempli par « Enregistrer le rapport » : ne pas modifier à la main.")
    for r in range(4, 5000):
        wp.cell(r, 2).number_format = 'DD/MM/YYYY'
        wp.cell(r, 22).number_format = wp.cell(r, 23).number_format = 'DD/MM/YYYY HH:MM'
    wr.column_dimensions['E'].width = 13
    for r in range(4, 2000):
        wr.cell(r, 5).number_format = 'DD/MM/YYYY'
        wr.cell(r, 8).number_format = 'DD/MM/YYYY HH:MM'
        wj.cell(r, 1).number_format = 'DD/MM/YYYY HH:MM'
    wme.conditional_formatting.add('A4:U100000', FormulaRule(formula=['$Q4="NC"'], font=Font(color=ROUGE)))

    # ------------------------------------------------ Paramètres et listes
    wpa = wb.create_sheet('Paramètres')
    wpa['A1'] = 'Paramètres'
    wpa['A1'].font = TITRE
    wpa.row_dimensions[2].height = 32
    wpa['A3'] = 'Mot de passe méthodiste'; wpa['B3'] = '1'
    wpa['A4'] = 'Quantité échantillon par défaut'; wpa['B4'] = 15
    for c in ('B3', 'B4'):
        wpa[c].border = CADRE
    wpa['A5'] = 'Le mot de passe protège la sortie du statut « Attente décision méthodes » et la réouverture des rapports Acceptés / Refusés.'
    wpa['A5'].font = ETIQ
    wpa.column_dimensions['A'].width = 32
    wpa.column_dimensions['B'].width = 12
    listes = [
        ('ListeOrigines', 'Origines', ["Contrôle d'entrée", 'Produit « déjà entré dans l\'usine »']),
        ('ListeDecisions', 'Décisions', ['Sans décision', 'Attente décision méthodes', 'Accepté', 'Refusé']),
        ('ListeTypesNC', 'Types NC', ['Type 1 - Réception documentaire', 'Type 2 - Réception', 'Type 3 - Contrôle']),
        ('ListeSousTypesNC', 'Sous-types NC', ['CCPU manquant', 'CCPU papier abîmé', 'Quantité manquante', 'Livraison de la mauvaise référence',
                                               'Mauvais conditionnement', 'Conditionnement abîmé', 'Mauvaise identification (carton, étiquette, etc.)',
                                               'Identification non lisible', 'Plan non respecté (dimension, géométrie, état de surface, marquage produit)',
                                               'Aspect visuel (chocs, rayures, criques, etc.)', 'CCPU erroné (manque information, mauvaise information)']),
        ('ListeCriticites', 'Criticités', ['Critique', 'Majeure', 'Mineure']),
        ('ListeTypesSpec', 'Types spéc.', ['Générale', 'Individuelle']),
        ('ListeOKNOK', 'OK / NOK', ['OK', 'NOK']),
    ]
    for k, (nom, titre, vals) in enumerate(listes):
        c = 4 + k
        wpa.cell(7, c, titre).font = GRAS
        wpa.cell(7, c).fill = ENTETE_FILL
        for i, v in enumerate(vals, 8):
            wpa.cell(i, c, v)
        wpa.column_dimensions[L(c)].width = 30 if k in (1, 2, 3) else 22
        nom_defini(wb, nom, "'Paramètres'!$%s$8:$%s$%d" % (L(c), L(c), 7 + len(vals)))
    wpa['D6'] = 'Listes proposées dans le rapport (vous pouvez les compléter)'
    wpa['D6'].font = ETIQ

    # ------------------------------------------------ Aide
    wai = wb.create_sheet('Aide')
    wai['A1'] = 'Mode d\'emploi'
    wai['A1'].font = TITRE
    wai.column_dimensions['A'].width = 120
    texte = [
        ('Démarrage', True),
        ("1. Si Excel affiche « Les macros ont été bloquées » : fermez le fichier, clic droit sur le fichier > Propriétés > cochez « Débloquer » > OK, puis rouvrez-le.", False),
        ("2. Si un bandeau jaune « Activer le contenu » apparaît : cliquez dessus. Les boutons des onglets s'installent alors automatiquement.", False),
        ("3. Enregistrez toujours le classeur au format .xlsm (classeur prenant en charge les macros).", False),
        ('', False),
        ('Faire un rapport de contrôle', True),
        ("1. Menu > « Nouveau rapport ». 2. Choisissez l'article (liste), renseignez le lot, la quantité échantillon, etc.", False),
        ("3. « Préparer le rapport » : les points de la gamme de la famille et les valeurs attendues (TS) s'affichent.", False),
        ("4. Saisissez les mesures pièce par pièce (avec une virgule), OK / NOK pour les points qualitatifs et pour les contrôles généraux (colonne Résultat).", False),
        ("   Les valeurs hors tolérance sont en rouge ; le nombre de non-conformités s'affiche par point (colonne Résultat) et par pièce (ligne « NC / pièce »).", False),
        ("5. « Enregistrer le rapport » : le rapport rejoint la liste du Menu et le classeur est enregistré.", False),
        ("   Décision « Accepté » ou « Refusé » : le rapport est figé ; « Rouvrir (méthodiste) » demande le mot de passe (onglet Paramètres).", False),
        ("6. Pour rouvrir un rapport : Menu > cliquez sur sa ligne > « Ouvrir le rapport sélectionné ».", False),
        ('', False),
        ('Référentiel', True),
        ("GAMSERIE : points de contrôle de chaque famille (code, spécification, moyen, méthode). Chaque point a un identifiant (ID) qui ne change jamais : ne le modifiez pas.", False),
        ("TS : valeurs de chaque article (Nominal, Tol -, Tol +, Type spéc. Générale / Individuelle), une ligne par point de la gamme. Famille et Code sont lus automatiquement.", False),
        ("Matrice : famille, fournisseur et références documentaires de chaque article. « Nouvel article » crée l'article et prépare ses lignes TS.", False),
        ("Documents : « Ajouter des PDF » ; un PDF nommé comme une référence de la Matrice (ex. PM-0081.pdf) s'affiche dans le rapport de l'article.", False),
        ("Révisions : avant de modifier une gamme ou des valeurs TS, « Ouvrir une révision » ; après modification, « Valider la révision » (valideur, motif).", False),
        ("Cohérence : liste les articles sans famille, les points sans valeur, les valeurs hors gamme et les modifications faites sans révision.", False),
        ("Paramètres > « Importer le référentiel d'un classeur » : reprend les onglets GAMSERIE, TS et Matrice d'un autre classeur (par exemple celui de l'application web).", False),
        ('', False),
        ('Travail à plusieurs', True),
        ("Un classeur avec macros ne s'ouvre en modification que par une personne à la fois (les autres l'ont en lecture seule). Fermez-le après usage.", False),
    ]
    for i, (t, gras) in enumerate(texte, 3):
        wai.cell(i, 1, t).font = GRAS if gras else Font(name='Calibri', size=11, color=ENCRE)
        wai.cell(i, 1).alignment = Alignment(wrap_text=True, vertical='top')

    # couleurs des onglets
    for nom, coul in [('Menu', BLEU), ('Rapport', BLEU), ('GAMSERIE', '7C3AED'), ('TS', '7C3AED'), ('Matrice', '7C3AED'), ('Documents', '7C3AED'),
                      ('Analyses', '0891B2'), ('Cohérence', 'D97706'), ('Révisions', '6B7280'), ('Historique', '9CA3AF'), ('Journal', '9CA3AF'),
                      ('Rapports', '9CA3AF'), ('Mesures', '9CA3AF'), ('Paramètres', '6B7280'), ('Aide', '15803D')]:
        wb[nom].sheet_properties.tabColor = coul
    wb.active = 0
    return wb


def main():
    tmp = tempfile.mkdtemp(prefix='ec-xlsm-')
    base = os.path.join(tmp, 'macros.xlsm')
    compiler_vba(base)
    wb = construire(base)
    wb.save(SORTIE)
    print('Classeur écrit :', SORTIE, os.path.getsize(SORTIE), 'octets')
    # copies des modules au format attendu par l'éditeur VBA d'Excel (Windows-1252, CRLF), en secours
    dest = os.path.join(ICI, 'modules-a-importer')
    os.makedirs(dest, exist_ok=True)
    for nom in sorted(os.listdir(os.path.join(ICI, 'vba'))):
        if nom.endswith('.bas'):
            texte = open(os.path.join(ICI, 'vba', nom), encoding='utf-8').read()
            with open(os.path.join(dest, nom), 'wb') as f:
                f.write(texte.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252'))
    print('Modules à importer :', dest)


if __name__ == '__main__':
    main()
