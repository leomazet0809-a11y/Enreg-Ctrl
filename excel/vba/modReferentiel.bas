Attribute VB_Name = "modReferentiel"
' =====================================================================
'  Référentiel : identifiants des points, création d'article,
'  lignes TS manquantes, cohérence, documents, import d'un classeur.
' =====================================================================
Option Explicit

' Donne un identifiant unique à chaque ligne GAMSERIE qui n'en a pas (ou dont l'identifiant est en double)
Public Function CompleterIds(Optional ByVal avecMessage As Boolean = True) As Long
    Dim ws As Worksheet, r As Long, fin As Long, vus As New Collection, pris As Collection, id As String, n As Long
    Set ws = Fe(F_GAM)
    Set pris = IdsGamme()
    fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
    For r = L_DEB To fin
        If Txt(ws.Cells(r, G_FAM).Value) <> "" Or Txt(ws.Cells(r, G_CODE).Value) <> "" Or Txt(ws.Cells(r, G_SPEC).Value) <> "" Then
            id = Txt(ws.Cells(r, G_ID).Value)
            If id = "" Or CollContient(vus, id) Then
                id = NouvelId(pris)
                ws.Cells(r, G_ID).Value = id
                n = n + 1
            End If
            vus.Add id, id
        End If
    Next r
    If avecMessage Then
        If n > 0 Then
            Informer n & " identifiant(s) attribué(s) aux points de contrôle."
        Else
            Informer "Tous les points de contrôle ont un identifiant."
        End If
    End If
    CompleterIds = n
End Function

Public Sub CompleterIdsBouton()
    CompleterIds True
End Sub

' Type de spécification le plus utilisé pour ce point par les autres articles de la famille
Private Function TypeSpecUsuel(ByVal famille As String, ByVal id As String, ByVal sauf As String) As String
    Dim ws As Worksheet, r As Long, fin As Long, ty As String, nG As Long, nI As Long
    Set ws = Fe(F_TS)
    fin = DerniereLigne(ws, T_ART)
    For r = L_DEB To fin
        If Txt(ws.Cells(r, T_ID).Value) = id And Txt(ws.Cells(r, T_ART).Value) <> sauf Then
            ty = Normaliser(Txt(ws.Cells(r, T_TYPE).Value))
            If ty = "generale" Then
                nG = nG + 1
            ElseIf ty = "individuelle" Then
                nI = nI + 1
            End If
        End If
    Next r
    If nG > nI Then
        TypeSpecUsuel = T_GEN
    ElseIf nI > 0 Then
        TypeSpecUsuel = T_IND
    End If
End Function

' Ajoute dans TS une ligne par point de la gamme absent pour cet article ; modele = article dont on copie les valeurs
Public Function PreparerTS(ByVal article As String, Optional ByVal modele As String = "") As Long
    Dim wg As Worksheet, wt As Worksheet, idx As Collection, lg As Variant, id As String, d As Long, rm As Long, n As Long, famille As String
    Set wg = Fe(F_GAM): Set wt = Fe(F_TS)
    famille = FamilleDe(article)
    If famille = "" Then Exit Function
    Set idx = IndexTS()
    d = DerniereLigne(wt, T_ART) + 1
    For Each lg In LignesGamme(famille)
        id = Txt(wg.Cells(lg, G_ID).Value)
        If id <> "" And Not CollContient(idx, article & "|" & id) Then
            wt.Cells(d, T_ID).Value = id
            wt.Cells(d, T_ART).Value = article
            rm = 0
            If modele <> "" Then rm = CollLire(idx, modele & "|" & id, 0)
            If rm > 0 Then
                wt.Cells(d, T_TYPE).Value = wt.Cells(rm, T_TYPE).Value
                wt.Cells(d, T_NOM).Value = wt.Cells(rm, T_NOM).Value
                wt.Cells(d, T_TM).Value = wt.Cells(rm, T_TM).Value
                wt.Cells(d, T_TP).Value = wt.Cells(rm, T_TP).Value
            Else
                wt.Cells(d, T_TYPE).Value = TypeSpecUsuel(famille, id, article)
            End If
            d = d + 1: n = n + 1
        End If
    Next lg
    PreparerTS = n
End Function

' ----- Nouvel article : ses lignes TS sont préparées à partir de la gamme de sa famille -----
Public Sub CreerArticle()
    Dim article As String, famille As String, four As String, modele As String, fams As Collection, f As Variant, liste As String, i As Long, rep As String
    Dim wm As Worksheet, rm As Long, n As Long, modeles As String, wt As Worksheet, r As Long, vus As New Collection, a As String
    CompleterIds False
    article = Trim(Demander("Référence du nouvel article :"))
    If article = "" Then Exit Sub
    If ArticleARenseigne(article) Then Alerter "L'article « " & article & " » existe déjà dans TS avec des valeurs.": Exit Sub
    Set fams = ListeFamilles()
    If fams.Count = 0 Then Alerter "Aucune famille de contrôle dans GAMSERIE.": Exit Sub
    i = 0
    For Each f In fams
        i = i + 1
        liste = liste & vbLf & "  " & i & " = " & f & " (" & LignesGamme(CStr(f)).Count & " points)"
    Next f
    rep = Trim(Demander("Famille de contrôle de « " & article & " » (numéro ou nom) :" & liste, , "1"))
    If rep = "" Then Exit Sub
    If IsNumeric(rep) Then
        If Val(rep) >= 1 And Val(rep) <= fams.Count Then famille = fams(CLng(Val(rep)))
    ElseIf CollContient(fams, rep) Then
        famille = rep
    End If
    If famille = "" Then Alerter "Famille inconnue : « " & rep & " ».": Exit Sub
    four = Trim(Demander("Code fournisseur (facultatif) :"))
    ' articles de la même famille ayant des valeurs : modèles possibles
    Set wt = Fe(F_TS)
    For r = L_DEB To DerniereLigne(wt, T_ART)
        a = Txt(wt.Cells(r, T_ART).Value)
        If a <> "" And a <> article And Not CollContient(vus, a) Then
            vus.Add a, a
            If FamilleDe(a) = famille And ArticleARenseigne(a) Then modeles = modeles & IIf(modeles = "", "", ", ") & a
        End If
    Next r
    If modeles <> "" Then
        modele = Trim(Demander("Copier les valeurs d'un article de la même famille ?" & vbLf & "Laissez vide pour des valeurs à saisir." & vbLf & vbLf & "Articles : " & modeles))
        If modele <> "" And Not CollContient(vus, modele) Then Alerter "Article modèle inconnu : « " & modele & " » ; les valeurs seront à saisir.": modele = ""
    End If
    EcranFige True
    Set wm = Fe(F_MAT)
    rm = LigneMatrice(article)
    If rm = 0 Then
        rm = DerniereLigne(wm, M_ART) + 1
        wm.Cells(rm, M_ART).Value = article
        wm.Cells(rm, M_FOUR).Value = four
    ElseIf four <> "" And Txt(wm.Cells(rm, M_FOUR).Value) = "" Then
        wm.Cells(rm, M_FOUR).Value = four
    End If
    wm.Cells(rm, M_FAM).Value = famille
    n = PreparerTS(article, modele)
    If LigneRevision(RV_TS, article) = 0 Then AjouterRevision RV_TS, article, 0, "Brouillon", "", "Création de l'article"
    Journaliser RV_TS, article, 0, "Article créé", "famille " & famille & IIf(modele <> "", ", valeurs copiées de " & modele, "")
    EcranFige False
    Fe(F_TS).Activate
    Informer "Article « " & article & " » créé : " & n & " point(s) de contrôle repris de la gamme « " & famille & " »." & vbLf & _
             IIf(modele <> "", "Valeurs copiées de " & modele & ".", "Saisissez les valeurs (Nominal, Tol -, Tol +) dans l'onglet TS.") & vbLf & _
             "Validez ensuite la révision 0 (bouton « Valider la révision »)."
End Sub

' L'article a-t-il au moins une valeur renseignée dans TS ?
Private Function ArticleARenseigne(ByVal article As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = Fe(F_TS)
    For r = L_DEB To DerniereLigne(ws, T_ART)
        If Txt(ws.Cells(r, T_ART).Value) = article Then
            If Txt(ws.Cells(r, T_NOM).Value) <> "" Or Txt(ws.Cells(r, T_TM).Value) <> "" Or Txt(ws.Cells(r, T_TP).Value) <> "" Then ArticleARenseigne = True: Exit Function
        End If
    Next r
End Function

' ----- Lignes TS manquantes pour tous les articles de la Matrice -----
Public Sub PreparerLignesTSManquantes()
    Dim wm As Worksheet, r As Long, n As Long, a As String
    CompleterIds False
    Set wm = Fe(F_MAT)
    EcranFige True
    For r = L_DEB To DerniereLigne(wm, M_ART)
        a = Txt(wm.Cells(r, M_ART).Value)
        If a <> "" Then n = n + PreparerTS(a)
    Next r
    EcranFige False
    If n > 0 Then
        Informer n & " ligne(s) ajoutée(s) dans TS : saisissez leurs valeurs."
    Else
        Informer "Aucune ligne TS manquante."
    End If
End Sub

' ----- Cohérence du paramétrage -----
Public Sub VerifierCoherence()
    Dim wc As Worksheet, wg As Worksheet, wt As Worksheet, wm As Worksheet, idx As Collection, arts As New Collection, a As Variant, fam As String
    Dim lg As Variant, id As String, rt As Long, manque As String, n As Long, d As Long, r As Long, ids As Collection, horsG As String, f As Variant
    Dim gamIds As Collection, nb As Long
    CompleterIds False
    Set wc = Fe(F_COH): Set wg = Fe(F_GAM): Set wt = Fe(F_TS): Set wm = Fe(F_MAT)
    EcranFige True
    wc.Range(wc.Cells(L_DEB, 1), wc.Cells(DerniereLigne(wc, 1) + 1, 4)).ClearContents
    d = L_DEB
    ' articles connus (Matrice et TS)
    For r = L_DEB To DerniereLigne(wm, M_ART)
        a = Txt(wm.Cells(r, M_ART).Value)
        If a <> "" Then
            If Not CollContient(arts, CStr(a)) Then arts.Add a, CStr(a)
        End If
    Next r
    For r = L_DEB To DerniereLigne(wt, T_ART)
        a = Txt(wt.Cells(r, T_ART).Value)
        If a <> "" Then
            If Not CollContient(arts, CStr(a)) Then arts.Add a, CStr(a)
        End If
    Next r
    Set idx = IndexTS()
    For Each a In arts
        fam = FamilleDe(CStr(a))
        If fam = "" Then
            Ecrire wc, d, "Article sans famille", a, "", IIf(LigneMatrice(CStr(a)) = 0, "Article absent de la Matrice", "Famille vide dans la Matrice")
        ElseIf LignesGamme(fam).Count = 0 Then
            Ecrire wc, d, "Famille inconnue", a, fam, "« " & fam & " » absente de GAMSERIE (vérifiez l'orthographe)"
        Else
            manque = "": nb = 0
            Set gamIds = New Collection
            For Each lg In LignesGamme(fam)
                id = Txt(wg.Cells(lg, G_ID).Value)
                If id <> "" Then
                    If Not CollContient(gamIds, id) Then gamIds.Add id, id
                    rt = CollLire(idx, a & "|" & id, 0)
                    If rt = 0 Then
                        nb = nb + 1: manque = manque & IIf(manque = "", "", ", ") & Libelle(wg, lg)
                    ElseIf Txt(wt.Cells(rt, T_NOM).Value) = "" Then
                        nb = nb + 1: manque = manque & IIf(manque = "", "", ", ") & Libelle(wg, lg)
                    End If
                End If
            Next lg
            If nb > 0 Then Ecrire wc, d, "Points sans valeur dans TS", a, fam, nb & " point(s) : " & manque
            horsG = ""
            For r = L_DEB To DerniereLigne(wt, T_ART)
                If Txt(wt.Cells(r, T_ART).Value) = a Then
                    id = Txt(wt.Cells(r, T_ID).Value)
                    If Not CollContient(gamIds, id) Then horsG = horsG & IIf(horsG = "", "", ", ") & IIf(id = "", "(sans ID, ligne " & r & ")", id)
                End If
            Next r
            If horsG <> "" Then Ecrire wc, d, "Valeurs TS hors gamme", a, fam, "Points absents de la gamme de la famille : " & horsG
            If ModifieSansRevision(RV_TS, CStr(a)) Then Ecrire wc, d, "Modifié sans révision", a, fam, "Valeurs TS modifiées alors que la " & LibelleRevision(RV_TS, CStr(a)) & " est figée : ouvrez une révision puis validez-la"
        End If
    Next a
    For Each f In ListeFamilles()
        If ModifieSansRevision(RV_GAM, CStr(f)) Then Ecrire wc, d, "Modifié sans révision", "", f, "Gamme modifiée alors que la " & LibelleRevision(RV_GAM, CStr(f)) & " est figée : ouvrez une révision puis validez-la"
    Next f
    For r = L_DEB To FinTableau(wg, G_ID, G_FAM, G_CODE, G_SPEC)
        If Txt(wg.Cells(r, G_FAM).Value) = "" And (Txt(wg.Cells(r, G_CODE).Value) <> "" Or Txt(wg.Cells(r, G_SPEC).Value) <> "") Then Ecrire wc, d, "Point sans famille", "", "", "GAMSERIE ligne " & r & " : " & Libelle(wg, r)
    Next r
    n = d - L_DEB
    Fe(F_MENU).Range("H7").Value = IIf(n = 0, "Cohérence : aucune incohérence", "Cohérence : " & n & " point(s) à vérifier")
    If n = 0 Then Ecrire wc, d, "Aucune incohérence", "", "", "Chaque article a une famille connue et chaque point de sa gamme a une valeur dans TS."
    wc.Range("A1").Value = "Cohérence du paramétrage - vérifiée le " & Format(Now, "dd/mm/yyyy hh:mm")
    EcranFige False
    wc.Activate
End Sub

Private Function Libelle(ByVal wg As Worksheet, ByVal r As Long) As String
    Dim c As String
    c = Txt(wg.Cells(r, G_CODE).Value)
    If c = "" Then c = Left(Txt(wg.Cells(r, G_SPEC).Value), 25)
    If c = "" Then c = Txt(wg.Cells(r, G_ID).Value)
    Libelle = c
End Function

Private Sub Ecrire(ByVal ws As Worksheet, ByRef d As Long, ByVal typ As String, ByVal art As Variant, ByVal fam As Variant, ByVal detail As String)
    ws.Cells(d, 1).Value = typ
    ws.Cells(d, 2).Value = art
    ws.Cells(d, 3).Value = fam
    ws.Cells(d, 4).Value = detail
    d = d + 1
End Sub

' ----- Documents PDF : référence = nom du fichier sans extension -----
Public Sub AjouterDocuments()
    Dim fichiers As Variant, i As Long, ws As Worksheet, chemin As String, ref As String, r As Long, k As Long, n As Long
    If ModeTest Then Exit Sub
    fichiers = Application.GetOpenFilename(FileFilter:="Documents PDF (*.pdf),*.pdf", Title:="Choisir des documents PDF", MultiSelect:=True)
    If Not IsArray(fichiers) Then Exit Sub
    Set ws = Fe(F_DOCS)
    For i = LBound(fichiers) To UBound(fichiers)
        chemin = CStr(fichiers(i))
        ref = Mid(chemin, InStrRev(chemin, "\") + 1)
        If LCase(Right(ref, 4)) = ".pdf" Then ref = Left(ref, Len(ref) - 4)
        r = 0
        For k = L_DEB To DerniereLigne(ws, D_REF)
            If LCase(Txt(ws.Cells(k, D_REF).Value)) = LCase(ref) Then r = k: Exit For
        Next k
        If r = 0 Then r = DerniereLigne(ws, D_REF) + 1
        ws.Cells(r, D_REF).Value = ref
        ws.Cells(r, D_CHEMIN).Value = chemin
        ws.Cells(r, D_DATE).Value = Now
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, D_REF), Address:=chemin, TextToDisplay:=ref
        On Error GoTo 0
        n = n + 1
    Next i
    Informer n & " document(s) ajouté(s). Ils s'affichent dans les rapports des articles dont la Matrice cite leur référence."
End Sub

' ----- Import du référentiel depuis un autre classeur (par exemple celui de l'application web) -----
Public Sub ImporterReferentiel()
    Dim fichier As Variant, wb As Workbook, n As String
    If ModeTest Then Exit Sub
    fichier = Application.GetOpenFilename(FileFilter:="Classeurs Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", Title:="Classeur contenant les onglets GAMSERIE, TS et Matrice")
    If VarType(fichier) = vbBoolean Then Exit Sub
    If Not Confirmer("Remplacer GAMSERIE, TS et Matrice par les onglets du même nom de ce classeur ?" & vbLf & CStr(fichier) & vbLf & vbLf & "Les rapports enregistrés sont conservés.") Then Exit Sub
    On Error GoTo echec
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=CStr(fichier), ReadOnly:=True, UpdateLinks:=0)
    n = ImporterClasseur(wb)
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Informer n
    Exit Sub
echec:
    Application.ScreenUpdating = True
    Alerter "Import impossible : " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' Importe les onglets GAMSERIE, TS et Matrice d'un classeur ouvert ; renvoie un compte rendu
Public Function ImporterClasseur(ByVal wb As Workbook) As String
    Dim src As Worksheet, cr As String, nG As Long, nT As Long, nM As Long
    If FeuilleExiste(wb, "GAMSERIE") Then nG = ImporterGamserie(wb.Worksheets("GAMSERIE")): cr = cr & vbLf & "  - GAMSERIE : " & nG & " ligne(s)"
    If FeuilleExiste(wb, "Matrice") Then nM = ImporterMatrice(wb.Worksheets("Matrice")): cr = cr & vbLf & "  - Matrice : " & nM & " article(s)"
    If FeuilleExiste(wb, "TS") Then nT = ImporterTS(wb.Worksheets("TS")): cr = cr & vbLf & "  - TS : " & nT & " valeur(s)"
    If cr = "" Then ImporterClasseur = "Aucun onglet GAMSERIE, TS ou Matrice dans ce classeur.": Exit Function
    CompleterIds False
    InitialiserRevisions
    Journaliser "Import", "", "", "Référentiel importé", Replace(cr, vbLf, " ")
    ImporterClasseur = "Référentiel importé :" & cr & vbLf & vbLf & "Vérifiez le résultat avec « Cohérence »."
End Function

Private Function FeuilleExiste(ByVal wb As Workbook, ByVal nom As String) As Boolean
    Dim s As Object
    For Each s In wb.Worksheets
        If LCase(s.Name) = LCase(nom) Then FeuilleExiste = True: Exit Function
    Next s
End Function

' Ligne d'en-têtes (parmi les 5 premières) contenant l'en-tête attendu, et index des colonnes
Private Function Entetes(ByVal ws As Worksheet, ByVal attendu As String, ByRef ligne As Long) As Collection
    Dim r As Long, c As Long, h As String, cols As Collection, dern As Long
    For r = 1 To 5
        Set cols = New Collection
        dern = ws.Cells(r, ws.Columns.Count).End(-4159).Column   ' -4159 = xlToLeft
        For c = 1 To dern
            h = Normaliser(Txt(ws.Cells(r, c).Value))
            If h <> "" Then
                If Not CollContient(cols, h) Then cols.Add c, h
            End If
        Next c
        If CollContient(cols, attendu) Then ligne = r: Set Entetes = cols: Exit Function
    Next r
    ligne = 0
    Set Entetes = New Collection
End Function

Private Function Col(ByVal cols As Collection, ParamArray noms() As Variant) As Long
    Dim i As Long
    For i = LBound(noms) To UBound(noms)
        If CollContient(cols, CStr(noms(i))) Then Col = cols(CStr(noms(i))): Exit Function
    Next i
    Col = 0
End Function

Private Function Cel(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then Cel = "" Else Cel = ws.Cells(r, c).Value
End Function

Private Function ImporterGamserie(ByVal src As Worksheet) As Long
    Dim cols As Collection, l0 As Long, r As Long, d As Long, ws As Worksheet, cI As Long, cF As Long, cC As Long, cCr As Long, cS As Long, cM As Long, cMe As Long, fin As Long
    Set cols = Entetes(src, "code", l0)
    If l0 = 0 Then Exit Function
    cI = Col(cols, "id", "identifiant"): cF = Col(cols, "famillemethode", "famille"): cC = Col(cols, "code")
    cCr = Col(cols, "criticite"): cS = Col(cols, "specification"): cM = Col(cols, "moyendecontrole", "moyen"): cMe = Col(cols, "methodedecontrole", "methode")
    Set ws = Fe(F_GAM)
    fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
    If fin >= L_DEB Then ws.Range(ws.Cells(L_DEB, 1), ws.Cells(fin, G_METH)).ClearContents
    d = L_DEB
    For r = l0 + 1 To src.Cells(src.Rows.Count, IIf(cF > 0, cF, cC)).End(-4162).Row
        If Txt(Cel(src, r, cF)) <> "" Or Txt(Cel(src, r, cC)) <> "" Or Txt(Cel(src, r, cS)) <> "" Then
            ws.Cells(d, G_ID).Value = Cel(src, r, cI)
            ws.Cells(d, G_FAM).Value = Cel(src, r, cF)
            ws.Cells(d, G_CODE).Value = Cel(src, r, cC)
            ws.Cells(d, G_CRIT).Value = Cel(src, r, cCr)
            ws.Cells(d, G_SPEC).Value = Cel(src, r, cS)
            ws.Cells(d, G_MOY).Value = Cel(src, r, cM)
            ws.Cells(d, G_METH).Value = Cel(src, r, cMe)
            d = d + 1
        End If
    Next r
    ImporterGamserie = d - L_DEB
    CompleterIds False
End Function

Private Function ImporterMatrice(ByVal src As Worksheet) As Long
    Dim cols As Collection, l0 As Long, r As Long, d As Long, ws As Worksheet, c As Long, k As Long, fin As Long, cs(1 To 12) As Long
    Set cols = Entetes(src, "article", l0)
    If l0 = 0 Then Exit Function
    cs(1) = Col(cols, "article", "reference"): cs(2) = Col(cols, "codefournisseur", "fournisseur"): cs(3) = Col(cols, "famille", "famillemethode")
    cs(4) = Col(cols, "matiere"): cs(5) = Col(cols, "liassedocumentaire", "liasse"): cs(6) = Col(cols, "refnorme")
    cs(7) = Col(cols, "refplanmethode"): cs(8) = Col(cols, "refplancontrole"): cs(9) = Col(cols, "refplandesituation")
    cs(10) = Col(cols, "refconditionnement"): cs(11) = Col(cols, "reftauxechantillon"): cs(12) = Col(cols, "refdefautheque")
    Set ws = Fe(F_MAT)
    fin = DerniereLigne(ws, M_ART)
    If fin >= L_DEB Then ws.Range(ws.Cells(L_DEB, 1), ws.Cells(fin, M_REFN)).ClearContents
    d = L_DEB
    For r = l0 + 1 To src.Cells(src.Rows.Count, cs(1)).End(-4162).Row
        If Txt(Cel(src, r, cs(1))) <> "" Then
            For k = 1 To 12
                ws.Cells(d, k).Value = Cel(src, r, cs(k))
            Next k
            d = d + 1
        End If
    Next r
    ImporterMatrice = d - L_DEB
End Function

Private Function ImporterTS(ByVal src As Worksheet) As Long
    Dim cols As Collection, l0 As Long, r As Long, d As Long, ws As Worksheet, fin As Long, wg As Worksheet
    Dim cI As Long, cA As Long, cC As Long, cF As Long, cTy As Long, cN As Long, cTm As Long, cTp As Long, id As String, art As String, fam As String, lg As Variant, code As String
    Set cols = Entetes(src, "article", l0)
    If l0 = 0 Then Exit Function
    cI = Col(cols, "id", "identifiant"): cA = Col(cols, "article", "reference"): cC = Col(cols, "code"): cF = Col(cols, "famille", "famillemethode")
    cTy = Col(cols, "typespec", "typespecification"): cN = Col(cols, "nominal", "nom"): cTm = Col(cols, "tol-", "tolmoins"): cTp = Col(cols, "tol+", "tolplus")
    Set ws = Fe(F_TS): Set wg = Fe(F_GAM)
    fin = DerniereLigne(ws, T_ART)
    If fin >= L_DEB Then
        ws.Range(ws.Cells(L_DEB, T_ID), ws.Cells(fin, T_ART)).ClearContents
        ws.Range(ws.Cells(L_DEB, T_TYPE), ws.Cells(fin, T_TP)).ClearContents
    End If
    d = L_DEB
    For r = l0 + 1 To src.Cells(src.Rows.Count, cA).End(-4162).Row
        art = Txt(Cel(src, r, cA))
        If art <> "" Then
            id = Txt(Cel(src, r, cI))
            If id = "" Or Not CollContient(IdsGamme(), id) Then
                ' sans identifiant : point retrouvé par son code dans la gamme de la famille
                fam = FamilleDe(art)
                If fam = "" Then fam = Txt(Cel(src, r, cF))
                code = Txt(Cel(src, r, cC)): id = ""
                For Each lg In LignesGamme(fam)
                    If MemeCode(Txt(wg.Cells(lg, G_CODE).Value), code) Then id = Txt(wg.Cells(lg, G_ID).Value): Exit For
                Next lg
            End If
            If id <> "" Then
                ws.Cells(d, T_ID).Value = id
                ws.Cells(d, T_ART).Value = art
                ws.Cells(d, T_TYPE).Value = Cel(src, r, cTy)
                ws.Cells(d, T_NOM).Value = Cel(src, r, cN)
                ws.Cells(d, T_TM).Value = Cel(src, r, cTm)
                ws.Cells(d, T_TP).Value = Cel(src, r, cTp)
                d = d + 1
            End If
        End If
    Next r
    ImporterTS = d - L_DEB
End Function
