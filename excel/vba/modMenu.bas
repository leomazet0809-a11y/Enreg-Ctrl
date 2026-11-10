Attribute VB_Name = "modMenu"
' =====================================================================
'  Ouverture du classeur, boutons, navigation, liste des rapports
'  enregistrés (onglet Menu) et analyses des dérives.
' =====================================================================
Option Explicit

Private Const MENU_L0 As Long = 9          ' première ligne de la liste des rapports (en-têtes ligne 8)
Private Const MENU_NBCOL As Long = 12

' Exécutée par Excel à l'ouverture du classeur (macros activées)
Public Sub Auto_Open()
    On Error Resume Next
    EcranFige True
    CompleterIds False
    InitialiserRevisions
    InstallerBoutons
    ActualiserMenu
    EcranFige False
    Fe(F_MENU).Activate
    Fe(F_MENU).Range("A1").Select
End Sub

' ----- boutons (formes reliées aux macros) -----
Public Sub InstallerBoutons()
    Dim x As Double
    x = 6
    x = Bouton(Fe(F_MENU), "Nouveau rapport", "AllerNouveauRapport", x, 3, True)
    x = Bouton(Fe(F_MENU), "Ouvrir le rapport sélectionné", "OuvrirRapportSelectionne", x, 3)
    x = Bouton(Fe(F_MENU), "Supprimer le rapport sélectionné", "SupprimerRapportSelectionne", x, 3, , True)
    x = Bouton(Fe(F_MENU), "Actualiser la liste", "ActualiserMenu", x, 3)
    x = 6
    x = Bouton(Fe(F_MENU), "GAMSERIE", "AllerGamserie", x, 5)
    x = Bouton(Fe(F_MENU), "TS", "AllerTS", x, 5)
    x = Bouton(Fe(F_MENU), "Matrice", "AllerMatrice", x, 5)
    x = Bouton(Fe(F_MENU), "Documents", "AllerDocuments", x, 5)
    x = Bouton(Fe(F_MENU), "Révisions", "AllerRevisions", x, 5)
    x = Bouton(Fe(F_MENU), "Cohérence", "VerifierCoherence", x, 5)
    x = Bouton(Fe(F_MENU), "Analyses des dérives", "AllerAnalyses", x, 5)
    x = Bouton(Fe(F_MENU), "Paramètres", "AllerParametres", x, 5)

    x = 6
    x = Bouton(Fe(F_RAPPORT), "Menu", "AllerMenu", x, 2)
    x = Bouton(Fe(F_RAPPORT), "Préparer le rapport", "PreparerRapport", x, 2, True)
    x = Bouton(Fe(F_RAPPORT), "Enregistrer le rapport", "EnregistrerRapport", x, 2, True)
    x = Bouton(Fe(F_RAPPORT), "Nouveau rapport", "NouveauRapport", x, 2)
    x = Bouton(Fe(F_RAPPORT), "Vider les mesures", "ViderMesures", x, 2)
    x = Bouton(Fe(F_RAPPORT), "Exporter en PDF", "ExporterRapportPDF", x, 2)
    x = Bouton(Fe(F_RAPPORT), "Rouvrir (méthodiste)", "RouvrirRapport", x, 2)

    x = 6
    x = Bouton(Fe(F_GAM), "Menu", "AllerMenu", x, 2)
    x = Bouton(Fe(F_GAM), "Ouvrir une révision", "OuvrirRevision", x, 2)
    x = Bouton(Fe(F_GAM), "Valider la révision", "ValiderRevision", x, 2, True)
    x = Bouton(Fe(F_GAM), "Abandonner la révision", "AbandonnerRevision", x, 2, , True)
    x = Bouton(Fe(F_GAM), "Compléter les ID", "CompleterIdsBouton", x, 2)
    x = Bouton(Fe(F_GAM), "Lignes TS manquantes", "PreparerLignesTSManquantes", x, 2)

    x = 6
    x = Bouton(Fe(F_TS), "Menu", "AllerMenu", x, 2)
    x = Bouton(Fe(F_TS), "Nouvel article", "CreerArticle", x, 2, True)
    x = Bouton(Fe(F_TS), "Lignes TS manquantes", "PreparerLignesTSManquantes", x, 2)
    x = Bouton(Fe(F_TS), "Ouvrir une révision", "OuvrirRevision", x, 2)
    x = Bouton(Fe(F_TS), "Valider la révision", "ValiderRevision", x, 2, True)
    x = Bouton(Fe(F_TS), "Abandonner la révision", "AbandonnerRevision", x, 2, , True)

    x = 6
    x = Bouton(Fe(F_MAT), "Menu", "AllerMenu", x, 2)
    x = Bouton(Fe(F_MAT), "Nouvel article", "CreerArticle", x, 2, True)

    x = 6
    x = Bouton(Fe(F_DOCS), "Menu", "AllerMenu", x, 2)
    x = Bouton(Fe(F_DOCS), "Ajouter des PDF", "AjouterDocuments", x, 2, True)

    x = Bouton(Fe(F_REV), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_REV), "Historique", "AllerHistorique", x, 2)
    x = Bouton(Fe(F_REV), "Journal", "AllerJournal", x, 2)
    x = Bouton(Fe(F_HIST), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_HIST), "Révisions", "AllerRevisions", x, 2)
    x = Bouton(Fe(F_JOUR), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_JOUR), "Révisions", "AllerRevisions", x, 2)
    x = Bouton(Fe(F_RAPS), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_MES), "Menu", "AllerMenu", 6, 2)

    x = Bouton(Fe(F_ANA), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_ANA), "Actualiser l'analyse", "ActualiserAnalyses", x, 2, True)
    x = Bouton(Fe(F_COH), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_COH), "Vérifier à nouveau", "VerifierCoherence", x, 2, True)
    x = Bouton(Fe(F_PAR), "Menu", "AllerMenu", 6, 2)
    x = Bouton(Fe(F_PAR), "Importer le référentiel d'un classeur", "ImporterReferentiel", x, 2)
    x = Bouton(Fe(F_PAR), "Réinstaller les boutons", "ReinstallerBoutons", x, 2)
End Sub

' Crée un bouton s'il n'existe pas ; renvoie la position du bouton suivant
Private Function Bouton(ByVal ws As Worksheet, ByVal texte As String, ByVal macro As String, ByVal x As Double, ByVal ligne As Long, Optional ByVal principal As Boolean = False, Optional ByVal danger As Boolean = False) As Double
    Dim nom As String, sh As Object, l As Double, h As Double, t As Double
    nom = "btn_" & macro & "_" & ligne
    l = 18 + Len(texte) * 6.2
    h = 24
    t = ws.Rows(ligne).Top + (ws.Rows(ligne).Height - h) / 2
    If t < 1 Then t = 1
    On Error Resume Next
    Set sh = ws.Shapes(nom)
    On Error GoTo 0
    If sh Is Nothing Then
        On Error Resume Next
        Set sh = ws.Shapes.AddShape(5, x, t, l, h)    ' 5 = rectangle à coins arrondis
        If sh Is Nothing Then Exit Function
        sh.Name = nom
        sh.OnAction = macro
        sh.Placement = 3                                ' ne bouge pas avec les cellules
        sh.Line.Visible = False
        If principal Then
            sh.Fill.ForeColor.RGB = RGB(37, 99, 235)
        ElseIf danger Then
            sh.Fill.ForeColor.RGB = RGB(254, 226, 226)
        Else
            sh.Fill.ForeColor.RGB = RGB(241, 243, 246)
        End If
        With sh.TextFrame
            .Characters.Text = texte
            .Characters.Font.Size = 10
            .Characters.Font.Bold = principal
            If principal Then
                .Characters.Font.Color = RGB(255, 255, 255)
            ElseIf danger Then
                .Characters.Font.Color = RGB(185, 28, 28)
            Else
                .Characters.Font.Color = RGB(17, 24, 39)
            End If
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
            .MarginLeft = 2: .MarginRight = 2
        End With
        On Error GoTo 0
    End If
    Bouton = sh.Left + sh.Width + 6
End Function

Public Sub ReinstallerBoutons()
    Dim ws As Worksheet, i As Long
    For Each ws In ThisWorkbook.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            If Left(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
        Next i
    Next ws
    InstallerBoutons
    Informer "Boutons réinstallés."
End Sub

' ----- navigation -----
Public Sub AllerMenu()
    ActualiserMenu
    Fe(F_MENU).Activate
End Sub
Public Sub AllerRapport()
    Fe(F_RAPPORT).Activate
End Sub
Public Sub AllerGamserie()
    Fe(F_GAM).Activate
End Sub
Public Sub AllerTS()
    Fe(F_TS).Activate
End Sub
Public Sub AllerMatrice()
    Fe(F_MAT).Activate
End Sub
Public Sub AllerDocuments()
    Fe(F_DOCS).Activate
End Sub
Public Sub AllerRevisions()
    Fe(F_REV).Activate
End Sub
Public Sub AllerHistorique()
    Fe(F_HIST).Activate
End Sub
Public Sub AllerJournal()
    Fe(F_JOUR).Activate
End Sub
Public Sub AllerAnalyses()
    Fe(F_ANA).Activate
End Sub
Public Sub AllerParametres()
    Fe(F_PAR).Activate
End Sub
Public Sub AllerNouveauRapport()
    NouveauRapport
End Sub

' ----- liste des rapports enregistrés (du plus récent au plus ancien) -----
Public Sub ActualiserMenu()
    Dim wm As Worksheet, wr As Worksheet, fin As Long, i As Long, j As Long, n As Long, t() As Variant, ordre() As Long, tmp As Long, a As Variant, b As Variant
    Set wm = Fe(F_MENU): Set wr = Fe(F_RAPS)
    fin = DerniereLigne(wr, 1)
    wm.Range(wm.Cells(MENU_L0, 1), wm.Cells(DerniereLigne(wm, 2) + 1, MENU_NBCOL)).ClearContents
    n = fin - L_DEB + 1
    If n <= 0 Then
        wm.Range("B7").Value = "Aucun rapport enregistré"
        wm.Cells(MENU_L0, 2).Value = "Aucun rapport enregistré : cliquez sur « Nouveau rapport »."
        Exit Sub
    End If
    ReDim ordre(1 To n)
    For i = 1 To n: ordre(i) = L_DEB + i - 1: Next i
    ' tri par date décroissante, puis création décroissante
    For i = 2 To n
        tmp = ordre(i): j = i - 1
        Do While j >= 1
            If PlusRecent(wr, tmp, ordre(j)) Then
                ordre(j + 1) = ordre(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        ordre(j + 1) = tmp
    Next i
    ReDim t(1 To n, 1 To MENU_NBCOL)
    For i = 1 To n
        t(i, 1) = ""
        t(i, 2) = wr.Cells(ordre(i), 2).Value      ' date
        t(i, 3) = wr.Cells(ordre(i), 3).Value      ' article
        t(i, 4) = wr.Cells(ordre(i), 4).Value      ' lot
        t(i, 5) = wr.Cells(ordre(i), 5).Value      ' commande
        t(i, 6) = wr.Cells(ordre(i), 6).Value      ' fournisseur
        t(i, 7) = wr.Cells(ordre(i), 7).Value      ' contrôleur
        t(i, 8) = wr.Cells(ordre(i), 10).Value     ' échantillon
        t(i, 9) = wr.Cells(ordre(i), 20).Value     ' conformes
        t(i, 10) = wr.Cells(ordre(i), 21).Value    ' non conformes
        t(i, 11) = IIf(Txt(wr.Cells(ordre(i), 13).Value) = "", "Sans décision", wr.Cells(ordre(i), 13).Value)
        t(i, 12) = wr.Cells(ordre(i), 1).Value     ' identifiant (colonne masquée)
    Next i
    wm.Range(wm.Cells(MENU_L0, 1), wm.Cells(MENU_L0 + n - 1, MENU_NBCOL)).Value = t
    wm.Range("B7").Value = n & " rapport(s) enregistré(s) - cliquez sur une ligne, puis sur « Ouvrir le rapport sélectionné »"
End Sub

Private Function PlusRecent(ByVal wr As Worksheet, ByVal a As Long, ByVal b As Long) As Boolean
    Dim da As Double, db As Double
    da = ValDate(wr.Cells(a, 2).Value): db = ValDate(wr.Cells(b, 2).Value)
    If da <> db Then PlusRecent = (da > db): Exit Function
    PlusRecent = (ValDate(wr.Cells(a, 22).Value) > ValDate(wr.Cells(b, 22).Value))
End Function

Private Function ValDate(ByVal v As Variant) As Double
    On Error Resume Next
    If IsDate(v) Then
        ValDate = CDbl(CDate(v))
    ElseIf IsNumeric(v) Then
        ValDate = CDbl(v)
    End If
End Function

Private Function IdSelectionne() As String
    Dim ws As Worksheet, r As Long
    Set ws = Fe(F_MENU)
    If ActiveSheet.Name <> ws.Name Then Exit Function
    r = ActiveCell.Row
    If r < MENU_L0 Then Exit Function
    IdSelectionne = Txt(ws.Cells(r, MENU_NBCOL).Value)
End Function

Public Sub OuvrirRapportSelectionne()
    Dim id As String
    id = IdSelectionne()
    If id = "" Then Alerter "Cliquez d'abord sur la ligne d'un rapport de la liste, puis sur « Ouvrir le rapport sélectionné ».": Exit Sub
    ChargerRapport id
End Sub

Public Sub SupprimerRapportSelectionne()
    Dim id As String, wr As Worksheet, wm As Worksheet, i As Long, desc As String
    id = IdSelectionne()
    If id = "" Then Alerter "Cliquez d'abord sur la ligne du rapport à supprimer.": Exit Sub
    Set wr = Fe(F_RAPS): Set wm = Fe(F_MES)
    For i = L_DEB To DerniereLigne(wr, 1)
        If Txt(wr.Cells(i, 1).Value) = id Then desc = Txt(wr.Cells(i, 3).Value) & " - lot " & Txt(wr.Cells(i, 4).Value)
    Next i
    If Not Confirmer("Supprimer définitivement le rapport " & desc & " ?") Then Exit Sub
    EcranFige True
    For i = DerniereLigne(wr, 1) To L_DEB Step -1
        If Txt(wr.Cells(i, 1).Value) = id Then wr.Rows(i).Delete
    Next i
    SupprimerMesures id
    If Txt(Fe(F_RAPPORT).Range(A_ID).Value) = id Then Fe(F_RAPPORT).Range(A_ETAT).Value = "Supprimé"
    ActualiserMenu
    EcranFige False
    SauverClasseur
End Sub

Private Sub SupprimerMesures(ByVal id As String)
    Dim wm As Worksheet, fin As Long, d As Variant, t() As Variant, i As Long, k As Long, n As Long
    Set wm = Fe(F_MES)
    fin = DerniereLigne(wm, 1)
    If fin < L_DEB Then Exit Sub
    d = wm.Range(wm.Cells(L_DEB, 1), wm.Cells(fin, MS_NB)).Value
    ReDim t(1 To UBound(d, 1), 1 To MS_NB)
    For i = 1 To UBound(d, 1)
        If Txt(d(i, 1)) <> id Then
            n = n + 1
            For k = 1 To MS_NB: t(n, k) = d(i, k): Next k
        End If
    Next i
    wm.Range(wm.Cells(L_DEB, 1), wm.Cells(fin, MS_NB)).ClearContents
    If n > 0 Then wm.Range(wm.Cells(L_DEB, 1), wm.Cells(L_DEB + UBound(d, 1) - 1, MS_NB)).Value = t
End Sub

' ----- Analyses des dérives : statistiques par article et par point, et détail des relevés -----
Public Sub ActualiserAnalyses()
    Dim wa As Worksheet, wm As Worksheet, fin As Long, d As Variant, i As Long, fArt As String, fLot As String, fFour As String, fPt As String
    Dim du As Double, au As Double, dt As Double, cle As String, stats As New Collection, ordre As New Collection, s As Variant, v As Double, ok As Boolean
    Dim r As Long, sortie() As Variant, n As Long, k As Long, det() As Variant, nd As Long, ligneDet As Long, moy As Double, et As Double
    Set wa = Fe(F_ANA): Set wm = Fe(F_MES)
    fArt = LCase(Txt(wa.Range("C4").Value)): fLot = LCase(Txt(wa.Range("C5").Value)): fFour = LCase(Txt(wa.Range("C6").Value)): fPt = LCase(Txt(wa.Range("C7").Value))
    du = ValDate(wa.Range("C8").Value): au = ValDate(wa.Range("C9").Value)
    EcranFige True
    fin = DerniereLigne(wa, 1) + 1
    If fin < 13 Then fin = 13
    wa.Range(wa.Cells(12, 1), wa.Cells(fin, 13)).ClearContents
    On Error Resume Next
    wa.Range(wa.Cells(12, 1), wa.Cells(fin, 13)).Interior.Pattern = -4142
    On Error GoTo 0
    fin = DerniereLigne(wm, 1)
    If fin < L_DEB Then wa.Range("A12").Value = "Aucun relevé enregistré.": EcranFige False: Exit Sub
    d = wm.Range(wm.Cells(L_DEB, 1), wm.Cells(fin, MS_NB)).Value
    ReDim det(1 To UBound(d, 1), 1 To 11)
    For i = 1 To UBound(d, 1)
        If Val(d(i, 15)) >= 1 Then                                  ' relevés pièce par pièce
            v = VersNombre(d(i, 16), ok)
            dt = ValDate(d(i, 21))
            If ok _
               And (fArt = "" Or InStr(LCase(Txt(d(i, 18))), fArt) > 0) _
               And (fLot = "" Or InStr(LCase(Txt(d(i, 19))), fLot) > 0) _
               And (fFour = "" Or InStr(LCase(Txt(d(i, 20))), fFour) > 0) _
               And (fPt = "" Or LCase(Txt(d(i, 3))) = fPt Or InStr(LCase(Txt(d(i, 4))), fPt) > 0) _
               And (du = 0 Or dt >= du) And (au = 0 Or dt <= au) Then
                cle = Txt(d(i, 18)) & "|" & Txt(d(i, 2))
                If Not CollContient(stats, cle) Then
                    stats.Add Array(Txt(d(i, 18)), Txt(d(i, 3)), Txt(d(i, 4)), d(i, 9), d(i, 10), d(i, 11), 0, 0#, 0#, 1E+308, -1E+308, 0), cle
                    ordre.Add cle
                End If
                s = stats(cle)
                s(6) = s(6) + 1: s(7) = s(7) + v: s(8) = s(8) + v * v
                If v < s(9) Then s(9) = v
                If v > s(10) Then s(10) = v
                If Txt(d(i, 17)) = "NC" Then s(11) = s(11) + 1
                s(3) = d(i, 9): s(4) = d(i, 10): s(5) = d(i, 11)       ' tolérances du relevé le plus récent
                stats.Remove cle: stats.Add s, cle
                nd = nd + 1
                det(nd, 1) = d(i, 21): det(nd, 2) = d(i, 18): det(nd, 3) = d(i, 19): det(nd, 4) = d(i, 20): det(nd, 5) = d(i, 3)
                det(nd, 6) = d(i, 4): det(nd, 7) = d(i, 15): det(nd, 8) = v: det(nd, 9) = d(i, 9): det(nd, 10) = d(i, 11): det(nd, 11) = d(i, 17)
            End If
        End If
    Next i
    If ordre.Count = 0 Then wa.Range("A12").Value = "Aucun relevé numérique ne correspond aux filtres.": EcranFige False: Exit Sub
    ' statistiques
    wa.Range("A12").Value = "Statistiques par article et par point de contrôle"
    wa.Range("A12").Font.Bold = True
    wa.Range("A13:M13").Value = Array("Article", "Code", "Spécification", "Mini", "Nominal", "Maxi", "Relevés", "Moyenne", "Min", "Max", "Écart-type", "Hors tolérance", "% hors tol.")
    wa.Range("A13:M13").Font.Bold = True
    ReDim sortie(1 To ordre.Count, 1 To 13)
    n = 0
    For Each s In ordre
        n = n + 1
        Dim x As Variant
        x = stats(CStr(s))
        moy = x(7) / x(6)
        If x(6) > 1 Then et = Sqr(Abs((x(8) - x(6) * moy * moy) / (x(6) - 1))) Else et = 0
        sortie(n, 1) = x(0): sortie(n, 2) = x(1): sortie(n, 3) = x(2): sortie(n, 4) = x(3): sortie(n, 5) = x(4): sortie(n, 6) = x(5)
        sortie(n, 7) = x(6): sortie(n, 8) = Round(moy, 4): sortie(n, 9) = x(9): sortie(n, 10) = x(10): sortie(n, 11) = Round(et, 4)
        sortie(n, 12) = x(11): sortie(n, 13) = Round(x(11) / x(6), 4)
    Next s
    wa.Range(wa.Cells(14, 1), wa.Cells(13 + n, 13)).Value = sortie
    For r = 14 To 13 + n
        If Val(wa.Cells(r, 12).Value) > 0 Then wa.Range(wa.Cells(r, 1), wa.Cells(r, 13)).Interior.Color = RGB(254, 226, 226)
    Next r
    ' détail des relevés
    ligneDet = 13 + n + 3
    wa.Cells(ligneDet, 1).Value = "Détail des relevés (" & nd & ")"
    wa.Cells(ligneDet, 1).Font.Bold = True
    wa.Range(wa.Cells(ligneDet + 1, 1), wa.Cells(ligneDet + 1, 11)).Value = Array("Date", "Article", "Lot", "Fournisseur", "Code", "Spécification", "Pièce", "Valeur", "Mini", "Maxi", "Résultat")
    wa.Range(wa.Cells(ligneDet + 1, 1), wa.Cells(ligneDet + 1, 11)).Font.Bold = True
    wa.Range(wa.Cells(ligneDet + 2, 1), wa.Cells(ligneDet + 1 + nd, 11)).Value = TronquerDet(det, nd)
    For r = ligneDet + 2 To ligneDet + 1 + nd
        If Txt(wa.Cells(r, 11).Value) = "NC" Then wa.Range(wa.Cells(r, 1), wa.Cells(r, 11)).Interior.Color = RGB(254, 226, 226)
    Next r
    EcranFige False
    wa.Activate
End Sub

Private Function TronquerDet(ByRef t() As Variant, ByVal n As Long) As Variant
    Dim r() As Variant, i As Long, k As Long
    ReDim r(1 To n, 1 To UBound(t, 2))
    For i = 1 To n
        For k = 1 To UBound(t, 2): r(i, k) = t(i, k): Next k
    Next i
    TronquerDet = r
End Function
