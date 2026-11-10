Attribute VB_Name = "modRapport"
' =====================================================================
'  Rapport de contrôle : préparation des lignes (gamme + TS), saisie,
'  enregistrement dans les onglets Rapports et Mesures, réouverture.
' =====================================================================
Option Explicit

' ----- cellules de l'en-tête (feuille Rapport) -----
Public Const A_ART As String = "C4", A_LOT As String = "F4", A_CDE As String = "I4", A_FOUR As String = "L4", A_CTRL As String = "O4"
Public Const A_DATE As String = "C6", A_QCDE As String = "F6", A_QREC As String = "I6", A_QECH As String = "L6", A_DUREE As String = "O6"
Public Const A_ORIG As String = "C8", A_DEC As String = "F8", A_TNC As String = "I8", A_STNC As String = "L8"
Public Const A_COM As String = "C10"
Public Const A_FAM As String = "C12", A_RGAM As String = "F12", A_RTS As String = "I12", A_ID As String = "L12", A_ETAT As String = "O12"
Public Const A_NBC As String = "C13", A_NBNC As String = "F13"
Private Const DOC_L0 As Long = 4      ' bloc « Documents applicables » : Q4:R10
Private Const DOC_COL As Long = 17

Private Function FR() As Worksheet
    Set FR = Fe(F_RAPPORT)
End Function

Public Function RapportFige() As Boolean
    RapportFige = (Txt(FR.Range(A_ETAT).Value) = "Figé")
End Function

Private Sub Deverrouiller()
    On Error Resume Next
    FR.Unprotect MDP_FIGE
End Sub

Private Sub Verrouiller()
    On Error Resume Next
    FR.Protect Password:=MDP_FIGE, DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

' Efface l'en-tête et le tableau
Private Sub ViderFormulaire()
    Dim ws As Worksheet, a As Variant
    Set ws = FR
    Deverrouiller
    For Each a In Array(A_ART, A_LOT, A_CDE, A_FOUR, A_CTRL, A_DATE, A_QCDE, A_QREC, A_QECH, A_DUREE, A_ORIG, A_DEC, A_TNC, A_STNC, A_COM, A_FAM, A_RGAM, A_RTS, A_ID, A_ETAT)
        ws.Range(a).ClearContents
    Next a
    ViderTableau
    ViderDocuments
End Sub

Private Sub ViderTableau()
    Dim ws As Worksheet
    Set ws = FR
    ws.Range(ws.Cells(R_L0, 1), ws.Cells(R_L0 + R_LMAX - 1, C_P1 + R_PMAX - 1)).ClearContents
End Sub

Private Sub ViderDocuments()
    Dim ws As Worksheet
    Set ws = FR
    On Error Resume Next
    ws.Range(ws.Cells(DOC_L0, DOC_COL), ws.Cells(DOC_L0 + 6, DOC_COL + 1)).Hyperlinks.Delete
    ws.Range(ws.Cells(DOC_L0, DOC_COL), ws.Cells(DOC_L0 + 6, DOC_COL + 1)).ClearContents
End Sub

' ----- Nouveau rapport -----
Public Sub NouveauRapport()
    Dim ws As Worksheet, q As Variant
    Set ws = FR
    If Txt(ws.Range(A_ART).Value) <> "" Or Txt(ws.Range(A_LOT).Value) <> "" Then
        If Txt(ws.Range(A_ETAT).Value) <> "Enregistré" And Txt(ws.Range(A_ETAT).Value) <> "Figé" Then
            If Not Confirmer("Le rapport affiché n'est pas enregistré. L'effacer pour commencer un nouveau rapport ?") Then Exit Sub
        End If
    End If
    EcranFige True
    ViderFormulaire
    ws.Range(A_DATE).Value = Date
    q = Fe(F_PAR).Range("B4").Value
    If Not EstNombre(q) Then q = 15
    ws.Range(A_QECH).Value = q
    ws.Range(A_DEC).Value = "Sans décision"
    ws.Range(A_ETAT).Value = "Nouveau"
    EcranFige False
    ws.Activate
    ws.Range(A_ART).Select
    If Not ModeTest Then Informer "Choisissez l'article, puis cliquez sur « Préparer le rapport » : les points de contrôle de sa gamme et les valeurs attendues s'affichent."
End Sub

' ----- lecture du tableau affiché (pour conserver les saisies) -----
' Relevés d'une ligne du rapport (tableau neuf à chaque appel)
Private Function ReleveLigne(ByVal ws As Worksheet, ByVal r As Long) As Variant
    Dim t() As Variant, j As Long
    ReDim t(1 To R_PMAX)
    For j = 1 To R_PMAX
        t(j) = ws.Cells(r, C_P1 + j - 1).Value
    Next j
    ReleveLigne = t
End Function

' Renvoie une collection de lignes ; chaque ligne = Array(id, code, spec, crit, moyen, methode, type, mini, nom, maxi, nmoyen, resultatGen, valeurs())
Private Function LireTableau() As Collection
    Dim ws As Worksheet, c As New Collection, r As Long, n As Long
    Set ws = FR
    For r = R_L0 To R_L0 + R_LMAX - 1
        If Txt(ws.Cells(r, C_ID).Value) = "" And Txt(ws.Cells(r, C_CODE).Value) = "" And Txt(ws.Cells(r, C_SPEC).Value) = "" Then Exit For
        n = n + 1
        c.Add Array(Txt(ws.Cells(r, C_ID).Value), Txt(ws.Cells(r, C_CODE).Value), Txt(ws.Cells(r, C_SPEC).Value), Txt(ws.Cells(r, C_CRIT).Value), _
                    Txt(ws.Cells(r, C_MOY).Value), Txt(ws.Cells(r, C_METH).Value), Txt(ws.Cells(r, C_TYPE).Value), ws.Cells(r, C_MINI).Value, _
                    ws.Cells(r, C_NOM).Value, ws.Cells(r, C_MAXI).Value, Txt(ws.Cells(r, C_NMOY).Value), _
                    IIf(Txt(ws.Cells(r, C_TYPE).Value) = T_GEN, Txt(ws.Cells(r, C_RES).Value), ""), ReleveLigne(ws, r))
    Next r
    Set LireTableau = c
End Function

' Écrit les lignes dans le tableau
Private Sub EcrireTableau(ByVal lignes As Collection)
    Dim ws As Worksheet, r As Long, j As Long, lg As Variant, v As Variant, i As Long, a As String, b As String
    Set ws = FR
    ViderTableau
    r = R_L0
    For Each lg In lignes
        If r > R_L0 + R_LMAX - 1 Then Exit For
        ws.Cells(r, C_ID).Value = lg(0)
        ws.Cells(r, C_CODE).Value = lg(1)
        ws.Cells(r, C_SPEC).Value = lg(2)
        ws.Cells(r, C_CRIT).Value = lg(3)
        ws.Cells(r, C_MOY).Value = lg(4)
        ws.Cells(r, C_METH).Value = lg(5)
        ws.Cells(r, C_TYPE).Value = lg(6)
        ws.Cells(r, C_MINI).Value = lg(7)
        ws.Cells(r, C_NOM).Value = lg(8)
        ws.Cells(r, C_MAXI).Value = lg(9)
        ws.Cells(r, C_NMOY).Value = lg(10)
        If lg(6) = T_GEN Then
            ws.Cells(r, C_RES).Value = lg(11)
        ElseIf lg(6) = T_IND Then
            a = ws.Cells(r, C_P1).Address(False, False)
            b = ws.Cells(r, C_P1 + R_PMAX - 1).Address(False, False)
            ws.Cells(r, C_RES).Formula = "=SUMPRODUCT(--ISNUMBER(" & a & ":" & b & "),--((" & a & ":" & b & "<$H" & r & ")+(" & a & ":" & b & ">$J" & r & ")>0))+COUNTIF(" & a & ":" & b & ",""NOK"")"
            v = lg(12)
            If IsArray(v) Then
                For j = LBound(v) To UBound(v)
                    If j >= 1 And j <= R_PMAX Then
                        If Txt(v(j)) <> "" Then ws.Cells(r, C_P1 + j - 1).Value = v(j)
                    End If
                Next j
            End If
        End If
        r = r + 1
    Next lg
End Sub

' Lignes du rapport d'après la gamme de la famille et les valeurs TS de l'article
Private Function LignesReferentiel(ByVal article As String, ByVal famille As String) As Collection
    Dim c As New Collection, wg As Worksheet, wt As Worksheet, idx As Collection, lg As Variant, id As String, rt As Long
    Dim ty As String, nomV As Variant, okN As Boolean, nom As Double, tm As Double, tp As Double, okM As Boolean, okP As Boolean, typ As String
    Dim mini As Variant, maxi As Variant, nomAff As Variant, vide() As Variant
    Set wg = Fe(F_GAM): Set wt = Fe(F_TS)
    Set idx = IndexTS()
    ReDim vide(1 To R_PMAX)
    For Each lg In LignesGamme(famille)
        id = Txt(wg.Cells(lg, G_ID).Value)
        rt = CollLire(idx, article & "|" & id, 0)
        ty = "": nomV = "": tm = 0: tp = 0
        If rt > 0 Then
            ty = Txt(wt.Cells(rt, T_TYPE).Value)
            nomV = wt.Cells(rt, T_NOM).Value
            tm = VersNombre(wt.Cells(rt, T_TM).Value, okM)
            tp = VersNombre(wt.Cells(rt, T_TP).Value, okP)
        End If
        If ty <> "" Then                                   ' type de spécification vide : point non déroulé
            nom = VersNombre(nomV, okN)
            mini = "": maxi = "": nomAff = ""
            If Normaliser(ty) = "generale" Then
                typ = T_GEN
            ElseIf Not okN Then
                typ = T_NA: nomAff = Txt(nomV)
            Else
                typ = T_IND
                If nom = 1 And tm = 0 And tp = 0 Then
                    nomAff = "OK / NOK"                     ' point qualitatif
                Else
                    nomAff = nom: mini = nom + tm: maxi = nom + tp
                End If
            End If
            c.Add Array(id, Txt(wg.Cells(lg, G_CODE).Value), Txt(wg.Cells(lg, G_SPEC).Value), Txt(wg.Cells(lg, G_CRIT).Value), _
                        Txt(wg.Cells(lg, G_MOY).Value), Txt(wg.Cells(lg, G_METH).Value), typ, mini, nomAff, maxi, "", "", vide)
        End If
    Next lg
    Set LignesReferentiel = c
End Function

' ----- Préparer le rapport -----
Public Sub PreparerRapport()
    Dim ws As Worksheet, article As String, famille As String, rm As Long, lignes As Collection, anciens As Collection, lg As Variant, k As Collection
    Dim nouv As New Collection, a As Variant, i As Long
    Set ws = FR
    If RapportFige() Then Alerter "Ce rapport est figé (décision prise). Cliquez sur « Rouvrir (méthodiste) » pour le modifier.": Exit Sub
    article = Txt(ws.Range(A_ART).Value)
    If article = "" Then Alerter "Choisissez d'abord l'article (cellule Article).": ws.Range(A_ART).Select: Exit Sub
    CompleterIds False
    rm = LigneMatrice(article)
    famille = FamilleDe(article)
    If rm = 0 Then Alerter "L'article « " & article & " » est absent de la Matrice : ajoutez-le (bouton « Nouvel article » de l'onglet TS ou Matrice).": Exit Sub
    If famille = "" Or LignesGamme(famille).Count = 0 Then Alerter "L'article « " & article & " » n'a pas de famille de contrôle connue dans GAMSERIE (colonne Famille de la Matrice).": Exit Sub
    If Val(ws.Range(A_QECH).Value) < 1 Then ws.Range(A_QECH).Value = 1
    If Val(ws.Range(A_QECH).Value) > R_PMAX Then Alerter "La quantité échantillon est limitée à " & R_PMAX & " pièces.": ws.Range(A_QECH).Value = R_PMAX
    EcranFige True
    ' saisies déjà faites, conservées par identifiant de point
    Set anciens = New Collection
    For Each lg In LireTableau()
        If lg(0) <> "" Then CollEcrire anciens, lg(0), lg
    Next lg
    Set lignes = LignesReferentiel(article, famille)
    For Each lg In lignes
        If CollContient(anciens, CStr(lg(0))) Then
            a = anciens(CStr(lg(0)))
            lg(10) = a(10): lg(11) = a(11): lg(12) = a(12)
        End If
        nouv.Add lg
    Next lg
    EcrireTableau nouv
    If Txt(ws.Range(A_FOUR).Value) = "" Then ws.Range(A_FOUR).Value = Fe(F_MAT).Cells(rm, M_FOUR).Value
    If Txt(ws.Range(A_DATE).Value) = "" Then ws.Range(A_DATE).Value = Date
    If Txt(ws.Range(A_DEC).Value) = "" Then ws.Range(A_DEC).Value = "Sans décision"
    ws.Range(A_FAM).Value = famille
    ws.Range(A_RGAM).Value = LibelleRevision(RV_GAM, famille)
    ws.Range(A_RTS).Value = LibelleRevision(RV_TS, article)
    If Txt(ws.Range(A_ETAT).Value) = "" Then ws.Range(A_ETAT).Value = "Nouveau"
    AfficherDocuments article
    EcranFige False
    If nouv.Count = 0 Then Alerter "Aucun point de contrôle à relever : renseignez la colonne « Type spéc. » (Générale ou Individuelle) des valeurs TS de l'article."
    If Not ModeTest Then ws.Cells(R_L0, C_P1).Select
End Sub

' Bloc « Documents applicables » : références de la Matrice, avec lien vers le PDF s'il est connu
Private Sub AfficherDocuments(ByVal article As String)
    Dim ws As Worksheet, wm As Worksheet, wd As Worksheet, rm As Long, c As Long, r As Long, ref As String, i As Long, chemin As String
    Set ws = FR: Set wm = Fe(F_MAT): Set wd = Fe(F_DOCS)
    ViderDocuments
    rm = LigneMatrice(article)
    If rm = 0 Then Exit Sub
    r = DOC_L0
    For c = M_REF1 To M_REFN
        ref = Txt(wm.Cells(rm, c).Value)
        If ref <> "" And UCase(ref) <> "N/A" And r <= DOC_L0 + 6 Then
            ws.Cells(r, DOC_COL).Value = Replace(Txt(wm.Cells(3, c).Value), "Réf. ", "")
            ws.Cells(r, DOC_COL + 1).Value = ref
            chemin = ""
            For i = L_DEB To DerniereLigne(wd, D_REF)
                If LCase(Txt(wd.Cells(i, D_REF).Value)) = LCase(ref) Then chemin = Txt(wd.Cells(i, D_CHEMIN).Value): Exit For
            Next i
            If chemin <> "" Then
                On Error Resume Next
                ws.Hyperlinks.Add Anchor:=ws.Cells(r, DOC_COL + 1), Address:=chemin, TextToDisplay:=ref
                On Error GoTo 0
            End If
            r = r + 1
        End If
    Next c
End Sub

' ----- Enregistrer le rapport -----
Public Sub EnregistrerRapport()
    Dim ws As Worksheet, wrp As Worksheet, wm As Worksheet, id As String, manque As String, r As Long, j As Long, ty As String
    Dim v As Variant, ok As Boolean, x As Double, erreurs As String, nbP As Long, ancienne As String, nouvelle As String
    Dim rr As Long, fin As Long, donnees As Variant, nb As Long, i As Long, k As Long, res As String, mini As Variant, maxi As Variant
    Dim sortie() As Variant, nOut As Long, garder As Long, nomV As Variant, okMi As Boolean, okMa As Boolean, mi As Double, ma As Double
    Set ws = FR
    If RapportFige() Then Alerter "Ce rapport est figé (décision prise). Cliquez sur « Rouvrir (méthodiste) » pour le modifier.": Exit Sub
    If Txt(ws.Cells(R_L0, C_ID).Value) = "" And Txt(ws.Cells(R_L0, C_CODE).Value) = "" Then Alerter "Préparez d'abord le rapport (bouton « Préparer le rapport »).": Exit Sub
    If Txt(ws.Range(A_ART).Value) = "" Then manque = manque & vbLf & "  - Article"
    If Txt(ws.Range(A_LOT).Value) = "" Then manque = manque & vbLf & "  - Lot"
    If Val(ws.Range(A_QECH).Value) < 1 Then manque = manque & vbLf & "  - Quantité échantillon"
    If manque <> "" Then Alerter "Le rapport ne peut pas être enregistré, il manque :" & manque: Exit Sub
    nbP = Val(ws.Range(A_QECH).Value)
    If nbP > R_PMAX Then nbP = R_PMAX

    ' 1. relevés : virgule ou point acceptés, OK / NOK en majuscules
    For r = R_L0 To R_L0 + R_LMAX - 1
        If Txt(ws.Cells(r, C_ID).Value) = "" And Txt(ws.Cells(r, C_CODE).Value) = "" Then Exit For
        ty = Txt(ws.Cells(r, C_TYPE).Value)
        If ty = T_GEN Then
            v = UCase(Txt(ws.Cells(r, C_RES).Value))
            If v <> "" And v <> "OK" And v <> "NOK" Then erreurs = erreurs & vbLf & "  - " & ws.Cells(r, C_CODE).Value & " : « " & ws.Cells(r, C_RES).Value & " » (OK ou NOK attendu)" Else ws.Cells(r, C_RES).Value = v
        ElseIf ty = T_IND Then
            For j = 1 To nbP
                v = ws.Cells(r, C_P1 + j - 1).Value
                If Txt(v) <> "" And Not IsNumeric(v) Then
                    If UCase(Txt(v)) = "OK" Or UCase(Txt(v)) = "NOK" Then
                        ws.Cells(r, C_P1 + j - 1).Value = UCase(Txt(v))
                    Else
                        x = VersNombre(v, ok)
                        If ok Then ws.Cells(r, C_P1 + j - 1).Value = x Else erreurs = erreurs & vbLf & "  - " & ws.Cells(r, C_CODE).Value & ", pièce " & j & " : « " & Txt(v) & " »"
                    End If
                ElseIf VarType(v) = vbString Then
                    x = VersNombre(v, ok)
                    If ok Then ws.Cells(r, C_P1 + j - 1).Value = x
                End If
            Next j
        End If
    Next r
    If erreurs <> "" Then Alerter "Valeurs non reconnues (saisissez un nombre, ou OK / NOK) :" & erreurs: Exit Sub

    ' 2. décision : sortie d'« Attente décision méthodes » réservée aux méthodistes
    id = Txt(ws.Range(A_ID).Value)
    Set wrp = Fe(F_RAPS)
    rr = 0
    If id <> "" Then
        For i = L_DEB To DerniereLigne(wrp, 1)
            If Txt(wrp.Cells(i, 1).Value) = id Then rr = i: Exit For
        Next i
    End If
    nouvelle = Txt(ws.Range(A_DEC).Value)
    If rr > 0 Then ancienne = Txt(wrp.Cells(rr, 13).Value)
    If ancienne = "Attente décision méthodes" And nouvelle <> ancienne Then
        If Not MotDePasseMethodiste("Sortie du statut « Attente décision méthodes » réservée aux méthodistes.") Then ws.Range(A_DEC).Value = ancienne: Exit Sub
    End If

    EcranFige True
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
    If id = "" Then id = NouvelIdRapport(): ws.Range(A_ID).Value = id
    If rr = 0 Then rr = DerniereLigne(wrp, 1) + 1: wrp.Cells(rr, 22).Value = Now
    wrp.Cells(rr, 1).Value = id
    wrp.Cells(rr, 2).Value = ws.Range(A_DATE).Value
    wrp.Cells(rr, 3).Value = ws.Range(A_ART).Value
    wrp.Cells(rr, 4).Value = ws.Range(A_LOT).Value
    wrp.Cells(rr, 5).Value = ws.Range(A_CDE).Value
    wrp.Cells(rr, 6).Value = ws.Range(A_FOUR).Value
    wrp.Cells(rr, 7).Value = ws.Range(A_CTRL).Value
    wrp.Cells(rr, 8).Value = ws.Range(A_QCDE).Value
    wrp.Cells(rr, 9).Value = ws.Range(A_QREC).Value
    wrp.Cells(rr, 10).Value = ws.Range(A_QECH).Value
    wrp.Cells(rr, 11).Value = ws.Range(A_DUREE).Value
    wrp.Cells(rr, 12).Value = ws.Range(A_ORIG).Value
    wrp.Cells(rr, 13).Value = nouvelle
    wrp.Cells(rr, 14).Value = ws.Range(A_TNC).Value
    wrp.Cells(rr, 15).Value = ws.Range(A_STNC).Value
    wrp.Cells(rr, 16).Value = ws.Range(A_COM).Value
    wrp.Cells(rr, 17).Value = ws.Range(A_FAM).Value
    wrp.Cells(rr, 18).Value = ws.Range(A_RGAM).Value
    wrp.Cells(rr, 19).Value = ws.Range(A_RTS).Value
    wrp.Cells(rr, 20).Value = ws.Range(A_NBC).Value
    wrp.Cells(rr, 21).Value = ws.Range(A_NBNC).Value
    wrp.Cells(rr, 23).Value = Now

    ' 3. mesures : on réécrit l'onglet sans les anciennes lignes de ce rapport, puis on ajoute les nouvelles
    Set wm = Fe(F_MES)
    fin = DerniereLigne(wm, 1)
    ReDim sortie(1 To (fin - L_DEB + 1) + R_LMAX * (R_PMAX + 1) + 1, 1 To MS_NB)
    nOut = 0
    If fin >= L_DEB Then
        donnees = wm.Range(wm.Cells(L_DEB, 1), wm.Cells(fin, MS_NB)).Value
        For i = 1 To UBound(donnees, 1)
            If Txt(donnees(i, 1)) <> id Then
                nOut = nOut + 1
                For k = 1 To MS_NB
                    sortie(nOut, k) = donnees(i, k)
                Next k
            End If
        Next i
    End If
    For r = R_L0 To R_L0 + R_LMAX - 1
        If Txt(ws.Cells(r, C_ID).Value) = "" And Txt(ws.Cells(r, C_CODE).Value) = "" Then Exit For
        ty = Txt(ws.Cells(r, C_TYPE).Value)
        mini = ws.Cells(r, C_MINI).Value: maxi = ws.Cells(r, C_MAXI).Value: nomV = ws.Cells(r, C_NOM).Value
        mi = VersNombre(mini, okMi): ma = VersNombre(maxi, okMa)
        If ty = T_GEN Then
            v = Txt(ws.Cells(r, C_RES).Value)
            res = IIf(v = "OK", "C", IIf(v = "NOK", "NC", ""))
            nOut = nOut + 1
            LigneMesure sortie, nOut, id, ws, r, 0, v, res
        Else
            ' ligne de structure (pièce 0) : le rapport garde toutes ses lignes, même sans mesure
            nOut = nOut + 1
            v = Empty: res = ""
            LigneMesure sortie, nOut, id, ws, r, 0, v, res
        End If
        If ty = T_IND Then
            For j = 1 To nbP
                v = ws.Cells(r, C_P1 + j - 1).Value
                If Txt(v) <> "" Then
                    If IsNumeric(v) And okMi And okMa Then
                        res = IIf(CDbl(v) < mi - 0.000000001 Or CDbl(v) > ma + 0.000000001, "NC", "C")
                    ElseIf UCase(Txt(v)) = "NOK" Then
                        res = "NC"
                    ElseIf UCase(Txt(v)) = "OK" Then
                        res = "C"
                    Else
                        res = ""
                    End If
                    nOut = nOut + 1
                    LigneMesure sortie, nOut, id, ws, r, j, v, res
                End If
            Next j
        End If
    Next r
    If fin >= L_DEB Then wm.Range(wm.Cells(L_DEB, 1), wm.Cells(fin, MS_NB)).ClearContents
    If nOut > 0 Then wm.Range(wm.Cells(L_DEB, 1), wm.Cells(L_DEB + nOut - 1, MS_NB)).Value = TronquerTableau(sortie, nOut)

    ' 4. état du rapport ; décision Accepté / Refusé : rapport figé
    If nouvelle = "Accepté" Or nouvelle = "Refusé" Then
        ws.Range(A_ETAT).Value = "Figé"
        Verrouiller
    Else
        ws.Range(A_ETAT).Value = "Enregistré"
    End If
    ActualiserMenu
    EcranFige False
    SauverClasseur
    Informer "Rapport enregistré" & IIf(ws.Range(A_ETAT).Value = "Figé", " et figé (décision « " & nouvelle & " »).", ".")
End Sub

Private Sub LigneMesure(ByRef s() As Variant, ByVal n As Long, ByVal id As String, ByVal ws As Worksheet, ByVal r As Long, ByVal piece As Long, ByVal v As Variant, ByVal res As String)
    Dim okN As Boolean, nom As Double, mini As Variant, maxi As Variant, okMi As Boolean, okMa As Boolean, mi As Double, ma As Double
    s(n, 1) = id
    s(n, 2) = ws.Cells(r, C_ID).Value
    s(n, 3) = ws.Cells(r, C_CODE).Value
    s(n, 4) = ws.Cells(r, C_SPEC).Value
    s(n, 5) = ws.Cells(r, C_CRIT).Value
    s(n, 6) = ws.Cells(r, C_MOY).Value
    s(n, 7) = ws.Cells(r, C_METH).Value
    s(n, 8) = ws.Cells(r, C_TYPE).Value
    mini = ws.Cells(r, C_MINI).Value: maxi = ws.Cells(r, C_MAXI).Value
    s(n, 9) = mini
    s(n, 10) = ws.Cells(r, C_NOM).Value
    s(n, 11) = maxi
    nom = VersNombre(ws.Cells(r, C_NOM).Value, okN)
    mi = VersNombre(mini, okMi): ma = VersNombre(maxi, okMa)
    If okN And okMi Then s(n, 12) = Round(mi - nom, 6) Else s(n, 12) = ""
    If okN And okMa Then s(n, 13) = Round(ma - nom, 6) Else s(n, 13) = ""
    s(n, 14) = ws.Cells(r, C_NMOY).Value
    s(n, 15) = piece
    s(n, 16) = v
    s(n, 17) = res
    s(n, 18) = ws.Range(A_ART).Value
    s(n, 19) = ws.Range(A_LOT).Value
    s(n, 20) = ws.Range(A_FOUR).Value
    s(n, 21) = ws.Range(A_DATE).Value
End Sub

Private Function TronquerTableau(ByRef s() As Variant, ByVal n As Long) As Variant
    Dim t() As Variant, i As Long, k As Long
    ReDim t(1 To n, 1 To UBound(s, 2))
    For i = 1 To n
        For k = 1 To UBound(s, 2)
            t(i, k) = s(i, k)
        Next k
    Next i
    TronquerTableau = t
End Function

' ----- Ouvrir un rapport enregistré -----
' Copie neuve des relevés d'une ligne (un tableau distinct par ligne)
Private Function ValeursLigne(ByRef vals() As Variant, ByVal q As Long) As Variant
    Dim t() As Variant, p As Long
    ReDim t(1 To R_PMAX)
    For p = 1 To R_PMAX
        t(p) = vals(q, p)
    Next p
    ValeursLigne = t
End Function

Public Sub ChargerRapport(ByVal id As String)
    Dim ws As Worksheet, wrp As Worksheet, wm As Worksheet, rr As Long, i As Long, fin As Long
    Dim cle As String, p As Long, dec As String
    Set ws = FR: Set wrp = Fe(F_RAPS): Set wm = Fe(F_MES)
    For i = L_DEB To DerniereLigne(wrp, 1)
        If Txt(wrp.Cells(i, 1).Value) = id Then rr = i: Exit For
    Next i
    If rr = 0 Then Alerter "Rapport introuvable.": Exit Sub
    EcranFige True
    ViderFormulaire
    ws.Range(A_ID).Value = id
    ws.Range(A_DATE).Value = wrp.Cells(rr, 2).Value
    ws.Range(A_ART).Value = wrp.Cells(rr, 3).Value
    ws.Range(A_LOT).Value = wrp.Cells(rr, 4).Value
    ws.Range(A_CDE).Value = wrp.Cells(rr, 5).Value
    ws.Range(A_FOUR).Value = wrp.Cells(rr, 6).Value
    ws.Range(A_CTRL).Value = wrp.Cells(rr, 7).Value
    ws.Range(A_QCDE).Value = wrp.Cells(rr, 8).Value
    ws.Range(A_QREC).Value = wrp.Cells(rr, 9).Value
    ws.Range(A_QECH).Value = wrp.Cells(rr, 10).Value
    ws.Range(A_DUREE).Value = wrp.Cells(rr, 11).Value
    ws.Range(A_ORIG).Value = wrp.Cells(rr, 12).Value
    dec = Txt(wrp.Cells(rr, 13).Value)
    ws.Range(A_DEC).Value = IIf(dec = "", "Sans décision", dec)
    ws.Range(A_TNC).Value = wrp.Cells(rr, 14).Value
    ws.Range(A_STNC).Value = wrp.Cells(rr, 15).Value
    ws.Range(A_COM).Value = wrp.Cells(rr, 16).Value
    ws.Range(A_FAM).Value = wrp.Cells(rr, 17).Value
    ws.Range(A_RGAM).Value = wrp.Cells(rr, 18).Value
    ws.Range(A_RTS).Value = wrp.Cells(rr, 19).Value
    ' lignes enregistrées (le rapport garde ses propres lignes, même si la gamme a changé depuis)
    Dim idx As New Collection, n As Long, meta() As Variant, vals() As Variant, resG() As String, q As Long, liste As New Collection
    ReDim meta(1 To R_LMAX): ReDim vals(1 To R_LMAX, 1 To R_PMAX): ReDim resG(1 To R_LMAX)
    fin = DerniereLigne(wm, 1)
    For i = L_DEB To fin
        If Txt(wm.Cells(i, 1).Value) = id Then
            cle = Txt(wm.Cells(i, 2).Value) & "#" & Txt(wm.Cells(i, 3).Value)
            If CollContient(idx, cle) Then
                q = idx(cle)
            ElseIf n < R_LMAX Then
                n = n + 1: q = n
                idx.Add q, cle
                meta(q) = Array(Txt(wm.Cells(i, 2).Value), Txt(wm.Cells(i, 3).Value), Txt(wm.Cells(i, 4).Value), Txt(wm.Cells(i, 5).Value), _
                                Txt(wm.Cells(i, 6).Value), Txt(wm.Cells(i, 7).Value), Txt(wm.Cells(i, 8).Value), wm.Cells(i, 9).Value, _
                                wm.Cells(i, 10).Value, wm.Cells(i, 11).Value, Txt(wm.Cells(i, 14).Value))
            Else
                q = 0
            End If
            If q > 0 Then
                p = Val(wm.Cells(i, 15).Value)
                If p = 0 Then
                    If Txt(wm.Cells(i, 16).Value) <> "" Then resG(q) = Txt(wm.Cells(i, 16).Value)
                ElseIf p >= 1 And p <= R_PMAX Then
                    vals(q, p) = wm.Cells(i, 16).Value
                End If
            End If
        End If
    Next i
    For q = 1 To n
        liste.Add Array(meta(q)(0), meta(q)(1), meta(q)(2), meta(q)(3), meta(q)(4), meta(q)(5), meta(q)(6), meta(q)(7), meta(q)(8), meta(q)(9), meta(q)(10), resG(q), ValeursLigne(vals, q))
    Next q
    EcrireTableau liste
    AfficherDocuments Txt(ws.Range(A_ART).Value)
    If dec = "Accepté" Or dec = "Refusé" Then
        ws.Range(A_ETAT).Value = "Figé"
        Verrouiller
    Else
        ws.Range(A_ETAT).Value = "Enregistré"
    End If
    EcranFige False
    ws.Activate
End Sub


' ----- Rapport figé : réouverture par un méthodiste -----
Public Function MotDePasseMethodiste(ByVal raison As String) As Boolean
    Dim mdp As String, attendu As String
    attendu = Txt(Fe(F_PAR).Range("B3").Value)
    mdp = Demander(raison & vbLf & vbLf & "Mot de passe :")
    If mdp = "" Then Exit Function
    If mdp <> attendu Then Alerter "Mot de passe incorrect.": Exit Function
    MotDePasseMethodiste = True
End Function

Public Sub RouvrirRapport()
    Dim ws As Worksheet
    Set ws = FR
    If Not RapportFige() Then Informer "Ce rapport n'est pas figé : il est modifiable.": Exit Sub
    If Not MotDePasseMethodiste("Réouverture d'un rapport figé (Accepté / Refusé), réservée aux méthodistes.") Then Exit Sub
    Deverrouiller
    ws.Range(A_ETAT).Value = "Rouvert"
    Informer "Rapport rouvert : vous pouvez le modifier, puis l'enregistrer."
End Sub

Public Sub ViderMesures()
    Dim ws As Worksheet
    Set ws = FR
    If RapportFige() Then Alerter "Ce rapport est figé.": Exit Sub
    If Not Confirmer("Effacer toutes les mesures saisies dans ce rapport ?") Then Exit Sub
    ws.Range(ws.Cells(R_L0, C_P1), ws.Cells(R_L0 + R_LMAX - 1, C_P1 + R_PMAX - 1)).ClearContents
    Dim r As Long
    For r = R_L0 To R_L0 + R_LMAX - 1
        If Txt(ws.Cells(r, C_TYPE).Value) = T_GEN Then ws.Cells(r, C_RES).ClearContents
    Next r
End Sub

' ----- Export PDF du rapport -----
Public Sub ExporterRapportPDF()
    Dim ws As Worksheet, nom As Variant, defaut As String
    Set ws = FR
    defaut = "Rapport_" & Txt(ws.Range(A_ART).Value) & "_" & Txt(ws.Range(A_LOT).Value) & ".pdf"
    defaut = Replace(Replace(Replace(defaut, "/", "-"), "\", "-"), ":", "-")
    If ModeTest Then Exit Sub
    nom = Application.GetSaveAsFilename(InitialFileName:=defaut, FileFilter:="PDF (*.pdf),*.pdf", Title:="Exporter le rapport en PDF")
    If VarType(nom) = vbBoolean Then Exit Sub
    On Error GoTo echec
    ws.ExportAsFixedFormat Type:=0, Filename:=CStr(nom), Quality:=0, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
echec:
    Alerter "Export PDF impossible : " & Err.Description
End Sub
