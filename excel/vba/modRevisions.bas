Attribute VB_Name = "modRevisions"
' =====================================================================
'  Révisions : 0, 1, 2, ... par famille (gamme GAMSERIE) et par article (TS).
'  Une révision applicable est figée (photo dans l'onglet Historique).
'  Toute modification se fait dans une révision en brouillon, validée ensuite.
' =====================================================================
Option Explicit

Public Const RV_GAM As String = "Gamme"
Public Const RV_TS As String = "TS"

' Révision la plus récente d'une famille / d'un article : numéro de ligne (0 si aucune)
Public Function LigneRevision(ByVal typ As String, ByVal cle As String) As Long
    Dim ws As Worksheet, r As Long, fin As Long, best As Long, rv As Long
    Set ws = Fe(F_REV)
    fin = DerniereLigne(ws, V_TYPE)
    best = 0: rv = -1
    For r = L_DEB To fin
        If Txt(ws.Cells(r, V_TYPE).Value) = typ And Txt(ws.Cells(r, V_CLE).Value) = cle Then
            If Val(ws.Cells(r, V_REV).Value) > rv Then rv = Val(ws.Cells(r, V_REV).Value): best = r
        End If
    Next r
    LigneRevision = best
End Function

' Libellé de la révision en vigueur, par exemple « rév. 2 » ou « rév. 3 (brouillon) »
Public Function LibelleRevision(ByVal typ As String, ByVal cle As String) As String
    Dim r As Long, ws As Worksheet
    r = LigneRevision(typ, cle)
    If r = 0 Then LibelleRevision = "sans révision": Exit Function
    Set ws = Fe(F_REV)
    LibelleRevision = "rév. " & ws.Cells(r, V_REV).Value
    If Txt(ws.Cells(r, V_STATUT).Value) = "Brouillon" Then LibelleRevision = LibelleRevision & " (brouillon)"
End Function

' Toute famille et tout article sans révision démarre en révision 0 applicable
Public Function InitialiserRevisions() As Long
    Dim f As Variant, n As Long, ws As Worksheet, r As Long, fin As Long, a As String, vus As New Collection
    For Each f In ListeFamilles()
        If LigneRevision(RV_GAM, CStr(f)) = 0 Then AjouterRevision RV_GAM, CStr(f), 0, "Applicable", "Reprise initiale", "Version initiale": n = n + 1
    Next f
    Set ws = Fe(F_TS)
    fin = DerniereLigne(ws, T_ART)
    For r = L_DEB To fin
        a = Txt(ws.Cells(r, T_ART).Value)
        If a <> "" Then
            If Not CollContient(vus, a) Then
                vus.Add a, a
                If LigneRevision(RV_TS, a) = 0 Then AjouterRevision RV_TS, a, 0, "Applicable", "Reprise initiale", "Version initiale": n = n + 1
            End If
        End If
    Next r
    InitialiserRevisions = n
End Function

Public Sub AjouterRevision(ByVal typ As String, ByVal cle As String, ByVal rev As Long, ByVal statut As String, ByVal par As String, ByVal motif As String)
    Dim ws As Worksheet, r As Long
    Set ws = Fe(F_REV)
    r = DerniereLigne(ws, V_TYPE) + 1
    ws.Cells(r, V_TYPE).Value = typ
    ws.Cells(r, V_CLE).Value = cle
    ws.Cells(r, V_REV).Value = rev
    ws.Cells(r, V_STATUT).Value = statut
    ws.Cells(r, V_OUVERT).Value = Now
    If statut = "Applicable" Then
        ws.Cells(r, V_DATE).Value = Date
        ws.Cells(r, V_PAR).Value = par
        ws.Cells(r, V_MOTIF).Value = motif
        Figer typ, cle, rev
        Journaliser typ, cle, rev, "Révision validée", motif
    Else
        ws.Cells(r, V_MOTIF).Value = motif
        Journaliser typ, cle, rev, "Brouillon ouvert", motif
    End If
End Sub

' Photo des données d'une famille / d'un article dans l'onglet Historique
Public Sub Figer(ByVal typ As String, ByVal cle As String, ByVal rev As Long)
    Dim wh As Worksheet, ws As Worksheet, r As Long, fin As Long, d As Long, lg As Variant
    Set wh = Fe(F_HIST)
    ' remplace une photo existante de la même révision
    fin = DerniereLigne(wh, H_TYPE)
    For r = fin To L_DEB Step -1
        If Txt(wh.Cells(r, H_TYPE).Value) = typ And Txt(wh.Cells(r, H_CLE).Value) = cle And Val(wh.Cells(r, H_REV).Value) = rev Then wh.Rows(r).Delete
    Next r
    d = DerniereLigne(wh, H_TYPE) + 1
    If typ = RV_GAM Then
        Set ws = Fe(F_GAM)
        For Each lg In LignesGamme(cle)
            wh.Cells(d, H_TYPE).Value = typ: wh.Cells(d, H_CLE).Value = cle: wh.Cells(d, H_REV).Value = rev
            wh.Cells(d, H_ID).Value = ws.Cells(lg, G_ID).Value
            wh.Cells(d, 5).Value = ws.Cells(lg, G_CODE).Value
            wh.Cells(d, 6).Value = ws.Cells(lg, G_CRIT).Value
            wh.Cells(d, 7).Value = ws.Cells(lg, G_SPEC).Value
            wh.Cells(d, 8).Value = ws.Cells(lg, G_MOY).Value
            wh.Cells(d, 9).Value = ws.Cells(lg, G_METH).Value
            d = d + 1
        Next lg
    Else
        Set ws = Fe(F_TS)
        fin = DerniereLigne(ws, T_ART)
        For r = L_DEB To fin
            If Txt(ws.Cells(r, T_ART).Value) = cle Then
                wh.Cells(d, H_TYPE).Value = typ: wh.Cells(d, H_CLE).Value = cle: wh.Cells(d, H_REV).Value = rev
                wh.Cells(d, H_ID).Value = ws.Cells(r, T_ID).Value
                wh.Cells(d, 10).Value = ws.Cells(r, T_TYPE).Value
                wh.Cells(d, 11).Value = ws.Cells(r, T_NOM).Value
                wh.Cells(d, 12).Value = ws.Cells(r, T_TM).Value
                wh.Cells(d, 13).Value = ws.Cells(r, T_TP).Value
                d = d + 1
            End If
        Next r
    End If
End Sub

' Signature des données actuelles (pour détecter une modification sans révision)
Public Function SignatureActuelle(ByVal typ As String, ByVal cle As String) As String
    Dim ws As Worksheet, s As String, lg As Variant, r As Long, fin As Long
    If typ = RV_GAM Then
        Set ws = Fe(F_GAM)
        For Each lg In LignesGamme(cle)
            s = s & Txt(ws.Cells(lg, G_ID).Value) & "~" & Txt(ws.Cells(lg, G_CODE).Value) & "~" & Txt(ws.Cells(lg, G_CRIT).Value) & "~" & Txt(ws.Cells(lg, G_SPEC).Value) & "~" & Txt(ws.Cells(lg, G_MOY).Value) & "~" & Txt(ws.Cells(lg, G_METH).Value) & "#"
        Next lg
    Else
        Set ws = Fe(F_TS)
        fin = DerniereLigne(ws, T_ART)
        For r = L_DEB To fin
            If Txt(ws.Cells(r, T_ART).Value) = cle Then s = s & Txt(ws.Cells(r, T_ID).Value) & "~" & Txt(ws.Cells(r, T_TYPE).Value) & "~" & Txt(ws.Cells(r, T_NOM).Value) & "~" & Txt(ws.Cells(r, T_TM).Value) & "~" & Txt(ws.Cells(r, T_TP).Value) & "#"
        Next r
    End If
    SignatureActuelle = s
End Function

Public Function SignaturePhoto(ByVal typ As String, ByVal cle As String, ByVal rev As Long) As String
    Dim wh As Worksheet, r As Long, fin As Long, s As String
    Set wh = Fe(F_HIST)
    fin = DerniereLigne(wh, H_TYPE)
    For r = L_DEB To fin
        If Txt(wh.Cells(r, H_TYPE).Value) = typ And Txt(wh.Cells(r, H_CLE).Value) = cle And Val(wh.Cells(r, H_REV).Value) = rev Then
            If typ = RV_GAM Then
                s = s & Txt(wh.Cells(r, H_ID).Value) & "~" & Txt(wh.Cells(r, 5).Value) & "~" & Txt(wh.Cells(r, 6).Value) & "~" & Txt(wh.Cells(r, 7).Value) & "~" & Txt(wh.Cells(r, 8).Value) & "~" & Txt(wh.Cells(r, 9).Value) & "#"
            Else
                s = s & Txt(wh.Cells(r, H_ID).Value) & "~" & Txt(wh.Cells(r, 10).Value) & "~" & Txt(wh.Cells(r, 11).Value) & "~" & Txt(wh.Cells(r, 12).Value) & "~" & Txt(wh.Cells(r, 13).Value) & "#"
            End If
        End If
    Next r
    SignaturePhoto = s
End Function

' Données modifiées alors que la révision en vigueur est applicable (donc figée)
Public Function ModifieSansRevision(ByVal typ As String, ByVal cle As String) As Boolean
    Dim r As Long, ws As Worksheet
    r = LigneRevision(typ, cle)
    If r = 0 Then Exit Function
    Set ws = Fe(F_REV)
    If Txt(ws.Cells(r, V_STATUT).Value) <> "Applicable" Then Exit Function
    ModifieSansRevision = (SignatureActuelle(typ, cle) <> SignaturePhoto(typ, cle, Val(ws.Cells(r, V_REV).Value)))
End Function

' Famille (feuille GAMSERIE) ou article (feuille TS) concerné par la cellule active
Private Function CibleRevision(ByRef typ As String, ByRef cle As String) As Boolean
    Dim ws As Worksheet, r As Long, defaut As String
    Set ws = ActiveSheet
    If ws.Name = F_GAM Then
        typ = RV_GAM
        r = ActiveCell.Row
        If r >= L_DEB Then defaut = Txt(ws.Cells(r, G_FAM).Value)
        cle = Demander("Famille de contrôle concernée :", , defaut)
    ElseIf ws.Name = F_TS Then
        typ = RV_TS
        r = ActiveCell.Row
        If r >= L_DEB Then defaut = Txt(ws.Cells(r, T_ART).Value)
        cle = Demander("Article concerné :", , defaut)
    Else
        Alerter "Placez-vous sur l'onglet GAMSERIE (révision d'une gamme) ou TS (révision des valeurs d'un article)."
        Exit Function
    End If
    cle = Trim(cle)
    CibleRevision = (cle <> "")
End Function

Public Sub OuvrirRevision()
    Dim typ As String, cle As String, r As Long, ws As Worksheet, rv As Long
    If Not CibleRevision(typ, cle) Then Exit Sub
    Set ws = Fe(F_REV)
    r = LigneRevision(typ, cle)
    If r > 0 Then
        If Txt(ws.Cells(r, V_STATUT).Value) = "Brouillon" Then
            Informer "La révision " & ws.Cells(r, V_REV).Value & " de « " & cle & " » est déjà ouverte en brouillon : vous pouvez modifier les données."
            Exit Sub
        End If
        rv = Val(ws.Cells(r, V_REV).Value) + 1
    Else
        rv = 0
    End If
    If Not Confirmer("Ouvrir la révision " & rv & " de « " & cle & " » en brouillon ?" & vbLf & vbLf & "Les modifications ne compteront qu'une fois la révision validée.") Then Exit Sub
    AjouterRevision typ, cle, rv, "Brouillon", "", ""
    Informer "Révision " & rv & " de « " & cle & " » ouverte en brouillon. Modifiez les données, puis cliquez sur « Valider la révision »."
End Sub

Public Sub ValiderRevision()
    Dim typ As String, cle As String, r As Long, ws As Worksheet, par As String, motif As String, d As String, rv As Long, i As Long, dt As Date
    If Not CibleRevision(typ, cle) Then Exit Sub
    Set ws = Fe(F_REV)
    r = LigneRevision(typ, cle)
    If r = 0 Then Alerter "Aucune révision pour « " & cle & " ».": Exit Sub
    If Txt(ws.Cells(r, V_STATUT).Value) <> "Brouillon" Then Alerter "Aucun brouillon à valider pour « " & cle & " » : ouvrez d'abord une révision.": Exit Sub
    rv = Val(ws.Cells(r, V_REV).Value)
    par = Trim(Demander("Révision " & rv & " de « " & cle & " »" & vbLf & "Validée par :"))
    If par = "" Then Alerter "Le nom du valideur est obligatoire.": Exit Sub
    motif = Trim(Demander("Motif de la révision :"))
    If motif = "" Then Alerter "Le motif est obligatoire.": Exit Sub
    d = Trim(Demander("Date d'effet (jj/mm/aaaa) :", , Format(Date, "dd/mm/yyyy")))
    If IsDate(d) Then dt = CDate(d) Else dt = Date
    ws.Cells(r, V_STATUT).Value = "Applicable"
    ws.Cells(r, V_DATE).Value = dt
    ws.Cells(r, V_PAR).Value = par
    ws.Cells(r, V_MOTIF).Value = motif
    For i = L_DEB To DerniereLigne(ws, V_TYPE)
        If i <> r And Txt(ws.Cells(i, V_TYPE).Value) = typ And Txt(ws.Cells(i, V_CLE).Value) = cle And Txt(ws.Cells(i, V_STATUT).Value) = "Applicable" Then ws.Cells(i, V_STATUT).Value = "Archivée"
    Next i
    Figer typ, cle, rv
    Journaliser typ, cle, rv, "Révision validée", motif
    If rv > 0 Then Journaliser typ, cle, rv - 1, "Révision archivée", "remplacée par la rév. " & rv
    SauverClasseur
    Informer "Révision " & rv & " de « " & cle & " » validée : elle s'applique aux nouveaux rapports."
End Sub

Public Sub AbandonnerRevision()
    Dim typ As String, cle As String, r As Long, ws As Worksheet, rv As Long
    If Not CibleRevision(typ, cle) Then Exit Sub
    Set ws = Fe(F_REV)
    r = LigneRevision(typ, cle)
    If r = 0 Then Exit Sub
    If Txt(ws.Cells(r, V_STATUT).Value) <> "Brouillon" Then Alerter "Aucun brouillon à abandonner pour « " & cle & " ».": Exit Sub
    rv = Val(ws.Cells(r, V_REV).Value)
    If rv > 0 Then
        If Not Confirmer("Abandonner le brouillon de la révision " & rv & " ?" & vbLf & "Les données de « " & cle & " » reviendront à la révision " & (rv - 1) & ".") Then Exit Sub
        Restaurer typ, cle, rv - 1
    Else
        If Not Confirmer("Abandonner le brouillon de la révision 0 ? Les données saisies sont conservées.") Then Exit Sub
    End If
    ws.Rows(r).Delete
    Journaliser typ, cle, rv, "Brouillon abandonné", ""
    Informer "Brouillon abandonné."
End Sub

' Remet les données d'une famille / d'un article dans l'état d'une révision figée
Private Sub Restaurer(ByVal typ As String, ByVal cle As String, ByVal rev As Long)
    Dim wh As Worksheet, ws As Worksheet, r As Long, fin As Long, d As Long
    Set wh = Fe(F_HIST)
    If typ = RV_GAM Then
        Set ws = Fe(F_GAM)
        fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
        For r = fin To L_DEB Step -1
            If Txt(ws.Cells(r, G_FAM).Value) = cle Then ws.Rows(r).Delete
        Next r
        d = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC) + 1
        fin = DerniereLigne(wh, H_TYPE)
        For r = L_DEB To fin
            If Txt(wh.Cells(r, H_TYPE).Value) = typ And Txt(wh.Cells(r, H_CLE).Value) = cle And Val(wh.Cells(r, H_REV).Value) = rev Then
                ws.Cells(d, G_ID).Value = wh.Cells(r, H_ID).Value
                ws.Cells(d, G_FAM).Value = cle
                ws.Cells(d, G_CODE).Value = wh.Cells(r, 5).Value
                ws.Cells(d, G_CRIT).Value = wh.Cells(r, 6).Value
                ws.Cells(d, G_SPEC).Value = wh.Cells(r, 7).Value
                ws.Cells(d, G_MOY).Value = wh.Cells(r, 8).Value
                ws.Cells(d, G_METH).Value = wh.Cells(r, 9).Value
                d = d + 1
            End If
        Next r
    Else
        Set ws = Fe(F_TS)
        fin = DerniereLigne(ws, T_ART)
        For r = fin To L_DEB Step -1
            If Txt(ws.Cells(r, T_ART).Value) = cle Then
                ws.Range(ws.Cells(r, T_ID), ws.Cells(r, T_ART)).ClearContents
                ws.Range(ws.Cells(r, T_TYPE), ws.Cells(r, T_TP)).ClearContents
            End If
        Next r
        CompacterTS
        d = DerniereLigne(ws, T_ART) + 1
        fin = DerniereLigne(wh, H_TYPE)
        For r = L_DEB To fin
            If Txt(wh.Cells(r, H_TYPE).Value) = typ And Txt(wh.Cells(r, H_CLE).Value) = cle And Val(wh.Cells(r, H_REV).Value) = rev Then
                ws.Cells(d, T_ID).Value = wh.Cells(r, H_ID).Value
                ws.Cells(d, T_ART).Value = cle
                ws.Cells(d, T_TYPE).Value = wh.Cells(r, 10).Value
                ws.Cells(d, T_NOM).Value = wh.Cells(r, 11).Value
                ws.Cells(d, T_TM).Value = wh.Cells(r, 12).Value
                ws.Cells(d, T_TP).Value = wh.Cells(r, 13).Value
                d = d + 1
            End If
        Next r
    End If
End Sub

' TS : remonte les lignes pour supprimer les trous (les colonnes Famille et Code sont des formules)
Public Sub CompacterTS()
    Dim ws As Worksheet, r As Long, fin As Long, d As Long, c As Long
    Set ws = Fe(F_TS)
    fin = DerniereLigne(ws, T_ART)
    d = L_DEB
    For r = L_DEB To fin
        If Txt(ws.Cells(r, T_ART).Value) <> "" Or Txt(ws.Cells(r, T_ID).Value) <> "" Then
            If r <> d Then
                ws.Cells(d, T_ID).Value = ws.Cells(r, T_ID).Value
                ws.Cells(d, T_ART).Value = ws.Cells(r, T_ART).Value
                For c = T_TYPE To T_TP
                    ws.Cells(d, c).Value = ws.Cells(r, c).Value
                Next c
                ws.Range(ws.Cells(r, T_ID), ws.Cells(r, T_ART)).ClearContents
                ws.Range(ws.Cells(r, T_TYPE), ws.Cells(r, T_TP)).ClearContents
            End If
            d = d + 1
        End If
    Next r
End Sub
