Attribute VB_Name = "modCommun"
' =====================================================================
'  Enregistrements Controle d'entree - outils communs
'  Noms des feuilles, positions, boites de dialogue, conversions.
' =====================================================================
Option Explicit

' ----- feuilles -----
Public Const F_MENU As String = "Menu"
Public Const F_RAPPORT As String = "Rapport"
Public Const F_GAM As String = "GAMSERIE"
Public Const F_TS As String = "TS"
Public Const F_MAT As String = "Matrice"
Public Const F_DOCS As String = "Documents"
Public Const F_REV As String = "Révisions"
Public Const F_HIST As String = "Historique"
Public Const F_JOUR As String = "Journal"
Public Const F_RAPS As String = "Rapports"
Public Const F_MES As String = "Mesures"
Public Const F_ANA As String = "Analyses"
Public Const F_COH As String = "Cohérence"
Public Const F_PAR As String = "Paramètres"

' ----- première ligne de données des feuilles-tableaux (titre ligne 1, boutons ligne 2, en-têtes ligne 3) -----
Public Const L_DEB As Long = 4

' ----- GAMSERIE -----
Public Const G_ID As Long = 1, G_FAM As Long = 2, G_CODE As Long = 3, G_CRIT As Long = 4
Public Const G_SPEC As Long = 5, G_MOY As Long = 6, G_METH As Long = 7
' ----- TS -----
Public Const T_ID As Long = 1, T_ART As Long = 2, T_FAM As Long = 3, T_CODE As Long = 4
Public Const T_TYPE As Long = 5, T_NOM As Long = 6, T_TM As Long = 7, T_TP As Long = 8
' ----- Matrice -----
Public Const M_ART As Long = 1, M_FOUR As Long = 2, M_FAM As Long = 3, M_REF1 As Long = 6, M_REFN As Long = 12
' ----- Documents -----
Public Const D_REF As Long = 1, D_CHEMIN As Long = 2, D_DATE As Long = 3
' ----- Révisions -----
Public Const V_TYPE As Long = 1, V_CLE As Long = 2, V_REV As Long = 3, V_STATUT As Long = 4
Public Const V_DATE As Long = 5, V_PAR As Long = 6, V_MOTIF As Long = 7, V_OUVERT As Long = 8
' ----- Historique -----
Public Const H_TYPE As Long = 1, H_CLE As Long = 2, H_REV As Long = 3, H_ID As Long = 4
' gamme : 5 code, 6 criticité, 7 spécification, 8 moyen, 9 méthode
' TS    : 10 type spéc., 11 nominal, 12 tol -, 13 tol +
' ----- Rapports -----
Public Const RP_NB As Long = 23
' 1 ID, 2 Date, 3 Article, 4 Lot, 5 Commande, 6 Fournisseur, 7 Contrôleur, 8 Qté commandée, 9 Qté réceptionnée,
' 10 Qté échantillon, 11 Durée, 12 Origine, 13 Décision, 14 Type NC, 15 Sous-type NC, 16 Commentaires,
' 17 Famille, 18 Rév. gamme, 19 Rév. TS, 20 Conformes, 21 Non conformes, 22 Créé le, 23 Modifié le
' ----- Mesures -----
Public Const MS_NB As Long = 21
' 1 ID rapport, 2 ID point, 3 Code, 4 Spécification, 5 Criticité, 6 Moyen, 7 Méthode, 8 Type spéc., 9 Mini,
' 10 Nominal, 11 Maxi, 12 Tol -, 13 Tol +, 14 N° moyen, 15 Pièce (0 = générale), 16 Valeur, 17 Résultat,
' 18 Article, 19 Lot, 20 Fournisseur, 21 Date

' ----- feuille Rapport -----
Public Const R_NCP As Long = 15     ' non-conformités par pièce
Public Const R_STP As Long = 16     ' statut de chaque pièce
Public Const R_ENT As Long = 17     ' en-têtes du tableau
Public Const R_L0 As Long = 18      ' première ligne du tableau
Public Const R_LMAX As Long = 200   ' nombre maximal de lignes
Public Const R_PMAX As Long = 100   ' nombre maximal de pièces
Public Const C_ID As Long = 1, C_CODE As Long = 2, C_SPEC As Long = 3, C_CRIT As Long = 4, C_MOY As Long = 5
Public Const C_METH As Long = 6, C_TYPE As Long = 7, C_MINI As Long = 8, C_NOM As Long = 9, C_MAXI As Long = 10
Public Const C_NMOY As Long = 11, C_RES As Long = 12, C_P1 As Long = 13

Public Const MDP_FIGE As String = "ec-fige"
Public Const T_GEN As String = "Générale"
Public Const T_IND As String = "Individuelle"
Public Const T_NA As String = "N/A"

' ----- mode test (vérifications automatiques sous LibreOffice) -----
Public ModeTest As Boolean
Public ReponsesTest As String     ' réponses préparées, séparées par « | »
Public MessagesTest As String

Public Sub TestDemarrer(Optional ByVal reponses As String = "")
    ModeTest = True
    MessagesTest = ""
    ReponsesTest = reponses
End Sub

Public Function TestMessages() As String
    TestMessages = MessagesTest
End Function

Private Function ReponseTest(ByVal defaut As String) As String
    Dim p As Long
    If ReponsesTest = "" Then
        ReponseTest = defaut
        Exit Function
    End If
    p = InStr(ReponsesTest, "|")
    If p = 0 Then
        ReponseTest = ReponsesTest
        ReponsesTest = ""
    Else
        ReponseTest = Left(ReponsesTest, p - 1)
        ReponsesTest = Mid(ReponsesTest, p + 1)
    End If
End Function

' ----- boîtes de dialogue -----
Public Function Demander(ByVal invite As String, Optional ByVal titre As String = "Enregistrements Contrôle d'entrée", Optional ByVal defaut As String = "") As String
    If ModeTest Then
        Demander = ReponseTest(defaut)
    Else
        Demander = InputBox(invite, titre, defaut)
    End If
End Function

Public Function Confirmer(ByVal message As String, Optional ByVal titre As String = "Enregistrements Contrôle d'entrée") As Boolean
    If ModeTest Then
        Confirmer = (LCase(ReponseTest("oui")) <> "non")
    Else
        Confirmer = (MsgBox(message, vbYesNo + vbQuestion, titre) = vbYes)
    End If
End Function

Public Sub Informer(ByVal message As String, Optional ByVal titre As String = "Enregistrements Contrôle d'entrée")
    If ModeTest Then
        MessagesTest = MessagesTest & "[i] " & message & vbLf
    Else
        MsgBox message, vbInformation, titre
    End If
End Sub

Public Sub Alerter(ByVal message As String, Optional ByVal titre As String = "Enregistrements Contrôle d'entrée")
    If ModeTest Then
        MessagesTest = MessagesTest & "[!] " & message & vbLf
    Else
        MsgBox message, vbExclamation, titre
    End If
End Sub

' ----- feuilles -----
Public Function Fe(ByVal nom As String) As Worksheet
    Set Fe = ThisWorkbook.Worksheets(nom)
End Function

' Dernière ligne non vide d'une colonne (au moins L_DEB - 1)
Public Function DerniereLigne(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(-4162).Row   ' -4162 = xlUp
    If r < L_DEB - 1 Then r = L_DEB - 1
    DerniereLigne = r
End Function

' Dernière ligne de données d'un tableau, d'après plusieurs colonnes
Public Function FinTableau(ByVal ws As Worksheet, ParamArray cols() As Variant) As Long
    Dim i As Long, r As Long, m As Long
    m = L_DEB - 1
    For i = LBound(cols) To UBound(cols)
        r = DerniereLigne(ws, CLng(cols(i)))
        If r > m Then m = r
    Next i
    FinTableau = m
End Function

' ----- textes et nombres -----
Public Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Txt = "": Exit Function
    If IsEmpty(v) Or IsNull(v) Then Txt = "": Exit Function
    Txt = Trim(CStr(v))
End Function

' Convertit une valeur (nombre, ou texte avec virgule ou point) en nombre ; ok = False si impossible
Public Function VersNombre(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String, i As Long, c As String, nbSep As Long, nbChiffres As Long
    ok = False
    VersNombre = 0
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbLong Or VarType(v) = vbSingle Or VarType(v) = vbCurrency Or VarType(v) = vbDecimal Then
        VersNombre = CDbl(v): ok = True: Exit Function
    End If
    s = Replace(Replace(Trim(CStr(v)), " ", ""), Chr(160), "")
    s = Replace(s, ",", ".")
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then
            nbChiffres = nbChiffres + 1
        ElseIf c = "." Then
            nbSep = nbSep + 1
        ElseIf (c = "-" Or c = "+") And i = 1 Then
            ' signe
        Else
            Exit Function
        End If
    Next i
    If nbChiffres = 0 Or nbSep > 1 Then Exit Function
    VersNombre = Val(s)
    ok = True
End Function

Public Function EstNombre(ByVal v As Variant) As Boolean
    Dim ok As Boolean
    VersNombre v, ok
    EstNombre = ok
End Function

' Texte sans accents, en minuscules, sans espaces ni ponctuation (comparaison d'en-têtes)
Public Function Normaliser(ByVal s As String) As String
    Dim a As String, b As String, i As Long, r As String, c As String
    a = "àâäáãéèêëíìîïóòôöõúùûüçñÀÂÄÁÃÉÈÊËÍÌÎÏÓÒÔÖÕÚÙÛÜÇÑ"
    b = "aaaaaeeeeiiiiooooouuuucnaaaaaeeeeiiiiooooouuuucn"
    s = LCase(Trim(s))
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If InStr(a, c) > 0 Then c = Mid(b, InStr(a, c), 1)
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = "+" Or c = "-" Then r = r & c
    Next i
    Normaliser = r
End Function

Public Function MemeCode(ByVal a As String, ByVal b As String) As Boolean
    MemeCode = (UCase(Trim(a)) = UCase(Trim(b)))
End Function

' ----- collections (clé -> valeur) -----
Public Function CollContient(ByVal c As Collection, ByVal cle As String) As Boolean
    Dim v As Variant
    On Error GoTo non
    v = c(cle)
    CollContient = True
    Exit Function
non:
    CollContient = False
End Function

Public Function CollLire(ByVal c As Collection, ByVal cle As String, Optional ByVal defaut As Variant = "") As Variant
    If CollContient(c, cle) Then CollLire = c(cle) Else CollLire = defaut
End Function

Public Sub CollEcrire(ByVal c As Collection, ByVal cle As String, ByVal v As Variant)
    If CollContient(c, cle) Then c.Remove cle
    c.Add v, cle
End Sub

' ----- identifiants -----
' Identifiant unique de 6 caractères pour un point de contrôle
Public Function NouvelId(Optional ByVal pris As Collection) As String
    Const ALPHA As String = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    Dim s As String, i As Long, essai As Long
    Randomize
    If pris Is Nothing Then Set pris = IdsGamme()
    For essai = 1 To 1000
        s = ""
        For i = 1 To 6
            s = s & Mid(ALPHA, Int(Rnd() * 32) + 1, 1)
        Next i
        If Not CollContient(pris, s) Then
            pris.Add s, s
            NouvelId = s
            Exit Function
        End If
    Next essai
    NouvelId = s
End Function

Public Function IdsGamme() As Collection
    Dim ws As Worksheet, r As Long, fin As Long, c As New Collection, id As String
    Set ws = Fe(F_GAM)
    fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
    For r = L_DEB To fin
        id = Txt(ws.Cells(r, G_ID).Value)
        If id <> "" Then
            If Not CollContient(c, id) Then c.Add id, id
        End If
    Next r
    Set IdsGamme = c
End Function

Public Function NouvelIdRapport() As String
    Randomize
    Dim t As Date
    t = Now
    NouvelIdRapport = "R" & Right("0" & (Year(t) Mod 100), 2) & Right("0" & Month(t), 2) & Right("0" & Day(t), 2) & _
                      Right("0" & Hour(t), 2) & Right("0" & Minute(t), 2) & Right("0" & Second(t), 2) & CStr(Int(Rnd() * 900) + 100)
End Function

' ----- journal des révisions -----
Public Sub Journaliser(ByVal typ As String, ByVal cle As String, ByVal rev As Variant, ByVal action As String, Optional ByVal detail As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = Fe(F_JOUR)
    r = DerniereLigne(ws, 1) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = typ
    ws.Cells(r, 3).Value = cle
    ws.Cells(r, 4).Value = rev
    ws.Cells(r, 5).Value = action
    ws.Cells(r, 6).Value = detail
End Sub

' ----- recherches dans le référentiel -----
Public Function LigneMatrice(ByVal article As String) As Long
    Dim ws As Worksheet, r As Long, fin As Long
    Set ws = Fe(F_MAT)
    fin = DerniereLigne(ws, M_ART)
    For r = L_DEB To fin
        If Txt(ws.Cells(r, M_ART).Value) = article And article <> "" Then LigneMatrice = r: Exit Function
    Next r
    LigneMatrice = 0
End Function

Public Function FamilleDe(ByVal article As String) As String
    Dim r As Long
    r = LigneMatrice(article)
    If r > 0 Then FamilleDe = Txt(Fe(F_MAT).Cells(r, M_FAM).Value) Else FamilleDe = ""
End Function

' Lignes (numéros) de la gamme d'une famille, dans l'ordre
Public Function LignesGamme(ByVal famille As String) As Collection
    Dim ws As Worksheet, r As Long, fin As Long, c As New Collection
    Set ws = Fe(F_GAM)
    fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
    For r = L_DEB To fin
        If Txt(ws.Cells(r, G_FAM).Value) = famille And famille <> "" Then c.Add r
    Next r
    Set LignesGamme = c
End Function

Public Function ListeFamilles() As Collection
    Dim ws As Worksheet, r As Long, fin As Long, c As New Collection, f As String
    Set ws = Fe(F_GAM)
    fin = FinTableau(ws, G_ID, G_FAM, G_CODE, G_SPEC)
    For r = L_DEB To fin
        f = Txt(ws.Cells(r, G_FAM).Value)
        If f <> "" Then
            If Not CollContient(c, f) Then c.Add f, f
        End If
    Next r
    Set ListeFamilles = c
End Function

' Index des lignes TS : clé "article|id" -> numéro de ligne
Public Function IndexTS() As Collection
    Dim ws As Worksheet, r As Long, fin As Long, c As New Collection, k As String
    Set ws = Fe(F_TS)
    fin = DerniereLigne(ws, T_ART)
    For r = L_DEB To fin
        k = Txt(ws.Cells(r, T_ART).Value) & "|" & Txt(ws.Cells(r, T_ID).Value)
        If Txt(ws.Cells(r, T_ART).Value) <> "" Then
            If Not CollContient(c, k) Then c.Add r, k
        End If
    Next r
    Set IndexTS = c
End Function

Public Function ValeurRenseignee(ByVal v As Variant) As Boolean
    ValeurRenseignee = (Txt(v) <> "")
End Function

Public Sub EcranFige(ByVal fige As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = Not fige
    If fige Then Application.Calculation = -4135 Else Application.Calculation = -4105   ' manuel / automatique
End Sub

Public Sub SauverClasseur()
    If ModeTest Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number <> 0 Then Alerter "Le classeur n'a pas pu être enregistré automatiquement : enregistrez-le (Ctrl + S)."
End Sub
