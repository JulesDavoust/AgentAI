Attribute VB_Name = "SimuMainframe"
Option Explicit



Private Const SHEET_MVS As String = "MVS"
Private Const SHEET_BASE As String = "MVS_Base"
Private Const NB_LIGNES As Long = 32
Private Const NB_COLS As Long = 80

' Coordonnees des champs (a garder identiques dans le module 2)
Private Const C_CHAMP As Long = 27
Private Const L_ISIN As Long = 6
Private Const LEN_ISIN As Long = 12
Private Const L_LIB As Long = 8
Private Const LEN_LIB As Long = 40
Private Const L_STAT As Long = 10
Private Const LEN_STAT As Long = 10
Private Const L_ENC As Long = 12
Private Const LEN_ENC As Long = 12
Private Const L_DATE As Long = 14
Private Const LEN_DATE As Long = 10
Private Const L_COM As Long = 16
Private Const LEN_COM As Long = 40
Private Const L_MSG As Long = 30
Private Const C_MSG As Long = 13
Private Const LEN_MSG As Long = 60

'--- PALETTE NATRAN ---
Private Function CFond() As Long:    CFond = RGB(8, 30, 60):      End Function
Private Function CTexte() As Long:   CTexte = RGB(235, 242, 250): End Function
Private Function CTitre() As Long:   CTitre = RGB(255, 255, 255): End Function
Private Function CLabel() As Long:   CLabel = RGB(0, 190, 175):   End Function
Private Function CSaisie() As Long:  CSaisie = RGB(0, 84, 159):   End Function
Private Function CSep() As Long:     CSep = RGB(0, 120, 140):     End Function
Private Function CMessage() As Long: CMessage = RGB(150, 210, 80): End Function

'==============================================================
' BOUTON 1 : OUVRIR LE MAINFRAME
' Cree la page dans le fichier si elle n'existe pas, sinon l'affiche.
'==============================================================
Public Sub OuvrirMainframe()
    Dim ws As Worksheet
    If Not FeuilleExiste(SHEET_MVS) Then
        CreerEcranMVS
    Else
        Set ws = ThisWorkbook.Worksheets(SHEET_MVS)
        ws.Visible = xlSheetVisible
        ws.Activate
        ws.Cells(L_ISIN, C_CHAMP).Select
    End If
End Sub

'==============================================================
' CREATION / REINITIALISATION DE L'ECRAN
'==============================================================
Public Sub CreerEcranMVS()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    If Not FeuilleExiste(SHEET_BASE) Then ReinitialiserBaseMVS
    Set ws = GetOrCreateSheet(SHEET_MVS)
    DessinerEcran ws
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    ws.Cells(L_ISIN, C_CHAMP).Select
    Application.ScreenUpdating = True
End Sub

Public Sub ReinitialiserBaseMVS()
    Dim ws As Worksheet, lignes As Variant, champs As Variant
    Dim i As Long, j As Long

    Set ws = GetOrCreateSheet(SHEET_BASE)
    ws.Cells.Clear
    ws.Cells.NumberFormat = "@"
    ws.Range("A1:F1").Value = Array("ISIN", "Libelle", "Statut", "Encours_MEUR", "Date_MAJ", "Commentaire")

    lignes = Array( _
        "FR0010000001;FONDS ACTIONS EUROPE;ACTIF;1250.5;01/10/2026;", _
        "FR0010000002;FONDS OBLIG CORPORATE;ACTIF;890.2;28/09/2026;RAS", _
        "FR0010000003;FONDS MONETAIRE EURO;ACTIF;2310.0;15/10/2026;", _
        "FR0010000004;FONDS DIVERSIFIE PRUDENT;ACTIF;455.8;03/10/2026;A SURVEILLER", _
        "FR0010000005;FONDS ACTIONS USA;FERME;120.4;30/06/2026;LIQUIDATION", _
        "FR0010000006;FONDS INFRA ENERGIE;ACTIF;780.0;20/10/2026;")

    For i = 0 To UBound(lignes)
        champs = Split(lignes(i), ";")
        For j = 0 To 5
            ws.Cells(i + 2, j + 1).Value = champs(j)
        Next j
    Next i
    ws.Range("H1").Value = ""
    ws.Columns("A:F").AutoFit
    On Error Resume Next
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

'==============================================================
' DESSIN DE L'ECRAN
'==============================================================
Private Sub DessinerEcran(ws As Worksheet)
    Dim zone As Range, c As Long

    ws.Cells.Clear
    ws.Cells.Interior.Color = CFond()
    ws.Range(ws.Columns(1), ws.Columns(NB_COLS)).ColumnWidth = 1.5
    ws.Range(ws.Rows(1), ws.Rows(36)).RowHeight = 15

    Set zone = ws.Range(ws.Cells(1, 1), ws.Cells(36, NB_COLS))
    With zone
        .NumberFormat = "@"
        .Font.Name = "Courier New"
        .Font.Size = 10
        .Font.Color = CTexte()
        .HorizontalAlignment = xlLeft
    End With


    Poser ws, 1, 2, "NATRAN - SIMULATION MAINFRAME MVS (3278 MODELE 3 - 32x80)", CTitre()
    Poser ws, 1, 62, "DATE " & Format(Date, "dd/mm/yyyy"), CTitre()
    Poser ws, 2, 1, String$(NB_COLS, "-"), CSep()
    Poser ws, 4, 3, "ECRAN FSIC - CONSULTATION / MISE A JOUR FONDS", CTitre()
    Poser ws, L_ISIN, 3, "CODE ISIN . . . . . :", CLabel()
    Poser ws, L_LIB, 3, "LIBELLE . . . . . . :", CLabel()
    Poser ws, L_STAT, 3, "STATUT . . . . . . :", CLabel()
    Poser ws, L_ENC, 3, "ENCOURS (MEUR) . . :", CLabel()
    Poser ws, L_DATE, 3, "DERNIERE MAJ . . . :", CLabel()
    Poser ws, L_COM, 3, "COMMENTAIRE . . . :", CLabel()
    Poser ws, 31, 1, String$(NB_COLS, "-"), CSep()
    Poser ws, L_MSG, 3, "MESSAGE :", CLabel()
    Poser ws, NB_LIGNES, 3, "ENTREE=Valider  PF3=Quitter  CLEAR=Effacer", CTitre()

    StyleSaisie ws, L_ISIN, LEN_ISIN
    StyleSaisie ws, L_COM, LEN_COM
    ws.Range(ws.Cells(L_MSG, C_MSG), ws.Cells(L_MSG, C_MSG + LEN_MSG - 1)).Font.Color = CMessage()

    For c = 1 To NB_COLS
        If c Mod 10 = 0 Then ws.Cells(34, c).Value = CStr(c \ 10)
        ws.Cells(35, c).Value = CStr(c Mod 10)
    Next c
    ws.Range(ws.Cells(34, 1), ws.Cells(35, NB_COLS)).Font.Color = RGB(120, 140, 170)
    Poser ws, 36, 1, "(regle de coordonnees, hors ecran)", RGB(120, 140, 170)

    Message ws, "PRET - SAISIR UN CODE ISIN PUIS ENTREE"
End Sub

Private Sub StyleSaisie(ws As Worksheet, l As Long, n As Long)
    With ws.Range(ws.Cells(l, C_CHAMP), ws.Cells(l, C_CHAMP + n - 1))
        .Interior.Color = CSaisie()
        .Font.Color = CTitre()
    End With
End Sub

'==============================================================
' TOUCHES DU TERMINAL (boutons optionnels sur la feuille MVS)
'==============================================================
' ENTREE : consultation de l'ISIN, ou mise a jour du commentaire
Public Sub MVS_Entree()
    Dim ws As Worksheet, wsB As Worksheet
    Dim isin As String, isinAff As String, com As String, r As Long

    On Error GoTo fin
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SHEET_MVS)
    Set wsB = ThisWorkbook.Worksheets(SHEET_BASE)

    NormaliserLigne ws, L_ISIN
    NormaliserLigne ws, L_COM
    isin = UCase$(Trim$(LireChamp(ws, L_ISIN, C_CHAMP, LEN_ISIN)))
    isinAff = CStr(wsB.Range("H1").Value)

    If isin = "" Then
        Message ws, "SAISIR UN CODE ISIN PUIS ENTREE"
        GoTo fin
    End If

    EcrireChamp ws, L_ISIN, C_CHAMP, LEN_ISIN, isin
    r = TrouverISIN(wsB, isin)

    If r = 0 Then
        EffacerSorties ws
        wsB.Range("H1").Value = ""
        Message ws, "ISIN INCONNU DANS FSIC : " & isin
    ElseIf isin <> isinAff Then
        EcrireChamp ws, L_LIB, C_CHAMP, LEN_LIB, CStr(wsB.Cells(r, 2).Value)
        EcrireChamp ws, L_STAT, C_CHAMP, LEN_STAT, CStr(wsB.Cells(r, 3).Value)
        EcrireChamp ws, L_ENC, C_CHAMP, LEN_ENC, CStr(wsB.Cells(r, 4).Value)
        EcrireChamp ws, L_DATE, C_CHAMP, LEN_DATE, CStr(wsB.Cells(r, 5).Value)
        EcrireChamp ws, L_COM, C_CHAMP, LEN_COM, CStr(wsB.Cells(r, 6).Value)
        wsB.Range("H1").Value = isin
        Message ws, "FSIC - ENREGISTREMENT TROUVE"
    Else
        com = Trim$(LireChamp(ws, L_COM, C_CHAMP, LEN_COM))
        If com <> Trim$(CStr(wsB.Cells(r, 6).Value)) Then
            wsB.Cells(r, 6).Value = com
            wsB.Cells(r, 5).Value = Format(Date, "dd/mm/yyyy")
            EcrireChamp ws, L_DATE, C_CHAMP, LEN_DATE, CStr(wsB.Cells(r, 5).Value)
            Message ws, "ENREGISTREMENT MIS A JOUR"
        Else
            Message ws, "AUCUNE MODIFICATION"
        End If
    End If

fin:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Erreur MVS_Entree : " & Err.Description, vbCritical
End Sub

Public Sub MVS_Effacer()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MVS)
    NormaliserLigne ws, L_ISIN
    EcrireChamp ws, L_ISIN, C_CHAMP, LEN_ISIN, ""
    EffacerSorties ws
    ThisWorkbook.Worksheets(SHEET_BASE).Range("H1").Value = ""
    Message ws, "ECRAN EFFACE"
End Sub

Public Sub MVS_PF3()
    MVS_Effacer
    Message ThisWorkbook.Worksheets(SHEET_MVS), "PF3 - SESSION TERMINEE (SIMULATION)"
End Sub

'==============================================================
' OUTILS ECRAN
'==============================================================

Private Sub Poser(ws As Worksheet, l As Long, c As Long, t As String, Optional couleur As Long = -1)
    Dim k As Long, fin As Long
    For k = 1 To Len(t)
        If c + k - 1 <= NB_COLS And Mid$(t, k, 1) <> " " Then ws.Cells(l, c + k - 1).Value = Mid$(t, k, 1)
    Next k
    If couleur <> -1 Then
        fin = c + Len(t) - 1
        If fin > NB_COLS Then fin = NB_COLS
        ws.Range(ws.Cells(l, c), ws.Cells(l, fin)).Font.Color = couleur
    End If
End Sub

Private Sub Message(ws As Worksheet, txt As String)
    EcrireChamp ws, L_MSG, C_MSG, LEN_MSG, txt
End Sub

Private Sub EffacerSorties(ws As Worksheet)
    EcrireChamp ws, L_LIB, C_CHAMP, LEN_LIB, ""
    EcrireChamp ws, L_STAT, C_CHAMP, LEN_STAT, ""
    EcrireChamp ws, L_ENC, C_CHAMP, LEN_ENC, ""
    EcrireChamp ws, L_DATE, C_CHAMP, LEN_DATE, ""
    EcrireChamp ws, L_COM, C_CHAMP, LEN_COM, ""
End Sub

Private Sub EcrireChamp(ws As Worksheet, l As Long, c As Long, n As Long, ByVal t As String)
    Dim k As Long
    ws.Range(ws.Cells(l, c), ws.Cells(l, c + n - 1)).ClearContents
    t = Left$(t, n)
    For k = 1 To Len(t)
        If Mid$(t, k, 1) <> " " Then ws.Cells(l, c + k - 1).Value = Mid$(t, k, 1)
    Next k
End Sub

Private Function BufferLigne(ws As Worksheet, ByVal l As Long) As String
    Dim arr As Variant, buf() As String, c As Long, k As Long, t As String
    ReDim buf(1 To NB_COLS)
    For c = 1 To NB_COLS: buf(c) = " ": Next c
    arr = ws.Range(ws.Cells(l, 1), ws.Cells(l, NB_COLS)).Value
    For c = 1 To NB_COLS
        t = CStr(arr(1, c))
        For k = 1 To Len(t)
            If c + k - 1 <= NB_COLS Then buf(c + k - 1) = Mid$(t, k, 1)
        Next k
    Next c
    BufferLigne = Join(buf, "")
End Function

Private Function LireChamp(ws As Worksheet, l As Long, c As Long, n As Long) As String
    LireChamp = Mid$(BufferLigne(ws, l), c, n)
End Function
Private Sub NormaliserLigne(ws As Worksheet, ByVal l As Long)
    Dim buf As String, c As Long
    buf = BufferLigne(ws, l)
    ws.Range(ws.Cells(l, 1), ws.Cells(l, NB_COLS)).ClearContents
    For c = 1 To NB_COLS
        If Mid$(buf, c, 1) <> " " Then ws.Cells(l, c).Value = Mid$(buf, c, 1)
    Next c
End Sub

'==============================================================
' BASE ET FEUILLES
'==============================================================
Private Function TrouverISIN(wsB As Worksheet, isin As String) As Long
    Dim r As Long, derniere As Long
    derniere = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To derniere
        If UCase$(Trim$(CStr(wsB.Cells(r, 1).Value))) = isin Then
            TrouverISIN = r
            Exit Function
        End If
    Next r
End Function

Private Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not ws Is Nothing
End Function

Private Function GetOrCreateSheet(nom As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If
    Set GetOrCreateSheet = ws
End Function

