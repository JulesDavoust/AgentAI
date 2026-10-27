Attribute VB_Name = "TestInterfaceMVS"
Option Explicit

Public Const TIMEOUT As Integer = 200

Private Const SHEET_MVS As String = "MVS"
Private Const NB_LIGNES As Long = 32
Private Const NB_COLS As Long = 80
Private Const AFFICHER_CURSEUR As Boolean = True    ' selectionne la cellule courante (visuel)

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

Private ecran As Worksheet
Private curLigne As Long
Private curCol As Long

'==============================================================
' API "TERMINAL" (meme signatures que le code d'origine)
'==============================================================
Public Function Init_MVS() As Boolean
    Set ecran = Nothing
    On Error Resume Next
    Set ecran = ThisWorkbook.Worksheets(SHEET_MVS)
    On Error GoTo 0

    If ecran Is Nothing Then
        MsgBox "Impossible de se connecter a MVS : feuille '" & SHEET_MVS & "' absente." & _
               vbCrLf & "Clique d'abord sur le bouton 'Ouvrir le mainframe'.", vbCritical
        Init_MVS = False
        Exit Function
    End If
    If Modele_Session <> 3 Then
        MsgBox "Votre terminal doit etre parametre en Modele 3 (32x80)", vbCritical
        Init_MVS = False
        Exit Function
    End If
    curLigne = 1: curCol = 1
    Init_MVS = True
End Function

Public Sub Attendre(ByVal Temps As Integer)
    DoEvents
End Sub

Public Function Lire(ByVal ligne As Integer, ByVal colonne As Integer, ByVal Taille As Integer) As String
    Dim s As String
    VerifierCoord ligne, colonne
    s = Mid$(BufferLigne(ligne), colonne, Taille)
    Lire = Left$(s & Space$(Taille), Taille)
    Call Attendre(TIMEOUT)
End Function

Public Sub Ecrire(ByVal ligne As Integer, ByVal colonne As Integer, ByVal Chaine As String)
    Dim k As Long
    Se_Positionner ligne, colonne
    NormaliserLigne ligne
    For k = 1 To Len(Chaine)
        If colonne + k - 1 > NB_COLS Then Exit For
        If Mid$(Chaine, k, 1) = " " Then
            ecran.Cells(ligne, colonne + k - 1).ClearContents
        Else
            ecran.Cells(ligne, colonne + k - 1).Value = Mid$(Chaine, k, 1)
        End If
    Next k
    curCol = colonne + Len(Chaine)
    If curCol > NB_COLS Then curCol = NB_COLS
    Call Attendre(TIMEOUT)
End Sub

Public Sub Touche(ByVal Cle As String)
    Select Case UCase$(Cle)
        Case "@E"
            AppelerMVS "MVS_Entree"
        Case "@C"
            AppelerMVS "MVS_Effacer"
        Case "@3"
            AppelerMVS "MVS_PF3"
        Case "@F"
            NormaliserLigne curLigne
            ecran.Range(ecran.Cells(curLigne, curCol), ecran.Cells(curLigne, NB_COLS)).ClearContents
        Case "", "@T"
        Case Else
            If Left$(Cle, 1) = "@" Then
                Debug.Print "Touche non geree : " & Cle
            Else
                Ecrire curLigne, curCol, Cle
            End If
    End Select
    Call Attendre(TIMEOUT)
End Sub

Public Function Recherche(ByVal texte As String) As Boolean
    Dim l As Long
    For l = 1 To NB_LIGNES
        If InStr(1, BufferLigne(l), texte, vbTextCompare) > 0 Then
            Recherche = True
            Exit Function
        End If
    Next l
End Function

Public Function Modele_Session() As Integer
    If ecran Is Nothing Then
        Modele_Session = 0
    Else
        Modele_Session = 3
    End If
End Function

Public Sub Se_Positionner(ByVal ligne As Integer, ByVal colonne As Integer)
    VerifierCoord ligne, colonne
    curLigne = ligne
    curCol = colonne
    If AFFICHER_CURSEUR Then
        On Error Resume Next
        If Not ActiveSheet Is ecran Then ecran.Activate
        ecran.Cells(ligne, colonne).Select
        On Error GoTo 0
    End If
    Call Attendre(TIMEOUT)
End Sub

'==============================================================
' BOUTON 2 : LIRE OU MODIFIER UNE ZONE DU MAINFRAME
' Demande : action (lire / ecrire), ligne, colonne, puis longueur ou texte.
'==============================================================
Public Sub InteragirMainframe()
    Dim choix As VbMsgBoxResult
    Dim ligne As Long, colonne As Long, longueur As Long, maxLong As Long
    Dim texte As String, lu As String

    If Not FeuilleMVSExiste() Then
        MsgBox "Le mainframe n'est pas encore ouvert." & vbCrLf & _
               "Clique d'abord sur le bouton 'Ouvrir le mainframe'.", vbExclamation, "Mainframe"
        Exit Sub
    End If
    If Not Init_MVS Then Exit Sub

    choix = MsgBox("Que veux-tu faire sur le mainframe ?" & vbCrLf & vbCrLf & _
                   "OUI = LIRE une zone" & vbCrLf & _
                   "NON = ECRIRE / MODIFIER une zone" & vbCrLf & _
                   "ANNULER = quitter", vbYesNoCancel + vbQuestion, "Mainframe - Lire / Modifier")
    If choix = vbCancel Then Exit Sub

    If Not DemanderEntier("Ligne (1 a " & NB_LIGNES & ") :", 1, NB_LIGNES, ligne) Then Exit Sub
    If Not DemanderEntier("Colonne (1 a " & NB_COLS & ") :", 1, NB_COLS, colonne) Then Exit Sub

    If choix = vbYes Then
        ' --- LECTURE ---
        maxLong = NB_COLS - colonne + 1
        If Not DemanderEntier("Longueur a lire (1 a " & maxLong & ") :", 1, maxLong, longueur, "10") Then Exit Sub
        lu = Lire(CInt(ligne), CInt(colonne), CInt(longueur))
        SelectionnerZone ligne, colonne, longueur
        MsgBox "Ligne " & ligne & ", colonne " & colonne & ", longueur " & longueur & vbCrLf & vbCrLf & _
               "Contenu : [" & lu & "]", vbInformation, "Mainframe - Lecture"
    Else
        ' --- ECRITURE ---
        texte = InputBox("Texte a ecrire en (" & ligne & "," & colonne & ") :" & vbCrLf & _
                         "(des espaces effacent la zone)", "Mainframe - Ecriture")
        If texte = "" Then Exit Sub
        maxLong = NB_COLS - colonne + 1
        If Len(texte) > maxLong Then
            texte = Left$(texte, maxLong)
            MsgBox "Texte tronque a " & maxLong & " caracteres (fin de ligne).", vbExclamation
        End If
        Ecrire CInt(ligne), CInt(colonne), texte
        SelectionnerZone ligne, colonne, Len(texte)
        If MsgBox("Texte ecrit sur le mainframe." & vbCrLf & _
                  "Valider avec ENTREE sur le mainframe ?", vbYesNo + vbQuestion, "Mainframe - Ecriture") = vbYes Then
            Touche "@E"
            MsgBox "Reponse du mainframe :" & vbCrLf & Trim$(Lire(L_MSG, C_MSG, LEN_MSG)), _
                   vbInformation, "Mainframe"
        End If
    End If
End Sub

'==============================================================
' DEMOS (boutons optionnels)
'==============================================================
Public Sub Demo_EcrireEtLire()
    Dim libelle As String, statut As String, encours As String
    If Not Init_MVS Then Exit Sub

    ConsulterISIN "FR0010000002"
    libelle = Trim$(Lire(L_LIB, C_CHAMP, LEN_LIB))
    statut = Trim$(Lire(L_STAT, C_CHAMP, LEN_STAT))
    encours = Trim$(Lire(L_ENC, C_CHAMP, LEN_ENC))
    MsgBox "Lecture mainframe :" & vbCrLf & libelle & vbCrLf & "Statut : " & statut & _
           vbCrLf & "Encours (MEUR) : " & encours, vbInformation, "Demo lecture"

    Se_Positionner L_COM, C_CHAMP
    Touche "@F"
    Ecrire L_COM, C_CHAMP, "MAJ PAR VBA " & Format(Now, "hh:nn:ss")
    Touche "@E"
    MsgBox "Reponse du mainframe : " & Trim$(Lire(L_MSG, C_MSG, LEN_MSG)), vbInformation, "Demo ecriture"
End Sub

Public Sub Demo_LireSaisieManuelle()
    If Not Init_MVS Then Exit Sub
    MsgBox "ISIN saisi : " & Trim$(Lire(L_ISIN, C_CHAMP, LEN_ISIN)) & vbCrLf & _
           "Commentaire : " & Trim$(Lire(L_COM, C_CHAMP, LEN_COM)) & vbCrLf & _
           "Message : " & Trim$(Lire(L_MSG, C_MSG, LEN_MSG)), vbInformation, "Lecture de l'ecran"
End Sub


Public Sub Demo_VerifierExistenceFSIC()
    Dim isin As String
    isin = UCase$(Trim$(InputBox("Code ISIN a verifier dans FSIC :", "FSIC", "FR0010000003")))
    If isin = "" Then Exit Sub
    If Not Init_MVS Then Exit Sub

    ConsulterISIN isin
    If Recherche("INCONNU") Then
        MsgBox isin & " n'existe pas dans FSIC.", vbExclamation, "FSIC"
    Else
        MsgBox isin & " existe dans FSIC :" & vbCrLf & Trim$(Lire(L_LIB, C_CHAMP, LEN_LIB)), vbInformation, "FSIC"
    End If
End Sub

Public Sub Demo_ExtractionVersFeuille()
    Dim liste As Variant, i As Long, wsX As Worksheet, r As Long

    liste = Array("FR0010000001", "FR0010000002", "FR0010000003", "FR0010000004", _
                  "FR0010000005", "FR0010000006", "FR0000000000")   ' le dernier est inconnu
    If Not Init_MVS Then Exit Sub

    On Error Resume Next
    Set wsX = ThisWorkbook.Worksheets("Extraction_MVS")
    On Error GoTo 0
    If wsX Is Nothing Then
        Set wsX = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsX.Name = "Extraction_MVS"
    End If
    wsX.Cells.Clear
    wsX.Range("A1:F1").Value = Array("ISIN", "Libelle", "Statut", "Encours_MEUR", "Date_MAJ", "Resultat")
    wsX.Range("A1:F1").Font.Bold = True

    For i = LBound(liste) To UBound(liste)
        r = i - LBound(liste) + 2
        ConsulterISIN CStr(liste(i))
        wsX.Cells(r, 1).Value = liste(i)
        If Recherche("INCONNU") Then
            wsX.Cells(r, 6).Value = "INCONNU"
        Else
            wsX.Cells(r, 2).Value = Trim$(Lire(L_LIB, C_CHAMP, LEN_LIB))
            wsX.Cells(r, 3).Value = Trim$(Lire(L_STAT, C_CHAMP, LEN_STAT))
            wsX.Cells(r, 4).Value = Val(Trim$(Lire(L_ENC, C_CHAMP, LEN_ENC)))
            wsX.Cells(r, 5).Value = Trim$(Lire(L_DATE, C_CHAMP, LEN_DATE))
            wsX.Cells(r, 6).Value = "OK"
        End If
    Next i

    wsX.Columns("A:F").AutoFit
    wsX.Activate
    MsgBox (UBound(liste) - LBound(liste) + 1) & " ISIN interroges sur le mainframe.", vbInformation, "Extraction"
End Sub

'==============================================================
' OUTILS
'==============================================================
Private Sub ConsulterISIN(ByVal isin As String)
    Se_Positionner L_ISIN, C_CHAMP
    Touche "@F"
    Ecrire L_ISIN, C_CHAMP, isin
    Touche "@E"
End Sub

Private Sub AppelerMVS(ByVal nomMacro As String)
    On Error Resume Next
    Application.Run nomMacro
    If Err.Number <> 0 Then
        MsgBox "Macro '" & nomMacro & "' introuvable ou en erreur : " & Err.Description & vbCrLf & _
               "Le module de simulation du mainframe (SimuMainframe) est-il present ?", vbCritical
        Err.Clear
    End If
End Sub

Private Sub VerifierCoord(ByVal ligne As Long, ByVal colonne As Long)
    If ligne < 1 Or ligne > NB_LIGNES Or colonne < 1 Or colonne > NB_COLS Then
        Err.Raise vbObjectError + 20, , "Coordonnees hors ecran : (" & ligne & "," & colonne & ")"
    End If
End Sub


Private Function BufferLigne(ByVal l As Long) As String
    Dim arr As Variant, buf() As String, c As Long, k As Long, t As String
    ReDim buf(1 To NB_COLS)
    For c = 1 To NB_COLS: buf(c) = " ": Next c
    arr = ecran.Range(ecran.Cells(l, 1), ecran.Cells(l, NB_COLS)).Value
    For c = 1 To NB_COLS
        t = CStr(arr(1, c))
        For k = 1 To Len(t)
            If c + k - 1 <= NB_COLS Then buf(c + k - 1) = Mid$(t, k, 1)
        Next k
    Next c
    BufferLigne = Join(buf, "")
End Function


Private Sub NormaliserLigne(ByVal l As Long)
    Dim buf As String, c As Long
    buf = BufferLigne(l)
    ecran.Range(ecran.Cells(l, 1), ecran.Cells(l, NB_COLS)).ClearContents
    For c = 1 To NB_COLS
        If Mid$(buf, c, 1) <> " " Then ecran.Cells(l, c).Value = Mid$(buf, c, 1)
    Next c
End Sub


Private Function DemanderEntier(ByVal invite As String, ByVal mini As Long, ByVal maxi As Long, _
                                ByRef valeur As Long, Optional ByVal defaut As String = "") As Boolean
    Dim s As String
    Do
        s = Trim$(InputBox(invite, "Mainframe", defaut))
        If s = "" Then Exit Function
        If IsNumeric(s) Then
            If CDbl(s) >= mini And CDbl(s) <= maxi And CDbl(s) = Int(CDbl(s)) Then
                valeur = CLng(s)
                DemanderEntier = True
                Exit Function
            End If
        End If
        MsgBox "Valeur invalide : saisir un entier entre " & mini & " et " & maxi & ".", vbExclamation
    Loop
End Function


Private Sub SelectionnerZone(ByVal l As Long, ByVal c As Long, ByVal n As Long)
    On Error Resume Next
    If Not ActiveSheet Is ecran Then ecran.Activate
    ecran.Range(ecran.Cells(l, c), ecran.Cells(l, c + n - 1)).Select
End Sub


Private Function FeuilleMVSExiste() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MVS)
    On Error GoTo 0
    FeuilleMVSExiste = Not ws Is Nothing
End Function

