Attribute VB_Name = "CréationCSV"
Option Explicit

'--- CONFIGURATION ---------------------------------------------
Public Const V3_CSV_FOLDER As String = "C:\Users\jdavoust\natran\VBATF_CSV\"   ' A MODIFIER
Public Const V3_CSV_PREFIX As String = "DAT_"
Public Const V3_DATE_FORMAT As String = "yyyymmdd"
Private Const SEP As String = ";"
Private Const SHEET_LOG As String = "Log"
Private Const INJECTER_LIGNE_INVALIDE As Boolean = False   ' True = ajoute un pb

' Nom complet du CSV
Public Function V3_NomCSV(d As Date) As String
    V3_NomCSV = V3_CSV_FOLDER & V3_CSV_PREFIX & Format(d, V3_DATE_FORMAT) & ".csv"
End Function

' Nom, latitude, longitude, volume de base (GWh)
Private Function Villes() As Variant
    Villes = Array( _
        Array("Lille", 50.63, 3.06, 95), _
        Array("Paris", 48.86, 2.35, 150), _
        Array("Strasbourg", 48.57, 7.75, 75), _
        Array("Nantes", 47.22, -1.55, 70), _
        Array("Bordeaux", 44.84, -0.58, 85), _
        Array("Toulouse", 43.6, 1.44, 80), _
        Array("Marseille", 43.3, 5.37, 100), _
        Array("Lyon", 45.76, 4.84, 110), _
        Array("Caen", 49.18, -0.37, 55), _
        Array("Dijon", 47.32, 5.04, 45))
End Function

'==============================================================
' POINTS D'ENTREE (boutons)
'==============================================================
Public Sub GenererCSVDuJourV3()
    On Error GoTo ErrHandler
    Dim chemin As String
    chemin = EcrireCSV(Date)
    MsgBox "CSV genere :" & vbCrLf & chemin, vbInformation, "VBATF V3 - Generation"
    Exit Sub
ErrHandler:
    LogMsg "ERREUR", "Generation CSV : " & Err.Number & " - " & Err.Description
    MsgBox "Erreur : " & Err.Description, vbCritical, "VBATF V3 - Generation"
End Sub

Public Sub GenererCSVDateV3()
    On Error GoTo ErrHandler
    Dim saisie As String, chemin As String
    saisie = Trim$(InputBox("Date du CSV a generer (jj/mm/aaaa) :", "VBATF V3", Format(Date, "dd/mm/yyyy")))
    If saisie = "" Then Exit Sub
    If Not IsDate(saisie) Then
        MsgBox "Date invalide : " & saisie, vbExclamation
        Exit Sub
    End If
    chemin = EcrireCSV(CDate(saisie))
    MsgBox "CSV genere :" & vbCrLf & chemin, vbInformation, "VBATF V3 - Generation"
    Exit Sub
ErrHandler:
    LogMsg "ERREUR", "Generation CSV : " & Err.Number & " - " & Err.Description
    MsgBox "Erreur : " & Err.Description, vbCritical, "VBATF V3 - Generation"
End Sub

'==============================================================
' ECRITURE DU CSV
'==============================================================
Private Function EcrireCSV(d As Date) As String
    Dim chemin As String, f As Integer, i As Long
    Dim listeVilles As Variant, v As Variant, vol As Double, temp As Double

    Randomize
    V3_CreerDossier V3_CSV_FOLDER
    chemin = V3_NomCSV(d)
    If Dir(chemin) <> "" Then LogMsg "INFO", "CSV existant ecrase : " & chemin

    listeVilles = Villes()
    f = FreeFile
    Open chemin For Output As #f
    Print #f, "Date" & SEP & "Ville" & SEP & "Latitude" & SEP & "Longitude" & SEP & "Volume_GWh" & SEP & "Temperature_C"
    For i = LBound(listeVilles) To UBound(listeVilles)
        v = listeVilles(i)
        vol = Round(v(3) * (0.85 + Rnd() * 0.3), 1)
        temp = Round(2 + Rnd() * 14, 1)
        Print #f, Format(d, "yyyy-mm-dd") & SEP & v(0) & SEP & nb(v(1)) & SEP & nb(v(2)) & SEP & nb(vol) & SEP & nb(temp)
    Next i
    If INJECTER_LIGNE_INVALIDE Then
        Print #f, Format(d, "yyyy-mm-dd") & SEP & "VilleInconnue" & SEP & "abc" & SEP & SEP & "xx" & SEP
    End If
    Close #f

    LogMsg "INFO", "CSV genere : " & chemin
    EcrireCSV = chemin
End Function

Private Function nb(v As Variant) As String
    nb = Trim$(Str$(CDbl(v)))
End Function

'==============================================================
' OUTILS (Public car reutilises par le module 2)
'==============================================================
Public Sub V3_CreerDossier(chemin As String)
    Dim parts As Variant, cur As String, i As Long
    parts = Split(chemin, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        End If
    Next i
End Sub

Private Sub LogMsg(niveau As String, msg As String)
    Dim ws As Worksheet, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_LOG
    End If
    If ws.Range("A1").Value = "" Then ws.Range("A1:C1").Value = Array("Horodatage", "Niveau", "Message")
    n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(n, 1).Value = Now
    ws.Cells(n, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(n, 2).Value = niveau
    ws.Cells(n, 3).Value = msg
End Sub

