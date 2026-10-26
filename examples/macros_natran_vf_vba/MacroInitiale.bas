Attribute VB_Name = "MacroInitiale"
Option Explicit


'--- CONFIGURATION ---------------------------------------------
Private Const SHEET_VISUEL As String = "Feuille1"
Private Const SHEET_DATA As String = "Feuille2"
Private Const SHEET_LOG As String = "Log"
Private Const TABLE_NAME As String = "tblDonnees"
Private Const FILE_TITLE As String = "VBA_Natran_Test"
Private Const FILE_SUFFIX As String = "VBATF"
Private Const MAIL_TO As String = "jdavoust201@headmind.com"   ' A modifier avec mail HDP
Private Const MAIL_CC As String = ""
Private Const SEND_MAIL As Boolean = True                   ' False = brouillon seulement

Private Function Regions() As Variant
    Regions = Array("Nord", "Sud", "Est", "Ouest", "Centre")
End Function

Private Function SaveFolder() As String
    SaveFolder = "C:\Users\jdavoust\natran\sortiesVBA\" ' A modifier avec path de chancun de vs 2
End Function

'==============================================================
' POINT D'ENTREE : lance tout le pipeline
'==============================================================
Public Sub LancerPipelineComplet()
    On Error GoTo ErrHandler
    Dim cheminFichier As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureStructure
    AjouterDonneesDuJour
    AppliquerFiltre            ' interactif, Annuler = pas de filtre
    CreerVisuel
    cheminFichier = SauvegarderVersion()
    EnvoyerMail cheminFichier

    LogMsg "INFO", "Pipeline termine : " & cheminFichier
    MsgBox "Pipeline termine." & vbCrLf & cheminFichier, vbInformation, "VBATF"

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogMsg "ERREUR", Err.Number & " - " & Err.Description
    MsgBox "Erreur : " & Err.Description, vbCritical, "VBATF"
    Resume Cleanup
End Sub

'==============================================================
' 1. STRUCTURE / DONNEES
'==============================================================
Private Sub EnsureStructure()
    Dim wsD As Worksheet, wsV As Worksheet, lo As ListObject
    Set wsD = GetOrCreateSheet(SHEET_DATA)
    Set wsV = GetOrCreateSheet(SHEET_VISUEL)

    On Error Resume Next
    Set lo = wsD.ListObjects(TABLE_NAME)
    On Error GoTo 0

    If lo Is Nothing Then

        Dim i As Long, r As Variant, ligne As Long
        wsD.Cells.Clear
        wsD.Range("A1:E1").Value = Array("Date", "Region", "Volume_GWh", "Temperature_C", "Statut")
        ligne = 2
        For i = 14 To 1 Step -1
            For Each r In Regions()
                wsD.Cells(ligne, 1).Value = Date - i
                wsD.Cells(ligne, 2).Value = r
                wsD.Cells(ligne, 3).Value = SimulerVolume(CStr(r))
                wsD.Cells(ligne, 4).Value = SimulerTemperature()
                wsD.Cells(ligne, 5).Value = "Historique"
                ligne = ligne + 1
            Next r
        Next i
        Set lo = wsD.ListObjects.Add(xlSrcRange, wsD.Range("A1").CurrentRegion, , xlYes)
        lo.Name = TABLE_NAME
        lo.TableStyle = "TableStyleMedium2"
        wsD.Columns("A").NumberFormat = "dd/mm/yyyy"
        wsD.Columns("A:E").AutoFit
        LogMsg "INFO", "Structure initiale creee"
    End If
End Sub

Public Sub AjouterDonneesDuJour()
    Dim lo As ListObject, r As Variant, nb As Long, newRow As ListRow
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)

    For Each r In Regions()
        If Not LigneExiste(lo, Date, CStr(r)) Then
            Set newRow = lo.ListRows.Add
            newRow.Range.Cells(1, 1).Value = Date
            newRow.Range.Cells(1, 2).Value = r
            newRow.Range.Cells(1, 3).Value = SimulerVolume(CStr(r))
            newRow.Range.Cells(1, 4).Value = SimulerTemperature()
            newRow.Range.Cells(1, 5).Value = "Phase de test"
            nb = nb + 1
        End If
    Next r
    LogMsg "INFO", nb & " ligne(s) ajoutee(s) pour le " & Format(Date, "dd/mm/yyyy")
End Sub

Private Function LigneExiste(lo As ListObject, d As Date, reg As String) As Boolean
    Dim c As ListRow
    For Each c In lo.ListRows
        If CLng(c.Range.Cells(1, 1).Value) = CLng(d) And c.Range.Cells(1, 2).Value = reg Then
            LigneExiste = True: Exit Function
        End If
    Next c
End Function

Private Function SimulerVolume(reg As String) As Double
    Dim base As Double
    Select Case reg
        Case "Nord":   base = 120
        Case "Sud":    base = 95
        Case "Est":    base = 80
        Case "Ouest":  base = 70
        Case Else:     base = 60
    End Select
    SimulerVolume = Round(base * (0.85 + Rnd() * 0.3), 1)
End Function

Private Function SimulerTemperature() As Double
    SimulerTemperature = Round(2 + Rnd() * 14, 1)
End Function

'==============================================================
' 2. FILTRE
'==============================================================
Public Sub AppliquerFiltre()
    Dim lo As ListObject, saisie As String
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)

    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData

    saisie = Trim$(InputBox("Filtrer sur quelle region ?" & vbCrLf & _
             "(" & Join(Regions(), ", ") & ")" & vbCrLf & "Vide = aucun filtre", "Filtre VBATF"))
    If saisie = "" Then Exit Sub

    If Not EstRegionValide(saisie) Then
        MsgBox "Region inconnue : " & saisie, vbExclamation
        Exit Sub
    End If
    lo.Range.AutoFilter Field:=2, Criteria1:=saisie
    LogMsg "INFO", "Filtre applique : " & saisie
End Sub

Private Function EstRegionValide(s As String) As Boolean
    Dim r As Variant
    For Each r In Regions()
        If StrComp(CStr(r), s, vbTextCompare) = 0 Then EstRegionValide = True: Exit Function
    Next r
End Function

'==============================================================
' 3. VISUEL (Feuille1) : KPI + TCD + graphique
'==============================================================
Public Sub CreerVisuel()
    Dim wsV As Worksheet, wsD As Worksheet, lo As ListObject
    Dim pc As PivotCache, pt As PivotTable, co As ChartObject

    Set wsV = ThisWorkbook.Worksheets(SHEET_VISUEL)
    Set wsD = ThisWorkbook.Worksheets(SHEET_DATA)
    Set lo = wsD.ListObjects(TABLE_NAME)

    Do While wsV.ChartObjects.Count > 0: wsV.ChartObjects(1).Delete: Loop
    Do While wsV.PivotTables.Count > 0: wsV.PivotTables(1).TableRange2.Clear: Loop
    wsV.Cells.Clear

    With wsV.Range("A1")
        .Value = "Suivi distribution gaz - mis a jour le " & Format(Now, "dd/mm/yyyy hh:nn")
        .Font.Bold = True: .Font.Size = 14
    End With
    wsV.Range("A2").Value = "Volume total (GWh)"
    wsV.Range("B2").Formula = "=SUM(" & TABLE_NAME & "[Volume_GWh])"
    wsV.Range("D2").Value = "Moyenne / jour (GWh)"
    wsV.Range("E2").Formula = "=SUM(" & TABLE_NAME & "[Volume_GWh])/SUMPRODUCT(1/COUNTIF(" & TABLE_NAME & "[Date]," & TABLE_NAME & "[Date]))"
    wsV.Range("G2").Value = "Region du pic"
    wsV.Range("A2,D2,G2").Font.Bold = True
    wsV.Range("B2,E2").NumberFormat = "#,##0.0"
    wsV.Range("H2").Formula = "=INDEX(" & TABLE_NAME & "[Region],MATCH(MAX(" & TABLE_NAME & "[Volume_GWh])," & TABLE_NAME & "[Volume_GWh],0))"

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=TABLE_NAME)
    Set pt = pc.CreatePivotTable(TableDestination:=wsV.Range("A5"), TableName:="TCD_Volume")
    With pt
        .PivotFields("Date").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlColumnField
        With .AddDataField(.PivotFields("Volume_GWh"), "Volume (GWh)", xlSum)
            .NumberFormat = "#,##0.0"
        End With
        .PivotFields("Date").NumberFormat = "dd/mm"
    End With

    Set co = wsV.ChartObjects.Add(Left:=wsV.Range("J5").Left, Top:=wsV.Range("J5").Top, Width:=520, Height:=300)
    With co.Chart
        .SetSourceData pt.TableRange1
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Volumes distribues par region (GWh)"
    End With

    wsV.Columns("A:H").AutoFit
    wsV.Activate
    LogMsg "INFO", "Visuel regenere"
End Sub

'==============================================================
' 4. SAUVEGARDE : titre_date_VBATF
'==============================================================
Private Function SauvegarderVersion() As String
    Dim dossier As String, nom As String, chemin As String
    dossier = SaveFolder()
    If Dir(dossier, vbDirectory) = "" Then MkDir dossier

    nom = FILE_TITLE & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & FILE_SUFFIX & ".xlsm"
    chemin = dossier & nom
    ThisWorkbook.SaveCopyAs chemin
    LogMsg "INFO", "Sauvegarde : " & chemin
    SauvegarderVersion = chemin
End Function

'==============================================================
' 5. MAIL (Outlook, late binding)
'==============================================================
Private Sub EnvoyerMail(cheminPJ As String)
    Dim ol As Object, mail As Object
    On Error GoTo MailErr
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    With mail
        .To = MAIL_TO
        .CC = MAIL_CC
        .Subject = "[PHASE DE TEST] Mise a jour suivi distribution - " & Format(Date, "dd/mm/yyyy")
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Ceci est un envoi automatique en phase de test." & vbCrLf & _
                "Le graphique et les donnees ont ete mis a jour." & vbCrLf & _
                "Fichier : " & Mid$(cheminPJ, InStrRev(cheminPJ, "\") + 1) & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & "Robot VBATF"
        .Attachments.Add cheminPJ
        If SEND_MAIL Then .Send Else .Display
    End With
    LogMsg "INFO", IIf(SEND_MAIL, "Mail envoye a ", "Brouillon cree pour ") & MAIL_TO
    Exit Sub
MailErr:
    LogMsg "ERREUR", "Mail impossible : " & Err.Description
End Sub

'==============================================================
' OUTILS
'==============================================================
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

Private Sub LogMsg(niveau As String, msg As String)
    Dim ws As Worksheet, n As Long
    Set ws = GetOrCreateSheet(SHEET_LOG)
    If ws.Range("A1").Value = "" Then ws.Range("A1:C1").Value = Array("Horodatage", "Niveau", "Message")
    n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(n, 1).Value = Now
    ws.Cells(n, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(n, 2).Value = niveau
    ws.Cells(n, 3).Value = msg
End Sub

