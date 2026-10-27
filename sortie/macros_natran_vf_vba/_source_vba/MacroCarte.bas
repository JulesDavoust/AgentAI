Attribute VB_Name = "MacroCarte"
Option Explicit


'--- CONFIGURATION ---------------------------------------------
Private Const SHEET_VISUEL As String = "Feuille1"
Private Const SHEET_DATA As String = "Feuille2"
Private Const SHEET_LOG As String = "Log"
Private Const TABLE_NAME As String = "tblDonnees"
Private Const FILE_TITLE As String = "VBA_Natran_Test"
Private Const FILE_SUFFIX As String = "VBATF_V2"
Private Const MAIL_TO As String = "jdavoust201@headmind.com"   ' A modifier avec mail HDP
Private Const MAIL_CC As String = ""
Private Const SEND_MAIL As Boolean = True                   ' False = brouillon seulement
Private Const OPEN_MAP As Boolean = True                    ' Ouvre la carte dans le navigateur

Private Function Regions() As Variant
    Regions = Array("Nord", "Sud", "Est", "Ouest", "Centre")
End Function

Private Function SaveFolder() As String
    SaveFolder = "C:\Users\jdavoust\natran\sortiesVBA\" ' A modifier avec path de chacun
End Function

Private Function CoordRegion(reg As String) As String
    ' Latitude,Longitude du point d'ancrage de chaque region
    Select Case reg
        Case "Nord":   CoordRegion = "50.63,3.06"     ' Lille
        Case "Sud":    CoordRegion = "43.30,5.37"     ' Marseille
        Case "Est":    CoordRegion = "48.57,7.75"     ' Strasbourg
        Case "Ouest":  CoordRegion = "47.22,-1.55"    ' Nantes
        Case "Centre": CoordRegion = "47.08,2.40"     ' Bourges
        Case Else:     CoordRegion = ""
    End Select
End Function

'==============================================================
' POINT D'ENTREE
'==============================================================
Public Sub LancerPipelineCarteV2()
    On Error GoTo ErrHandler
    Dim cheminCarte As String, cheminXlsm As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureStructure
    AjouterDonnees
    AppliquerFiltreV2
    cheminCarte = GenererCarteHTML()
    CreerSynthese cheminCarte
    cheminXlsm = SauvegarderVersion()
    EnvoyerMail cheminXlsm, cheminCarte

    If OPEN_MAP Then OuvrirFichier cheminCarte

    LogMsg "INFO", "Pipeline V2 termine : " & cheminXlsm
    MsgBox "Pipeline V2 termine." & vbCrLf & cheminXlsm & vbCrLf & cheminCarte, vbInformation, "VBATF V2"

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogMsg "ERREUR", Err.Number & " - " & Err.Description
    MsgBox "Erreur : " & Err.Description, vbCritical, "VBATF V2"
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

Private Sub AjouterDonnees()
    Dim lo As ListObject, r As Variant, nb As Long, newRow As ListRow
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)

    For Each r In Regions()
        If Not LigneExiste(lo, Date, CStr(r)) Then
            Set newRow = lo.ListRows.Add
            newRow.Range.Cells(1, 1).Value = Date
            newRow.Range.Cells(1, 2).Value = r
            newRow.Range.Cells(1, 3).Value = SimulerVolume(CStr(r))
            newRow.Range.Cells(1, 4).Value = SimulerTemperature()
            newRow.Range.Cells(1, 5).Value = "Phase de test V2"
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
Public Sub AppliquerFiltreV2()
    Dim lo As ListObject, saisie As String
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)

    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData

    saisie = Trim$(InputBox("Region a mettre en avant sur la carte ?" & vbCrLf & _
             "(" & Join(Regions(), ", ") & ")" & vbCrLf & "Vide = toutes les regions", "Filtre VBATF V2"))
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
Private Function RegionFiltree() As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)
    On Error GoTo fin
    If lo.AutoFilter.FilterMode Then
        If lo.AutoFilter.Filters(2).On Then
            RegionFiltree = Replace(lo.AutoFilter.Filters(2).Criteria1, "=", "")
        End If
    End If
fin:
End Function

'==============================================================
' 3. CARTE INTERACTIVE (HTML autonome, carte SVG, aucun appel reseau)
'==============================================================
Private Function GenererCarteHTML() As String
    Dim lo As ListObject, arr As Variant, i As Long
    Dim donnees As String, coords As String, r As Variant
    Dim h As String, chemin As String, f As Integer, regPre As String

    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 1, , "Table de donnees vide"

    arr = lo.DataBodyRange.Value
    For i = 1 To UBound(arr, 1)
        donnees = donnees & "['" & Format(arr(i, 1), "yyyy-mm-dd") & "','" & arr(i, 2) & "'," & _
                  Num(arr(i, 3)) & "," & Num(arr(i, 4)) & "],"
    Next i
    donnees = Left$(donnees, Len(donnees) - 1)

    For Each r In Regions()
        coords = coords & "'" & r & "':[" & CoordRegion(CStr(r)) & "],"
    Next r
    coords = Left$(coords, Len(coords) - 1)

    regPre = RegionFiltree()
    If regPre = "" Then regPre = "*"

    ' --- HTML  ---
    Ajoute h, "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    Ajoute h, "<title>Suivi distribution gaz - carte</title>"
    Ajoute h, "<style>html,body{height:100%;margin:0;font-family:Arial,sans-serif;background:#fff}"
    Ajoute h, "#panel{padding:8px;background:#f2f2f2;border-bottom:1px solid #ccc}"
    Ajoute h, "#panel *{margin-right:8px}#svg{width:100%;height:calc(100% - 78px);display:block}"
    Ajoute h, "#info{padding:6px 8px;height:26px;font-size:14px;border-top:1px solid #ccc;background:#fafafa}"
    Ajoute h, "circle{cursor:pointer}circle:hover{stroke-width:3}</style></head><body>"
    Ajoute h, "<div id='panel'><b>Distribution gaz</b> Date <select id='selDate'></select>"
    Ajoute h, "Region <select id='selReg'></select><button onclick='play()'>Lecture / Pause</button>"
    Ajoute h, "<span id='tot'></span><small>Genere le " & Format(Now, "dd/mm/yyyy hh:nn") & " - " & FILE_SUFFIX & "</small></div>"
    Ajoute h, "<svg id='svg' viewBox='0 0 640 640' preserveAspectRatio='xMidYMid meet'>"
    Ajoute h, "<path id='fr' fill='#eaf0e0' stroke='#6b7a5e' stroke-width='1.5' stroke-linejoin='round'/>"
    Ajoute h, "<path id='co' fill='#eaf0e0' stroke='#6b7a5e' stroke-width='1.5' stroke-linejoin='round'/>"
    Ajoute h, "<g id='g'></g></svg>"
    Ajoute h, "<div id='info'>Cliquer sur une region pour le detail.</div><script>"
    Ajoute h, "var DATA=[" & donnees & "];"
    Ajoute h, "var COORD={" & coords & "};"
    ' Contour simplifie de la France metropolitaine [longitude, latitude]
    Ajoute h, "var FR=[[2.55,51.09],[1.85,50.95],[1.60,50.73],[1.55,50.35],[1.38,50.06],[1.08,49.93],[0.37,49.76],[0.10,49.50],"
    Ajoute h, "[0.25,49.43],[0.07,49.35],[-0.25,49.30],[-0.62,49.34],[-1.10,49.40],[-1.27,49.69],[-1.95,49.72],[-1.80,49.38],"
    Ajoute h, "[-1.60,48.83],[-1.50,48.64],[-1.85,48.68],[-2.02,48.65],[-2.32,48.69],[-2.75,48.55],[-3.05,48.78],[-3.45,48.83],"
    Ajoute h, "[-3.98,48.73],[-4.77,48.33],[-4.73,48.04],[-4.37,47.80],[-3.92,47.87],[-3.37,47.73],[-3.12,47.48],[-2.75,47.55],"
    Ajoute h, "[-2.45,47.45],[-2.52,47.29],[-2.20,47.27],[-2.10,47.11],[-2.20,46.95],[-1.93,46.68],[-1.78,46.49],[-1.15,46.15],"
    Ajoute h, "[-1.25,45.95],[-1.05,45.62],[-1.07,45.57],[-1.20,45.00],[-1.25,44.65],[-1.29,44.20],[-1.45,43.65],[-1.56,43.48],"
    Ajoute h, "[-1.78,43.37],[-1.45,43.05],[-0.75,42.95],[-0.30,42.85],[0.20,42.75],[0.67,42.85],[1.00,42.80],[1.45,42.60],"
    Ajoute h, "[1.75,42.50],[2.00,42.35],[2.50,42.35],[3.05,42.43],[3.03,42.70],[3.05,43.05],[3.70,43.40],[4.20,43.50],"
    Ajoute h, "[4.80,43.35],[5.35,43.30],[5.93,43.10],[6.75,43.40],[7.00,43.55],[7.27,43.70],[7.53,43.78],[7.60,44.15],"
    Ajoute h, "[6.95,44.40],[7.05,44.65],[6.75,45.00],[7.05,45.25],[6.80,45.55],[6.85,45.95],[6.80,46.38],[6.25,46.30],"
    Ajoute h, "[6.05,46.20],[6.30,46.55],[6.40,46.80],[6.95,47.20],[7.15,47.50],[7.55,47.58],[7.55,48.00],[7.80,48.60],"
    Ajoute h, "[8.20,48.98],[7.40,49.15],[6.35,49.45],[5.85,49.55],[5.45,49.50],[5.05,49.65],[4.85,49.95],[4.20,50.25],"
    Ajoute h, "[3.60,50.35],[3.25,50.75],[2.60,50.85]];"
    Ajoute h, "var CORSE=[[9.15,41.38],[9.40,41.60],[9.55,42.10],[9.45,42.70],[9.36,43.00],[9.30,42.70],[8.75,42.57],"
    Ajoute h, "[8.60,42.28],[8.72,41.92],[8.90,41.67]];"
    Ajoute h, "var K=60,C=0.688;"
    Ajoute h, "function X(lon){return (lon+5.2)*C*K;}function Y(lat){return (51.4-lat)*K;}"
    Ajoute h, "function path(a){return a.map(function(p,i){return (i?'L':'M')+X(p[0]).toFixed(1)+' '+Y(p[1]).toFixed(1);}).join('')+'Z';}"
    Ajoute h, "document.getElementById('fr').setAttribute('d',path(FR));"
    Ajoute h, "document.getElementById('co').setAttribute('d',path(CORSE));"
    Ajoute h, "var NS='http://www.w3.org/2000/svg',g=document.getElementById('g');"
    Ajoute h, "var dates=[],regs=[];"
    Ajoute h, "DATA.forEach(function(r){if(dates.indexOf(r[0])<0)dates.push(r[0]);if(regs.indexOf(r[1])<0)regs.push(r[1]);});"
    Ajoute h, "dates.sort();"
    Ajoute h, "var selD=document.getElementById('selDate'),selR=document.getElementById('selReg');"
    Ajoute h, "dates.forEach(function(d){selD.add(new Option(d,d));});"
    Ajoute h, "selR.add(new Option('Toutes','*'));regs.forEach(function(r){selR.add(new Option(r,r));});"
    Ajoute h, "selD.value=dates[dates.length-1];selR.value='" & regPre & "';"
    Ajoute h, "var MAXV=Math.max.apply(null,DATA.map(function(r){return r[2];}));"
    Ajoute h, "function color(v){var x=v/MAXV;return x>0.85?'#b30000':x>0.65?'#e34a33':x>0.45?'#fc8d59':'#fdcc8a';}"
    Ajoute h, "function el(n,a){var e=document.createElementNS(NS,n);for(var k in a)e.setAttribute(k,a[k]);return e;}"
    Ajoute h, "function draw(){while(g.firstChild)g.removeChild(g.firstChild);var d=selD.value,f=selR.value,tot=0;"
    Ajoute h, "DATA.forEach(function(r){if(r[0]!==d)return;if(f!=='*'&&r[1]!==f)return;"
    Ajoute h, "var c=COORD[r[1]];if(!c)return;tot+=r[2];"
    Ajoute h, "var rad=12+34*r[2]/MAXV,cx=X(c[1]),cy=Y(c[0]);"
    Ajoute h, "var ci=el('circle',{cx:cx,cy:cy,r:rad,fill:color(r[2]),'fill-opacity':0.8,stroke:'#333','stroke-width':1});"
    Ajoute h, "var tt=el('title',{});tt.textContent=r[1]+' : '+r[2].toFixed(1)+' GWh';ci.appendChild(tt);"
    Ajoute h, "ci.onclick=function(){document.getElementById('info').innerText=r[1]+' - '+d+' : '+r[2].toFixed(1)+' GWh, temperature '+r[3].toFixed(1)+' C';};"
    Ajoute h, "g.appendChild(ci);"
    Ajoute h, "var t=el('text',{x:cx,y:cy+4,'text-anchor':'middle','font-size':12,'font-weight':'bold',fill:'#222','pointer-events':'none'});"
    Ajoute h, "t.textContent=r[1];g.appendChild(t);});"
    Ajoute h, "document.getElementById('tot').innerText='Total affiche : '+tot.toFixed(1)+' GWh';}"
    Ajoute h, "selD.onchange=draw;selR.onchange=draw;var timer=null;"
    Ajoute h, "function play(){if(timer){clearInterval(timer);timer=null;return;}"
    Ajoute h, "var i=dates.indexOf(selD.value);"
    Ajoute h, "timer=setInterval(function(){i=(i+1)%dates.length;selD.value=dates[i];draw();},800);}"
    Ajoute h, "draw();</script></body></html>"

    ' Ecriture du fichier
    chemin = EnsureFolder() & FILE_TITLE & "_Carte_" & Format(Now, "yyyymmdd_hhnnss") & "_" & FILE_SUFFIX & ".html"
    f = FreeFile
    Open chemin For Output As #f
    Print #f, h
    Close #f

    LogMsg "INFO", "Carte generee : " & chemin
    GenererCarteHTML = chemin
End Function

Private Sub Ajoute(ByRef sb As String, ByVal ligne As String)
    sb = sb & ligne & vbLf
End Sub

Private Function Num(v As Variant) As String
    Num = Trim$(Str$(CDbl(v)))
End Function

'==============================================================
' 4. SYNTHESE (Feuille1) : KPI + tableau du dernier jour + lien carte
'==============================================================
Private Sub CreerSynthese(cheminCarte As String)
    Dim wsV As Worksheet, i As Long, r As Variant, ligne As Long

    Set wsV = ThisWorkbook.Worksheets(SHEET_VISUEL)

    Do While wsV.ChartObjects.Count > 0: wsV.ChartObjects(1).Delete: Loop
    Do While wsV.PivotTables.Count > 0: wsV.PivotTables(1).TableRange2.Clear: Loop
    wsV.Hyperlinks.Delete
    wsV.Cells.Clear

    With wsV.Range("A1")
        .Value = "Suivi distribution gaz (V2 carte) - mis a jour le " & Format(Now, "dd/mm/yyyy hh:nn")
        .Font.Bold = True: .Font.Size = 14
    End With

    wsV.Range("A2").Value = "Volume total (GWh)"
    wsV.Range("B2").Formula = "=SUM(" & TABLE_NAME & "[Volume_GWh])"
    wsV.Range("D2").Value = "Moyenne / jour (GWh)"
    wsV.Range("E2").Formula = "=SUM(" & TABLE_NAME & "[Volume_GWh])/SUMPRODUCT(1/COUNTIF(" & TABLE_NAME & "[Date]," & TABLE_NAME & "[Date]))"
    wsV.Range("G2").Value = "Region du pic"
    wsV.Range("H2").Formula = "=INDEX(" & TABLE_NAME & "[Region],MATCH(MAX(" & TABLE_NAME & "[Volume_GWh])," & TABLE_NAME & "[Volume_GWh],0))"
    wsV.Range("A2,D2,G2").Font.Bold = True
    wsV.Range("B2,E2").NumberFormat = "#,##0.0"

    wsV.Range("A3").Value = "Dernier jour"
    wsV.Range("B3").Formula = "=MAX(" & TABLE_NAME & "[Date])"
    wsV.Range("B3").NumberFormat = "dd/mm/yyyy"
    wsV.Range("A4:C4").Value = Array("Region", "Volume du jour (GWh)", "Temp. moyenne (C)")
    wsV.Range("A4:C4").Font.Bold = True
    wsV.Range("A4:C4").Interior.Color = RGB(221, 235, 247)

    ligne = 5
    For Each r In Regions()
        wsV.Cells(ligne, 1).Value = r
        wsV.Cells(ligne, 2).Formula = "=SUMIFS(" & TABLE_NAME & "[Volume_GWh]," & TABLE_NAME & "[Region],A" & ligne & "," & TABLE_NAME & "[Date],$B$3)"
        wsV.Cells(ligne, 3).Formula = "=IFERROR(AVERAGEIFS(" & TABLE_NAME & "[Temperature_C]," & TABLE_NAME & "[Region],A" & ligne & "," & TABLE_NAME & "[Date],$B$3),"""")"
        ligne = ligne + 1
    Next r
    wsV.Range(wsV.Cells(5, 2), wsV.Cells(ligne - 1, 3)).NumberFormat = "#,##0.0"
    wsV.Range(wsV.Cells(5, 2), wsV.Cells(ligne - 1, 2)).FormatConditions.AddDatabar

    wsV.Hyperlinks.Add Anchor:=wsV.Cells(ligne + 1, 1), Address:=cheminCarte, _
        TextToDisplay:="Ouvrir la carte interactive"
    wsV.Cells(ligne + 2, 1).Value = "Fichier : " & cheminCarte
    wsV.Cells(ligne + 2, 1).Font.Italic = True

    wsV.Columns("A:H").AutoFit
    wsV.Columns("A").ColumnWidth = 24
    wsV.Activate
    LogMsg "INFO", "Synthese regeneree"
End Sub

'==============================================================
' 5. SAUVEGARDE : titre_date_VBATF_V2
'==============================================================
Private Function EnsureFolder() As String
    Dim dossier As String
    dossier = SaveFolder()
    If Dir(dossier, vbDirectory) = "" Then MkDir dossier
    EnsureFolder = dossier
End Function

Private Function SauvegarderVersion() As String
    Dim chemin As String
    chemin = EnsureFolder() & FILE_TITLE & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & FILE_SUFFIX & ".xlsm"
    ThisWorkbook.SaveCopyAs chemin
    LogMsg "INFO", "Sauvegarde : " & chemin
    SauvegarderVersion = chemin
End Function

'==============================================================
' 6. MAIL (Outlook, late binding) : xlsm + carte en pieces jointes
'==============================================================
Private Sub EnvoyerMail(cheminXlsm As String, cheminCarte As String)
    Dim ol As Object, mail As Object
    On Error GoTo MailErr
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    With mail
        .To = MAIL_TO
        .CC = MAIL_CC
        .Subject = "[PHASE DE TEST V2] Mise a jour suivi distribution - " & Format(Date, "dd/mm/yyyy")
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Ceci est un envoi automatique en phase de test (version carte interactive)." & vbCrLf & _
                "Les donnees et la carte ont ete mises a jour." & vbCrLf & _
                "Classeur : " & Mid$(cheminXlsm, InStrRev(cheminXlsm, "\") + 1) & vbCrLf & _
                "Carte : " & Mid$(cheminCarte, InStrRev(cheminCarte, "\") + 1) & _
                " (a ouvrir dans un navigateur, fonctionne hors ligne)" & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & "Robot VBATF"
        .Attachments.Add cheminXlsm
        .Attachments.Add cheminCarte
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
Private Sub OuvrirFichier(chemin As String)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink chemin
    If Err.Number <> 0 Then LogMsg "ERREUR", "Ouverture impossible : " & Err.Description
End Sub

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

