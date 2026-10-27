Attribute VB_Name = "UtilisationCSV"
Option Explicit

'--- CONFIGURATION ---------------------------------------------
' IMPORTANT : CSV_FOLDER, CSV_PREFIX et DATE_FORMAT doivent etre
' identiques a ceux du module de creation du CSV.
Private Const CSV_FOLDER As String = "C:\Users\jdavoust\natran\VBATF_CSV\"   ' A MODIFIER
Private Const CSV_PREFIX As String = "DAT_"
Private Const DATE_FORMAT As String = "yyyymmdd"
Private Const CSV_SEP As String = ";"

Private Const SHEET_DATA As String = "Donnees_V3"
Private Const SHEET_VISUEL As String = "Carte_V3"
Private Const SHEET_LOG As String = "Log"
Private Const TABLE_NAME As String = "tblCSV_V3"
Private Const FILE_TITLE As String = "VBA_Natran_Test"
Private Const FILE_SUFFIX As String = "VBATF_V3"
Private Const MAIL_TO As String = "tom.ferrandis@hotmail.com"   ' A modifier avec mail HDP
Private Const MAIL_CC As String = ""
Private Const SEND_MAIL As Boolean = True                   ' False = brouillon seulement
Private Const OPEN_MAP As Boolean = True                    ' Ouvre la carte dans le navigateur

Private Function SaveFolder() As String
    SaveFolder = "C:\Users\jdavoust\natran\sortiesVBA\" ' A modifier avec path de chacun
End Function

Private Function NomCSV(d As Date) As String
    NomCSV = CSV_FOLDER & CSV_PREFIX & Format(d, DATE_FORMAT) & ".csv"
End Function

'==============================================================
' POINT D'ENTREE
'==============================================================
Public Sub LancerPipelineCSVV3()
    On Error GoTo ErrHandler
    Dim cheminCSV As String, cheminCarte As String, cheminXlsm As String
    Dim villePre As String, nbLignes As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    cheminCSV = NomCSV(Date)
    If Dir(cheminCSV) = "" Then
        Err.Raise vbObjectError + 10, , "CSV du jour introuvable : " & cheminCSV & _
            vbCrLf & "Genere d'abord le CSV avec le module de creation."
    End If

    nbLignes = ImporterCSV(cheminCSV)
    villePre = DemanderVille()
    cheminCarte = GenererCarteHTML(villePre)
    CreerSynthese cheminCarte, cheminCSV
    cheminXlsm = SauvegarderVersion()
    EnvoyerMail cheminXlsm, cheminCarte, cheminCSV

    If OPEN_MAP Then OuvrirFichier cheminCarte

    LogMsg "INFO", "Pipeline V3 termine : " & nbLignes & " ligne(s) valide(s), " & cheminXlsm
    MsgBox "Pipeline V3 termine." & vbCrLf & nbLignes & " ville(s) importee(s)" & vbCrLf & _
           cheminXlsm & vbCrLf & cheminCarte, vbInformation, "VBATF V3"

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    LogMsg "ERREUR", Err.Number & " - " & Err.Description
    MsgBox "Erreur : " & Err.Description, vbCritical, "VBATF V3"
    Resume Cleanup
End Sub

'==============================================================
' 1. IMPORT DU CSV -> feuille Donnees_V3 (tableau tblCSV_V3)
'==============================================================
Private Function ImporterCSV(chemin As String) As Long
    Dim wsD As Worksheet, lo As ListObject, f As Integer
    Dim ligne As String, champs As Variant, n As Long, nbInvalides As Long

    Set wsD = GetOrCreateSheet(SHEET_DATA)
    Do While wsD.ListObjects.Count > 0: wsD.ListObjects(1).Delete: Loop
    wsD.Cells.Clear
    wsD.Range("A1:F1").Value = Array("Date", "Ville", "Latitude", "Longitude", "Volume_GWh", "Temperature_C")

    n = 1
    f = FreeFile
    Open chemin For Input As #f
    If Not EOF(f) Then Line Input #f, ligne          ' en-tete ignoree
    Do While Not EOF(f)
        Line Input #f, ligne
        If Len(Trim$(ligne)) > 0 Then
            champs = Split(ligne, CSV_SEP)
            If LigneValide(champs) Then
                n = n + 1
                wsD.Cells(n, 1).Value = DateSerial(CInt(Left$(champs(0), 4)), CInt(Mid$(champs(0), 6, 2)), CInt(Mid$(champs(0), 9, 2)))
                wsD.Cells(n, 2).Value = Trim$(champs(1))
                wsD.Cells(n, 3).Value = Val(Trim$(champs(2)))
                wsD.Cells(n, 4).Value = Val(Trim$(champs(3)))
                wsD.Cells(n, 5).Value = Val(Trim$(champs(4)))
                wsD.Cells(n, 6).Value = Val(Trim$(champs(5)))
            Else
                nbInvalides = nbInvalides + 1
                LogMsg "AVERTISSEMENT", "Ligne ignoree : " & ligne
            End If
        End If
    Loop
    Close #f

    If n = 1 Then Err.Raise vbObjectError + 11, , "Aucune ligne valide dans " & chemin

    Set lo = wsD.ListObjects.Add(xlSrcRange, wsD.Range(wsD.Cells(1, 1), wsD.Cells(n, 6)), , xlYes)
    lo.Name = TABLE_NAME
    lo.TableStyle = "TableStyleMedium2"
    wsD.Columns("A").NumberFormat = "dd/mm/yyyy"
    wsD.Columns("A:F").AutoFit

    LogMsg "INFO", "CSV importe : " & (n - 1) & " valide(s), " & nbInvalides & " ignoree(s) - " & chemin
    ImporterCSV = n - 1
End Function

Private Function LigneValide(champs As Variant) As Boolean
    Dim lat As Double, lon As Double
    If UBound(champs) <> 5 Then Exit Function
    ' Date au format yyyy-mm-dd
    If Len(Trim$(champs(0))) <> 10 Then Exit Function
    If Mid$(champs(0), 5, 1) <> "-" Or Mid$(champs(0), 8, 1) <> "-" Then Exit Function
    If Not (IsNumeric(Left$(champs(0), 4)) And IsNumeric(Mid$(champs(0), 6, 2)) And IsNumeric(Mid$(champs(0), 9, 2))) Then Exit Function
    If Trim$(champs(1)) = "" Then Exit Function
    If Not (EstNombre(champs(2)) And EstNombre(champs(3)) And EstNombre(champs(4)) And EstNombre(champs(5))) Then Exit Function
    lat = Val(Trim$(champs(2))): lon = Val(Trim$(champs(3)))
    If lat < 41 Or lat > 52 Or lon < -5.5 Or lon > 10 Then Exit Function
    LigneValide = True
End Function

Private Function EstNombre(s As Variant) As Boolean
    Dim t As String
    t = Trim$(CStr(s))
    If t = "" Then Exit Function
    If InStr(t, ",") > 0 Then Exit Function
    EstNombre = IsNumeric(Replace(t, ".", Application.International(xlDecimalSeparator)))
End Function

'==============================================================
' 2. CHOIX DE LA VILLE A METTRE EN AVANT
'==============================================================
Private Function DemanderVille() As String
    Dim lo As ListObject, arr As Variant, i As Long, liste As String, saisie As String
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)
    arr = lo.ListColumns("Ville").DataBodyRange.Value
    For i = 1 To UBound(arr, 1)
        liste = liste & arr(i, 1) & ", "
    Next i
    liste = Left$(liste, Len(liste) - 2)

    saisie = Trim$(InputBox("Ville a mettre en avant sur la carte ?" & vbCrLf & _
             "(" & liste & ")" & vbCrLf & "Vide = toutes les villes", "VBATF V3"))
    If saisie = "" Then Exit Function

    For i = 1 To UBound(arr, 1)
        If StrComp(CStr(arr(i, 1)), saisie, vbTextCompare) = 0 Then
            DemanderVille = CStr(arr(i, 1))
            LogMsg "INFO", "Ville selectionnee : " & DemanderVille
            Exit Function
        End If
    Next i
    MsgBox "Ville inconnue : " & saisie & vbCrLf & "Toutes les villes seront affichees.", vbExclamation
End Function

'==============================================================
' 3. CARTE INTERACTIVE (HTML autonome, carte SVG, aucun appel reseau)
'==============================================================
Private Function GenererCarteHTML(villePre As String) As String
    Dim lo As ListObject, arr As Variant, i As Long
    Dim donnees As String, h As String, chemin As String, f As Integer
    Dim sel As String, dateDonnees As String

    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 1, , "Table de donnees vide"

    ' Donnees -> tableau JS : ['Caen',49.18,-0.37,120.5,7.2],...
    arr = lo.DataBodyRange.Value
    dateDonnees = Format(arr(1, 1), "dd/mm/yyyy")
    For i = 1 To UBound(arr, 1)
        donnees = donnees & "['" & Replace(CStr(arr(i, 2)), "'", "\'") & "'," & _
                  Num(arr(i, 3)) & "," & Num(arr(i, 4)) & "," & Num(arr(i, 5)) & "," & Num(arr(i, 6)) & "],"
    Next i
    donnees = Left$(donnees, Len(donnees) - 1)

    sel = villePre
    If sel = "" Then sel = "*"

    Ajoute h, "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    Ajoute h, "<title>Suivi distribution gaz - carte</title>"
    Ajoute h, "<style>html,body{height:100%;margin:0;font-family:Arial,sans-serif;background:#fff}"
    Ajoute h, "#panel{padding:8px;background:#f2f2f2;border-bottom:1px solid #ccc}"
    Ajoute h, "#panel *{margin-right:8px}#svg{width:100%;height:calc(100% - 78px);display:block}"
    Ajoute h, "#info{padding:6px 8px;height:26px;font-size:14px;border-top:1px solid #ccc;background:#fafafa}"
    Ajoute h, "circle{cursor:pointer}circle:hover{stroke-width:3}</style></head><body>"
    Ajoute h, "<div id='panel'><b>Distribution gaz - donnees du " & dateDonnees & "</b> Ville <select id='selV'></select>"
    Ajoute h, "<span id='tot'></span><small>Taille et couleur = volume | Genere le " & Format(Now, "dd/mm/yyyy hh:nn") & " - " & FILE_SUFFIX & "</small></div>"
    Ajoute h, "<svg id='svg' viewBox='0 0 640 640' preserveAspectRatio='xMidYMid meet'>"
    Ajoute h, "<path id='fr' fill='#eaf0e0' stroke='#6b7a5e' stroke-width='1.5' stroke-linejoin='round'/>"
    Ajoute h, "<path id='co' fill='#eaf0e0' stroke='#6b7a5e' stroke-width='1.5' stroke-linejoin='round'/>"
    Ajoute h, "<g id='g'></g></svg>"
    Ajoute h, "<div id='info'>Cliquer sur une ville pour le detail.</div><script>"
    Ajoute h, "var DATA=[" & donnees & "];"   ' ville, lat, lon, volume, temperature
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
    Ajoute h, "var selV=document.getElementById('selV');"
    Ajoute h, "var names=DATA.map(function(r){return r[0];}).sort();"
    Ajoute h, "selV.add(new Option('Toutes','*'));names.forEach(function(n){selV.add(new Option(n,n));});"
    Ajoute h, "selV.value='" & Replace(sel, "'", "\'") & "';"
    Ajoute h, "var MAXV=Math.max.apply(null,DATA.map(function(r){return r[3];}));"
    Ajoute h, "function color(v){var x=v/MAXV;return x>0.85?'#b30000':x>0.65?'#e34a33':x>0.45?'#fc8d59':'#fdcc8a';}"
    Ajoute h, "function el(n,a){var e=document.createElementNS(NS,n);for(var k in a)e.setAttribute(k,a[k]);return e;}"
    Ajoute h, "function draw(){while(g.firstChild)g.removeChild(g.firstChild);var f=selV.value,tot=0;"
    Ajoute h, "DATA.forEach(function(r){if(f!=='*'&&r[0]!==f)return;tot+=r[3];"
    Ajoute h, "var rad=8+26*r[3]/MAXV,cx=X(r[2]),cy=Y(r[1]);"
    Ajoute h, "var ci=el('circle',{cx:cx,cy:cy,r:rad,fill:color(r[3]),'fill-opacity':0.8,stroke:'#333','stroke-width':1});"
    Ajoute h, "var tt=el('title',{});tt.textContent=r[0]+' : '+r[3].toFixed(1)+' GWh';ci.appendChild(tt);"
    Ajoute h, "ci.onclick=function(){document.getElementById('info').innerText=r[0]+' : '+r[3].toFixed(1)+' GWh, temperature '+r[4].toFixed(1)+' C';};"
    Ajoute h, "g.appendChild(ci);"
    Ajoute h, "var t=el('text',{x:cx,y:cy+rad+13,'text-anchor':'middle','font-size':11,'font-weight':'bold',fill:'#222',stroke:'#fff','stroke-width':3,'paint-order':'stroke','pointer-events':'none'});"
    Ajoute h, "t.textContent=r[0];g.appendChild(t);});"
    Ajoute h, "document.getElementById('tot').innerText='Total affiche : '+tot.toFixed(1)+' GWh';}"
    Ajoute h, "selV.onchange=draw;draw();</script></body></html>"

    CreerDossier SaveFolder()
    chemin = SaveFolder() & FILE_TITLE & "_Carte_" & Format(Now, "yyyymmdd_hhnnss") & "_" & FILE_SUFFIX & ".html"
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
' 4. SYNTHESE (feuille Carte_V3)
'==============================================================
Private Sub CreerSynthese(cheminCarte As String, cheminCSV As String)
    Dim wsV As Worksheet, lo As ListObject, i As Long, nb As Long, ligne As Long

    Set wsV = GetOrCreateSheet(SHEET_VISUEL)
    Set lo = ThisWorkbook.Worksheets(SHEET_DATA).ListObjects(TABLE_NAME)
    nb = lo.ListRows.Count

    Do While wsV.ChartObjects.Count > 0: wsV.ChartObjects(1).Delete: Loop
    Do While wsV.PivotTables.Count > 0: wsV.PivotTables(1).TableRange2.Clear: Loop
    wsV.Hyperlinks.Delete
    wsV.Cells.Clear

    With wsV.Range("A1")
        .Value = "Suivi distribution gaz (V3 CSV) - mis a jour le " & Format(Now, "dd/mm/yyyy hh:nn")
        .Font.Bold = True: .Font.Size = 14
    End With

    wsV.Range("A2").Value = "Volume total (GWh)"
    wsV.Range("B2").Formula = "=SUM(" & TABLE_NAME & "[Volume_GWh])"
    wsV.Range("D2").Value = "Moyenne / ville (GWh)"
    wsV.Range("E2").Formula = "=AVERAGE(" & TABLE_NAME & "[Volume_GWh])"
    wsV.Range("G2").Value = "Ville du pic"
    wsV.Range("H2").Formula = "=INDEX(" & TABLE_NAME & "[Ville],MATCH(MAX(" & TABLE_NAME & "[Volume_GWh])," & TABLE_NAME & "[Volume_GWh],0))"
    wsV.Range("A2,D2,G2").Font.Bold = True
    wsV.Range("B2,E2").NumberFormat = "#,##0.0"

    wsV.Range("A3").Value = "Date des donnees"
    wsV.Range("B3").Formula = "=MAX(" & TABLE_NAME & "[Date])"
    wsV.Range("B3").NumberFormat = "dd/mm/yyyy"

    wsV.Range("A4:C4").Value = Array("Ville", "Volume (GWh)", "Temperature (C)")
    wsV.Range("A4:C4").Font.Bold = True
    wsV.Range("A4:C4").Interior.Color = RGB(221, 235, 247)

    ligne = 5
    For i = 1 To nb
        wsV.Cells(ligne, 1).Formula = "=INDEX(" & TABLE_NAME & "[Ville]," & i & ")"
        wsV.Cells(ligne, 2).Formula = "=INDEX(" & TABLE_NAME & "[Volume_GWh]," & i & ")"
        wsV.Cells(ligne, 3).Formula = "=INDEX(" & TABLE_NAME & "[Temperature_C]," & i & ")"
        ligne = ligne + 1
    Next i
    wsV.Range(wsV.Cells(5, 2), wsV.Cells(ligne - 1, 3)).NumberFormat = "#,##0.0"
    wsV.Range(wsV.Cells(5, 2), wsV.Cells(ligne - 1, 2)).FormatConditions.AddDatabar

    wsV.Hyperlinks.Add Anchor:=wsV.Cells(ligne + 1, 1), Address:=cheminCarte, TextToDisplay:="Ouvrir la carte interactive"
    wsV.Cells(ligne + 2, 1).Value = "Source CSV : " & cheminCSV
    wsV.Cells(ligne + 3, 1).Value = "Carte : " & cheminCarte
    wsV.Range(wsV.Cells(ligne + 2, 1), wsV.Cells(ligne + 3, 1)).Font.Italic = True

    wsV.Columns("A:H").AutoFit
    wsV.Columns("A").ColumnWidth = 24
    wsV.Activate
    LogMsg "INFO", "Synthese regeneree"
End Sub

'==============================================================
' 5. SAUVEGARDE : titre_date_VBATF_V3
'==============================================================
Private Function SauvegarderVersion() As String
    Dim chemin As String
    CreerDossier SaveFolder()
    chemin = SaveFolder() & FILE_TITLE & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & FILE_SUFFIX & ".xlsm"
    ThisWorkbook.SaveCopyAs chemin
    LogMsg "INFO", "Sauvegarde : " & chemin
    SauvegarderVersion = chemin
End Function

'==============================================================
' 6. MAIL (Outlook, late binding) : xlsm + carte + CSV source
'==============================================================
Private Sub EnvoyerMail(cheminXlsm As String, cheminCarte As String, cheminCSV As String)
    Dim ol As Object, mail As Object
    On Error GoTo MailErr
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    With mail
        .To = MAIL_TO
        .CC = MAIL_CC
        .Subject = "[PHASE DE TEST V3] Mise a jour suivi distribution - " & Format(Date, "dd/mm/yyyy")
        .Body = "Bonjour," & vbCrLf & vbCrLf & _
                "Ceci est un envoi automatique en phase de test (version CSV + carte interactive)." & vbCrLf & _
                "Les donnees ont ete importees depuis : " & Mid$(cheminCSV, InStrRev(cheminCSV, "\") + 1) & vbCrLf & _
                "Classeur : " & Mid$(cheminXlsm, InStrRev(cheminXlsm, "\") + 1) & vbCrLf & _
                "Carte : " & Mid$(cheminCarte, InStrRev(cheminCarte, "\") + 1) & _
                " (a ouvrir dans un navigateur, fonctionne hors ligne)" & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & "Robot VBATF"
        .Attachments.Add cheminXlsm
        .Attachments.Add cheminCarte
        .Attachments.Add cheminCSV
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
' Cree le dossier et tous les dossiers parents manquants
Private Sub CreerDossier(chemin As String)
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

