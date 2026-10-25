Attribute VB_Name = "ExportVBA"
Option Explicit

' Exporte tous les modules VBA du classeur actif dans un dossier "<NomDuClasseur>_vba"
' créé à côté du classeur. Ce dossier peut ensuite être donné à vba2py :
'     python main.py traduire "C:\...\Ventes_vba" -o sortie
'
' Prérequis (une seule fois) : Fichier > Options > Centre de gestion de la confidentialité >
' Paramètres du Centre de gestion de la confidentialité > Paramètres des macros >
' cocher « Accès approuvé au modèle d'objet du projet VBA ».
'
' Conseil : placez ce module dans votre classeur de macros personnelles (PERSONAL.XLSB),
' activez le classeur à exporter, puis lancez ExporterModulesVBA (Alt+F8).

Public Sub ExporterModulesVBA()
    Dim wb As Workbook
    Dim composant As Object
    Dim dossier As String
    Dim extension As String
    Dim nombre As Long

    Set wb = ActiveWorkbook
    If wb.Path = "" Then
        MsgBox "Enregistrez d'abord le classeur.", vbExclamation
        Exit Sub
    End If

    dossier = wb.Path & Application.PathSeparator & _
              Left(wb.Name, InStrRev(wb.Name, ".") - 1) & "_vba"
    If Dir(dossier, vbDirectory) = "" Then MkDir dossier

    On Error GoTo Erreur
    For Each composant In wb.VBProject.VBComponents
        Select Case composant.Type
            Case 1: extension = ".bas"        ' module standard
            Case 2, 100: extension = ".cls"   ' module de classe, ThisWorkbook, feuilles
            Case 3: extension = ".frm"        ' UserForm (+ fichier .frx créé automatiquement)
            Case Else: extension = ""
        End Select
        If extension <> "" Then
            composant.Export dossier & Application.PathSeparator & composant.Name & extension
            nombre = nombre + 1
        End If
    Next composant

    MsgBox nombre & " module(s) exporté(s) dans :" & vbCrLf & dossier, vbInformation
    Exit Sub

Erreur:
    MsgBox "Export impossible : " & Err.Description & vbCrLf & vbCrLf & _
           "Vérifiez que « Accès approuvé au modèle d'objet du projet VBA » est activé.", vbCritical
End Sub
