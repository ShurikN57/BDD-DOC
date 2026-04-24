Attribute VB_Name = "zDocOuvertureExcel"
Option Explicit

' =============================================
' Ouverture d'un document Excel et ciblage
' d'une ligne / cellule selon les métadonnées
'
' Contrat :
' - filePath  : chemin local ou file://...
' - sheetInfo : format attendu "XLS:NomOnglet"
' - lineNum   : numéro de ligne Excel (> 0)
' - searchText / fullText : texte à rechercher sur la ligne cible
'
' Comportement :
' - ouvre le classeur s'il n'est pas déjà ouvert
' - active l'onglet demandé
' - tente de trouver searchText puis fullText sur la ligne cible
' - sinon se positionne sur A{ligne}
' =============================================

Private Const XL_WINDOW_STATE_MAXIMIZE As Long = -4137
Private Const XL_FIND_LOOKIN_VALUES As Long = -4163
Private Const XL_LOOKAT_WHOLE As Long = 1

Public Sub OpenExcelAtSheetAndLine(ByVal filePath As String, ByVal sheetInfo As String, _
                                   ByVal lineNum As Variant, Optional ByVal searchText As String = "", _
                                   Optional ByVal fullText As String = "")

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim wb As Object
    Dim foundCell As Object
    Dim localPath As String
    Dim sheetName As String
    Dim targetLine As Long
    Dim oldEvents As Boolean
    Dim eventsStateSaved As Boolean

    On Error GoTo ErrHandler

    localPath = FileUrlToWindowsPath(filePath)

    If Trim$(localPath) = "" Then
        MsgBox "Chemin Excel vide ou invalide.", vbExclamation
        Exit Sub
    End If

    If Len(Dir(localPath)) = 0 Then
        MsgBox "Fichier Excel introuvable :" & vbCrLf & localPath, vbExclamation
        Exit Sub
    End If

    sheetName = ExtractSheetName(sheetInfo)
    If Len(sheetName) = 0 Then
        MsgBox "Nom d'onglet invalide en colonne K :" & vbCrLf & sheetInfo, vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(lineNum) Then
        MsgBox "Numéro de ligne Excel invalide en colonne L.", vbExclamation
        Exit Sub
    End If

    targetLine = CLng(lineNum)
    If targetLine <= 0 Then
        MsgBox "Le numéro de ligne Excel doit être supérieur à 0.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrHandler

    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
    End If

    xlApp.Visible = True
    xlApp.WindowState = XL_WINDOW_STATE_MAXIMIZE

    For Each wb In xlApp.Workbooks
        If StrComp(wb.FullName, localPath, vbTextCompare) = 0 Then
            Set xlWb = wb
            Exit For
        End If
    Next wb

    If xlWb Is Nothing Then
        oldEvents = xlApp.EnableEvents
        eventsStateSaved = True
        xlApp.EnableEvents = False
        Set xlWb = xlApp.Workbooks.Open(localPath)
        xlApp.EnableEvents = oldEvents
        eventsStateSaved = False
    End If

    On Error Resume Next
    Set xlWs = xlWb.Worksheets(sheetName)
    On Error GoTo ErrHandler

    If xlWs Is Nothing Then
        MsgBox "Onglet introuvable dans le fichier Excel :" & vbCrLf & sheetName, vbExclamation
        Exit Sub
    End If

    xlWs.Activate

    If targetLine > xlWs.Rows.Count Then
        MsgBox "Le numéro de ligne demandé dépasse la taille de la feuille.", vbExclamation
        Exit Sub
    End If

    If Trim$(searchText) <> "" Then
        Set foundCell = xlWs.Rows(targetLine).Find(What:=searchText, LookIn:=XL_FIND_LOOKIN_VALUES, LookAt:=XL_LOOKAT_WHOLE)
    End If

    If foundCell Is Nothing And Trim$(fullText) <> "" Then
        Set foundCell = xlWs.Rows(targetLine).Find(What:=fullText, LookIn:=XL_FIND_LOOKIN_VALUES, LookAt:=XL_LOOKAT_WHOLE)
    End If

    If Not foundCell Is Nothing Then
        xlApp.GoTo foundCell, True
        foundCell.Select
    Else
        xlApp.GoTo xlWs.Range("A" & targetLine), True
        xlWs.Rows(targetLine).Select
    End If

    Exit Sub

ErrHandler:
    On Error Resume Next
    If eventsStateSaved Then xlApp.EnableEvents = oldEvents
    MsgBox "Erreur Excel : " & Err.description, vbExclamation

End Sub

Private Function ExtractSheetName(ByVal sheetInfo As String) As String

    Dim s As String

    s = Trim$(sheetInfo)

    If UCase$(Left$(s, 4)) = "XLS:" Then
        ExtractSheetName = Mid$(s, 5)
    Else
        ExtractSheetName = ""
    End If

End Function
