Attribute VB_Name = "zDocOuvertureWord"
Option Explicit

' =============================================
' Ouverture d'un document Word à proximité
' d'une page cible puis recherche/surlignage d'un texte
'
' Contrat :
' - filePath : chemin local ou file://...
' - pageNum  : numéro de page attendu (colonne J)
' - searchText : texte RF/REF à rechercher
'
' Comportement :
' - ouvre le document
' - calcule le nombre de pages
' - cherche sur la page cible puis autour : +1, -1, +2, -2 ... jusqu'à ±5
' =============================================

Private Const WD_WINDOW_STATE_MAXIMIZE As Long = 1
Private Const WD_STATISTIC_PAGES As Long = 2

Private Function RecupererInstanceWord(ByRef wdApp As Object) As Boolean

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        On Error GoTo Fin
        Set wdApp = CreateObject("Word.Application")
    End If

    RecupererInstanceWord = Not wdApp Is Nothing
    Exit Function

Fin:
    Err.Clear
    Set wdApp = Nothing

End Function

Public Sub OpenWordAtPageAndHighlight(ByVal filePath As String, ByVal pageNum As Variant, ByVal searchText As String)

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim localPath As String
    Dim pageTarget As Long
    Dim totalPages As Long
    Dim ok As Boolean

    On Error GoTo ErrHandler

    localPath = FileUrlToWindowsPath(filePath)

    If Trim$(localPath) = "" Then
        MsgBox "Chemin Word vide ou invalide.", vbExclamation
        Exit Sub
    End If

    If Len(Dir(localPath)) = 0 Then
        MsgBox "Fichier Word introuvable :" & vbCrLf & localPath, vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(pageNum) Then
        MsgBox "Le numéro de page en colonne J n'est pas valide.", vbExclamation
        Exit Sub
    End If

    pageTarget = CLng(pageNum)
    If pageTarget <= 0 Then
        MsgBox "Le numéro de page doit être supérieur à 0.", vbExclamation
        Exit Sub
    End If

    If Not RecupererInstanceWord(wdApp) Then
        MsgBox "Impossible d'ouvrir Microsoft Word sur ce poste.", vbExclamation
        Exit Sub
    End If

    wdApp.Visible = True
    wdApp.WindowState = WD_WINDOW_STATE_MAXIMIZE

    Set wdDoc = wdApp.Documents.Open(localPath)

    totalPages = wdDoc.ComputeStatistics(WD_STATISTIC_PAGES)
    If totalPages <= 0 Then
        MsgBox "Impossible de déterminer le nombre de pages du document Word.", vbExclamation
        Exit Sub
    End If

    If pageTarget > totalPages Then pageTarget = totalPages

    ok = HighlightTextOnSpecificPage(wdDoc, pageTarget, searchText, totalPages)

    If Not ok Then
        MsgBox "Texte non trouvé autour de la page " & pageTarget & " (recherche ±5) :" & vbCrLf & searchText, vbInformation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'ouverture/recherche Word :" & vbCrLf & Err.description, vbExclamation

End Sub

' =============================================
' Recherche sur la page cible puis à proximité :
' +1, -1, +2, -2 ... jusqu'à ±5
' =============================================
Private Function HighlightTextOnSpecificPage(ByVal wdDoc As Object, ByVal pageNum As Long, _
                                             ByVal rawSearch As String, ByVal totalPages As Long) As Boolean

    Dim pageRange As Object
    Dim p As Long
    Dim order(1 To 11) As Long
    Dim i As Long

    order(1) = pageNum
    order(2) = pageNum + 1
    order(3) = pageNum - 1
    order(4) = pageNum + 2
    order(5) = pageNum - 2
    order(6) = pageNum + 3
    order(7) = pageNum - 3
    order(8) = pageNum + 4
    order(9) = pageNum - 4
    order(10) = pageNum + 5
    order(11) = pageNum - 5

    For i = 1 To 11
        p = order(i)

        If p >= 1 And p <= totalPages Then
            Set pageRange = GetRangeOfPage(wdDoc, p, totalPages)

            If Not pageRange Is Nothing Then
                If HighlightOnRange(pageRange, rawSearch) Then
                    HighlightTextOnSpecificPage = True
                    Exit Function
                End If
            End If
        End If
    Next i

    HighlightTextOnSpecificPage = False

End Function

' =============================================
' Surlignage via Find Word :
' 1) motif souple avec espaces optionnels
' 2) texte brut exact
' =============================================
Private Function HighlightOnRange(ByVal pageRange As Object, ByVal searchText As String) As Boolean

    Dim oFind As Object
    Dim pattern As String

    If Trim$(searchText) = "" Then Exit Function

    pattern = BuildSearchPattern(searchText)
    Set oFind = pageRange.Find

    With oFind
        .ClearFormatting
        .Text = pattern
        .MatchCase = False
        .MatchWildcards = True
        .Forward = True
        .Wrap = 0

        If .Execute Then
            pageRange.HighlightColorIndex = 7
            pageRange.Select
            HighlightOnRange = True
            Exit Function
        End If
    End With

    With oFind
        .ClearFormatting
        .Text = searchText
        .MatchCase = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = 0

        If .Execute Then
            pageRange.HighlightColorIndex = 7
            pageRange.Select
            HighlightOnRange = True
            Exit Function
        End If
    End With

End Function

' =============================================
' Construit un motif de recherche souple
' pour RF/REF concaténées avec espaces variables
' =============================================
Public Function BuildSearchPattern(ByVal ref As String) As String

    Dim bloc1 As String
    Dim bloc2 As String
    Dim bloc3 As String
    Dim bloc4 As String
    Dim lenRef As Long
    Dim lenBloc3 As Long
    Dim lenBloc4 As Long
    Dim last3 As String

    lenRef = Len(ref)

    If lenRef < 9 Or lenRef > 12 Then
        BuildSearchPattern = ref
        Exit Function
    End If

    bloc1 = Left$(ref, 1)
    bloc2 = Mid$(ref, 2, 3)

    last3 = Right$(ref, 3)
    If Left$(last3, 2) Like "[A-Za-z][A-Za-z]" Then
        lenBloc4 = 3
    Else
        lenBloc4 = 2
    End If

    lenBloc3 = lenRef - 1 - 3 - lenBloc4

    If lenBloc3 < 3 Or lenBloc3 > 5 Then
        BuildSearchPattern = ref
        Exit Function
    End If

    bloc3 = Mid$(ref, 5, lenBloc3)
    bloc4 = Right$(ref, lenBloc4)

    BuildSearchPattern = bloc1 & "[ ]*" & bloc2 & "[ ]*" & bloc3 & "[ ]*" & bloc4

End Function

' =============================================
' Récupère la plage correspondant à une page Word
' =============================================
Private Function GetRangeOfPage(ByVal wdDoc As Object, ByVal pageNum As Long, _
                                ByVal totalPages As Long) As Object

    Dim rStart As Object
    Dim rEnd As Object

    On Error GoTo ErrHandler

    If pageNum < 1 Or pageNum > totalPages Then Exit Function

    Set rStart = wdDoc.GoTo(What:=1, Which:=1, Count:=pageNum)

    If pageNum < totalPages Then
        Set rEnd = wdDoc.GoTo(What:=1, Which:=1, Count:=pageNum + 1)
        Set GetRangeOfPage = wdDoc.Range(Start:=rStart.Start, End:=rEnd.Start - 1)
    Else
        Set GetRangeOfPage = wdDoc.Range(Start:=rStart.Start, End:=wdDoc.Content.End)
    End If

    Exit Function

ErrHandler:
    Set GetRangeOfPage = Nothing

End Function
