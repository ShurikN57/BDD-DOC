Attribute VB_Name = "AuditPerformance"
Option Explicit

' ============================================================
' AUDIT PERFORMANCE CLASSEUR
' ============================================================
' Analyse :
' - UsedRange
' - Nb cellules avec formule
' - Nb cellules avec formule volatile potentielle
' - Nb règles de MFC
' - Nb formes
' - Nb hyperliens
' - Nb validations de données
' - Nb commentaires / notes
' - Nb cellules fusionnées
' - Nb objets OLE / contrôles
' - Nb noms définis
'
' Sortie :
' - feuille AUDIT_PERF recréée à chaque lancement
' ============================================================

Private Const SHEET_AUDIT As String = "AUDIT_PERF"

Public Sub AuditPerformanceClasseur()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAudit As Worksheet
    Dim nextRow As Long

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevDisplayAlerts As Boolean

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    SupprimerFeuilleAuditSiExiste wb
    Set wsAudit = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsAudit.Name = SHEET_AUDIT

    PreparerFeuilleAudit wsAudit
    nextRow = 2

    For Each ws In wb.Worksheets
        If ws.Name <> SHEET_AUDIT Then
            AnalyserFeuille ws, wsAudit, nextRow
            nextRow = nextRow + 1
        End If
    Next ws

    EcrireBlocSynthese wb, wsAudit, nextRow + 2
    MettreEnFormeAudit wsAudit

SortiePropre:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'audit performance : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' ============================================================
' Analyse d'une feuille
' ============================================================
Private Sub AnalyserFeuille(ByVal ws As Worksheet, ByVal wsAudit As Worksheet, ByVal outRow As Long)

    Dim ur As Range
    Dim nbRows As Long
    Dim nbCols As Long
    Dim nbCells As Double

    Dim nbFormules As Double
    Dim nbFormulesVolatiles As Double
    Dim nbMFC As Long
    Dim nbShapes As Long
    Dim nbHyperlinks As Long
    Dim nbValidations As Double
    Dim nbCommentaires As Double
    Dim nbFusions As Double
    Dim nbOLE As Long
    Dim score As Long
    Dim usedAddr As String

    Set ur = SafeUsedRange(ws)

    If ur Is Nothing Then
        nbRows = 0
        nbCols = 0
        nbCells = 0
        usedAddr = ""
    Else
        nbRows = ur.Rows.Count
        nbCols = ur.Columns.Count
        nbCells = CDbl(nbRows) * CDbl(nbCols)
        usedAddr = ur.Address(False, False)
    End If

    nbFormules = CompterFormules(ws)
    nbFormulesVolatiles = CompterFormulesVolatiles(ws)
    nbMFC = CompterMFC(ws)
    nbShapes = ws.Shapes.Count
    nbHyperlinks = ws.Hyperlinks.Count
    nbValidations = CompterValidations(ws)
    nbCommentaires = CompterCommentaires(ws)
    nbFusions = CompterFusions(ws)
    nbOLE = ws.OLEObjects.Count

    score = EvaluerScoreRalentissement(nbCells, nbFormules, nbFormulesVolatiles, nbMFC, nbShapes, nbHyperlinks, nbValidations, nbCommentaires, nbFusions, nbOLE)

    wsAudit.Cells(outRow, 1).Value = ws.Name
    wsAudit.Cells(outRow, 2).Value = usedAddr
    wsAudit.Cells(outRow, 3).Value = nbRows
    wsAudit.Cells(outRow, 4).Value = nbCols
    wsAudit.Cells(outRow, 5).Value = nbCells
    wsAudit.Cells(outRow, 6).Value = nbFormules
    wsAudit.Cells(outRow, 7).Value = nbFormulesVolatiles
    wsAudit.Cells(outRow, 8).Value = nbMFC
    wsAudit.Cells(outRow, 9).Value = nbShapes
    wsAudit.Cells(outRow, 10).Value = nbHyperlinks
    wsAudit.Cells(outRow, 11).Value = nbValidations
    wsAudit.Cells(outRow, 12).Value = nbCommentaires
    wsAudit.Cells(outRow, 13).Value = nbFusions
    wsAudit.Cells(outRow, 14).Value = nbOLE
    wsAudit.Cells(outRow, 15).Value = score
    wsAudit.Cells(outRow, 16).Value = DiagnosticFeuille(nbCells, nbFormules, nbFormulesVolatiles, nbMFC, nbShapes, nbHyperlinks, nbValidations, nbCommentaires, nbFusions, nbOLE)

End Sub

' ============================================================
' Feuille audit
' ============================================================
Private Sub PreparerFeuilleAudit(ByVal wsAudit As Worksheet)

    wsAudit.Cells(1, 1).Value = "Feuille"
    wsAudit.Cells(1, 2).Value = "UsedRange"
    wsAudit.Cells(1, 3).Value = "Nb lignes"
    wsAudit.Cells(1, 4).Value = "Nb colonnes"
    wsAudit.Cells(1, 5).Value = "Nb cellules"
    wsAudit.Cells(1, 6).Value = "Nb formules"
    wsAudit.Cells(1, 7).Value = "Nb formules volatiles"
    wsAudit.Cells(1, 8).Value = "Nb règles MFC"
    wsAudit.Cells(1, 9).Value = "Nb formes"
    wsAudit.Cells(1, 10).Value = "Nb hyperliens"
    wsAudit.Cells(1, 11).Value = "Nb validations"
    wsAudit.Cells(1, 12).Value = "Nb commentaires/notes"
    wsAudit.Cells(1, 13).Value = "Nb fusions"
    wsAudit.Cells(1, 14).Value = "Nb OLE/contrôles"
    wsAudit.Cells(1, 15).Value = "Score risque"
    wsAudit.Cells(1, 16).Value = "Diagnostic"

End Sub

Private Sub MettreEnFormeAudit(ByVal wsAudit As Worksheet)

    Dim lastRow As Long
    Dim rng As Range

    lastRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Set rng = wsAudit.Range("A1:P" & lastRow)

    With wsAudit.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    rng.Columns.AutoFit
    wsAudit.Activate
    wsAudit.Range("A1").Select
    ActiveWindow.FreezePanes = False
    wsAudit.Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Private Sub SupprimerFeuilleAuditSiExiste(ByVal wb As Workbook)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_AUDIT)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Delete
    End If

End Sub

' ============================================================
' Synthèse classeur
' ============================================================
Private Sub EcrireBlocSynthese(ByVal wb As Workbook, ByVal wsAudit As Worksheet, ByVal startRow As Long)

    Dim nbNoms As Long
    Dim nbLiensExternes As Long

    nbNoms = wb.Names.Count
    nbLiensExternes = CompterLiensExternes(wb)

    wsAudit.Cells(startRow, 1).Value = "SYNTHESE CLASSEUR"
    wsAudit.Cells(startRow, 1).Font.Bold = True
    wsAudit.Cells(startRow + 1, 1).Value = "Nb feuilles analysées"
    wsAudit.Cells(startRow + 1, 2).Value = wb.Worksheets.Count - 1

    wsAudit.Cells(startRow + 2, 1).Value = "Nb noms définis"
    wsAudit.Cells(startRow + 2, 2).Value = nbNoms

    wsAudit.Cells(startRow + 3, 1).Value = "Nb liens externes"
    wsAudit.Cells(startRow + 3, 2).Value = nbLiensExternes

End Sub

' ============================================================
' Compteurs
' ============================================================
Private Function SafeUsedRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set SafeUsedRange = ws.UsedRange
    On Error GoTo 0
End Function

Private Function CompterFormules(ByVal ws As Worksheet) As Double

    Dim rng As Range

    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If rng Is Nothing Then
        CompterFormules = 0
    Else
        CompterFormules = rng.CountLarge
    End If

End Function

Private Function CompterFormulesVolatiles(ByVal ws As Worksheet) As Double

    Dim rng As Range
    Dim area As Range
    Dim arr As Variant
    Dim i As Long
    Dim j As Long
    Dim f As String
    Dim total As Double

    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If rng Is Nothing Then Exit Function

    For Each area In rng.Areas
        arr = area.Formula

        If IsArray(arr) Then
            For i = 1 To UBound(arr, 1)
                For j = 1 To UBound(arr, 2)
                    f = UCase$(CStr(arr(i, j)))
                    If EstFormuleVolatile(f) Then total = total + 1
                Next j
            Next i
        Else
            f = UCase$(CStr(arr))
            If EstFormuleVolatile(f) Then total = total + 1
        End If
    Next area

    CompterFormulesVolatiles = total

End Function

Private Function EstFormuleVolatile(ByVal f As String) As Boolean

    If InStr(f, "INDIRECT(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "DECALER(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "OFFSET(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "INDIRECT(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "AUJOURDHUI(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "TODAY(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "MAINTENANT(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "NOW(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "ALEA(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "RAND(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "ALEA.ENTRE.BORNES(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "RANDBETWEEN(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "CELL(") > 0 Then EstFormuleVolatile = True: Exit Function
    If InStr(f, "INFO(") > 0 Then EstFormuleVolatile = True: Exit Function

End Function

Private Function CompterMFC(ByVal ws As Worksheet) As Long

    Dim ur As Range
    Dim c As Range
    Dim total As Long

    On Error GoTo Fin

    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    For Each c In ur.Cells
        total = total + c.FormatConditions.Count
    Next c

Fin:
    CompterMFC = total

End Function

Private Function CompterValidations(ByVal ws As Worksheet) As Double

    Dim ur As Range
    Dim c As Range
    Dim total As Double

    On Error GoTo Fin

    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    For Each c In ur.Cells
        On Error Resume Next
        If c.Validation.Type <> xlValidateInputOnly Then total = total + 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo Fin
    Next c

Fin:
    CompterValidations = total

End Function

Private Function CompterCommentaires(ByVal ws As Worksheet) As Double

    Dim ur As Range
    Dim c As Range
    Dim total As Double

    On Error GoTo Fin

    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    For Each c In ur.Cells
        On Error Resume Next
        If Not c.Comment Is Nothing Then total = total + 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo Fin

        On Error Resume Next
        If Not c.CommentThreaded Is Nothing Then total = total + 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo Fin
    Next c

Fin:
    CompterCommentaires = total

End Function

Private Function CompterFusions(ByVal ws As Worksheet) As Double

    Dim ur As Range
    Dim c As Range
    Dim total As Double

    On Error GoTo Fin

    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    For Each c In ur.Cells
        If c.MergeCells Then total = total + 1
    Next c

Fin:
    CompterFusions = total

End Function

Private Function CompterLiensExternes(ByVal wb As Workbook) As Long

    Dim arr As Variant

    On Error Resume Next
    arr = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0

    If IsEmpty(arr) Then
        CompterLiensExternes = 0
    Else
        CompterLiensExternes = UBound(arr) - LBound(arr) + 1
    End If

End Function

' ============================================================
' Score / diagnostic
' ============================================================
Private Function EvaluerScoreRalentissement(ByVal nbCells As Double, _
                                            ByVal nbFormules As Double, _
                                            ByVal nbVolatiles As Double, _
                                            ByVal nbMFC As Long, _
                                            ByVal nbShapes As Long, _
                                            ByVal nbHyperlinks As Long, _
                                            ByVal nbValidations As Double, _
                                            ByVal nbCommentaires As Double, _
                                            ByVal nbFusions As Double, _
                                            ByVal nbOLE As Long) As Long

    Dim score As Long

    If nbCells > 100000 Then score = score + 2
    If nbCells > 500000 Then score = score + 3

    If nbFormules > 1000 Then score = score + 2
    If nbFormules > 10000 Then score = score + 3

    If nbVolatiles > 0 Then score = score + 3
    If nbVolatiles > 100 Then score = score + 3

    If nbMFC > 100 Then score = score + 2
    If nbMFC > 1000 Then score = score + 3

    If nbShapes > 20 Then score = score + 1
    If nbShapes > 100 Then score = score + 2

    If nbHyperlinks > 500 Then score = score + 1
    If nbValidations > 1000 Then score = score + 1
    If nbCommentaires > 100 Then score = score + 1
    If nbFusions > 100 Then score = score + 1
    If nbOLE > 0 Then score = score + 2

    EvaluerScoreRalentissement = score

End Function

Private Function DiagnosticFeuille(ByVal nbCells As Double, _
                                   ByVal nbFormules As Double, _
                                   ByVal nbVolatiles As Double, _
                                   ByVal nbMFC As Long, _
                                   ByVal nbShapes As Long, _
                                   ByVal nbHyperlinks As Long, _
                                   ByVal nbValidations As Double, _
                                   ByVal nbCommentaires As Double, _
                                   ByVal nbFusions As Double, _
                                   ByVal nbOLE As Long) As String

    Dim msg As String

    If nbCells > 500000 Then msg = msg & "UsedRange très large; "
    If nbFormules > 10000 Then msg = msg & "beaucoup de formules; "
    If nbVolatiles > 0 Then msg = msg & "formules volatiles; "
    If nbMFC > 1000 Then msg = msg & "beaucoup de MFC; "
    If nbShapes > 100 Then msg = msg & "beaucoup de formes; "
    If nbHyperlinks > 1000 Then msg = msg & "beaucoup d'hyperliens; "
    If nbValidations > 5000 Then msg = msg & "beaucoup de validations; "
    If nbCommentaires > 100 Then msg = msg & "beaucoup de commentaires; "
    If nbFusions > 100 Then msg = msg & "beaucoup de cellules fusionnées; "
    If nbOLE > 0 Then msg = msg & "objets OLE/contrôles présents; "

    If msg = "" Then
        msg = "RAS majeur"
    Else
        msg = Left$(msg, Len(msg) - 2)
    End If

    DiagnosticFeuille = msg

End Function




