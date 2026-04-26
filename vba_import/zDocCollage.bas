Attribute VB_Name = "zDocCollage"
Option Explicit

Public CollageValeursEnCours As Boolean
Private DerniereCopieValeurUniqueValide As Boolean
Private DerniereCopieValeurUnique As Variant
Private DerniereCopieNbRows As Long
Private DerniereCopieNbCols As Long

' =============================================
' 0.1 ActiverCollageValeursRecherche
' =============================================
Public Sub ActiverCollageValeursRecherche()
    Application.OnKey "^c", "'" & ThisWorkbook.Name & "'!CopierValeursRecherche"
    Application.OnKey "^v", "'" & ThisWorkbook.Name & "'!CollerValeursRecherche"
End Sub

' =============================================
' 0.2 DesactiverCollageValeursRecherche
' =============================================
Public Sub DesactiverCollageValeursRecherche()
    Application.OnKey "^c"
    Application.OnKey "^v"
End Sub

' =============================================
' 1. CollerValeursRecherche
' =============================================
Public Sub CollerValeursRecherche()

    Dim ws As Worksheet
    Dim wsTitres As Worksheet
    Dim cible As Range
    Dim zoneAutorisee As Range
    Dim zoneVisible As Range
    Dim cibleFinale As Range
    Dim cibleReelleCollage As Range
    Dim areaP As Range
    Dim zoneRecherche As Range
    Dim zoneConf As Range
    Dim cellR As Range
    Dim cellV As Range
    Dim titre As String
    Dim vCheck As String
    Dim lastRow As Long
    Dim bValeurUnique As Boolean
    Dim valeurUnique As Variant
    Dim prevEnableEvents As Boolean

    On Error GoTo FinAvecErreur

    prevEnableEvents = Application.EnableEvents

    If ModeDeveloppeurActif Then
        Application.CommandBars.ExecuteMso "Paste"
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        Application.CommandBars.ExecuteMso "Paste"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    If Not ws Is ActiveSheet Then
        Application.CommandBars.ExecuteMso "Paste"
        Exit Sub
    End If

    lastRow = DerniereLigneUtileMain()
    Set zoneAutorisee = ConstruireZoneAutoriseeCollage(ws, lastRow)

    If zoneAutorisee Is Nothing Then GoTo Fin

    If Selection.Cells.CountLarge = 1 Then
        Set cible = ActiveCell
    Else
        Set cible = Selection
    End If

    If Not PlageEntierementAutorisee(cible, zoneAutorisee) Then
        MsgBox "Collage interdit dans cette zone.", vbExclamation
        GoTo Fin
    End If

    If cible.Cells.CountLarge = 1 Then
        Set cibleFinale = cible
    Else
        If ws.FilterMode Then
            On Error Resume Next
            Set zoneVisible = cible.SpecialCells(xlCellTypeVisible)
            On Error GoTo FinAvecErreur

            If zoneVisible Is Nothing Then
                MsgBox "Aucune cellule visible sélectionnée.", vbExclamation
                GoTo Fin
            End If

            Set cibleFinale = zoneVisible
        Else
            Set cibleFinale = cible
        End If
    End If

    Set cibleReelleCollage = ConstruirePlageReelleCollage(ws, cibleFinale)

    If cibleReelleCollage Is Nothing Then
        MsgBox "Impossible de déterminer la taille réelle du collage. Recopiez puis recollez.", vbExclamation
        GoTo Fin
    End If

    If Not PlageEntierementAutorisee(cibleReelleCollage, zoneAutorisee) Then
        MsgBox "Collage interdit dans cette zone.", vbExclamation
        GoTo Fin
    End If

    If Application.CutCopyMode = 0 Then
        MsgBox "Le contenu copié a été perdu. Recopiez puis recollez.", vbExclamation
        GoTo Fin
    End If

    bValeurUnique = False

    ' ===== Cas filtre + sélection discontinue =====
    If cibleReelleCollage.Areas.Count > 1 Then

        If DerniereCopieValeurUniqueValide Then
            bValeurUnique = True
            valeurUnique = DerniereCopieValeurUnique
        Else
            valeurUnique = LireValeurUniqueCopieeDepuisPressePapiers(bValeurUnique)
        End If

        If Not bValeurUnique Then
            MsgBox "Sous filtre, seul le collage d'une valeur unique sur plusieurs lignes visibles est autorisé." & vbCrLf & vbCrLf & _
                   "Pour un collage multi-cellules, retirez le filtre ou collez sur une zone continue.", vbExclamation
            GoTo Fin
        End If
    End If

    SauvegarderEtat cibleReelleCollage
    CollageValeursEnCours = True

    If cibleReelleCollage.Areas.Count > 1 Then
        For Each areaP In cibleReelleCollage.Areas
            areaP.Value = valeurUnique
        Next areaP
    Else
        cibleReelleCollage.PasteSpecial Paste:=xlPasteValues
    End If

    CollageValeursEnCours = False
    Application.CutCopyMode = False
    NettoyerBordureSelectionApresCollage ws, cibleReelleCollage

    Set zoneRecherche = Intersect(cibleReelleCollage, ws.Range(PLAGE_RECHERCHE))

    If Not zoneRecherche Is Nothing Then

        Set wsTitres = ThisWorkbook.Worksheets(SHEET_TITRES)
        ws.Range(PLAGE_RECHERCHE).Interior.Color = COLOR_RECHERCHE_FOND

        For Each cellR In zoneRecherche.Cells

            titre = CStr(wsTitres.Cells(ROW_TITRES, cellR.Column).Value)

            If Trim$(CStr(cellR.Value)) = "" Then
                cellR.Value = titre
                cellR.Font.Color = COLOR_PLACEHOLDER
                cellR.Font.Bold = False

            ElseIf CStr(cellR.Value) = titre Then
                cellR.Font.Color = COLOR_PLACEHOLDER
                cellR.Font.Bold = False

            Else
                cellR.Font.Color = COLOR_TEXTE_NOIR
                cellR.Font.Bold = True
            End If

        Next cellR

        If cibleReelleCollage.Cells.CountLarge = 1 Then
            If Not Intersect(cibleReelleCollage.Cells(1, 1), ws.Range(PLAGE_RECHERCHE)) Is Nothing Then
                cibleReelleCollage.Cells(1, 1).Interior.Color = COLOR_RECHERCHE_ACTIVE
            End If
        End If

    End If

    Set zoneConf = Intersect(cibleReelleCollage, ws.Columns(COL_CONF))

    If Not zoneConf Is Nothing Then
        For Each cellV In zoneConf.Cells
            If cellV.Row >= ROW_START Then
                vCheck = LCase$(Trim$(CStr(cellV.Value)))

                If vCheck <> "" _
                   And vCheck <> LCase$(VAL_CONF_1) _
                   And vCheck <> LCase$(VAL_CONF_2) _
                   And vCheck <> LCase$(VAL_CONF_3) Then

                    AnnulerDerniereAction
                    MsgBox "Valeur non autorisée en colonne " & COL_CONF & "." & vbCrLf & MSG_VALEURS_CONF, vbExclamation
                    GoTo Fin
                End If
            End If
        Next cellV

        Application.Run "'" & ThisWorkbook.Name & "'!" & ws.CodeName & ".RafraichirCouleursConformiteSurLignes", zoneConf.Address
    End If

Fin:
    CollageValeursEnCours = False
    Application.CutCopyMode = False
    If Not ws Is Nothing Then NettoyerBordureSelectionApresCollage ws, cibleReelleCollage
    Application.EnableEvents = prevEnableEvents
    Exit Sub

FinAvecErreur:
    CollageValeursEnCours = False
    Application.CutCopyMode = False
    If Not ws Is Nothing Then NettoyerBordureSelectionApresCollage ws, cibleReelleCollage
    Application.EnableEvents = prevEnableEvents
    MsgBox "Erreur lors du collage : " & Err.description, vbExclamation

End Sub

' =============================================
' 1-bis. NettoyerBordureSelectionApresCollage
' =============================================
Private Sub NettoyerBordureSelectionApresCollage(ByVal ws As Worksheet, ByVal rngCollee As Range)

    Dim area As Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rowMin As Long
    Dim rowMax As Long
    Dim lig As Long
    Dim rngLigne As Range

    On Error GoTo Fin

    If ws Is Nothing Then Exit Sub
    If rngCollee Is Nothing Then Exit Sub
    If Not ws Is ActiveSheet Then Exit Sub

    rowMin = 0
    rowMax = 0

    For Each area In rngCollee.Areas
        firstRow = area.Row
        lastRow = area.Row + area.Rows.Count - 1

        If rowMin = 0 Or firstRow < rowMin Then rowMin = firstRow
        If rowMax = 0 Or lastRow > rowMax Then rowMax = lastRow
    Next area

    If rowMin = 0 Or rowMax = 0 Then Exit Sub

    If rowMin > ROW_START Then rowMin = rowMin - 1
    rowMax = rowMax + 1
    If rowMax > ws.Rows.Count Then rowMax = ws.Rows.Count

    Application.ScreenUpdating = False

    For lig = rowMin To rowMax
        Set rngLigne = ws.Range(ws.Cells(lig, 1), ws.Cells(lig, NB_COL_UI))

        With rngLigne
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = COLOR_BORDURE_BLEUE
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = COLOR_BORDURE_BLEUE
            End With
        End With
    Next lig

Fin:
    Application.ScreenUpdating = True

End Sub

' =============================================
' 2. ConstruireZoneAutoriseeCollage
' =============================================
Private Function ConstruireZoneAutoriseeCollage(ByVal ws As Worksheet, ByVal lastRow As Long) As Range

    Dim rng As Range
    Dim rngTemp As Range

    If PLAGE_COLLER_RECHERCHE <> "" Then
        Set rng = ws.Range(PLAGE_COLLER_RECHERCHE)
    End If

    If PLAGE_COLLER_EDITABLE <> "" Then
        Set rngTemp = Intersect(ws.Range(PLAGE_COLLER_EDITABLE), ws.Rows(ROW_START & ":" & lastRow))
        If Not rngTemp Is Nothing Then
            If rng Is Nothing Then
                Set rng = rngTemp
            Else
                Set rng = Union(rng, rngTemp)
            End If
        End If
    End If

    If PLAGE_COLLER_SUIVI <> "" Then
        Set rngTemp = Intersect(ws.Range(PLAGE_COLLER_SUIVI), ws.Rows(ROW_START & ":" & lastRow))
        If Not rngTemp Is Nothing Then
            If rng Is Nothing Then
                Set rng = rngTemp
            Else
                Set rng = Union(rng, rngTemp)
            End If
        End If
    End If

    Set ConstruireZoneAutoriseeCollage = rng

End Function

' =============================================
' 2-bis. PlageEntierementAutorisee
' =============================================
Private Function PlageEntierementAutorisee(ByVal rngTest As Range, ByVal zoneAutorisee As Range) As Boolean

    Dim rngInter As Range

    If rngTest Is Nothing Then Exit Function
    If zoneAutorisee Is Nothing Then Exit Function

    Set rngInter = Intersect(rngTest, zoneAutorisee)
    If rngInter Is Nothing Then Exit Function

    PlageEntierementAutorisee = (rngInter.CountLarge = rngTest.CountLarge)

End Function

' =============================================
' 3. LireValeurUniqueCopieeDepuisPressePapiers
' =============================================
Private Function LireValeurUniqueCopieeDepuisPressePapiers(ByRef bValeurUnique As Boolean) As Variant

    Dim dataObj As Object
    Dim txt As String

    bValeurUnique = False
    LireValeurUniqueCopieeDepuisPressePapiers = Empty

    On Error GoTo Fin

    Set dataObj = CreateObject("Forms.DataObject")
    dataObj.GetFromClipboard
    txt = dataObj.GetText

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Do While Len(txt) > 0 And Right$(txt, 1) = vbLf
        txt = Left$(txt, Len(txt) - 1)
    Loop

    If InStr(txt, vbTab) = 0 And InStr(txt, vbLf) = 0 Then
        bValeurUnique = True
        LireValeurUniqueCopieeDepuisPressePapiers = txt
    End If

Fin:

End Function

' =============================================
' 4. CopierValeursRecherche
' =============================================
Public Sub CopierValeursRecherche()

    On Error GoTo Fin

    DerniereCopieValeurUniqueValide = False
    DerniereCopieValeurUnique = Empty
    DerniereCopieNbRows = 0
    DerniereCopieNbCols = 0

    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            DerniereCopieNbRows = Selection.Rows.Count
            DerniereCopieNbCols = Selection.Columns.Count

            If Selection.Cells.CountLarge = 1 Then
                DerniereCopieValeurUniqueValide = True
                DerniereCopieValeurUnique = Selection.Cells(1, 1).Value
            End If
        End If
    End If

    Application.CommandBars.ExecuteMso "Copy"

Fin:

End Sub

' =============================================
' 4-bis. LireDimensionsCopieesDepuisPressePapiers
' =============================================
Private Sub LireDimensionsCopieesDepuisPressePapiers(ByRef nbRows As Long, ByRef nbCols As Long)

    Dim dataObj As Object
    Dim txt As String
    Dim lignes() As String
    Dim cellules() As String
    Dim i As Long
    Dim maxCols As Long

    nbRows = 0
    nbCols = 0

    On Error GoTo Fin

    Set dataObj = CreateObject("Forms.DataObject")
    dataObj.GetFromClipboard
    txt = dataObj.GetText

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Do While Len(txt) > 0 And Right$(txt, 1) = vbLf
        txt = Left$(txt, Len(txt) - 1)
    Loop

    If Len(txt) = 0 Then GoTo Fin

    lignes = Split(txt, vbLf)
    nbRows = UBound(lignes) - LBound(lignes) + 1

    For i = LBound(lignes) To UBound(lignes)
        cellules = Split(lignes(i), vbTab)
        If UBound(cellules) - LBound(cellules) + 1 > maxCols Then
            maxCols = UBound(cellules) - LBound(cellules) + 1
        End If
    Next i

    nbCols = maxCols

Fin:

End Sub

' =============================================
' 5. ConstruirePlageReelleCollage
' =============================================
Private Function ConstruirePlageReelleCollage(ByVal ws As Worksheet, ByVal cibleFinale As Range) As Range

    Dim rngResult As Range
    Dim rngTopLeft As Range
    Dim nbRows As Long
    Dim nbCols As Long

    If cibleFinale Is Nothing Then Exit Function

    If ws.FilterMode Or cibleFinale.Areas.Count > 1 Then
        Set ConstruirePlageReelleCollage = cibleFinale
        Exit Function
    End If

    nbRows = DerniereCopieNbRows
    nbCols = DerniereCopieNbCols

    If nbRows <= 0 Or nbCols <= 0 Then
        LireDimensionsCopieesDepuisPressePapiers nbRows, nbCols
    End If

    If nbRows <= 0 Or nbCols <= 0 Then Exit Function

    Set rngTopLeft = cibleFinale.Cells(1, 1)

    If cibleFinale.Cells.CountLarge = 1 Then
        Set rngResult = rngTopLeft.Resize(nbRows, nbCols)

    ElseIf cibleFinale.Columns.Count = 1 And nbCols > 1 Then
        Set rngResult = cibleFinale.Resize(cibleFinale.Rows.Count, nbCols)

    ElseIf cibleFinale.Rows.Count = 1 And nbRows > 1 Then
        Set rngResult = cibleFinale.Resize(nbRows, cibleFinale.Columns.Count)

    Else
        Set rngResult = cibleFinale
    End If

    Set ConstruirePlageReelleCollage = rngResult

End Function




