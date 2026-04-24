Attribute VB_Name = "zDocImportPowerQuery"
Option Explicit

' ============================================================
' IMPORT EXCEL PQ -> BDD-DOC
' Source : fichier Excel Power Query final
' Cible  : ThisWorkbook (BDD-DOC)
' ============================================================

Private Const NOM_CLASSEUR_SOURCE As String = "comparaison-PowerQuerry-24-04"   ' à adapter au nom réel sans extension
Private Const NOM_ONGLET_SOURCE As String = "REF-RF"
Private Const NOM_FEUILLE_ARCHIVE As String = "ID_supprimes_conformes"

' ============================================================
' MACRO PRINCIPALE
' ============================================================
Public Sub ImporterDepuisFichierPowerQuery()

    Dim wbSource As Workbook
    Dim wbCible As Workbook
    Dim wsSource As Worksheet
    Dim wsCible As Worksheet
    Dim wsArchive As Worksheet

    Dim dictAncien As Object
    Dim dictNouveau As Object

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean

    On Error GoTo ErrHandler

    Set wbCible = ThisWorkbook
    Set wsCible = wbCible.Worksheets(SHEET_MAIN)
    Set wsArchive = GetOrCreateSheet(wbCible, NOM_FEUILLE_ARCHIVE)

    Set wbSource = GetWorkbookByBaseName(NOM_CLASSEUR_SOURCE)
    If wbSource Is Nothing Then
        Err.Raise vbObjectError + 2000, , _
            "Classeur source introuvable : " & NOM_CLASSEUR_SOURCE & vbCrLf & _
            "Ouvre d'abord le fichier Excel Power Query final."
    End If

    Set wsSource = GetWorksheetSafe(wbSource, NOM_ONGLET_SOURCE)
    If wsSource Is Nothing Then
        Err.Raise vbObjectError + 2001, , _
            "Onglet source introuvable : " & NOM_ONGLET_SOURCE
    End If

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set dictAncien = ChargerAncienneBase(wsCible)
    Set dictNouveau = ChargerIDsDepuisFeuille(wsSource, 2)

    ArchiverLignesDisparuesAvecConformite wsArchive, dictAncien, dictNouveau
    RemplacerBaseDepuisSource wsCible, wsSource
    ReinjecterColonnesSuivi wsCible, dictAncien

    MsgBox "Import depuis le fichier Power Query terminé.", vbInformation

SortiePropre:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Exit Sub

ErrHandler:
    MsgBox "Erreur ImporterDepuisFichierPowerQuery : " & Err.description, vbCritical
    Resume SortiePropre

End Sub

' ============================================================
' CHARGE L'ANCIENNE BASE
' dict(ID) = Array(Date, Nom, Conf, Obs, LigneAAB)
' ============================================================
Private Function ChargerAncienneBase(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim idxID As Long, idxDate As Long, idxNom As Long, idxConf As Long, idxObs As Long
    Dim idVal As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    idxID = ColNum(COL_ID)
    idxDate = ColNum(COL_DATE)
    idxNom = ColNum(COL_NOM)
    idxConf = ColNum(COL_CONF)
    idxObs = ColNum(COL_OBS)

    lastRow = ws.Cells(ws.Rows.Count, idxID).End(xlUp).Row
    If lastRow < ROW_START Then
        Set ChargerAncienneBase = dict
        Exit Function
    End If

    arr = ws.Range(ws.Cells(ROW_START, 1), ws.Cells(lastRow, NB_COL_TABLE)).Value2

    For i = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(i, idxID)))
        If Len(idVal) > 0 Then
            dict(idVal) = Array( _
                arr(i, idxDate), _
                arr(i, idxNom), _
                arr(i, idxConf), _
                arr(i, idxObs), _
                ExtraireLigne(arr, i, NB_COL_TABLE))
        End If
    Next i

    Set ChargerAncienneBase = dict

End Function

' ============================================================
' CHARGE LES ID DEPUIS LA FEUILLE SOURCE
' ============================================================
Private Function ChargerIDsDepuisFeuille(ByVal ws As Worksheet, ByVal firstDataRow As Long) As Object

    Dim dict As Object
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim idxID As Long
    Dim idVal As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    idxID = ColNum(COL_ID)

    lastRow = ws.Cells(ws.Rows.Count, idxID).End(xlUp).Row
    If lastRow < firstDataRow Then
        Set ChargerIDsDepuisFeuille = dict
        Exit Function
    End If

    arr = ws.Range(ws.Cells(firstDataRow, idxID), ws.Cells(lastRow, idxID)).Value2

    For i = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(i, 1)))
        If Len(idVal) > 0 Then dict(idVal) = True
    Next i

    Set ChargerIDsDepuisFeuille = dict

End Function

' ============================================================
' ARCHIVE DES LIGNES DISPARUES AVEC CONFORMITE
' ============================================================
Private Sub ArchiverLignesDisparuesAvecConformite(ByVal wsArchive As Worksheet, ByVal dictAncien As Object, ByVal dictNouveau As Object)

    Dim k As Variant
    Dim info As Variant
    Dim confVal As String
    Dim nextRow As Long

    If wsArchive.Cells(1, 1).Value = "" Then
        ThisWorkbook.Worksheets(SHEET_MAIN).Rows(ROW_HEADER).Copy Destination:=wsArchive.Rows(1)
    End If

    nextRow = wsArchive.Cells(wsArchive.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    For Each k In dictAncien.Keys
        If Not dictNouveau.Exists(CStr(k)) Then
            info = dictAncien(k)
            confVal = Trim$(CStr(info(2)))

            If Len(confVal) > 0 Then
                wsArchive.Range(wsArchive.Cells(nextRow, 1), wsArchive.Cells(nextRow, NB_COL_TABLE)).Value = info(4)
                nextRow = nextRow + 1
            End If
        End If
    Next k

End Sub

' ============================================================
' REMPLACE BASE PAR LA FEUILLE SOURCE
' ============================================================
Private Sub RemplacerBaseDepuisSource(ByVal wsCible As Worksheet, ByVal wsSource As Worksheet)

    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim arr As Variant

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastColSource = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRowSource < 2 Then
        Err.Raise vbObjectError + 2002, , "La feuille source est vide."
    End If

    If lastColSource < NB_COL_TABLE Then
        Err.Raise vbObjectError + 2003, , _
            "La feuille source n'a pas assez de colonnes. Attendu : " & NB_COL_TABLE
    End If

    wsCible.Range(wsCible.Cells(ROW_START, 1), wsCible.Cells(wsCible.Rows.Count, NB_COL_TABLE)).ClearContents

    arr = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, NB_COL_TABLE)).Value2
    wsCible.Cells(ROW_START, 1).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

' ============================================================
' REINJECTION Y:AB
' ============================================================
Private Sub ReinjecterColonnesSuivi(ByVal wsBase As Worksheet, ByVal dictAncien As Object)

    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim idxID As Long, idxDate As Long, idxNom As Long, idxConf As Long, idxObs As Long
    Dim idVal As String
    Dim info As Variant

    idxID = ColNum(COL_ID)
    idxDate = ColNum(COL_DATE)
    idxNom = ColNum(COL_NOM)
    idxConf = ColNum(COL_CONF)
    idxObs = ColNum(COL_OBS)

    lastRow = wsBase.Cells(wsBase.Rows.Count, idxID).End(xlUp).Row
    If lastRow < ROW_START Then Exit Sub

    arr = wsBase.Range(wsBase.Cells(ROW_START, 1), wsBase.Cells(lastRow, NB_COL_TABLE)).Value2

    For i = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(i, idxID)))
        If Len(idVal) > 0 Then
            If dictAncien.Exists(idVal) Then
                info = dictAncien(idVal)
                arr(i, idxDate) = info(0)
                arr(i, idxNom) = info(1)
                arr(i, idxConf) = info(2)
                arr(i, idxObs) = info(3)
            End If
        End If
    Next i

    wsBase.Range(wsBase.Cells(ROW_START, 1), wsBase.Cells(lastRow, NB_COL_TABLE)).Value = arr

End Sub

' ============================================================
' HELPERS
' ============================================================
Private Function GetWorkbookByBaseName(ByVal baseName As String) As Workbook

    Dim wb As Workbook
    Dim NomSansExtension As String

    For Each wb In Application.Workbooks
        NomSansExtension = wb.Name
        If InStrRev(NomSansExtension, ".") > 0 Then
            NomSansExtension = Left$(NomSansExtension, InStrRev(NomSansExtension, ".") - 1)
        End If

        If StrComp(NomSansExtension, baseName, vbTextCompare) = 0 Then
            Set GetWorkbookByBaseName = wb
            Exit Function
        End If
    Next wb

End Function

Private Function GetWorksheetSafe(ByVal wb As Workbook, ByVal nomOnglet As String) As Worksheet

    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(nomOnglet)
    On Error GoTo 0

End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal nomFeuille As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(nomFeuille)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = nomFeuille
    End If

End Function

Private Function ColNum(ByVal colLetter As String) As Long
    ColNum = ThisWorkbook.Worksheets(SHEET_MAIN).Range(colLetter & "1").Column
End Function

Private Function ExtraireLigne(ByRef arr As Variant, ByVal rowIndex As Long, ByVal nbCols As Long) As Variant

    Dim ligne() As Variant
    Dim j As Long

    ReDim ligne(1 To 1, 1 To nbCols)

    For j = 1 To nbCols
        ligne(1, j) = arr(rowIndex, j)
    Next j

    ExtraireLigne = ligne

End Function

