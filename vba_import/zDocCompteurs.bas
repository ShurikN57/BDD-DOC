Attribute VB_Name = "zDocCompteurs"
' =============================================
' Remplace les 4 formules lourdes de MENU DEROULANT J1:M1
'
' J1 = nombre de RF uniques visibles (colonne A)
' K1 = nombre d'ID uniques visibles (colonne U)
' L1 = nombre de REF uniques visibles (colonne I)
' M1 = concatenation J1 & K1 & L1
' =============================================

Private m_nbRF_Total As Long
Private m_nbID_Total As Long
Private m_nbREF_Total As Long
Private m_Initialise As Boolean

' =============================================
' Calcul complet (ouverture + filtre)
' =============================================
Public Sub MettreAJourCompteurs()

    Dim ws As Worksheet
    Dim wsMenu As Worksheet
    Dim lastRow As Long
    Dim dictRF As Object
    Dim dictID As Object
    Dim dictREF As Object
    Dim nbRF As Long
    Dim nbID As Long
    Dim nbREF As Long
    Dim bFiltre As Boolean
    Dim rngVisibleRows As Range
    Dim area As Range
    Dim firstLig As Long
    Dim lastLigArea As Long
    Dim arrA As Variant
    Dim arrU As Variant
    Dim arrI As Variant

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set wsMenu = ThisWorkbook.Worksheets(SHEET_MENU_DEROULANT)

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastRow = ws.Cells(ws.Rows.Count, COL_RF).End(xlUp).Row
    If lastRow < ROW_START Then
        EcrireCompteurs wsMenu, 0, 0, 0
        GoTo SortiePropre
    End If

    Set dictRF = CreateObject("Scripting.Dictionary")
    Set dictID = CreateObject("Scripting.Dictionary")
    Set dictREF = CreateObject("Scripting.Dictionary")

    dictRF.CompareMode = vbTextCompare
    dictID.CompareMode = vbTextCompare
    dictREF.CompareMode = vbTextCompare

    bFiltre = ws.FilterMode

    If Not bFiltre Then
        arrA = ws.Range(COL_RF & ROW_START & ":" & COL_RF & lastRow).Value2
        arrU = ws.Range(COL_ID & ROW_START & ":" & COL_ID & lastRow).Value2
        arrI = ws.Range(COL_REF & ROW_START & ":" & COL_REF & lastRow).Value2

        CompterUniques3Colonnes arrA, arrU, arrI, dictRF, dictID, dictREF
    Else
        On Error Resume Next
        Set rngVisibleRows = ws.Range(COL_REF & ROW_START & ":" & COL_REF & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo ErrHandler

        If Not rngVisibleRows Is Nothing Then
            For Each area In rngVisibleRows.Areas
                firstLig = area.Row
                lastLigArea = area.Row + area.Rows.Count - 1

                arrA = ws.Range(COL_RF & firstLig & ":" & COL_RF & lastLigArea).Value2
                arrU = ws.Range(COL_ID & firstLig & ":" & COL_ID & lastLigArea).Value2
                arrI = ws.Range(COL_REF & firstLig & ":" & COL_REF & lastLigArea).Value2

                CompterUniques3Colonnes arrA, arrU, arrI, dictRF, dictID, dictREF
            Next area
        End If
    End If

    nbRF = dictRF.Count
    nbID = dictID.Count
    nbREF = dictREF.Count

    If Not bFiltre Then
        m_nbRF_Total = nbRF
        m_nbID_Total = nbID
        m_nbREF_Total = nbREF
        m_Initialise = True
    End If

    EcrireCompteurs wsMenu, nbRF, nbID, nbREF

SortiePropre:
    Set dictRF = Nothing
    Set dictID = Nothing
    Set dictREF = Nothing
    Set rngVisibleRows = Nothing

    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la mise à jour des compteurs : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' Comptage simultané des 3 colonnes
' =============================================
Private Sub CompterUniques3Colonnes(ByVal arrA As Variant, ByVal arrU As Variant, ByVal arrI As Variant, _
                                    ByVal dictRF As Object, ByVal dictID As Object, ByVal dictREF As Object)

    Dim i As Long
    Dim valRF As String
    Dim valID As String
    Dim valREF As String

    If Not IsArray(arrI) Then
        valRF = Trim$(CStr(arrA))
        valID = Trim$(CStr(arrU))
        valREF = Trim$(CStr(arrI))

        If valRF <> "" Then
            If Not dictRF.Exists(valRF) Then dictRF.Add valRF, 1
        End If

        If valID <> "" Then
            If Not dictID.Exists(valID) Then dictID.Add valID, 1
        End If

        If valREF <> "" Then
            If Not dictREF.Exists(valREF) Then dictREF.Add valREF, 1
        End If

        Exit Sub
    End If

    For i = 1 To UBound(arrI, 1)
        valRF = Trim$(CStr(arrA(i, 1)))
        valID = Trim$(CStr(arrU(i, 1)))
        valREF = Trim$(CStr(arrI(i, 1)))

        If valRF <> "" Then
            If Not dictRF.Exists(valRF) Then dictRF.Add valRF, 1
        End If

        If valID <> "" Then
            If Not dictID.Exists(valID) Then dictID.Add valID, 1
        End If

        If valREF <> "" Then
            If Not dictREF.Exists(valREF) Then dictREF.Add valREF, 1
        End If
    Next i

End Sub

' =============================================
' Restauration instantanee (effacer filtres)
' =============================================
Public Sub RestaurerCompteursInitiaux()

    Dim wsMenu As Worksheet

    On Error GoTo ErrHandler

    If Not m_Initialise Then
        MettreAJourCompteurs
        Exit Sub
    End If

    Set wsMenu = ThisWorkbook.Worksheets(SHEET_MENU_DEROULANT)
    EcrireCompteurs wsMenu, m_nbRF_Total, m_nbID_Total, m_nbREF_Total
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la restauration des compteurs : " & Err.description, vbExclamation

End Sub

' =============================================
' Ecriture dans MENU DEROULANT
' =============================================
Private Sub EcrireCompteurs(ByVal wsMenu As Worksheet, ByVal nbRF As Long, ByVal nbID As Long, ByVal nbREF As Long)

    wsMenu.Range("J1").Value = nbRF & "  RF "
    wsMenu.Range("K1").Value = nbID & " ID "
    wsMenu.Range("L1").Value = nbREF & " REF Uniques"
    wsMenu.Range("M1").Value = nbRF & "  RF   |  " & nbID & " ID   | " & nbREF & " REF Uniques "

End Sub

' =============================================
' Wrappers BDD-DOC pour les boutons de la feuille
'
' Contrat d'appel :
' - les formes Excel doivent appeler uniquement ces wrappers
' - ces wrappers délèguent la logique aux modules spécialisés
'
' Affectations conseillées :
' - Bouton "Appliquer" -> AppliquerFiltresDoc
' - Bouton "Effacer"   -> EffacerFiltresDoc
'
' Dépendances :
' - AppliquerFiltresDoc dépend de :
'     BoutonAppliquerFiltres.AppliquerFiltres
'     zDocCompteurs.MettreAJourCompteurs
' - EffacerFiltresDoc dépend de :
'     BoutonEffacerFiltres.EffacerFiltres
'     zDocCompteurs.RestaurerCompteursInitiaux
' - InitialiserPlaceholdersFeuillePrincipale dépend de :
'     Base.InitialiserPlaceholders
' =============================================

Public Sub AppliquerFiltresDoc()
    AppliquerFiltres
    MettreAJourCompteurs
End Sub

Public Sub EffacerFiltresDoc()
    EffacerFiltres
    RestaurerCompteursInitiaux
End Sub

Public Sub InitialiserPlaceholdersFeuillePrincipale()
    Base.InitialiserPlaceholders
End Sub



