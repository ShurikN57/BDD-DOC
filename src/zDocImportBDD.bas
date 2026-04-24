Attribute VB_Name = "zDocImportBDD"
Option Explicit

' =============================================
' Synchronisation BDD-DOC agents -> BDD-DOC perso
' =============================================

Private Const NOM_CLASSEUR_SOURCE As String = "BDD-DOC-24-04"
Private Const NOM_ONGLET_SOURCE As String = "Base"

Private Const NOM_CLASSEUR_CIBLE As String = "BDD-DOC"
Private Const NOM_ONGLET_CIBLE As String = "Base"

Private Const ONGLET_ID_ABSENTS As String = "ID_absents"
Private Const ONGLET_ID_DOUBLONS As String = "ID_doublons"
Private Const ONGLET_ECARTS As String = "Ecarts_valeurs"

Private Const NOM_FORME_ACTUALISATION As String = "Actualisation"

Private Const FERMER_APRES_SYNCHRO As Boolean = True

' =============================================
' 1. SynchroniserDonneesAgents
' =============================================
Public Sub SynchroniserDonneesAgents()

    Dim wbSource As Workbook
    Dim wbCible As Workbook
    Dim wsSource As Worksheet
    Dim wsCible As Worksheet

    Dim lastRowSource As Long
    Dim lastRowCible As Long

    Dim dictCible As Object
    Dim dictCount As Object

    Dim i As Long
    Dim idVal As String
    Dim confSource As String

    Dim arrSourceID As Variant
    Dim arrSourceYZAB As Variant
    Dim arrCibleID As Variant
    Dim arrCibleYZAB As Variant

    Dim wsAbs As Worksheet
    Dim wsDoublons As Worksheet
    Dim wsEcarts As Worksheet

    Dim rowAbs As Long
    Dim rowDoublons As Long
    Dim rowEcarts As Long

    Dim nbMaj As Long
    Dim nbAbs As Long
    Dim nbDoublons As Long
    Dim nbEcarts As Long
    Dim nbIgnorees As Long

    Dim ligneCible As Long
    Dim yzabSource As Variant
    Dim yzabCible As Variant

    Dim rngConfImpactee As Range
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim cibleDeprotegee As Boolean

    On Error GoTo ErrHandler

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbSource = GetWorkbookByBaseName(NOM_CLASSEUR_SOURCE)
    If wbSource Is Nothing Then
        MsgBox "Classeur source introuvable : " & NOM_CLASSEUR_SOURCE, vbExclamation
        GoTo SortiePropre
    End If

    Set wbCible = GetWorkbookByBaseName(NOM_CLASSEUR_CIBLE)
    If wbCible Is Nothing Then
        MsgBox "Classeur cible introuvable : " & NOM_CLASSEUR_CIBLE, vbExclamation
        GoTo SortiePropre
    End If

    Set wsSource = GetWorksheetSafe(wbSource, NOM_ONGLET_SOURCE)
    If wsSource Is Nothing Then
        MsgBox "Onglet source introuvable : " & NOM_ONGLET_SOURCE, vbExclamation
        GoTo SortiePropre
    End If

    Set wsCible = GetWorksheetSafe(wbCible, NOM_ONGLET_CIBLE)
    If wsCible Is Nothing Then
        MsgBox "Onglet cible introuvable : " & NOM_ONGLET_CIBLE, vbExclamation
        GoTo SortiePropre
    End If

    ' ===== Déprotection de la feuille cible =====
    On Error Resume Next
    wsCible.Unprotect Password:=MDP_DEV
    On Error GoTo ErrHandler
    cibleDeprotegee = (Not wsCible.ProtectContents)

    Set wsAbs = PreparerOngletRapport(wbCible, ONGLET_ID_ABSENTS)
    Set wsDoublons = PreparerOngletRapport(wbCible, ONGLET_ID_DOUBLONS)
    Set wsEcarts = PreparerOngletRapport(wbCible, ONGLET_ECARTS)

    InitialiserRapportAbsents wsAbs
    InitialiserRapportDoublons wsDoublons
    InitialiserRapportEcarts wsEcarts

    rowAbs = 2
    rowDoublons = 2
    rowEcarts = 2

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, COL_ID).End(xlUp).Row
    If lastRowSource < ROW_START Then
        MsgBox "Aucune donnée source à traiter.", vbInformation
        GoTo SortiePropre
    End If

    lastRowCible = wsCible.Cells(wsCible.Rows.Count, COL_ID).End(xlUp).Row
    If lastRowCible < ROW_START Then
        MsgBox "Aucune donnée cible à comparer.", vbExclamation
        GoTo SortiePropre
    End If

    arrSourceID = wsSource.Range(COL_ID & ROW_START & ":" & COL_ID & lastRowSource).Value2
    arrSourceYZAB = wsSource.Range(COL_DATE & ROW_START & ":" & COL_OBS & lastRowSource).Value2

    arrCibleID = wsCible.Range(COL_ID & ROW_START & ":" & COL_ID & lastRowCible).Value2
    arrCibleYZAB = wsCible.Range(COL_DATE & ROW_START & ":" & COL_OBS & lastRowCible).Value2

    Set dictCible = CreateObject("Scripting.Dictionary")
    Set dictCount = CreateObject("Scripting.Dictionary")
    dictCible.CompareMode = vbTextCompare
    dictCount.CompareMode = vbTextCompare

    ConstruireIndexCible arrCibleID, dictCible, dictCount

    For i = 1 To UBound(arrSourceID, 1)

        idVal = Trim$(CStr(arrSourceID(i, 1)))
        confSource = Trim$(CStr(arrSourceYZAB(i, 3)))

        If idVal <> "" And confSource <> "" Then

            yzabSource = ExtraireYZABDepuisArray(arrSourceYZAB, i)

            If Not dictCount.Exists(idVal) Then

                EcrireRapportAbsent wsAbs, rowAbs, idVal, i + ROW_START - 1, yzabSource
                rowAbs = rowAbs + 1
                nbAbs = nbAbs + 1

            ElseIf CLng(dictCount(idVal)) > 1 Then

                EcrireRapportDoublon wsDoublons, rowDoublons, idVal, i + ROW_START - 1, CStr(dictCible(idVal)), yzabSource
                rowDoublons = rowDoublons + 1
                nbDoublons = nbDoublons + 1

            Else

                ligneCible = CLng(dictCible(idVal))
                yzabCible = ExtraireYZABDepuisArray(arrCibleYZAB, ligneCible - ROW_START + 1)

                If YZABEstVide(yzabCible) Then
                    wsCible.Range(COL_DATE & ligneCible & ":" & COL_OBS & ligneCible).Value = yzabSource
                    AjouterCelluleConformiteImpactee wsCible, ligneCible, rngConfImpactee
                    nbMaj = nbMaj + 1

                ElseIf YZABEgaux(yzabSource, yzabCible) Then
                    AjouterCelluleConformiteImpactee wsCible, ligneCible, rngConfImpactee
                    nbIgnorees = nbIgnorees + 1

                Else
                    EcrireRapportEcart wsEcarts, rowEcarts, idVal, i + ROW_START - 1, ligneCible, yzabSource, yzabCible
                    rowEcarts = rowEcarts + 1
                    nbEcarts = nbEcarts + 1
                End If

            End If
        End If
    Next i

    If Not rngConfImpactee Is Nothing Then
        Application.Run "'" & wbCible.Name & "'!" & wsCible.CodeName & ".RafraichirCouleursConformiteSurLignes", rngConfImpactee.Address
    End If

    AjusterRapports wsAbs
    AjusterRapports wsDoublons
    AjusterRapports wsEcarts

    MettreAJourTexteActualisation wbCible, NOM_ONGLET_CIBLE, NOM_FORME_ACTUALISATION
    NettoyerContexteApresSynchronisation wbCible, wsCible
    EnregistrerJournalSynchro wbCible, nbMaj, nbAbs, nbDoublons, nbEcarts, nbIgnorees

    If FERMER_APRES_SYNCHRO Then
        MsgBox "Synchronisation terminée." & vbCrLf & vbCrLf & _
                "Mises à jour : " & nbMaj & vbCrLf & _
                "ID absents : " & nbAbs & vbCrLf & _
                "ID doublons : " & nbDoublons & vbCrLf & _
                "Écarts valeurs : " & nbEcarts & vbCrLf & _
                "Déjà identiques : " & nbIgnorees & vbCrLf & vbCrLf & _
                "Sauvegarde du fichier en cours." & vbCrLf & _
                "Veuillez rouvrir BDD-DOC.", vbInformation

        FinaliserEtFermerApresSynchronisation wbSource, wbCible, wsCible, cibleDeprotegee, _
                                          prevCalculation, prevEnableEvents, prevScreenUpdating
        Exit Sub
    End If

    MsgBox _
        "Synchronisation terminée." & vbCrLf & vbCrLf & _
        "Mises à jour : " & nbMaj & vbCrLf & _
        "ID absents : " & nbAbs & vbCrLf & _
        "ID doublons : " & nbDoublons & vbCrLf & _
        "Écarts valeurs : " & nbEcarts & vbCrLf & _
        "Déjà identiques : " & nbIgnorees, _
        vbInformation

SortiePropre:
    On Error Resume Next
    If cibleDeprotegee Then
        wsCible.Protect Password:=MDP_DEV, UserInterfaceOnly:=True, _
                         AllowFiltering:=True, AllowSorting:=True
    End If
    On Error GoTo 0

    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur SynchroniserDonneesAgents : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' 2. FinaliserEtFermerApresSynchronisation
' =============================================
Private Sub FinaliserEtFermerApresSynchronisation(ByVal wbSource As Workbook, _
                                                  ByVal wbCible As Workbook, _
                                                  ByVal wsCible As Worksheet, _
                                                  ByVal cibleDeprotegee As Boolean, _
                                                  ByVal prevCalculation As XlCalculation, _
                                                  ByVal prevEnableEvents As Boolean, _
                                                  ByVal prevScreenUpdating As Boolean)

    On Error Resume Next

    If cibleDeprotegee Then
        wsCible.Protect Password:=MDP_DEV, UserInterfaceOnly:=True, _
                        AllowFiltering:=True, AllowSorting:=True
    End If

    Application.CutCopyMode = False
    Application.Calculation = prevCalculation
    Application.ScreenUpdating = prevScreenUpdating

    ' Nettoyage manuel avant fermeture
    DesactiverCollageValeursRecherche
    Application.OnKey "^l"
    Application.OnKey "%{F11}"

    ' Important : on bloque les événements pour empêcher Workbook_BeforeClose
    Application.EnableEvents = False

    wbCible.Save

    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If

    wbCible.Close SaveChanges:=True

    On Error GoTo 0

End Sub

' =============================================
' 3. ConstruireIndexCible
' =============================================
Private Sub ConstruireIndexCible(ByVal arrCibleID As Variant, ByVal dictCible As Object, ByVal dictCount As Object)

    Dim i As Long
    Dim idVal As String

    For i = 1 To UBound(arrCibleID, 1)
        idVal = Trim$(CStr(arrCibleID(i, 1)))

        If idVal <> "" Then
            If Not dictCount.Exists(idVal) Then
                dictCount.Add idVal, 1
                dictCible.Add idVal, i + ROW_START - 1
            Else
                dictCount(idVal) = CLng(dictCount(idVal)) + 1
                dictCible(idVal) = CStr(dictCible(idVal)) & "," & CStr(i + ROW_START - 1)
            End If
        End If
    Next i

End Sub

' =============================================
' 4. ExtraireYZABDepuisArray
' =============================================
Private Function ExtraireYZABDepuisArray(ByVal arr As Variant, ByVal indexLigne As Long) As Variant

    Dim t(1 To 1, 1 To 4) As Variant

    t(1, 1) = arr(indexLigne, 1)
    t(1, 2) = arr(indexLigne, 2)
    t(1, 3) = arr(indexLigne, 3)
    t(1, 4) = arr(indexLigne, 4)

    ExtraireYZABDepuisArray = t

End Function

' =============================================
' 5. YZABEstVide
' =============================================
Private Function YZABEstVide(ByVal yzab As Variant) As Boolean

    YZABEstVide = _
        Trim$(CStr(yzab(1, 1))) = "" And _
        Trim$(CStr(yzab(1, 2))) = "" And _
        Trim$(CStr(yzab(1, 3))) = "" And _
        Trim$(CStr(yzab(1, 4))) = ""

End Function

' =============================================
' 6. YZABEgaux
' =============================================
Private Function YZABEgaux(ByVal a As Variant, ByVal b As Variant) As Boolean

    YZABEgaux = _
        NormaliserValeur(a(1, 1)) = NormaliserValeur(b(1, 1)) And _
        NormaliserValeur(a(1, 2)) = NormaliserValeur(b(1, 2)) And _
        NormaliserValeur(a(1, 3)) = NormaliserValeur(b(1, 3)) And _
        NormaliserValeur(a(1, 4)) = NormaliserValeur(b(1, 4))

End Function

' =============================================
' 7. NormaliserValeur
' =============================================
Private Function NormaliserValeur(ByVal v As Variant) As String

    If IsError(v) Then
        NormaliserValeur = "#ERREUR#"
    ElseIf IsDate(v) Then
        NormaliserValeur = Format$(Int(CDate(v)), "dd/mm/yyyy")
    Else
        NormaliserValeur = Trim$(CStr(v))
    End If

End Function

' =============================================
' 8. AjouterCelluleConformiteImpactee
' =============================================
Private Sub AjouterCelluleConformiteImpactee(ByVal ws As Worksheet, ByVal lig As Long, ByRef rngConfImpactee As Range)

    Dim rngCell As Range

    If lig < ROW_START Then Exit Sub

    Set rngCell = ws.Range(COL_CONF & lig)

    If rngConfImpactee Is Nothing Then
        Set rngConfImpactee = rngCell
    Else
        Set rngConfImpactee = Union(rngConfImpactee, rngCell)
    End If

End Sub

' =============================================
' 9. MettreAJourTexteActualisation
' =============================================
Private Sub MettreAJourTexteActualisation(ByVal wb As Workbook, ByVal nomOnglet As String, ByVal nomForme As String)

    Dim ws As Worksheet

    On Error GoTo Fin

    Set ws = wb.Worksheets(nomOnglet)

    With ws.Shapes(nomForme).TextFrame
        .Characters.Text = "Dernière actualisation : " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
                           "Source : " & NOM_CLASSEUR_SOURCE
    End With

Fin:

End Sub

' =============================================
' 10. PreparerOngletRapport
' =============================================
Private Function PreparerOngletRapport(ByVal wb As Workbook, ByVal nomOnglet As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(nomOnglet)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nomOnglet
    Else
        ws.Cells.Clear
    End If

    Set PreparerOngletRapport = ws

End Function

' =============================================
' 11. InitialiserRapportAbsents
' =============================================
Private Sub InitialiserRapportAbsents(ByVal ws As Worksheet)

    ws.Range("A1:G1").Value = Array("ID", "Ligne source", "Date source", "Nom source", "Conformité source", "Observation source", "Motif")

End Sub

' =============================================
' 12. InitialiserRapportDoublons
' =============================================
Private Sub InitialiserRapportDoublons(ByVal ws As Worksheet)

    ws.Range("A1:H1").Value = Array("ID", "Ligne source", "Lignes cible", "Date source", "Nom source", "Conformité source", "Observation source", "Motif")

End Sub

' =============================================
' 13. InitialiserRapportEcarts
' =============================================
Private Sub InitialiserRapportEcarts(ByVal ws As Worksheet)

    ws.Range("A1:K1").Value = Array("ID", "Ligne source", "Ligne cible", "Date source", "Nom source", "Conformité source", "Observation source", "Date cible", "Nom cible", "Conformité cible", "Observation cible")

End Sub

' =============================================
' 14. EcrireRapportAbsent
' =============================================
Private Sub EcrireRapportAbsent(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal yzabSource As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = yzabSource(1, 1)
    ws.Cells(r, 4).Value = yzabSource(1, 2)
    ws.Cells(r, 5).Value = yzabSource(1, 3)
    ws.Cells(r, 6).Value = yzabSource(1, 4)
    ws.Cells(r, 7).Value = "ID absent du fichier cible"

End Sub

' =============================================
' 15. EcrireRapportDoublon
' =============================================
Private Sub EcrireRapportDoublon(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal lignesCible As String, ByVal yzabSource As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = lignesCible
    ws.Cells(r, 4).Value = yzabSource(1, 1)
    ws.Cells(r, 5).Value = yzabSource(1, 2)
    ws.Cells(r, 6).Value = yzabSource(1, 3)
    ws.Cells(r, 7).Value = yzabSource(1, 4)
    ws.Cells(r, 8).Value = "ID présent plusieurs fois dans la cible"

End Sub

' =============================================
' 16. EcrireRapportEcart
' =============================================
Private Sub EcrireRapportEcart(ByVal ws As Worksheet, ByVal r As Long, ByVal idVal As String, ByVal ligneSource As Long, ByVal ligneCible As Long, ByVal yzabSource As Variant, ByVal yzabCible As Variant)

    ws.Cells(r, 1).Value = idVal
    ws.Cells(r, 2).Value = ligneSource
    ws.Cells(r, 3).Value = ligneCible

    ws.Cells(r, 4).Value = yzabSource(1, 1)
    ws.Cells(r, 5).Value = yzabSource(1, 2)
    ws.Cells(r, 6).Value = yzabSource(1, 3)
    ws.Cells(r, 7).Value = yzabSource(1, 4)

    ws.Cells(r, 8).Value = yzabCible(1, 1)
    ws.Cells(r, 9).Value = yzabCible(1, 2)
    ws.Cells(r, 10).Value = yzabCible(1, 3)
    ws.Cells(r, 11).Value = yzabCible(1, 4)

End Sub

' =============================================
' 17. AjusterRapports
' =============================================
Private Sub AjusterRapports(ByVal ws As Worksheet)

    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

End Sub

' =============================================
' 18. GetWorkbookByBaseName
' =============================================
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

' =============================================
' 19. GetWorksheetSafe
' =============================================
Private Function GetWorksheetSafe(ByVal wb As Workbook, ByVal nomOnglet As String) As Worksheet

    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(nomOnglet)
    On Error GoTo 0

End Function

' =============================================
' 20. NettoyerContexteApresSynchronisation
' =============================================
Private Sub NettoyerContexteApresSynchronisation(ByVal wbCible As Workbook, ByVal wsCible As Worksheet)

    On Error Resume Next

    Application.CutCopyMode = False
    DoEvents

    wbCible.Activate
    wsCible.Activate
    wsCible.Range("A1").Select

    Application.CutCopyMode = False

    On Error GoTo 0

End Sub

' =============================================
' 21. EnregistrerJournalSynchro
' =============================================
Private Sub EnregistrerJournalSynchro(ByVal wb As Workbook, _
                                      ByVal nbMaj As Long, _
                                      ByVal nbAbs As Long, _
                                      ByVal nbDoublons As Long, _
                                      ByVal nbEcarts As Long, _
                                      ByVal nbIgnorees As Long)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo Fin

    Set ws = GetOrCreateSheetSynchro(wb)

    If Trim$(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:H1").Value = Array("Date", "Heure", "Source", "Mises à jour", "ID absents", "ID doublons", "Écarts valeurs", "Déjà identiques")
        ws.Rows(1).Font.Bold = True
    End If

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Date
    ws.Cells(nextRow, 1).NumberFormat = "dd/mm/yyyy"

    ws.Cells(nextRow, 2).Value = Time
    ws.Cells(nextRow, 2).NumberFormat = "hh:mm:ss"

    ws.Cells(nextRow, 3).Value = NOM_CLASSEUR_SOURCE
    ws.Cells(nextRow, 4).Value = nbMaj
    ws.Cells(nextRow, 5).Value = nbAbs
    ws.Cells(nextRow, 6).Value = nbDoublons
    ws.Cells(nextRow, 7).Value = nbEcarts
    ws.Cells(nextRow, 8).Value = nbIgnorees

    ws.Columns("A:H").AutoFit

Fin:

End Sub

' =============================================
' 22. GetOrCreateSheetSynchro
' =============================================
Private Function GetOrCreateSheetSynchro(ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("Synchro")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Synchro"
    End If

    Set GetOrCreateSheetSynchro = ws

End Function