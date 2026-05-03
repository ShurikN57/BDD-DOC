Attribute VB_Name = "zDocImportPowerQuery"
Option Explicit

' ============================================================
' IMPORT EXCEL PQ -> BDD-DOC
' Source : fichier Excel Power Query final
' Cible  : ThisWorkbook (BDD-DOC)
' ============================================================

' Valeurs par défaut proposées dans le UserForm
Private Const NOM_CLASSEUR_DEFAULT As String = "PQ-30-04"
Private Const NOM_ONGLET_DEFAULT   As String = "PQ-DOC"
Private Const NOM_FEUILLE_ARCHIVE  As String = "PQ_ID_Supprimes"
Private Const NOM_FEUILLE_SYNCHRO  As String = "Synchro"

' ============================================================
' MACRO PRINCIPALE
' ============================================================
Public Sub ImporterDepuisFichierPowerQuery()

    Dim wbSource   As Workbook
    Dim wbCible    As Workbook
    Dim wsSource   As Worksheet
    Dim wsCible    As Worksheet
    Dim wsArchive  As Worksheet

    Dim dictAncien  As Object
    Dim dictNouveau As Object

    Dim oldCalc    As XlCalculation
    Dim oldScreen  As Boolean
    Dim oldEvents  As Boolean
    Dim oldAlerts  As Boolean
    Dim etatApplicationSauve As Boolean

    Dim nomClasseur As String
    Dim nomOnglet   As String

    Dim nbArchives As Long
    Dim nbNonConformes As Long
    Dim nbLignesImportees As Long
    Dim nbConformitesReinjectees As Long
    Dim nbDoublonsSource As Long
    Dim nbNouveauxIDSansSuivi As Long

    Dim msg As String

    On Error GoTo ErrHandler

    ' ===== 1. Saisie via UserForm =====
    Dim frm As UF_ImportPQ
    Set frm = New UF_ImportPQ

    frm.InitialiserImportPQ NOM_CLASSEUR_DEFAULT, NOM_ONGLET_DEFAULT
    frm.Show vbModal

    Dim bConfirmed As Boolean
    bConfirmed = frm.Confirmed
    nomClasseur = frm.nomClasseur
    nomOnglet = frm.nomOnglet

    Unload frm
    Set frm = Nothing

    If Not bConfirmed Then Exit Sub

    ' ===== 2. Résolution des objets =====
    Set wbCible = ThisWorkbook
    Set wsCible = wbCible.Worksheets(SHEET_MAIN)
    Set wsArchive = GetOrCreateSheet(wbCible, NOM_FEUILLE_ARCHIVE)

    Set wbSource = GetWorkbookByBaseName(nomClasseur)
    If wbSource Is Nothing Then
        Err.Raise vbObjectError + 2000, , _
            "Classeur source introuvable : " & nomClasseur & vbCrLf & _
            "Ouvre d'abord le fichier Excel Power Query final."
    End If

    Set wsSource = GetWorksheetSafe(wbSource, nomOnglet)
    If wsSource Is Nothing Then
        Err.Raise vbObjectError + 2001, , _
            "Onglet source introuvable : " & nomOnglet
    End If

    ' ===== 3. Validation source AVANT tout effacement =====
    ValiderSourceOuErreur wsSource

    ' ===== 4. Sauvegarde état application =====
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    etatApplicationSauve = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ===== 5. Import =====
    Set dictAncien = ChargerAncienneBase(wsCible)

    nbLignesImportees = CompterLignesImportees(wsSource)
    Set dictNouveau = ChargerIDsDepuisFeuille(wsSource, 2, nbDoublonsSource)
    nbNouveauxIDSansSuivi = CompterNouveauxIDSansSuivi(dictAncien, dictNouveau)

    ArchiverLignesDisparuesAvecConformite wsArchive, dictAncien, dictNouveau, nbArchives, nbNonConformes

    RemplacerBaseDepuisSource wsCible, wsSource
    ReinjecterColonnesSuivi wsCible, dictAncien, nbConformitesReinjectees

    ' ===== 6. Rafraîchissement couleurs conformité =====
    Base.RafraichirCouleursConformite

    ' ===== 7. Mise à jour shape actualisation =====
    MettreAJourShapeImportPQ nomClasseur

    ' ===== 8. Journal Synchro J:R =====
    EnregistrerJournalImportPQ wbCible, nomClasseur, nbLignesImportees, _
                               nbConformitesReinjectees, nbArchives, _
                               nbNonConformes, nbDoublonsSource, _
                               nbNouveauxIDSansSuivi

    ' ===== 9. Rapport de fin =====
    msg = "Import depuis le fichier Power Query terminé." & vbCrLf & vbCrLf & _
          "Lignes importées : " & nbLignesImportees & vbCrLf & _
          "Conformités réinjectées : " & nbConformitesReinjectees & vbCrLf & _
          "Lignes archivées avec conformité : " & nbArchives

    If nbNonConformes > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Attention : " & nbNonConformes & " ID disparu(s) SANS conformité ont été ignoré(s) et non archivé(s)."
    End If

    If nbDoublonsSource > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Attention : " & nbDoublonsSource & " ID doublon(s) détecté(s) dans la source PQ."
    End If

    MsgBox msg, vbInformation

SortiePropre:
    If etatApplicationSauve Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
    End If
    Exit Sub

ErrHandler:
    MsgBox "Erreur ImporterDepuisFichierPowerQuery : " & Err.description, vbCritical
    Resume SortiePropre

End Sub

' ============================================================
' MISE A JOUR SHAPE ACTUALISATION
' ============================================================
Private Sub MettreAJourShapeImportPQ(ByVal nomSource As String)

    Dim ws As Worksheet
    Dim titre As String

    On Error GoTo Fin

    titre = "Dernier Import PQ :"

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    With ws.Shapes("DernierImportPQ").TextFrame
        .Characters.Text = titre & vbCrLf & _
                           Format(Now, "dd/mm/yyyy à hh:mm") & vbCrLf & _
                           "Source : " & nomSource
        ' Couleur rose (#E59EDD) sur le titre uniquement
        .Characters(1, Len(titre)).Font.Color = RGB(229, 158, 221)
    End With

Fin:

End Sub

' ============================================================
' VALIDATION SOURCE (avant tout effacement)
' ============================================================
Private Sub ValiderSourceOuErreur(ByVal wsSource As Worksheet)

    Dim lastRowSource As Long
    Dim lastColSource As Long

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastColSource = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRowSource < 2 Then
        Err.Raise vbObjectError + 2002, , "La feuille source est vide."
    End If

    If lastColSource < NB_COL_TABLE Then
        Err.Raise vbObjectError + 2003, , _
            "La feuille source n'a pas assez de colonnes." & vbCrLf & _
            "Colonnes trouvées : " & lastColSource & " — Attendu : " & NB_COL_TABLE
    End If

    ' Contrôle des colonnes structurantes.
    ' On compare les titres source ligne 1 avec les titres de Base ligne ROW_HEADER.
    ValiderTitreColonneSource wsSource, ColNum(COL_RF)
    ValiderTitreColonneSource wsSource, ColNum(COL_REF)
    ValiderTitreColonneSource wsSource, ColNum(COL_ID)
    ValiderTitreColonneSource wsSource, ColNum(COL_DATE)
    ValiderTitreColonneSource wsSource, ColNum(COL_NOM)
    ValiderTitreColonneSource wsSource, ColNum(COL_CONF)
    ValiderTitreColonneSource wsSource, ColNum(COL_OBS)

End Sub

' ============================================================
' VALIDE UN TITRE DE COLONNE SOURCE
' ============================================================
Private Sub ValiderTitreColonneSource(ByVal wsSource As Worksheet, ByVal idxCol As Long)

    Dim wsBase As Worksheet
    Dim titreAttendu As String
    Dim titreSource As String
    Dim nomColonne As String

    Set wsBase = ThisWorkbook.Worksheets(SHEET_MAIN)

    titreAttendu = NormaliserTitreColonne(wsBase.Cells(ROW_HEADER, idxCol).Value)
    titreSource = NormaliserTitreColonne(wsSource.Cells(1, idxCol).Value)

    nomColonne = Replace(wsBase.Cells(1, idxCol).Address(False, False), "1", "")

    If Len(titreAttendu) = 0 Then
        Err.Raise vbObjectError + 2010, , _
            "Titre attendu vide dans Base pour la colonne " & nomColonne & "."
    End If

    If titreSource <> titreAttendu Then
        Err.Raise vbObjectError + 2011, , _
            "Structure source invalide en colonne " & nomColonne & "." & vbCrLf & _
            "Titre attendu : " & wsBase.Cells(ROW_HEADER, idxCol).Value & vbCrLf & _
            "Titre trouvé : " & wsSource.Cells(1, idxCol).Value
    End If

End Sub

' ============================================================
' NORMALISE UN TITRE DE COLONNE
' ============================================================
Private Function NormaliserTitreColonne(ByVal valeur As Variant) As String

    Dim s As String

    s = CStr(valeur)
    s = Replace(s, ChrW$(160), " ")
    s = Trim$(s)

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormaliserTitreColonne = LCase$(s)

End Function



' ============================================================
' CHARGE L'ANCIENNE BASE
' dict(ID) = Array(Date, Nom, Conf, Obs, LigneAAB)
' ============================================================
Private Function ChargerAncienneBase(ByVal ws As Worksheet) As Object

    Dim dict    As Object
    Dim lastRow As Long
    Dim arr     As Variant
    Dim i       As Long
    Dim idxID   As Long
    Dim idxDate As Long
    Dim idxNom  As Long
    Dim idxConf As Long
    Dim idxObs  As Long
    Dim idVal   As String

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
' Remonte aussi le nombre d'ID doublons source
' ============================================================
Private Function ChargerIDsDepuisFeuille(ByVal ws As Worksheet, _
                                         ByVal firstDataRow As Long, _
                                         ByRef nbDoublonsSource As Long) As Object

    Dim dict    As Object
    Dim lastRow As Long
    Dim arr     As Variant
    Dim i       As Long
    Dim idxID   As Long
    Dim idVal   As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    nbDoublonsSource = 0

    idxID = ColNum(COL_ID)

    lastRow = ws.Cells(ws.Rows.Count, idxID).End(xlUp).Row
    If lastRow < firstDataRow Then
        Set ChargerIDsDepuisFeuille = dict
        Exit Function
    End If

    arr = ws.Range(ws.Cells(firstDataRow, idxID), ws.Cells(lastRow, idxID)).Value2

    For i = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(i, 1)))

        If Len(idVal) > 0 Then
            If dict.Exists(idVal) Then
                nbDoublonsSource = nbDoublonsSource + 1
            Else
                dict.Add idVal, True
            End If
        End If
    Next i

    Set ChargerIDsDepuisFeuille = dict

End Function

' ============================================================
' ARCHIVE DES LIGNES DISPARUES AVEC CONFORMITE
' Remonte nbArchives et nbNonConformes à l'appelant
' ============================================================
Private Sub ArchiverLignesDisparuesAvecConformite(ByVal wsArchive As Worksheet, _
                                                   ByVal dictAncien As Object, _
                                                   ByVal dictNouveau As Object, _
                                                   ByRef nbArchives As Long, _
                                                   ByRef nbNonConformes As Long)

    Dim k       As Variant
    Dim info    As Variant
    Dim confVal As String
    Dim nextRow As Long
    Dim ligneData As Variant

    nbArchives = 0
    nbNonConformes = 0

    ' En-tête si feuille archive vide
    If wsArchive.Cells(1, 1).Value = "" Then
        ThisWorkbook.Worksheets(SHEET_MAIN).Rows(ROW_HEADER).Copy _
            Destination:=wsArchive.Rows(1)
    End If

    nextRow = wsArchive.Cells(wsArchive.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    For Each k In dictAncien.Keys
        If Not dictNouveau.Exists(CStr(k)) Then
            info = dictAncien(k)
            confVal = Trim$(CStr(info(2)))

            If Len(confVal) > 0 Then
                ' Sécurité taille avant écriture
                ligneData = info(4)
                If IsArray(ligneData) Then
                    If UBound(ligneData, 2) = NB_COL_TABLE Then
                        wsArchive.Range( _
                            wsArchive.Cells(nextRow, 1), _
                            wsArchive.Cells(nextRow, NB_COL_TABLE) _
                        ).Value = ligneData
                        nextRow = nextRow + 1
                        nbArchives = nbArchives + 1
                    End If
                End If
            Else
                nbNonConformes = nbNonConformes + 1
            End If
        End If
    Next k

End Sub

' ============================================================
' REMPLACE BASE PAR LA FEUILLE SOURCE
' Validation effectuée en amont par ValiderSourceOuErreur
' ============================================================
Private Sub RemplacerBaseDepuisSource(ByVal wsCible As Worksheet, ByVal wsSource As Worksheet)

    Dim lastRowSource As Long
    Dim lastRowCible As Long
    Dim lastRowNouveau As Long
    Dim lastRowEffacer As Long

    Dim arr           As Variant
    Dim i             As Long
    Dim idxNum        As Long
    Dim nbLignes      As Long
    Dim nbColonnes    As Long

    idxNum = ColNum(COL_NUM)

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastRowCible = wsCible.Cells(wsCible.Rows.Count, 1).End(xlUp).Row

    If lastRowCible < ROW_START Then lastRowCible = ROW_START

    arr = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, NB_COL_TABLE)).Value2

    nbLignes = UBound(arr, 1)
    nbColonnes = UBound(arr, 2)

    lastRowNouveau = ROW_START + nbLignes - 1

    lastRowEffacer = lastRowCible
    If lastRowNouveau > lastRowEffacer Then lastRowEffacer = lastRowNouveau

    ' Effacement limité à la zone réellement utile :
    ' ancienne base ou nouvelle base, selon la plus grande.
    wsCible.Range(wsCible.Cells(ROW_START, 1), _
                  wsCible.Cells(lastRowEffacer, NB_COL_TABLE)).ClearContents

    ' Colonne NUM en texte uniquement sur la zone utile
    wsCible.Range(wsCible.Cells(ROW_START, idxNum), _
                  wsCible.Cells(lastRowNouveau, idxNum)).NumberFormat = "@"

    ' Colonne NUM : conserver exactement le texte Power Query
    For i = 1 To nbLignes
        If Len(Trim$(CStr(arr(i, idxNum)))) > 0 Then
            arr(i, idxNum) = CStr(arr(i, idxNum))
        End If
    Next i

    wsCible.Cells(ROW_START, 1).Resize(nbLignes, nbColonnes).Value = arr

End Sub

' ============================================================
' REINJECTION Y:AB uniquement (pas réécriture de A:AB entier)
' Remonte nbConformitesReinjectees
' ============================================================
Private Sub ReinjecterColonnesSuivi(ByVal wsBase As Worksheet, _
                                    ByVal dictAncien As Object, _
                                    ByRef nbConformitesReinjectees As Long)

    Dim lastRow As Long
    Dim arrID   As Variant
    Dim arrSuivi As Variant
    Dim i       As Long
    Dim idxID   As Long
    Dim idxDate As Long
    Dim idxNom  As Long
    Dim idxConf As Long
    Dim idxObs  As Long
    Dim colSuiviDeb As Long
    Dim nbColsSuivi As Long
    Dim idVal   As String
    Dim info    As Variant

    nbConformitesReinjectees = 0

    idxID = ColNum(COL_ID)
    idxDate = ColNum(COL_DATE)
    idxNom = ColNum(COL_NOM)
    idxConf = ColNum(COL_CONF)
    idxObs = ColNum(COL_OBS)

    ' Colonnes suivi : Y à AB (indices relatifs dans le bloc suivi)
    colSuiviDeb = idxDate                        ' colonne Y (index absolu)
    nbColsSuivi = idxObs - idxDate + 1           ' Y, Z, AA, AB = 4 colonnes

    lastRow = wsBase.Cells(wsBase.Rows.Count, idxID).End(xlUp).Row
    If lastRow < ROW_START Then Exit Sub

    ' Lecture matricielle : colonne ID + bloc suivi Y:AB
    arrID = wsBase.Range(wsBase.Cells(ROW_START, idxID), _
                         wsBase.Cells(lastRow, idxID)).Value2

    arrSuivi = wsBase.Range(wsBase.Cells(ROW_START, colSuiviDeb), _
                            wsBase.Cells(lastRow, colSuiviDeb + nbColsSuivi - 1)).Value2

    For i = 1 To UBound(arrID, 1)
        idVal = Trim$(CStr(arrID(i, 1)))

        If Len(idVal) > 0 Then
            If dictAncien.Exists(idVal) Then
                info = dictAncien(idVal)

                ' info(0)=Date  info(1)=Nom  info(2)=Conf  info(3)=Obs
                arrSuivi(i, 1) = info(0)   ' Y = Date
                arrSuivi(i, 2) = info(1)   ' Z = Nom
                arrSuivi(i, 3) = info(2)   ' AA = Conf
                arrSuivi(i, 4) = info(3)   ' AB = Obs

                If Trim$(CStr(info(2))) <> "" Then
                    nbConformitesReinjectees = nbConformitesReinjectees + 1
                End If
            End If
        End If
    Next i

    ' Ecriture uniquement du bloc Y:AB
    wsBase.Range(wsBase.Cells(ROW_START, colSuiviDeb), _
                 wsBase.Cells(lastRow, colSuiviDeb + nbColsSuivi - 1)).Value = arrSuivi

End Sub

' ============================================================
' COMPTE LES LIGNES IMPORTEES SOURCE PQ
' ============================================================
Private Function CompterLignesImportees(ByVal wsSource As Worksheet) As Long

    Dim lastRow As Long

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        CompterLignesImportees = 0
    Else
        CompterLignesImportees = lastRow - 1
    End If

End Function

' ============================================================
' COMPTE LES NOUVEAUX ID SANS SUIVI
' ============================================================
Private Function CompterNouveauxIDSansSuivi(ByVal dictAncien As Object, _
                                            ByVal dictNouveau As Object) As Long

    Dim k As Variant
    Dim total As Long

    total = 0

    For Each k In dictNouveau.Keys
        If Not dictAncien.Exists(CStr(k)) Then
            total = total + 1
        End If
    Next k

    CompterNouveauxIDSansSuivi = total

End Function

' ============================================================
' ENREGISTRE JOURNAL IMPORT PQ DANS SYNCHRO J:R
' ============================================================
Private Sub EnregistrerJournalImportPQ(ByVal wb As Workbook, _
                                       ByVal nomSourcePQ As String, _
                                       ByVal nbLignesImportees As Long, _
                                       ByVal nbConformitesReinjectees As Long, _
                                       ByVal nbArchives As Long, _
                                       ByVal nbNonConformes As Long, _
                                       ByVal nbDoublonsSource As Long, _
                                       ByVal nbNouveauxIDSansSuivi As Long)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo Fin

    Set ws = GetOrCreateSheet(wb, NOM_FEUILLE_SYNCHRO)

    ' On écrit uniquement les titres si la zone est vide.
    ' Aucune mise en forme, aucun effacement.
    If Trim$(CStr(ws.Range("J1").Value)) = "" Then
        ws.Range("J1:R1").Value = Array( _
            "Date", _
            "Heure", _
            "Source PQ", _
            "Nb lignes importées", _
            "Conformités réinjectées", _
            "ID disparus avec conformité", _
            "ID disparus sans conformité", _
            "ID doublons source", _
            "Nouveaux ID sans suivi" _
        )
    End If

    nextRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, "J").Value = Date
    ws.Cells(nextRow, "K").Value = Time
    ws.Cells(nextRow, "L").Value = nomSourcePQ
    ws.Cells(nextRow, "M").Value = nbLignesImportees
    ws.Cells(nextRow, "N").Value = nbConformitesReinjectees
    ws.Cells(nextRow, "O").Value = nbArchives
    ws.Cells(nextRow, "P").Value = nbNonConformes
    ws.Cells(nextRow, "Q").Value = nbDoublonsSource
    ws.Cells(nextRow, "R").Value = nbNouveauxIDSansSuivi

Fin:

End Sub

' ============================================================
' HELPERS
' ============================================================
Private Function GetWorkbookByBaseName(ByVal baseName As String) As Workbook

    Dim wb             As Workbook
    Dim nomSansExtension As String

    For Each wb In Application.Workbooks
        nomSansExtension = wb.Name
        If InStrRev(nomSansExtension, ".") > 0 Then
            nomSansExtension = Left$(nomSansExtension, InStrRev(nomSansExtension, ".") - 1)
        End If
        If StrComp(nomSansExtension, baseName, vbTextCompare) = 0 Then
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
    Dim j       As Long

    ReDim ligne(1 To 1, 1 To nbCols)

    For j = 1 To nbCols
        ligne(1, j) = arr(rowIndex, j)
    Next j

    ExtraireLigne = ligne

End Function
