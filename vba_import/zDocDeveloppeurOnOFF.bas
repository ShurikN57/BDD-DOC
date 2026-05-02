Attribute VB_Name = "zDocDeveloppeurOnOFF"
Option Explicit

' =============================================
'            Mode DeveloppeurONOFF
' =============================================

Public ModeDeveloppeurActif As Boolean

' =============================================
' OUTILS
' =============================================
Private Function ConstruirePlageEditable(ByVal ws As Worksheet, ByVal lastRow As Long) As Range

    Dim rng As Range
    Dim rngTemp As Range

    If PLAGE_RECHERCHE <> "" Then
        Set rng = ws.Range(PLAGE_RECHERCHE)
    End If

    If PLAGE_EDITABLE_MAIN <> "" Then
        Set rngTemp = Intersect(ws.Range(PLAGE_EDITABLE_MAIN), ws.Rows(ROW_START & ":" & lastRow))
        If Not rngTemp Is Nothing Then
            If rng Is Nothing Then
                Set rng = rngTemp
            Else
                Set rng = Union(rng, rngTemp)
            End If
        End If
    End If

    If PLAGE_EDITABLE_SUIVI <> "" Then
        Set rngTemp = Intersect(ws.Range(PLAGE_EDITABLE_SUIVI), ws.Rows(ROW_START & ":" & lastRow))
        If Not rngTemp Is Nothing Then
            If rng Is Nothing Then
                Set rng = rngTemp
            Else
                Set rng = Union(rng, rngTemp)
            End If
        End If
    End If

    If PLAGE_EDITABLE_AIDE <> "" Then
        Set rngTemp = Intersect(ws.Range(PLAGE_EDITABLE_AIDE), ws.Rows(ROW_START & ":" & lastRow))
        If Not rngTemp Is Nothing Then
            If rng Is Nothing Then
                Set rng = rngTemp
            Else
                Set rng = Union(rng, rngTemp)
            End If
        End If
    End If

    Set ConstruirePlageEditable = rng

End Function

Private Sub AppliquerValidationConformite()

    Dim ws As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    lastRow = DerniereLigneUtileMain()

    ws.Unprotect Password:=MDP_DEV

    With ws.Range(COL_VALIDATION_CONF & ROW_START & ":" & COL_VALIDATION_CONF & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="='" & SHEET_MENU_DEROULANT & "'!$A$1:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Valeur non autorisée"
        .ErrorMessage = MSG_VALEURS_CONF
    End With

    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'application de la validation conformité : " & Err.description, vbExclamation

End Sub

Private Function ObtenirVBProjectSiAccessible(ByRef vbProj As Object) As Boolean

    On Error GoTo ErrHandler
    Set vbProj = ThisWorkbook.VBProject
    ObtenirVBProjectSiAccessible = Not vbProj Is Nothing
    Exit Function

ErrHandler:
    Debug.Print "[zDocDeveloppeurOnOFF] Accès VBProject indisponible : " & Err.Number & " - " & Err.description
    Err.Clear
    Set vbProj = Nothing

End Function

Private Sub DeprotegerFeuilleSansErreur(ByVal ws As Worksheet)

    On Error GoTo ErrHandler
    ws.Unprotect Password:=MDP_DEV
    ws.EnableSelection = xlNoRestrictions
    Exit Sub

ErrHandler:
    Debug.Print "[zDocDeveloppeurOnOFF] Déprotection feuille impossible (" & ws.Name & ") : " & _
                Err.Number & " - " & Err.description
    Err.Clear

End Sub

Private Sub DeprotegerClasseurSansErreur()

    On Error GoTo ErrHandler
    ThisWorkbook.Unprotect Password:=MDP_DEV
    Exit Sub

ErrHandler:
    Debug.Print "[zDocDeveloppeurOnOFF] Déprotection classeur impossible : " & Err.Number & " - " & Err.description
    Err.Clear

End Sub

Private Function MotDePasseValidePourBouton(ByVal contexte As String) As Boolean

    Dim MDP As String

    MDP = InputBox("Mot de passe développeur :", contexte)

    If MDP <> MDP_DEV Then
        MsgBox "Mot de passe incorrect.", vbCritical
        MotDePasseValidePourBouton = False
    Else
        MotDePasseValidePourBouton = True
    End If

End Function

' =============================================
' BLOCAGE / DEBLOCAGE ALT+F11
' =============================================
Public Sub BloquerAltF11()
    Application.OnKey "%{F11}", "AccesFerme"
End Sub

Public Sub DebloquerAltF11()
    Application.OnKey "%{F11}"
End Sub

Public Sub AccesFerme()
    MsgBox "Accès non autorisé.", vbCritical
End Sub

Public Sub OuvrirEditeurVBA()
    Application.VBE.MainWindow.Visible = True
End Sub

' =============================================
' MODE DEVELOPPEUR ON
' =============================================
Public Sub ModeDeveloppeur_ON()

    Dim MDP As String
    Dim ws As Worksheet
    Dim vbProj As Object

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

      If Not MotDePasseValidePourBouton("Mode développeur") Then Exit Sub

    MDP = MDP_DEV

    On Error GoTo ErrHandler

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    DebloquerAltF11

    If ObtenirVBProjectSiAccessible(vbProj) Then
        If vbProj.Protection = 1 Then
            SendKeys MDP & "~"
            Application.VBE.MainWindow.Visible = True
        End If
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"

    For Each ws In ThisWorkbook.Worksheets
        DeprotegerFeuilleSansErreur ws
    Next ws

    DeprotegerClasseurSansErreur

    AppliquerValidationConformite

    ModeDeveloppeurActif = True
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    MsgBox "Mode développeur ACTIVÉ." & vbCrLf & _
           "- Feuilles déprotégées" & vbCrLf & _
           "- Structure du classeur déprotégée", vbInformation

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de l'activation du mode développeur : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

' =============================================
' MODE DEVELOPPEUR OFF
' =============================================
Public Sub ModeDeveloppeur_OFF(Optional ByVal Silencieux As Boolean = False)

    Dim ws As Worksheet
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim rngEditable As Range

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation

    On Error GoTo ErrHandler

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    BloquerAltF11

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ModeDeveloppeurActif = False
    lastRow = DerniereLigneUtileMain()
    Set wsMain = ThisWorkbook.Worksheets(SHEET_MAIN)

    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

    With wsMain
        .Unprotect Password:=MDP_DEV
        .Cells.Locked = True

        Set rngEditable = ConstruirePlageEditable(wsMain, lastRow)
        If Not rngEditable Is Nothing Then
            rngEditable.Locked = False
        End If
    End With

    AppliquerValidationConformite

    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=MDP_DEV, UserInterfaceOnly:=True, _
                   AllowFiltering:=True, AllowSorting:=True
        ws.EnableSelection = xlNoRestrictions
    Next ws

    ThisWorkbook.Protect Password:=MDP_DEV, Structure:=True
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    If Not Silencieux Then
        MsgBox "Mode développeur DÉSACTIVÉ." & vbCrLf & _
               "- Feuilles protégées" & vbCrLf & _
               "- Structure du classeur protégée", vbInformation
    End If

SortiePropre:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Erreur lors de la désactivation du mode développeur : " & Err.description, vbExclamation
    Resume SortiePropre

End Sub

Public Sub ModeDeveloppeur_OFF_Bouton()
    ModeDeveloppeur_OFF False
End Sub



