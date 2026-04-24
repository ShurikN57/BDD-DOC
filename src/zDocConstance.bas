Attribute VB_Name = "zDocConstance"
Option Explicit

' =============================================
' FEUILLES
' =============================================
Public Const SHEET_MAIN As String = "Base"
Public Const SHEET_TITRES As String = "titres"
Public Const SHEET_MENU_DEROULANT As String = "MENU DEROULANT"

' =============================================
' MOT DE PASSE / SECURITE
' =============================================
' IMPORTANT :
' - Ne pas versionner de secret en clair dans le code.
' - Définir la variable d'environnement BDD_DOC_DEV_PASSWORD
'   sur les postes autorisés.
Private Const ENV_MDP_DEV As String = "BDD_DOC_DEV_PASSWORD"

Public Function MotDePasseDeveloppeur() As String

    Static cacheMdp As String

    If Len(cacheMdp) > 0 Then
        MotDePasseDeveloppeur = cacheMdp
        Exit Function
    End If

    cacheMdp = Trim$(Environ$(ENV_MDP_DEV))
    MotDePasseDeveloppeur = cacheMdp

End Function

Public Function ExigerMotDePasseDeveloppeur(Optional ByVal contexte As String = "") As Boolean

    If Len(MotDePasseDeveloppeur()) > 0 Then
        ExigerMotDePasseDeveloppeur = True
        Exit Function
    End If

    MsgBox "Mot de passe développeur non configuré." & vbCrLf & _
           "Définissez la variable d'environnement " & ENV_MDP_DEV & "." & _
           IIf(Len(contexte) > 0, vbCrLf & "Contexte : " & contexte, ""), vbExclamation

End Function

' =============================================
' SESSION
' =============================================
Public Const CELL_NOM_SESSION As String = "A1"
Public Const CELL_DATE_SESSION As String = "B1"

' =============================================
' STRUCTURE
' =============================================
Public Const ROW_START As Long = 5
Public Const ROW_RECHERCHE As Long = 2
Public Const ROW_HEADER As Long = 4
Public Const ROW_TITRES As Long = 1

Public Const COL_FIRST As String = "A"
Public Const COL_LAST As String = "AB"
Public Const COL_LAST_RECHERCHE As String = "AB"

Public Const NB_COL_RECHERCHE As Long = 28
Public Const NB_COL_UI As Long = 28
Public Const NB_COL_TABLE As Long = 28

Public Const PLAGE_RECHERCHE As String = "A2:AB2"

' =============================================
' COLONNES DOC - IDENTIFICATION RF
' =============================================
Public Const COL_RF As String = "A"
Public Const COL_TR As String = "B"
Public Const COL_SYST As String = "C"
Public Const COL_NUM As String = "D"
Public Const COL_BIGR As String = "E"
Public Const COL_C1 As String = "F"
Public Const COL_C2 As String = "G"

' =============================================
' COLONNES DOC - OUVERTURE / ANALYSE
' =============================================
Public Const COL_CHEMIN_ANALYSE As String = "H"
Public Const COL_REF As String = "I"
Public Const COL_PAGE As String = "J"
Public Const COL_FORMAT_PAGE As String = "K"
Public Const COL_LIGNE_EXCEL As String = "L"
Public Const COL_TYPE_DOC As String = "M"
Public Const COL_SOURCE As String = "N"
Public Const COL_LIENS_DOCUMENTS As String = "O"

' =============================================
' COLONNES DOC - POWER QUERY / FORMULES
' =============================================
Public Const COL_LETTRE_TR As String = "P"
Public Const COL_LETTRE_NUM As String = "Q"
Public Const COL_CAR_NUM As String = "R"
Public Const COL_DEBUT_NUM As String = "S"
Public Const COL_NUM_IJK As String = "T"
Public Const COL_ID As String = "U"
Public Const COL_NB_PAGES As String = "V"
Public Const COL_PAGES As String = "W"
Public Const COL_DOC As String = "X"

' =============================================
' SUIVI AGENT / CONTROLE
' =============================================
Public Const COL_DATE As String = "Y"
Public Const COL_NOM As String = "Z"
Public Const COL_CONF As String = "AA"
Public Const COL_OBS As String = "AB"

' =============================================
' FACTORISATION - EDITION / PROTECTION
' =============================================
' BDD-DOC : seules les colonnes AA:AB sont éditables par l'agent.
' Les colonnes A:X sont en lecture seule (données Power Query).
' Les colonnes Y:Z sont remplies automatiquement par VBA.
' PLAGE_EDITABLE_MAIN et PLAGE_EDITABLE_AIDE restent vides
' car elles sont testées par les modules partagés (DeveloppeurOnOFF).
Public Const PLAGE_EDITABLE_MAIN As String = ""      ' pas de zone éditable principale (pas de saisie RF)
Public Const PLAGE_EDITABLE_SUIVI As String = "AA:AB" ' conformité + observations
Public Const PLAGE_EDITABLE_AIDE As String = ""       ' pas de colonne aide dans BDD-DOC

Public Const COL_VALIDATION_CONF As String = "AA"

' =============================================
' FACTORISATION - COLLAGE
' =============================================
' BDD-DOC : le collage n'est autorisé que dans la barre de recherche
' et dans les colonnes suivi (AA:AB).
' PLAGE_COLLER_EDITABLE reste vide car les données A:X ne sont pas
' modifiables par collage (modules partagés testent cette constante).
Public Const PLAGE_COLLER_RECHERCHE As String = "A2:AB2"
Public Const PLAGE_COLLER_EDITABLE As String = ""     ' pas de collage dans les données RF
Public Const PLAGE_COLLER_SUIVI As String = "AA:AB"

' =============================================
' VALEURS CONFORMITE
' =============================================
Public Const VAL_CONF_1 As String = "conforme"
Public Const VAL_CONF_2 As String = "non conforme"
Public Const VAL_CONF_3 As String = "lien incorrect"

Public Const MSG_VALEURS_CONF As String = "Valeurs autorisées : conforme, non conforme ou lien incorrect."

' =============================================
' BOUTON RECHERCHE RF PRINCIPAL
' =============================================
Public Const COL_RF_CONCAT As String = "A"
Public Const COL_RECHERCHE_EXACTE As String = COL_RF_CONCAT

' =============================================
' BOUTON PREMIERE LIGNE VIDE / COLONNE MASQUEE
' =============================================
Public Const COL_PREMIERE_LIGNE_VIDE As String = "AA"
Public Const COLONNES_MASQUEES As String = "P:X"

' =============================================
' BOUTON ZOOM
' =============================================
Public Const ZOOM_ECRAN_PRINCIPAL As Long = 89
Public Const ZOOM_ECRAN_SECONDAIRE As Long = 104
' =============================================
' USERFORM UF-AGENT
' =============================================
Public Const COL_AGENTS As String = "C"
Public Const ROW_AGENTS_START As Long = 2
Public Const ROW_AGENTS_END As Long = 44

' =============================================
' COMPORTEMENT
' =============================================
' Flags utilisés par les modules partagés pour activer/désactiver
' des fonctionnalités selon le classeur (BDD-RF ou BDD-DOC).
Public Const HAS_MAIN_EDIT_ZONE As Boolean = False     ' BDD-DOC : pas de saisie RF (colonnes A:X en lecture seule)
Public Const HAS_AIDE_ZONE As Boolean = False           ' BDD-DOC : pas de colonne aide
Public Const HAS_DOC_DOUBLECLICK As Boolean = True      ' BDD-DOC : double-clic ouvre le document lié
Public Const HAS_RECHERCHE_EXACTE As Boolean = True
Public Const HAS_COLONNES_MASQUEES As Boolean = True    ' colonnes P:X masquables
Public Const HAS_PREMIERE_LIGNE_VIDE As Boolean = True

' =============================================
' COULEURS
' =============================================
Public Const COLOR_RECHERCHE_FOND As Long = 16510410
Public Const COLOR_RECHERCHE_ACTIVE As Long = 15781618
Public Const COLOR_PLACEHOLDER As Long = 9868950
Public Const COLOR_TEXTE_NOIR As Long = 0
Public Const COLOR_BORDURE_BLEUE As Long = 10318348
Public Const COLOR_BORDURE_VIOLETTE As Long = 9644960
Public Const COLOR_ERREUR_ROUGE As Long = 9869055

' =============================================
' COULEURS CONFORMITE
' =============================================
Public Const COLOR_CONF_CONFORME As Long = 10675893      ' RGB(181, 230, 162) / #B5E6A2
Public Const COLOR_CONF_NON_CONFORME As Long = 14524132  ' RGB(228, 158, 221) / #E49EDD
Public Const COLOR_CONF_LIEN_INCORRECT As Long = 7531262 ' RGB(254, 234, 114) / #FEEA72

' =============================================
' SEUILS / SECURITE
' =============================================
Public Const MAX_SELECTION_CHANGE As Long = 500
Public Const SEUIL_CONFIRMATION_MASSE As Long = 2000
Public Const SEUIL_BLOCAGE_MASSE As Long = 15000
