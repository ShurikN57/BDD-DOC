Attribute VB_Name = "zDocWindows"
Option Explicit

' =============================================
' Conversion des liens fichier vers chemins Windows
' Contrat :
' - accepte les formats locaux standards :
'     file:///C:/dossier/fichier.pdf
'     C:\dossier\fichier.pdf
'     \\serveur\partage\fichier.pdf
' - accepte aussi les formats UNC :
'     file://serveur/partage/fichier.pdf
' - renvoie le texte d'origine si aucun cas reconnu
' =============================================

Public Function FileUrlToWindowsPath(ByVal fileUrl As String) As String

    Dim s As String
    Dim rest As String
    Dim p As Long
    Dim host As String
    Dim pathPart As String

    s = Trim$(fileUrl)
    If s = "" Then Exit Function

    s = Replace(s, "\", "/")

    ' ===== Déjà un chemin UNC =====
    If Left$(s, 2) = "//" Then
        FileUrlToWindowsPath = Replace(s, "/", "\")
        Exit Function
    End If

    ' ===== Déjà un chemin local Windows =====
    If Len(s) >= 3 Then
        If Mid$(s, 2, 2) = ":/" Or Mid$(s, 2, 2) = ":\\" Then
            s = UrlDecodeUtf8(s)
            FileUrlToWindowsPath = Replace(s, "/", "\")
            Exit Function
        End If
    End If

    ' ===== Cas file:///C:/... =====
    If LCase$(Left$(s, 8)) = "file:///" Then
        rest = Mid$(s, 9)
        rest = UrlDecodeUtf8(rest)
        rest = Replace(rest, "/", "\")
        FileUrlToWindowsPath = rest
        Exit Function
    End If

    ' ===== Cas file://serveur/partage/... =====
    If LCase$(Left$(s, 7)) = "file://" Then
        rest = Mid$(s, 8)

        p = InStr(1, rest, "/")
        If p > 0 Then
            host = Left$(rest, p - 1)
            pathPart = Mid$(rest, p + 1)

            host = Trim$(host)
            pathPart = UrlDecodeUtf8(pathPart)
            pathPart = Replace(pathPart, "/", "\")

            If host <> "" Then
                FileUrlToWindowsPath = "\\" & host & "\" & pathPart
                Exit Function
            End If
        End If
    End If

    ' ===== Cas file:C:/... ou file:/C:/... =====
    If LCase$(Left$(s, 5)) = "file:" Then
        rest = Mid$(s, 6)

        Do While Left$(rest, 1) = "/"
            rest = Mid$(rest, 2)
        Loop

        rest = UrlDecodeUtf8(rest)

        If Len(rest) >= 2 Then
            If Mid$(rest, 2, 1) = ":" Then
                FileUrlToWindowsPath = Replace(rest, "/", "\")
                Exit Function
            End If
        End If
    End If

    FileUrlToWindowsPath = fileUrl

End Function

Public Function UrlDecodeUtf8(ByVal txt As String) As String

    Dim i As Long
    Dim b() As Byte
    Dim n As Long
    Dim hexVal As String
    Dim ch As String

    If txt = "" Then
        UrlDecodeUtf8 = ""
        Exit Function
    End If

    ReDim b(0 To Len(txt) - 1)
    i = 1

    Do While i <= Len(txt)
        ch = Mid$(txt, i, 1)

        If ch = "%" And i + 2 <= Len(txt) Then
            hexVal = Mid$(txt, i + 1, 2)

            If hexVal Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                b(n) = CByte("&H" & hexVal)
                n = n + 1
                i = i + 3
            Else
                b(n) = AscW(ch) And &HFF
                n = n + 1
                i = i + 1
            End If

        ElseIf ch = "+" Then
            b(n) = 32
            n = n + 1
            i = i + 1

        Else
            b(n) = AscW(ch) And &HFF
            n = n + 1
            i = i + 1
        End If
    Loop

    If n = 0 Then
        UrlDecodeUtf8 = ""
        Exit Function
    End If

    ReDim Preserve b(0 To n - 1)
    UrlDecodeUtf8 = StrConv(b, vbUnicode)

End Function
