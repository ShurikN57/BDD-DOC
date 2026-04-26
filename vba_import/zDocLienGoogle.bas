Attribute VB_Name = "zDocLienGoogle"
' =============================================
'           Ouverture PDF vers Chrome
' =============================================

Option Explicit

Public Function GetChromePath() As String
    Dim cmd As String, exe As String

    ' 1) ChromeHTML (le plus fréquent)
    cmd = RegReadSafe("HKEY_CLASSES_ROOT\ChromeHTML\shell\open\command\")
    exe = ExtractExeFromCommand(cmd)
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    cmd = RegReadSafe("HKEY_CURRENT_USER\Software\Classes\ChromeHTML\shell\open\command\")
    exe = ExtractExeFromCommand(cmd)
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    ' 2) App Paths
    cmd = RegReadSafe("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")
    exe = ExtractExeFromCommand(cmd)
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    cmd = RegReadSafe("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")
    exe = ExtractExeFromCommand(cmd)
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    cmd = RegReadSafe("HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")
    exe = ExtractExeFromCommand(cmd)
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    ' 3) Uninstall (parfois utile)
    cmd = RegReadSafe("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome\InstallLocation")
    If cmd <> "" Then
        exe = cmd & "\Application\chrome.exe"
        If FileExists(exe) Then GetChromePath = exe: Exit Function
    End If

    cmd = RegReadSafe("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome\InstallLocation")
    If cmd <> "" Then
        exe = cmd & "\Application\chrome.exe"
        If FileExists(exe) Then GetChromePath = exe: Exit Function
    End If

    cmd = RegReadSafe("HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome\InstallLocation")
    If cmd <> "" Then
        exe = cmd & "\Application\chrome.exe"
        If FileExists(exe) Then GetChromePath = exe: Exit Function
    End If

    ' 4) Chemins standards (fallback)
    exe = Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe"
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    exe = Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe"
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    exe = Environ$("LocalAppData") & "\Google\Chrome\Application\chrome.exe"
    If FileExists(exe) Then GetChromePath = exe: Exit Function

    GetChromePath = ""
End Function

Private Function RegReadSafe(ByVal keyPath As String) As String
    On Error GoTo Fin
    RegReadSafe = CreateObject("WScript.Shell").RegRead(keyPath)
    Exit Function
Fin:
    RegReadSafe = ""
End Function

Private Function ExtractExeFromCommand(ByVal cmd As String) As String
    Dim p1 As Long, p2 As Long
    cmd = Trim$(cmd)
    If cmd = "" Then ExtractExeFromCommand = "": Exit Function

    ' Cas: "C:\...\chrome.exe" -- "%1"
    If Left$(cmd, 1) = """" Then
        p2 = InStr(2, cmd, """")
        If p2 > 2 Then
            ExtractExeFromCommand = Mid$(cmd, 2, p2 - 2)
            Exit Function
        End If
    End If

    ' Cas: C:\...\chrome.exe -- "%1"
    p1 = InStr(1, cmd, ".exe", vbTextCompare)
    If p1 > 0 Then
        ExtractExeFromCommand = Left$(cmd, p1 + 3)
        Exit Function
    End If

    ExtractExeFromCommand = ""
End Function

Private Function FileExists(ByVal p As String) As Boolean
    If Len(p) = 0 Then
        FileExists = False
    Else
        FileExists = (Dir(p) <> "")
    End If
End Function

