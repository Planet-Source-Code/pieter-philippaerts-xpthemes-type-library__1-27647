Attribute VB_Name = "mdlMain"
'ThemeTest created by The KPD-Team
'Copyright (c) 2001, The KPD-Team
'Visit our site at http://www.allapi.net/
'or email us at KPDTeam@allapi.net

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Returns True if themes are supported
Public Function AreThemesSupported() As Boolean
    Dim hLib As Long
    hLib = LoadLibrary("uxtheme.dll")
    If hLib <> 0 Then FreeLibrary hLib
    AreThemesSupported = Not (hLib = 0)
End Function
'Returns True if the current vwindows version is Windows NT 5.1
Public Function IsWindowsXP() As Boolean
    Dim OsInfo As OSVERSIONINFO
    GetVersionEx OsInfo
    If OsInfo.dwPlatformId = 2 Then 'It's NT
        IsWindowsXP = (OsInfo.dwMajorVersion = 5 And OsInfo.dwMinorVersion = 1) 'Version 5.1
    End If
End Function
