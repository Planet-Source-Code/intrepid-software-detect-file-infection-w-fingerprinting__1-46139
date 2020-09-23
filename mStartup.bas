Attribute VB_Name = "mStartup"
Declare Function SHGetPathFromIDList Lib "shell32" (Pidl As Long, ByVal FolderPath As String) As Long
    Global Const CSIDL_COMMON_STARTUP = 24
    Global Const MAX_PATH = 260
Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As Long, ByVal nFolder As Long, Pidl As Long) As Long
Public Function StartupMenu() As String
    Dim lpStartupPath As String * MAX_PATH
    Dim Pidl As Long
    Dim hResult As Long
    
    hResult = SHGetSpecialFolderLocation(0, CSIDL_COMMON_STARTUP, Pidl)


    If hResult = 0 Then
        hResult = SHGetPathFromIDList(ByVal Pidl, lpStartupPath)


        If hResult = 1 Then
            lpStartupPath = Left(lpStartupPath, InStr(lpStartupPath, Chr(0)) - 1)
            StartupMenu = lpStartupPath
        End If
    End If
End Function
