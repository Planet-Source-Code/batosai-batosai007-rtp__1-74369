Attribute VB_Name = "mod_function"
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As SAVEFILENAME) As Long
Private Type SAVEFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Function ShowSave(hwnd As Long, Optional extFile As String = "Application files|*.exe") As String
    Dim OFName As SAVEFILENAME
    extFile = Replace(extFile, "|", Chr(0))
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = extFile
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Open"
    OFName.flags = 0
    If GetSaveFileName(OFName) Then
       ShowSave = Trim$(OFName.lpstrFile)
    Else
       ShowSave = ""
    End If
End Function

