VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MaxLen = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const vbStar = "*"
Private Const vbAllFiles = "*.*"
Private Const vbBackslash = "\"
Private Const vbKeyDot = 46

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLen
    cShortFileName As String * 14
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Private Declare Function FindFirstFile Lib _
    "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib _
    "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib _
    "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileSize Lib _
    "kernel32" (ByVal hFile As Long, _
    lpFileSizeHigh As Long) As Long
Private Declare Function GetFileAttributes Lib _
    "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileName As String) As Long
Private Declare Function OpenFile Lib _
    "kernel32.dll" (ByVal lpFileName As String, _
    ByRef lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle As Long) As Long
Private Declare Function PathIsDirectory Lib _
    "shlwapi.dll" Alias "PathIsDirectoryA" _
    (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib _
    "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long
Private Declare Sub CloseHandle Lib _
    "kernel32" (ByVal hPass As Long)
Private Declare Function SetFileAttributes Lib _
    "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileName As String, _
    ByVal dwFileAttributes As Long) As Long
Private Declare Function DeleteFile Lib _
    "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long

Dim Wfd As WIN32_FIND_DATA, _
    hItem As Long, hFile As Long, _
    hScan As Long, nProgFiles As Long
Dim nAllFiles As Double
Dim FileSpec As String, UseFileSpec As Integer




Private Function StripNulls(sStr As String) As String

    StripNulls = Left$(sStr, InStr(1, sStr, Chr$(0)) - 1)
    
End Function

'--------------------------------------------------------------------------------------------------'
'Begin Calculate
Private Function asSearchFile(sPath As String)
   Dim ass As Integer
    Dim dirs As Integer, dirbuff() As String, _
        i As Integer
  
    DoEvents
    
    
    
    hItem = FindFirstFile(sPath & vbAllFiles, Wfd)
    
    If hItem <> INVALID_HANDLE_VALUE Then
        Do
            If (Wfd.dwFileAttributes And _
                FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(Wfd.cFileName) <> vbKeyDot Then
                    If (dirs Mod 10) = 0 Then _
                        ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls _
                        (Wfd.cFileName)
                End If
            ElseIf Not UseFileSpec Then
                nAllFiles = nAllFiles + 1
                
            End If
        Loop While FindNextFile(hItem, Wfd)
        Call FindClose(hItem)
    End If

    If UseFileSpec Then
       
    End If
    
    For i = 1 To dirs: asSearchFile sPath & dirbuff(i) & _
        vbBackslash: Next i
     Text1.Text = nAllFiles
     
End Function


'End Calculate
'--------------------------------------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------------'
'Begin Scanning File


Private Sub Command1_Click()
Dim a As String
nAllFiles = 0
a = Text2.Text
asSearchFile a
End Sub
