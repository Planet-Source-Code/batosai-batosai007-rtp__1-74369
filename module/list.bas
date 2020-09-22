Attribute VB_Name = "list"
Public dofilemon As Boolean
Public dopromp  As Boolean
Public dowpromp As Boolean
Public doStartup As Boolean
Public doprose As Boolean
Public doalamat As String
Public SystemFiles() As String


Public Sub loadSettings()
doalamat = GetSetting("Triyan", "Ganteng", "alamatbat")
 doStartup = GetSetting("Triyan", "Ganteng", "start", 1)
 dofilemon = GetSetting("Triyan", "Ganteng", "filemon", 1)
 dopromp = GetSetting("Triyan", "Ganteng", "promp", 1)
 dowpromp = GetSetting("Triyan", "Ganteng", "wpromp", 1)
 doprose = GetSetting("Triyan", "Ganteng", "prosemon", 1)
 If dofilemon Then
  Form1.Timer1.Enabled = True
 Else
  Form1.Timer1.Enabled = False
 End If
If doprose Then
Form1.Check1.Value = vbChecked
Form1.cek.Enabled = True
    If dopromp Then
    Form1.cek.Enabled = True
    Else
    Form1.cek.Enabled = False
     End If
 
    If dowpromp Then
     Form1.cek.Enabled = True
    Else
    
     End If
Else
Form1.cek.Enabled = False
Form1.Check1.Value = vbUnchecked
End If
sapi:
End Sub

Public Function isSystemFile(ByVal Filename As String) As Boolean
On Error GoTo bailout
 isSystemFile = False
 Filename = LCase(Filename)
 If UBound(SystemFiles) = -1 Then Exit Function
 For i = 0 To UBound(SystemFiles)
' Debug.Print filename, SystemFiles(i)
  If Filename = SystemFiles(i) Then isSystemFile = True
 Next
bailout:
 If err.Number <> 0 Then isSystemFile = True '<< if a failure mark all files as protected (we cannot risk closing a system files)
End Function
Public Sub LoadSystemFilesINI()
ReDim SystemFiles(0) As String
Dim l() As String, data As String
Dim ff As Long

ff = FreeFile
Open App.path & "\system.pif" For Input As #ff
    data = Input(LOF(ff), #ff)
    Close #ff

l() = Split(data, vbCrLf)
For i = 0 To UBound(l)
 If Trim(l(i)) <> "" Then
  l(i) = LCase(l(i))
  l(i) = Replace(l(i), "{apppath}", App.path & "\")
  l(i) = Replace(l(i), "{windows}", Environ("windir") & "\")
  l(i) = Replace(l(i), "\\", "\")
  l(i) = Replace(l(i), " ", "")
  l(i) = LCase(l(i))
  
  ReDim Preserve SystemFiles(UBound(SystemFiles) + 1) As String
  SystemFiles(UBound(SystemFiles)) = l(i)
 End If
Next
End Sub

Public Function Block(ByVal Filename As String)
 Filename = LCase(Filename)
 If Not isBlocked(Filename) Then
  Dim ff As Long
  ff = FreeFile
  
  Open App.path & "\blocked.ini" For Append As #ff
    Print #ff, Filename
    Close #ff
  LoadBlockedFilesINI
 End If
End Function

Public Function Unblock(ByVal Filename As String)
 Filename = LCase(Filename)
 If isBlocked(Filename) Then
  Dim ff As Long: ff = FreeFile
  
  Open App.path & "\blocked.ini" For Output As #ff
  For i = 0 To UBound(BlockedFiles)
    If BlockedFiles(i) <> Filename Then Print #ff, BlockedFiles(i)
  Next
    Close #ff
    LoadBlockedFilesINI
 End If
End Function
Public Function isBlocked(ByVal Filename As String) As Boolean
On Error GoTo bailout
 Filename = LCase(Filename)
 
 For i = 0 To UBound(BlockedFiles)
  If BlockedFiles(i) = Filename Then isBlocked = True
 Next
bailout:
End Function
Public Sub LoadBlockedFilesINI()

End Sub

