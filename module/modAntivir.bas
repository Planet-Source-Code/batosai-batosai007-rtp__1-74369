Attribute VB_Name = "modAntivir"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Public Sub BuildUI()
   
    On Error Resume Next
    
        With Form1
            .lblText(3).Caption = AV.Signature.SignatureDate
            If CDate(AV.Signature.SignatureDate) < Date Then
                .lblText(3).ForeColor = vbRed
                
            End If
            .lblText(5).Caption = AV.Signature.SignatureCount
                
        End With 'FRMMAIN
    
    On Error GoTo 0

End Sub

Public Function CheckFile(ByVal strFilename As String) As Boolean
On Error GoTo err
  Dim strResult As String
  Dim temp()    As String
  Dim path      As String
  Dim c         As Collection

    CheckFile = False
    If UCase$(Mid$(strFilename, Len(strFilename) - 4, 4)) = ".ZIP" Then
       ' CheckDLL
       ' path = UnzipFile2RandomPath(strFilename)
        'modAntivir2.FullPathSearch path, c, , , , True
       ' DoEvents
       ' DelTree path
     Else 'NOT UCASE$(MID$(STRFILENAME,...
        If GetFileOI(strFilename) Then
            strResult = Search(strFilename)
            If strResult <> "NOTHING" Then
                With Virus
                    .Filename = strFilename
                    .Reason = strResult
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                Log "Virus found: " & Virus.Reason & " in " & Virus.Filename, 1, True
                 Form1.Label7.Caption = Form1.Label7.Caption + 1
                 frmAlert.Show
                CheckFile = True
            End If
        End If
    
    
        
        
        
        
        If IsFileaScript(strFilename) Then
        
            If SearchScript(strFilename) Then
                With Virus
                    .Filename = strFilename
                    .Reason = LoadResString(151)
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
               Form1.Label7.Caption = Form1.Label7.Caption + 1
               frmAlert.Show
                CheckFile = True
            End If
             
        End If
       
      
       
    End If
    Form1.Label7.Caption = Form1.Label7.Caption + 1
    
      BuildUI
    DoEvents
Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.Checkfile", strFilename
End Function

Public Function FileExist(ByVal strFilename As String) As Boolean


    On Error Resume Next
    FileExist = True
    If FileLen(strFilename) = 0 Then
        FileExist = False
    End If
    On Error GoTo 0

End Function

Public Function FileText(ByVal strFilename As String) As String
On Error GoTo err:
  Dim Handle As Long

    Handle = FreeFile
    Open strFilename For Binary As #Handle
    FileText = Space$(LOF(Handle))
    Get #Handle, , FileText
    Close #Handle
Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.Filetext", strFilename
End Function

Private Function IsWinNT() As Boolean
On Error GoTo err:
  Dim myOS As OSVERSIONINFO

    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)

Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.IsWinNT"

End Function

Public Sub KeepOnTop(f As Form)

  Const SWP_NOMOVE   As Long = 2
  Const SWP_NOSIZE   As Long = 1
  Const HWND_TOPMOST As Long = -1

    SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Function LoadIcon(size As Long, _
                         ByVal strFilename As String) As IPictureDisp

On Error GoTo err:
  Dim Result    As Long

  Dim file      As String
  Dim Unkown    As IUnknown
  Dim Icon      As IconType
  Dim CLSID     As CLSIdType
  Dim ShellInfo As ShellFileInfoType
    file = strFilename
    Call SHGetFileInfo(file, 0, ShellInfo, Len(ShellInfo), size)
    With Icon
        .cbSize = Len(Icon)
        .picType = vbPicTypeIcon
        .hIcon = ShellInfo.hIcon
    End With 'Icon
    CLSID.Id(8) = &HC0
    CLSID.Id(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    Set LoadIcon = Unkown
    
Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.LoadIcon", size & ":" & strFilename
End Function

Public Sub Main()

On Error GoTo err:
    If App.PrevInstance Then
        MsgBox "Only one instance allowed!", vbOKOnly, "Error"
        End
    End If
    With AV
        .AVname = "Batosai 007"
        .Signature.SignatureFilename = App.path & "\signatures.db"
        .Signature.SignatureOnlineFilename = "batosai007@ne.its.ac.id"
    End With 'AV
       ' Set the attributes,..
    BuildSigns
    CheckExe
    RegisterFile ".batosai", "Tahanan Perang", "Anti Virus", App.path & "\" & App.exeName & ".exe /R %1", App.path & "\tahanan.ico"   '"This file is secured by "
 
      Select Case UCase$(Left$(Command, 2))
     Case "/S"
        CheckFile (Mid$(Command, 3, Len(Command) - 3))
     Case vbNullString
     If Len(Dir(App.path & "\system.pif")) < 1 Then
         frmSplash.buat
        End If
        BuildUI
        frmSplash.Show
     Case "/B"
      If Len(Dir(App.path & "\system.pif")) < 1 Then
         frmSplash.buat
        End If
        frmSplash.Show
        frmSplash.Check1.Value = 1
        
     Case "/C"
           
     Case "/F"
      
     Case "/R"
           Case Else
           End Select
Exit Sub
err:
 ErrorFunc err.Number, err.Description, "modAntivir.Main"
End Sub

Public Sub RemoveFile(ByVal strFilename As String)

On Error GoTo err:
  Dim Files As String
  Dim SFO   As SHFILEOPSTRUCT

    DoEvents
    Files = strFilename & Chr$(0)
    Files = Files & Chr$(0)
    With SFO
        .hwnd = frmAlert.hwnd
        .wFunc = FO_DELETE
        .pFrom = Files
        .pTo = "" & Chr$(0)
    End With 'SFO
    Call SHFileOperation(SFO)
Exit Sub
err:
 ErrorFunc err.Number, err.Description, "modAntivir.RemoveFile", strFilename
End Sub
Public Function test123(ByVal strFilename As String, pidta As Long) As Boolean
On Error GoTo err
  Dim strResult As String
  Dim temp()    As String
  Dim path      As String
  Dim c         As Collection
 Dim lstProc As New processlist
   
    If UCase$(Mid$(strFilename, Len(strFilename) - 4, 4)) = ".ZIP" Then
       ' CheckDLL
       ' path = UnzipFile2RandomPath(strFilename)
        'modAntivir2.FullPathSearch path, c, , , , True
       ' DoEvents
       ' DelTree path
     Else 'NOT UCASE$(MID$(STRFILENAME,...
        If GetFileOI(strFilename) Then
            strResult = Search(strFilename)
            If strResult <> "NOTHING" Then
                With Virus
                    .Filename = strFilename
                    .Reason = strResult
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                Log "Virus found: " & Virus.Reason & " in " & Virus.Filename, 1, True
                frmAlert.Show
               
               lstProc.KillProcess pidta
                
            End If
            ResumeThreads pidta
        End If
        If IsFileaScript(strFilename) Then
        
            If SearchScript(strFilename) Then
                With Virus
                    .Filename = strFilename
                   
                    temp = Split(.Filename, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                 frmAlert.Show
                 lstProc.KillProcess pidta
               
            End If
             ResumeThreads pidta
        End If
       
      
       
    End If
      BuildUI
    DoEvents
Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.Checkfile", strFilename
End Function



