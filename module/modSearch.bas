Attribute VB_Name = "modSearch"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sign()               As String    'The Signatures will be loaded into this array
Private SignStr()            As String
Private SignVirusStringType() As String * 1
Private SignVirusName()      As String

Public Sub BuildSigns()

  'This builds the Signature - Array
  
  Dim sIn        As String
  Dim swords()   As String
  Dim X          As Long
  Dim Y          As Long

  Dim Data()     As String
    sIn = FileText(AV.Signature.SignatureFilename)
    swords = Split(sIn, vbCrLf)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        ReDim Preserve Sign(0 To X) As String
        ReDim Preserve SignVirusStringType(0 To X) As String * 1
        ReDim Preserve SignVirusName(0 To X) As String
        Data = Split(swords(X) & ":" & ":", ":")
        Sign(X) = UCase(Data(0))
        SignVirusStringType(X) = Data(1)
        SignVirusName(X) = Data(2)
        Y = X + 1
    Next X
    ReDim Preserve Sign(0 To X + 1) As String
    Sign(X + 1) = "#END#"
    AV.Signature.SignatureDate = Sign(0)
    AV.Signature.SignatureCount = UBound(swords) - 1

Exit Sub

    err
    MsgBox "An error has occured while loading the signature File!" & vbCrLf & "This could be caused by an empty or damaged file!" & vbCrLf & _
     "The error message was: " & err.Description, vbCritical + vbOKOnly, LoadResString(140)

End Sub
Private Function FindTermInFile(strFilename As String, strString As String, strFiletext As String) As Boolean

    FindTermInFile = False
    If InStr(1, strFiletext, strString, vbTextCompare) <> 0 Then FindTermInFile = True
    
End Function


Public Function Search(ByVal strFilename As String) As String

On Error GoTo err:
  Dim Current  As Long
  Dim crc      As String
  Dim strFiletext As String
  Dim Zeilen()   As String
Dim abca As String
    crc = CalcCRC(strFilename)
    strFiletext = Replace(CStr(FileText(strFilename)), "ß", "-")
    Debug.Print strFilename
    For Current = 1 To 4096
        If Sign(Current) = "#END#" Or LenB(Sign(Current)) = 0 Then
            GoTo Finish
        End If
        If SignVirusStringType(Current) = "E" Then
            If crc = Sign(Current) Then
                DoEvents
                Search = SignVirusName(Current)
                Log Search & " - " & strFilename, 1
                Exit Function
            
            Else: Search = "NOTHING"
            End If
        ElseIf SignVirusStringType(Current) = "S" Then
        
         abca = StrCheck(strFilename, hex2ascii(Sign(Current)))
If abca = True Then

                DoEvents
                Search = SignVirusName(Current)
                Log Search & " - " & strFilename, 1
                Exit Function
            Else: Search = "NOTHING"
            End If
        End If
        DoEvents
    Next Current
    Search = "NOTHING"
Finish:
Exit Function
err:
 ErrorFunc err.Number, err.Description, "modAntivir.Search", strFilename

End Function

Public Function SearchScript(ByVal strFilename As String) As Boolean

 
End Function

Function StrCheck(MyPath As String, StrText As String) As Boolean
On Error Resume Next
Dim filedata As String
Dim a As Integer
Open MyPath For Binary As #1
filedata = Space$(LOF(1))
Get #1, , filedata
If InStr(1, filedata, StrText) > 0 Then
StrCheck = True
Else
StrCheck = False
End If
Close #1
End Function
