VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batosai Prototipe program"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   2040
      Top             =   2400
   End
   Begin batosai007rtps.DMSXpButton DMSXpButton2 
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "sitry"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin batosai007rtps.DMSXpButton DMSXpButton1 
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Close"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Allow active File Monitoring"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Run on Startup"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Allow Monitor Process"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame aktif 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monitor Process"
      Height          =   975
      Left            =   3240
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&With prompt"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Withour prompt"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Timer tmrProcessRefresh 
      Interval        =   1
      Left            =   2400
      Top             =   6000
   End
   Begin VB.Timer cek 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   360
      Top             =   6000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Antivirus Info"
      Height          =   975
      Left            =   2880
      TabIndex        =   0
      Top             =   2880
      Width           =   2535
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   3360
         Top             =   1440
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Files Scan:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Data base:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "0.0.0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Database Release:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstProcess 
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL  OPTION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   240
      Picture         =   "Form1.frx":2A8B2
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label rubah 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Menu pop 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mShow 
         Caption         =   "&Show Form"
      End
      Begin VB.Menu mQua 
         Caption         =   "&Quarantina"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu batas 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONUP = &H205
Const WM_DBLLEFTBUTTON = &H203
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Boolean

Dim nid As NOTIFYICONDATA
Dim iewindow As InternetExplorer
Private currentwindows     As New ShellWindows

Private Sub cek_Timer()
If rubah = Label2 Then Exit Sub
If rubah.Caption < Label2.Caption Then kecil
If rubah.Caption > Label2.Caption Then
If Label2.Caption <> "" Then opt
End If
End Sub

Private Sub Check1_Click()
If (Check1.Value = vbChecked) Then
cek.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
 SaveSetting "Triyan", "Ganteng", "prosemon", "1"
Else
cek.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
 SaveSetting "Triyan", "Ganteng", "prosemon", "0"
End If
End Sub

Private Sub Check2_Click()
If (Check2.Value = vbChecked) Then
            SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname, App.path & "\" & App.exeName & ".exe /B", 1
             SaveSetting "Triyan", "Ganteng", "start", "1"

                    Else 'NOT LBLTEXT(8).CAPTION...
            DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname
             SaveSetting "Triyan", "Ganteng", "start", "0"

                   End If
End Sub

Private Sub Check3_Click()
If (Check3.Value = vbChecked) Then
Timer1.Enabled = True
 SaveSetting "Triyan", "Ganteng", "filemon", "1"
Else
Timer1.Enabled = False
SaveSetting "Triyan", "Ganteng", "filemon", "0"
End If

End Sub



Private Sub Command1_Click()
     
End Sub

Private Sub DMSXpButton1_Click()
On Error GoTo a

Me.Visible = False
If Check1.Value = (Check1.Value = vbChecked) Then
If Check3.Value = (Check3.Value = vbChecked) Then
SetAttr (App.path & "\system.pif"), vbReadOnly + vbSystem
End
End If
End If


a:
End Sub

Private Sub DMSXpButton2_Click()
With nid
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle
.szTip = "Right Click to show menu" & vbNullChar
.cbSize = Len(nid)
End With
 Shell_NotifyIcon NIM_ADD, nid
 
End Sub

Private Sub Form_Load()
Dim abc As Integer
 Call LoadSystemFilesINI
 Call setupColumbs
 Call RefreshProcessess
 Call LoadBlockedFilesINI
 rubah.Caption = lstProcess.ListItems.count
 
 abc = rubah.Caption
 Label3.Caption = lstProcess.ListItems(abc).SubItems(2)
 Label2.Caption = rubah.Caption
Timer1.Enabled = False
DMSXpButton2_Click
loadSettings
upgradechecks
End Sub

Private Sub mAbout_Click()
frmAbout.Show
End Sub

Private Sub mExit_Click()
If MsgBox(" Yakin mau keluar", vbExclamation + vbYesNo, "Warning") = vbYes Then
     End
     Else
     Exit Sub
    End If

End Sub

Private Sub mQua_Click()
frmSecFiles.Show
End Sub

Private Sub mShow_Click()
Form1.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
 SaveSetting "Triyan", "Ganteng", "promp", "1"
 SaveSetting "Triyan", "Ganteng", "wpromp", "0"
 Else
 SaveSetting "Triyan", "Ganteng", "promp", "0"
 SaveSetting "Triyan", "Ganteng", "wpromp", "1"

 End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
 SaveSetting "Triyan", "Ganteng", "promp", "0"
 SaveSetting "Triyan", "Ganteng", "wpromp", "1"
 Else
 SaveSetting "Triyan", "Ganteng", "promp", "1"
 SaveSetting "Triyan", "Ganteng", "wpromp", "0"

 End If
End Sub

Private Sub Timer1_Timer()
Dim buffer, ValidData As String
Dim c As Collection
Dim currentlocation As String
    On Error Resume Next

    Timer1.Enabled = False
    For Each iewindow In currentwindows
        DoEvents
        If iewindow.Busy Then
            GoTo busysignal
        End If
        currentlocation = iewindow.LocationURL
        ValidData = InStr(1, buffer, iewindow.LocationName & "|" & iewindow.LocationURL & "|")
        If ValidData = 0 Then
            If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "%60", "`")
                 currentlocation = Replace(currentlocation, "%23", "#")
                 currentlocation = Replace(currentlocation, "%25", "%")
                 currentlocation = Replace(currentlocation, "%5E", "^")
                 currentlocation = Replace(currentlocation, "%26", "&")
                 currentlocation = Replace(currentlocation, "%7D", "}")
                  currentlocation = Replace(currentlocation, "%7B", "{")
                   currentlocation = Replace(currentlocation, "%5B", "[")
                 currentlocation = Replace(currentlocation, "%5D", "]")
                 currentlocation = Replace(currentlocation, "/", "\")
                 currentlocation = Replace(currentlocation, "\\", "\")
                   FullPathSearch currentlocation, c
                   Debug.Print currentlocation
            End If
        End If
busysignal:
        
    Next
    Timer1.Enabled = True
    On Error GoTo 0
End Sub

Private Sub Timer2_Timer()
DMSXpButton2_Click
End Sub

Private Sub tmrProcessRefresh_Timer()
 Call RefreshProcessess
 rubah.Caption = lstProcess.ListItems.count
End Sub
Private Function besar()

Dim anka1 As String
Dim anka2 As Integer
Dim per As Integer
Dim abc As Integer
Dim i As Integer
Dim regiu As String
regiu = GetSetting("Triyan", "Ganteng", "alamatbat")
regiu = LCase$(regiu)
anka1 = rubah.Caption
per = Label2.Caption
anka2 = per + 1
 For i = anka2 To anka1
 If (lstProcess.ListItems(i).SubItems(1)) <> regiu Then
  If Not isSystemFile(lstProcess.ListItems(i).SubItems(1)) Then
 SuspendThreads (lstProcess.ListItems(i).SubItems(2))
 Set frm = New frmNew
 frm.lblPID = (lstProcess.ListItems(i).SubItems(2))
 frm.lblParent = (lstProcess.ListItems(i).SubItems(1))
 frm.lblProcName = Title(frm.lblParent)
 frm.lblFilename = frm.lblProcName
 frm.Picture1.Picture = LoadIcon(Large, frm.lblParent)
          frm.Show
 End If
 End If
 Next i
 Label2.Caption = rubah.Caption
 abc = rubah.Caption
 Label3.Caption = lstProcess.ListItems(abc).SubItems(2)
tes:
End Function
Private Function kecil()
Label2.Caption = rubah.Caption

End Function
Sub setupColumbs()
 'lstProcess.ColumnHeaders.Add , , "Icon"
 'lstProcess.ColumnHeaders.Add , , "Threat"
 lstProcess.ColumnHeaders.Add , , " ", 10
 lstProcess.ColumnHeaders.Add , , "Image", 4000
 'lstProcess.ColumnHeaders.Add , , "Path"
 lstProcess.ColumnHeaders.Add(, , "Memory", 1000).Alignment = lvwColumnRight


 
End Sub

Sub RefreshProcessess()
 Dim i As Integer
 Dim Apps As New processlist
 Dim procInfo As New ProcessInfo
 
 'lstProcess.ListItems.Clear
 For i = 1 To lstProcess.ListItems.count '- 1
  lstProcess.ListItems(i).Tag = "0"
 Next
 
 For i = 0 To Apps.processCount - 1
  procInfo.newQuery Apps.ProcessHandle(i)
  'With lstProcess.ListItems.Add(, "k" & procInfo.ProcessID, procInfo.ImageName)
  '
  '  .ListSubItems.Add , , procInfo.filename
  '  .ListSubItems.Add , , FormatNumber((procInfo.MemoryUsage / 1000), 3, vbUseDefault, vbUseDefault, vbTrue) & " K"
  '  .ListSubItems.Add , , procInfo.StartTime
  '  .ListSubItems.Add , , procInfo.KernalTime
  'End With
  On Error Resume Next
   lstProcess.ListItems("k" & procInfo.ProcessID).Tag = "1"
   If err.Number = 35601 Then '<< element not found
        With lstProcess.ListItems.Add(, "k" & procInfo.ProcessID)
             
            
             .ListSubItems.Add , , " "
             .ListSubItems.Add , , " "
             
        End With
   End If
  err.Clear
  On Error GoTo 0
  
  With lstProcess.ListItems("k" & procInfo.ProcessID)
    .Tag = "1"
    .ListSubItems(1) = procInfo.Filename
  
    .ListSubItems(2) = procInfo.ProcessID
   
   
    
  End With
 Next

 For i = 1 To lstProcess.ListItems.count '- 1
  If lstProcess.ListItems(i).Tag = "0" Then lstProcess.ListItems.Remove i
  If i >= lstProcess.ListItems.count Then Exit For
 Next
 
 Set procInfo = Nothing
 Set Apps = Nothing
End Sub
Private Function cec()

Dim anka1 As String
Dim anka2 As Integer
Dim per As Integer
Dim abc As Integer
Dim i As Integer
anka1 = rubah.Caption
per = Label2.Caption
anka2 = per + 1
 For i = anka2 To anka1
  If Not isSystemFile(lstProcess.ListItems(i).SubItems(1)) Then
SuspendThreads (lstProcess.ListItems(i).SubItems(2))
test123 (lstProcess.ListItems(i).SubItems(1)), (lstProcess.ListItems(i).SubItems(2))
 End If
 Next i
 Label2.Caption = rubah.Caption
 abc = rubah.Caption
 Label3.Caption = lstProcess.ListItems(abc).SubItems(2)
tes:
End Function
Private Function opt()
If (Option1.Value = True) Then
besar
Else
cec
End If
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngMsg As Single
sngMsg = X / Screen.TwipsPerPixelX
If sngMsg = WM_RBUTTONUP Then
Me.PopupMenu pop
End If
If sngMsg = WM_DBLLEFTBUTTON Then
pangil
End If
End Sub
Public Sub upgradechecks()
 If doStartup Then Check2.Value = vbChecked Else Check2.Value = vbUnchecked
 If dofilemon Then Check3.Value = vbChecked Else Check3.Value = vbUnchecked
 If dopromp Then Option1.Value = True Else Option1.Value = False
 If dowpromp Then Option2.Value = True Else Option2.Value = False
End Sub
Private Function pangil()
On Error Resume Next

 CreateStringValue HKEY_LOCAL_MACHINE, "Software\classes\exefile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
   CreateStringValue HKEY_CLASSES_ROOT, "exefile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
Shell doalamat, vbNormalFocus
End Function
