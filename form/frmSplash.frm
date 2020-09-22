VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6060
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   6060
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox pid 
      Height          =   450
      Left            =   5640
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ListBox pros 
      Height          =   450
      Left            =   3960
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H003A4342&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5310
      TabIndex        =   0
      Top             =   2480
      Width           =   5370
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   1275
         TabIndex        =   1
         Top             =   0
         Width           =   1335
         Begin VB.Label Label1 
            BackColor       =   &H00008000&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Monotype Corsiva"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   7140
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   5760
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnt As Integer

Option Explicit

Private Sub lblCopyright_Click()

End Sub

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
test123 pros.list(a), pid.list(a)
Next a
End Sub



Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Picture2.Height = Picture1.Height
Picture2.BackColor = RGB(222, 125, 125)
Picture2.Width = 0
Timer1.Enabled = True
Timer1.Interval = 1
Command1_Click
End Sub

Private Sub Timer1_Timer()


Picture2.Refresh
Picture1.Refresh
Cnt = Cnt + 1
Label1.Caption = (Cnt) & "%"
If Cnt > 100 Then
Picture2.Width = Picture1.Width
Cnt = 0
frmSplash.Hide

cek1

Picture2.Width = Picture1.Width
Timer1.Enabled = False
Else
Picture2.Width = (Cnt / 100) * Picture1.Width
End If
End Sub
Function cek1()
If (Check1.Value = vbChecked) Then
Form1.Visible = False

Else
Form1.Show

End If

End Function
