VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Access Attempt!"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   600
      Top             =   1080
   End
   Begin batosai007rtps.DMSXpButton Command1 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Allow"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin batosai007rtps.DMSXpButton Command2 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Terminate"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3720
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "New Access Attempt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   1200
      TabIndex        =   13
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALERT!"
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
      Height          =   435
      Left            =   360
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblProcName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[Process Name]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PID         :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPID 
      BackStyle       =   0  'Transparent
      Caption         =   "[PID]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Path        :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblParent 
      BackStyle       =   0  'Transparent
      Caption         =   "[Parent PID]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label lblFilename 
      BackStyle       =   0  'Transparent
      Caption         =   "[File Name]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Process Will be terminate in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tPID As Long
Dim tPName As String
Dim tempBypass As Boolean


Private Sub Command1_Click()
test123 lblParent.Caption, lblPID.Caption

Unload Me
End Sub

Private Sub Command2_Click()
Dim lstProc As New processlist
lstProc.KillProcess lblPID.Caption
    
    Unload Me
End Sub

Private Sub Form_Load()
KeepOnTop Me
End Sub

Private Sub tmr_Timer()
lblCount.Caption = lblCount.Caption - 1
    If lblCount.Caption = "0" Then
        Call Command2_Click
    End If
End Sub





