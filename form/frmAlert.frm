VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batosai Antivirus ALERT!!!!!!!!"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8445
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAlert.frx":2A8B2
   ScaleHeight     =   2865
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin batosai007rtps.DMSXpButton cmdSecure 
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   2280
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
      Caption         =   "&Quarantina"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin batosai007rtps.DMSXpButton cmdIgnore 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2280
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
      Caption         =   "&Ignore"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin batosai007rtps.DMSXpButton cmdRemove 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   2280
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
      Caption         =   "&Remove"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H007A7C7E&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   6960
      ScaleHeight     =   555
      ScaleWidth      =   525
      TabIndex        =   3
      Top             =   840
      Width           =   525
   End
   Begin VB.Line lline 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   8040
      X2              =   8040
      Y1              =   2040
      Y2              =   480
   End
   Begin VB.Line lline 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   480
      X2              =   8040
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lline 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   2040
   End
   Begin VB.Line lline 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   480
      X2              =   8040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Virus found!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub BuildAlert()

    On Error Resume Next
    Virus.Filename = Replace(Virus.Filename, "\\", "\")
    lblText(1).Caption = Virus.Filename
    lblText(1).ToolTipText = Virus.Filename & "  (" & FileLen(Virus.Filename) & " Bytes )"
    lblText(2).Caption = Virus.Reason
    lblText(8).Caption = FileLen(Virus.Filename) & " Bytes"
    If Virus.Type = Executable Then
        lblText(7).Caption = "Executable File"
    End If
    If Virus.Type = Script Then
        lblText(7).Caption = "Script"
    End If
    picIcon.Picture = LoadIcon(Large, Virus.Filename)
    On Error GoTo 0

End Sub

Private Sub cmdIgnore_Click()

    Log "Alert ignored: " & Virus.Reason, 2
    Unload Me

End Sub

Private Sub cmdRemove_Click()

    Log "File removed: " & Virus.Filename, 2
   
    RemoveFile (Virus.Filename)
Unload Me
End Sub





Private Sub cmdSecure_Click()
Dim sXor As New clsSimpleXOR

    On Error Resume Next
    sXor.EncryptFile Virus.Filename, Virus.Filename, "Batosai_cakep_getoloh"
    Set sXor = Nothing
    MkDir App.path & "\Tahanan\"
    SetAttr (App.path & "\Tahanan\"), vbHidden + vbSystem
    FileCopy Virus.Filename, App.path & "\Tahanan\" & Mid$(Virus.FileNameShort, 1, Len(Virus.FileNameShort)) & ".batosai"
    SetAttr (Virus.Filename), vbNormal
    Kill Virus.Filename
    With frmSecFiles
        .Visible = False
        .Show
        SaveSetting "Triyan", "Ganteng", "Quarintine", .flSec.ListCount
         End With 'frmSecFiles
    Unload frmSecFiles
    Log "File moved to quarintine: " & Virus.Filename, 2
    On Error GoTo 0
    Unload Me
End Sub

Private Sub Form_Load()
Debug.Print Time
  Dim R1  As RECT
  Dim R2  As RECT
  Dim TPP As Integer

    TPP = Screen.TwipsPerPixelX
    Call SetRect(R1, Screen.Width / TPP, Screen.Height / TPP, Screen.Width / TPP, Screen.Height / TPP)
    Call SetRect(R2, 0, 0, Me.Width / TPP, Me.Height / TPP)
    Call DrawAnimatedRects(Me.hwnd, IDANI_CLOSE Or IDANI_CAPTION, R1, R2)
    BuildAlert
    KeepOnTop Me
    
    DoEvents
    BeepAlert
End Sub

Private Sub BeepAlert()
    Beep 4000, 220
    Beep 3000, 200
    Beep 4000, 220
    Beep 3000, 200
        Beep 4000, 220
    Beep 3000, 200

End Sub

Private Sub hpOnline_Click()

  'http://www.viruslist.com/eng/viruslistfind.html?findTxt=code+red

   ' Call ShellExecute(Me.hwnd, "Open", "http://www.viruslist.com/eng/viruslistfind.html?findTxt=" & Replace(Virus.Reason, " ", "+"), vbNullString, "c:\", 1)

End Sub


