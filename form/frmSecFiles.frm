VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSecFiles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Files Tahanan"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "frmSecFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin batosai007rtps.DMSXpButton cmdRemove 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
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
   Begin batosai007rtps.DMSXpButton cmdDesecure 
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "&Desecure"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.FileListBox flSec 
      Height          =   1455
      Left            =   8640
      Pattern         =   "*.batosai"
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvQuarintine 
      Height          =   3375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   6455
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Extention"
         Object.Width           =   3810
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "File Quarantine"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   0
      Picture         =   "frmSecFiles.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSecFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdDesecure_Click()
Dim tensi As String
Dim sXor As New clsSimpleXOR
Dim cop As String
Dim cop1 As String
Dim alamat As String
tensi = "All Files"

    If MsgBox("Yakin neh Maw bebasin Tahanan?", vbYesNo + vbCritical) = vbYes Then
        Log "File desecured: " & flSec.Filename, 2
        alamat = ShowSave(0, tensi)
        If alamat = "" Then Exit Sub
        
        cop = App.path & "\Tahanan\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1) & ".batosai"
      cop1 = App.path & "\Tahanan\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1)
            FileCopy cop, cop1
            
            Kill cop
            sXor.DecryptFile cop1, cop1, "Batosai_cakep_getoloh"
             FileCopy cop1, alamat
            
            Kill cop1
         'APP
        Set sXor = Nothing
        flSec.Refresh
        Call Form_Load
        SaveSetting "Triyan", "Ganteng", "Quarintine", flSec.ListCount
        
    End If

End Sub

Private Sub cmdRemove_Click()
Dim cop As String
cop = App.path & "\Tahanan\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1) & ".batosai"

    RemoveFile cop

End Sub



Private Sub Form_Load()

  Dim CurEntry As ListItem
  Dim Counter  As Long

    lvQuarintine.ListItems.Clear
    Me.flSec.path = App.path & "\Tahanan\"
    flSec.Refresh
    For Counter = 0 To flSec.ListCount - 1
        Set CurEntry = Me.lvQuarintine.ListItems.Add
        CurEntry.Text = Mid$(flSec.list(Counter), 1, InStr(1, flSec.list(Counter), ".") - 1)
        CurEntry.SubItems(1) = Mid$(flSec.list(Counter), InStr(1, flSec.list(Counter), ".") + 1, 3)
    Next '  COUNTER

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Triyan", "Ganteng", "Quarintine", flSec.ListCount
    

End Sub

