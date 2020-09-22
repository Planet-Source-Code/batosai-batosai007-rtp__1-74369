VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Worm Hunter Generation 2"
   ClientHeight    =   4875
   ClientLeft      =   -1395
   ClientTop       =   -780
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30500
      Left            =   0
      ScaleHeight     =   30495
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   720
         Top             =   2880
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   240
         Top             =   2880
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&CLOSE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         MouseIcon       =   "frmAbout.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   29880
         Width           =   735
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1560
         X2              =   5280
         Y1              =   29640
         Y2              =   29640
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BATOSAI007@GMAIL.COM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1920
         MouseIcon       =   "frmAbout.frx":0A1C
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   29280
         Width           =   2835
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0D6E
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   43
         Top             =   28200
         Width           =   6540
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0DFF
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   240
         TabIndex        =   42
         Top             =   25800
         Width           =   6540
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1410
         X2              =   5490
         Y1              =   25440
         Y2              =   25440
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.planetsourcecode.com"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1372
         MouseIcon       =   "frmAbout.frx":0F25
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   25080
         Width           =   4170
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLANET SOURCE CODE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2272
         TabIndex        =   40
         Top             =   24720
         Width           =   2370
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1830
         X2              =   5070
         Y1              =   24480
         Y2              =   24480
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http:/WWW.JASAKOM.COM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1845
         MouseIcon       =   "frmAbout.frx":1277
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   24120
         Width           =   3000
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JASAKOM - SITE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   38
         Top             =   23760
         Width           =   1680
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1560
         X2              =   5400
         Y1              =   23520
         Y2              =   23520
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://VIROLGI.INFO"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2280
         MouseIcon       =   "frmAbout.frx":15C9
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   23160
         Width           =   2175
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIROLOGI -  SITE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   36
         Top             =   22800
         Width           =   1695
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DON'T FORGET TO VISIT"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1890
         TabIndex        =   35
         Top             =   21240
         Width           =   3135
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1080
         X2              =   5760
         Y1              =   22560
         Y2              =   22560
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.friendster.com/batosai007"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1200
         MouseIcon       =   "frmAbout.frx":191B
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   22200
         Width           =   4620
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MY FRIENDSTER"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         TabIndex        =   33
         Top             =   21840
         Width           =   1725
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HUTCHY"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3060
         TabIndex        =   32
         Top             =   19920
         Width           =   915
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIEUMI"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3045
         TabIndex        =   31
         Top             =   19560
         Width           =   780
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MERIAM 06"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   30
         Top             =   19200
         Width           =   1245
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GREETINGS TO:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2445
         TabIndex        =   29
         Top             =   18720
         Width           =   2025
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AND ALL THE OTHER PEOPLE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2002
         TabIndex        =   28
         Top             =   18120
         Width           =   2910
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFIF"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3240
         TabIndex        =   27
         Top             =   17760
         Width           =   450
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUKMA"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   26
         Top             =   17400
         Width           =   795
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SYAMSU"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   25
         Top             =   17040
         Width           =   930
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADIMAS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   24
         Top             =   16680
         Width           =   870
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BETA TESTERS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2490
         TabIndex        =   23
         Top             =   16200
         Width           =   1935
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2190
         X2              =   4710
         Y1              =   15960
         Y2              =   15960
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.allapi.net"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2182
         MouseIcon       =   "frmAbout.frx":1C6D
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   15600
         Width           =   2550
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VARIOUS FUNCTIONS - KPD-TEAM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   21
         Top             =   15240
         Width           =   3570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sniper6oo@hotmail.com "
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2520
         MouseIcon       =   "frmAbout.frx":1FBF
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   14760
         Width           =   1815
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   2280
         X2              =   4560
         Y1              =   15000
         Y2              =   15000
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHOR UNKNOWN - REALLY COOL SPLASH SCREEN"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   750
         TabIndex        =   19
         Top             =   14400
         Width           =   5400
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63457&lngWId=1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         MouseIcon       =   "frmAbout.frx":2311
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   13920
         Width           =   6705
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   105
         X2              =   6840
         Y1              =   14175
         Y2              =   14160
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHOR UNKNOWN - VB6 TO VB5 FUNCTIONS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   17
         Top             =   13560
         Width           =   4800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1770
         X2              =   5130
         Y1              =   13320
         Y2              =   13320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " http://vbaccelerator.com"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1770
         MouseIcon       =   "frmAbout.frx":2663
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   12960
         Width           =   3045
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRY CODE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   15
         Top             =   12600
         Width           =   1710
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63457&lngWId=1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         MouseIcon       =   "frmAbout.frx":29B5
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   12120
         Width           =   6705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   360
         X2              =   6600
         Y1              =   12420
         Y2              =   12420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BATOSAI USES CODES FROM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1695
         TabIndex        =   13
         Top             =   11280
         Width           =   3795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Noel A. Dacara"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   12
         Top             =   11760
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batosai"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   11
         Top             =   8160
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAPHICS  DRAWN BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         TabIndex        =   10
         Top             =   7680
         Width           =   3105
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batosai"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   9
         Top             =   6960
         Width           =   915
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNED  BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batosai Research Center Program"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   7
         Top             =   5760
         Width           =   4020
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Batosai .inc"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2880
         TabIndex        =   6
         Top             =   4920
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Batosai"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2985
         TabIndex        =   5
         Top             =   4320
         Width           =   945
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CODED BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GENERATION II"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2430
         TabIndex        =   3
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "007"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BATOSAI"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   2640
         Picture         =   "frmAbout.frx":2D07
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OK this form is our about form. All controls
' are contained in a picture box, so we can scroll
' the whole thing without too much code.

' I know there is a scrollhdc API, but this is simpler
' and scrolls at a rate that stops most flickering

' Get the cursor position, so we can detect
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' This works out the position of our control
' on the screen. This works hand-in-hand with
' the functions above.
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
' Dim our R1 as a new rect, or you
' can get a byref error. I sat here for an hour
' trying to work out what the byref error was.
' Turns out I forgot to dim my rect \:)
Dim R1 As RECT
Dim P1 As POINTAPI

Private Sub Form_Load()
' Load our picture for the start of the about box

End Sub

Private Sub Label16_Click()
' RunHyper is a hyperlink runner.
' Saves on code :)
RunHyper Label6.Caption

End Sub


Private Sub Image1_Click()
Label41_Click
End Sub

Private Sub Label12_Click()
Unload Me
End Sub

Private Sub Label19_Click()
RunHyper Label19

End Sub

Private Sub Label21_Click()
RunHyper Label21.Caption
End Sub

Private Sub RunHyper(Hyperlink As String)
' Run our hyperlink using an extended shell command
lngRet = ShellExecute(0&, "Open", Hyperlink, "", vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub Label23_Click()
RunHyper Label23

End Sub

Private Sub Label27_Click()
RunHyper Label27
End Sub

Private Sub Label41_Click()
RunHyper Label41
End Sub

Private Sub Label44_Click()
RunHyper Label44
End Sub

Private Sub Label46_Click()
RunHyper Label46
End Sub

Private Sub Label51_Click()
' oooh! A Special link. Sends an e-mail to me
RunHyper "mailto:" & Label51
End Sub

Private Sub Picture1_Click()
Label41_Click
End Sub

Private Sub Timer1_Timer()
' Slowly scrolls our credits
' Doesn't flicker on my computer, not sure about others
Picture1.Top = Picture1.Top - 50

If Picture1.Top < -Picture1.Height - Me.Height Then Picture1.Top = Me.Height


End Sub

Private Sub Timer2_Timer()
Dim R1 As RECT
Dim P1 As POINTAPI
' Get the location of our form
GetWindowRect Me.hwnd, R1
' Get the location of our cursor
GetCursorPos P1

' Is our cursor over our form? If so, then stop scrolling
If P1.X < R1.Right And P1.X > R1.Left And P1.Y < R1.Bottom And P1.Y > R1.Top Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

End Sub
