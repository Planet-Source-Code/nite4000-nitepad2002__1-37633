VERSION 5.00
Begin VB.Form frmPics 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   HelpContextID   =   420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picsaveall 
      Height          =   495
      Left            =   1080
      Picture         =   "frmPics.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   34
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox pictxt 
      Height          =   495
      Left            =   720
      Picture         =   "frmPics.frx":0342
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   33
      Top             =   2520
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox pictime 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":0444
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   2520
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picolor 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":0786
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   31
      Top             =   1320
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picfind 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":0AC8
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   1920
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picredo 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":0BCA
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picprintsetup 
      Height          =   495
      Left            =   1800
      Picture         =   "frmPics.frx":0F0C
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picpreview 
      Height          =   495
      Left            =   2160
      Picture         =   "frmPics.frx":124E
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picabout 
      Height          =   495
      Left            =   720
      Picture         =   "frmPics.frx":1590
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   4320
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox pictipday 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":18D2
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   4320
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox pichelp 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":1C14
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   4320
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picpic 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":1D16
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   2520
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox pichtml 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":2058
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   3120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox piccalculator 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":239A
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   3120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox piccamera 
      Height          =   495
      Left            =   720
      Picture         =   "frmPics.frx":26DC
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   3120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox piccascade 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":27DE
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   3720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picarrange 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":28E0
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   3720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picundo 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":29E2
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picdelete 
      Height          =   495
      Left            =   1800
      Picture         =   "frmPics.frx":2D24
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picpaste 
      Height          =   495
      Left            =   1440
      Picture         =   "frmPics.frx":2E26
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   360
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   7
         Top             =   0
         WhatsThisHelpID =   420
         Width           =   15
      End
   End
   Begin VB.PictureBox piccopy 
      Height          =   495
      Left            =   1080
      Picture         =   "frmPics.frx":2F28
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox piccut 
      Height          =   495
      Left            =   720
      Picture         =   "frmPics.frx":302A
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   720
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picprint 
      Height          =   495
      Left            =   1440
      Picture         =   "frmPics.frx":312C
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picsave 
      Height          =   495
      Left            =   720
      Picture         =   "frmPics.frx":322E
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picopen 
      Height          =   495
      Left            =   360
      Picture         =   "frmPics.frx":3330
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.PictureBox picNew 
      Height          =   495
      Left            =   0
      Picture         =   "frmPics.frx":3432
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Help"
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Top             =   4320
      WhatsThisHelpID =   420
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Windows"
      Height          =   375
      Left            =   2640
      TabIndex        =   29
      Top             =   3720
      WhatsThisHelpID =   420
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Tools"
      Height          =   375
      Left            =   2640
      TabIndex        =   28
      Top             =   3120
      WhatsThisHelpID =   420
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Insert"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   2640
      WhatsThisHelpID =   420
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Search"
      Height          =   375
      Left            =   2640
      TabIndex        =   26
      Top             =   2040
      WhatsThisHelpID =   420
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Format"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   1440
      WhatsThisHelpID =   420
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   840
      WhatsThisHelpID =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   120
      WhatsThisHelpID =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
