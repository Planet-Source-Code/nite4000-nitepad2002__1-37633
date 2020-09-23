VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form paintmain 
   Caption         =   "NitePaint Pro"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9090
   HelpContextID   =   2150
   Icon            =   "paintmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tm 
      Interval        =   1
      Left            =   8280
      Top             =   6840
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7950
      WhatsThisHelpID =   2150
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptb2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   606
      TabIndex        =   23
      Top             =   7305
      WhatsThisHelpID =   2150
      Width           =   9090
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   64
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   0
         X2              =   64
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Label lbdColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   165
         WhatsThisHelpID =   2150
         Width           =   300
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   51
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   50
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   19
         Left            =   3720
         TabIndex        =   49
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   4
         Left            =   1800
         TabIndex        =   48
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   3
         Left            =   1800
         TabIndex        =   47
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   21
         Left            =   3960
         TabIndex        =   46
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   6
         Left            =   2040
         TabIndex        =   45
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   5
         Left            =   2040
         TabIndex        =   44
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   23
         Left            =   4200
         TabIndex        =   43
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   8
         Left            =   2280
         TabIndex        =   42
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   7
         Left            =   2280
         TabIndex        =   41
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   25
         Left            =   4440
         TabIndex        =   40
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   10
         Left            =   2520
         TabIndex        =   39
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   9
         Left            =   2520
         TabIndex        =   38
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   12
         Left            =   2760
         TabIndex        =   37
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   11
         Left            =   2760
         TabIndex        =   36
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   26
         Left            =   4440
         TabIndex        =   35
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   14
         Left            =   3000
         TabIndex        =   34
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   13
         Left            =   3000
         TabIndex        =   33
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   24
         Left            =   4200
         TabIndex        =   32
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   16
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   15
         Left            =   3240
         TabIndex        =   30
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   22
         Left            =   3960
         TabIndex        =   29
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   18
         Left            =   3480
         TabIndex        =   28
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   17
         Left            =   3480
         TabIndex        =   27
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   20
         Left            =   3720
         TabIndex        =   26
         Top             =   360
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.Label lbbgColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   240
         WhatsThisHelpID =   2150
         Width           =   300
      End
   End
   Begin MSComctlLib.ImageList il4 
      Left            =   4320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":1CC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il3 
      Left            =   4080
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":299E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":2D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":306E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":33D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":373E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":3AA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il2 
      Left            =   4920
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":3E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":47CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":5186
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":5B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":64FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptb1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   15
      Top             =   570
      WhatsThisHelpID =   2150
      Width           =   945
      Begin VB.ComboBox cbzoom 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "paintmain.frx":6EBA
         Left            =   45
         List            =   "paintmain.frx":6ECD
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         WhatsThisHelpID =   2150
         Width           =   855
      End
      Begin MSComctlLib.Toolbar tbesize 
         Height          =   945
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         WhatsThisHelpID =   2150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1667
         ButtonWidth     =   582
         ButtonHeight    =   556
         Style           =   1
         ImageList       =   "il3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbfs 
         Height          =   1170
         Left            =   30
         TabIndex        =   22
         Top             =   5640
         Visible         =   0   'False
         WhatsThisHelpID =   2150
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2064
         ButtonWidth     =   1508
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "il4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbsize 
         Height          =   315
         Left            =   30
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         WhatsThisHelpID =   2150
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         ButtonWidth     =   1508
         ButtonHeight    =   556
         Style           =   1
         ImageList       =   "il2"
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar tbtools 
         Height          =   495
         Left            =   30
         TabIndex        =   16
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         ButtonWidth     =   1508
         ButtonHeight    =   767
         AllowCustomize  =   0   'False
         ImageList       =   "il"
         _Version        =   393216
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   0
         X2              =   64
         Y1              =   225
         Y2              =   225
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   0
         X2              =   64
         Y1              =   224
         Y2              =   224
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   64
         Y1              =   3
         Y2              =   3
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   64
         Y1              =   2
         Y2              =   2
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   6720
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":6EF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":7623
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":7D53
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":8483
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":8BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":92E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":9A17
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":A147
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":A877
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":AFA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":B6D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":BE07
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":C537
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":CC67
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":D397
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":DAC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":E1F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":E927
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":F057
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":F787
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":FEB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":105E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "paintmain.frx":10D17
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcontainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   7215
      Left            =   1080
      ScaleHeight     =   477
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   600
      WhatsThisHelpID =   2150
      Width           =   7935
      Begin VB.HScrollBar hs 
         Height          =   210
         LargeChange     =   200
         Left            =   4440
         Max             =   0
         SmallChange     =   200
         TabIndex        =   5
         Top             =   6240
         WhatsThisHelpID =   2150
         Width           =   615
      End
      Begin VB.CommandButton cdcenter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   6
         Top             =   6120
         WhatsThisHelpID =   2150
         Width           =   225
      End
      Begin VB.VScrollBar vs 
         Height          =   495
         LargeChange     =   200
         Left            =   5520
         Max             =   0
         SmallChange     =   200
         TabIndex        =   4
         Top             =   5520
         WhatsThisHelpID =   2150
         Width           =   210
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   4
         Left            =   0
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   14
         Top             =   2520
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   5
         Left            =   7680
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   13
         Top             =   2400
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   7
         Left            =   3480
         MousePointer    =   7  'Size N S
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   12
         Top             =   5280
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   3240
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   11
         Top             =   0
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   7680
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   10
         Top             =   0
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   8
         Left            =   7680
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   9
         Top             =   5280
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   6
         Left            =   0
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   8
         Top             =   5280
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox psizing 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   0
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   7
         Top             =   0
         WhatsThisHelpID =   2150
         Width           =   120
      End
      Begin VB.PictureBox pimage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   5175
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   505
         TabIndex        =   3
         Top             =   120
         WhatsThisHelpID =   2150
         Width           =   7575
         Begin VB.PictureBox prs 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   90
            Left            =   6000
            MousePointer    =   8  'Size NW SE
            ScaleHeight     =   4
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   4
            TabIndex        =   20
            Top             =   1800
            Visible         =   0   'False
            WhatsThisHelpID =   2150
            Width           =   90
         End
         Begin VB.PictureBox pregion 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   3600
            MousePointer    =   99  'Custom
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   121
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            WhatsThisHelpID =   2150
            Width           =   1815
         End
         Begin VB.Shape dBox 
            DrawMode        =   6  'Mask Pen Not
            Height          =   2655
            Left            =   2400
            Top             =   2160
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Line dLn 
            BorderColor     =   &H00E0E0E0&
            DrawMode        =   6  'Mask Pen Not
            Index           =   2
            Visible         =   0   'False
            X1              =   80
            X2              =   120
            Y1              =   232
            Y2              =   48
         End
         Begin VB.Line gLn 
            BorderColor     =   &H00000000&
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   256
            Y1              =   16
            Y2              =   16
         End
         Begin VB.Line dLn 
            BorderColor     =   &H00E0E0E0&
            DrawMode        =   6  'Mask Pen Not
            Index           =   1
            Visible         =   0   'False
            X1              =   56
            X2              =   96
            Y1              =   232
            Y2              =   40
         End
      End
      Begin VB.Shape boxSizing 
         BorderStyle     =   3  'Dot
         Height          =   615
         Left            =   960
         Top             =   75000
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   2150
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   1005
      ButtonWidth     =   794
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "il"
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.TextBox tx 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "X :     -  Y :"
         Top             =   90
         Visible         =   0   'False
         WhatsThisHelpID =   2150
         Width           =   2655
      End
   End
End
Attribute VB_Name = "paintmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'any variables
Dim Lp, Ctr, Srt, Tmr
'grids spacing and counter variables
Dim gCtr As Integer
Dim gSpace As Integer
'pimage sizing variables
Dim PX, PY As Integer
Dim IgWidth, IgHeight As Integer
'drawing points variables
Dim PX1, PY1, PX2, PY2 As Integer
'pollyline variables
Dim plStarted As Boolean
'rectangle,ovel fill variables
Dim dbFill As Boolean
Dim dbBorder As Boolean
'erasingwidth variable
Dim eWidth As Integer
'to show zooming or not , to show grids or not
'to know which tool was selected
Dim SelectedTool As Integer
Dim Zooming As Boolean
Dim Zoom As Integer
Dim Grids As Boolean
'the draw width , draw color and background color variables
Dim dWidth As Integer
Dim dColor, bgColor As Long
'undo variables
Dim BmphWnd(3) As Long
Dim BmphDC(3) As Long
Dim OldWidth(3), OldHeight(3) As Integer
Dim BmpCtr As Integer
'get picture tool variable
Dim RegWidth, RegHeight As Integer
Dim ReghWnd, ReghDC As Long
Dim RegOldWidth, RegOldHeight As Integer
Dim Moved As Boolean
Dim SelDone As Boolean
'clipbaord variables
Dim cbFmt As Long



Private Sub cbzoom_Click()

Select Case cbzoom.ListIndex
      Case 0
       Zoom = 50
      Case 1
       Zoom = 100
      Case 2
       Zoom = 200
      Case 3
       Zoom = 250
      Case 4
       Zoom = 500
       
End Select

End Sub

Private Sub Form_Load()

'bulid the main toolbar
tb.Buttons.Add , , , 3
tb.Buttons.Add , , , , 1: tb.Buttons(2).Enabled = False
tb.Buttons.Add , , , , 2: tb.Buttons(3).Enabled = False
tb.Buttons.Add , , , , 3: tb.Buttons(4).Enabled = False
tb.Buttons.Add , , , 3
tb.Buttons.Add , , , , 4: tb.Buttons(6).Enabled = False
tb.Buttons.Add , , , , 5: tb.Buttons(7).Enabled = False
tb.Buttons.Add , , , , 6
tb.Buttons.Add , , , 3
tb.Buttons.Add , , , , 7: tb.Buttons(10).Enabled = False
tb.Buttons.Add , , , , 8: tb.Buttons(11).Enabled = False
tb.Buttons.Add , , , 3
tb.Buttons.Add , , , , 9
tb.Buttons.Add , , , , 10
tb.Buttons.Add , , , 3
'build the tools toolbar
tbtools.Buttons.Add , , , 2, 11
tbtools.Buttons.Add , , , 2, 12
tbtools.Buttons.Add , , , 2, 13
tbtools.Buttons.Add , , , 2, 14
tbtools.Buttons.Add , , , 2, 15
tbtools.Buttons.Add , , , 2, 16
tbtools.Buttons.Add , , , 2, 17
tbtools.Buttons.Add , , , 2, 18
tbtools.Buttons.Add , , , 2, 19
tbtools.Buttons.Add , , , 2, 20
tbtools.Buttons.Add , , , 2, 21
tbtools.Buttons.Add , , , 1, 22
tbtools.Buttons.Add , , , 1, 23
'build the size toolbar
tbsize.Buttons.Add , , , 2, 1: tbsize.Buttons(1).Value = tbrPressed
tbsize.Buttons.Add , , , 2, 2
tbsize.Buttons.Add , , , 2, 3
tbsize.Buttons.Add , , , 2, 4
tbsize.Buttons.Add , , , 2, 5

'sets variables for a start
gSpace = 15 'grids spacing
IgWidth = pimage.Width
IgHeight = pimage.Height
pimage.FillStyle = 1
pimage.DrawStyle = 0
'sets the drawwidth to 1 and give first button on the size toolbar pressed value
pimage.DrawWidth = 1: tbsize.Buttons(1).Value = tbrPressed: dWidth = 1
tbesize.Buttons(1).Value = tbrPressed: eWidth = 2
tbfs.Buttons(1).Value = tbrPressed: dbFill = False
Zoom = 50
cbzoom.ListIndex = 0
dColor = lbdColor.BackColor
bgColor = lbbgColor.BackColor
For Lp = 1 To lbc.Count
   lbc(Lp).MousePointer = 99
   lbc(Lp).MouseIcon = LoadResPicture(215, 2)
Next Lp
lbdColor.MouseIcon = LoadResPicture(215, 2)
lbbgColor.MouseIcon = LoadResPicture(215, 2)
pregion.MouseIcon = LoadResPicture(30987, 2)
Call pimage_Resize


End Sub

Private Sub Form_Resize()

On Error GoTo er

pcontainer.Left = ptb1.Width + 2
pcontainer.Top = tb.Height + 2
pcontainer.Width = Me.ScaleWidth - (ptb1.ScaleWidth) '+ ptb2.ScaleWidth + 15)
pcontainer.Height = Me.ScaleHeight - (tb.Height + ptb2.ScaleHeight + sb.Height + 2)
vs.Top = 0
vs.Left = pcontainer.ScaleWidth - vs.Width
vs.Height = pcontainer.ScaleHeight - 15
hs.Left = 0
hs.Top = pcontainer.ScaleHeight - hs.Height
hs.Width = pcontainer.ScaleWidth - 15

cdcenter.Top = vs.Height
cdcenter.Left = hs.Width

er:
  Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lbc_DblClick(Index As Integer)

Select Case Index
      Case 19, 20, 21, 22, 23, 24, 25, 26, 27
       dColor = GetColor 'gets the color dialog when dbliclick any of the white labels
       lbc(Index).BackColor = dColor
       lbdColor.BackColor = dColor
       pimage.FillColor = dColor
       pimage.ForeColor = dColor
      Case Else
       Exit Sub
       
End Select
         
End Sub

Private Sub lbc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  dColor = lbc(Index).BackColor
  lbdColor.BackColor = dColor
  pimage.FillColor = dColor
  pimage.ForeColor = dColor
End If
If Button = 2 Then
  lbbgColor.BackColor = lbc(Index).BackColor
  bgColor = lbc(Index).BackColor
End If

End Sub





Private Sub pcontainer_Resize()

Call pimage_Resize

End Sub

Private Sub pimage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
  Select Case SelectedTool
        Case 6
         plStarted = False
        Case 3, 11
         pregion.Visible = False
         dBox.BorderStyle = 1
         dBox.Visible = False
         prs.Visible = False
         Call DeleteDC(ReghDC)
         Call DeleteObject(ReghWnd)
         tbtools.Buttons(3).Value = tbrUnpressed
         tbtools.Buttons(11).Value = tbrUnpressed
         SelectedTool = 0
         Moved = False
         SelDone = False
         
  End Select
End If

If Button = 1 Then
  Select Case SelectedTool
        Case 1 'eraser
         CreatUndo
         pimage.DrawWidth = eWidth
         pimage.Line (X, Y)-(X + 1, Y + 1), bgColor
         PX1 = X: PY1 = Y
        Case 2
         lbdColor.BackColor = pimage.Point(X, Y)
         dColor = lbdColor.BackColor
         pimage.FillColor = dColor
         pimage.FillColor = dColor
        Case 3
         If SelDone = True Then
           BitBlt pimage.hdc, pregion.Left, pregion.Top, _
                  pregion.Width, pregion.Height, _
                  pregion.hdc, 0, 0, SRCCOPY
           pimage.Refresh
           pregion.Visible = False
           dBox.BorderStyle = 1
           dBox.Visible = False
           prs.Visible = False
           'tbtools.Buttons(3).Value = tbrUnpressed
           SelDone = False
           Moved = False
           tb.Buttons(6).Enabled = False
           tb.Buttons(7).Enabled = False
           Exit Sub
           Else
           SelDone = True
         End If
         dBox.BorderStyle = 3
         dBox.Shape = 0
         dBox.Top = X
         dBox.Left = Y
         dBox.Width = 1
         dBox.Height = 1
         PX1 = X: PY1 = Y
         PX2 = X + 1: PY2 = Y + 1
         dBox.Visible = True
         pregion.Visible = False
         prs.Visible = False
         tb.Buttons(6).Enabled = True
         tb.Buttons(7).Enabled = True
        Case 4 'pen
         CreatUndo
         pimage.Line (X, Y)-(X + 1, Y + 1), dColor
         PX1 = X: PY1 = Y
        Case 5 'line
         CreatUndo
         dLn(1).X1 = X: dLn(1).Y1 = Y
         dLn(1).X2 = X + 1: dLn(1).Y2 = Y + 1
         dLn(1).Visible = True
        Case 6 'pollyline
         If plStarted = False Then
           CreatUndo
           dLn(1).X1 = X: dLn(1).Y1 = Y
           dLn(1).X2 = X + 1: dLn(1).Y2 = Y + 1
           plStarted = True
           Else
           dLn(1).X1 = dLn(1).X2: dLn(1).Y1 = dLn(1).Y2
         End If
         dLn(1).Visible = True
        Case 7 'rectangle
         CreatUndo
         dBox.Shape = 0
         dBox.Left = X: dBox.Top = Y
         dBox.Width = 1: dBox.Height = 1
         PX1 = X: PY1 = Y
         PX2 = X + 1: PY2 = Y + 1
         dBox.Visible = True
         If dbFill = True Then pimage.FillStyle = 0
        Case 8 'ovel
         CreatUndo
         dBox.Shape = 2
         dBox.Left = X: dBox.Top = Y
         dBox.Width = 1: dBox.Height = 1
         PX1 = X: PY1 = Y
         PX2 = X + 1: PY2 = Y + 1
         dBox.Visible = True
         If dbFill = True Then pimage.FillStyle = 0
        Case 9 'fill
         CreatUndo
         pimage.FillStyle = 0
         Call ExtFloodFill(pimage.hdc, X, Y, pimage.Point(X, Y), 1)
         pimage.FillStyle = 1
        Case 10
         CreatUndo
         dLn(1).X1 = X: dLn(1).X2 = X
         dLn(1).Y1 = 0: dLn(1).Y2 = pimage.ScaleHeight
         dLn(2).Y1 = Y: dLn(2).Y2 = Y
         dLn(2).X1 = 0: dLn(2).X2 = pimage.ScaleWidth
         dLn(1).BorderStyle = 3: dLn(2).BorderStyle = 3
         dLn(1).Visible = True: dLn(2).Visible = True
        Case 11
         CreatUndo
         BitBlt pimage.hdc, pregion.Left, pregion.Top, _
                pregion.Width, pregion.Height, _
                pregion.hdc, 0, 0, SRCCOPY
         pimage.Refresh
         pregion.Visible = False
         dBox.BorderStyle = 1
         dBox.Visible = False
         prs.Visible = False
         tbtools.Buttons(11).Value = tbrUnpressed
         SelectedTool = 0
         tb.Buttons(6).Enabled = False
         tb.Buttons(7).Enabled = False
         
  End Select
End If

End Sub

Private Sub pimage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Zooming = True Then
  fzoom.pzoom.Cls
  Call StretchBlt(fzoom.pzoom.hdc, 0, 0, Zoom, Zoom, pimage.hdc, X, Y, 10, 10, SRCCOPY)
  fzoom.pzoom.Refresh
End If

If Button = 1 Then
  Select Case SelectedTool
        Case 1 'eraser
         pimage.Line (PX1, PY1)-(X, Y), bgColor
         PX1 = X: PY1 = Y
        Case 2
         If pimage.Point(X, Y) = -1 Then
           lbdColor.BackColor = dColor
           Else
           lbdColor.BackColor = pimage.Point(X, Y)
           dColor = lbdColor.BackColor
           pimage.FillColor = dColor
           pimage.FillColor = dColor
         End If
        Case 3
         PX2 = X: PY2 = Y
         If PX2 > PX1 Then
           dBox.Left = PX1
           dBox.Width = (PX2 - PX1) + 1
         End If
         If PX2 < PX1 Then
           dBox.Left = PX2
           dBox.Width = (PX1 - PX2) + 1
         End If
         If PY2 > PY1 Then
           dBox.Top = PY1
           dBox.Height = (PY2 - PY1) + 1
         End If
         If PY2 < PY1 Then
           dBox.Top = PY2
           dBox.Height = (PY1 - PY2) + 1
         End If
        Case 4 'pen
         pimage.Line (PX1, PY1)-(X, Y), dColor
         PX1 = X: PY1 = Y
        Case 5 'line
         dLn(1).X2 = X: dLn(1).Y2 = Y
        Case 6 'pollyline
         dLn(1).X2 = X: dLn(1).Y2 = Y
        Case 7 'rectangle
         PX2 = X: PY2 = Y
         If PX2 > PX1 Then
           dBox.Left = PX1
           dBox.Width = (PX2 - PX1) + 1
         End If
         If PX2 < PX1 Then
           dBox.Left = PX2
           dBox.Width = (PX1 - PX2) + 1
         End If
         If PY2 > PY1 Then
           dBox.Top = PY1
           dBox.Height = (PY2 - PY1) + 1
         End If
         If PY2 < PY1 Then
           dBox.Top = PY2
           dBox.Height = (PY1 - PY2) + 1
         End If
        Case 8 'ovel
         PX2 = X: PY2 = Y
         If PX2 > PX1 Then
           dBox.Left = PX1
           dBox.Width = (PX2 - PX1) + 1
         End If
         If PX2 < PX1 Then
           dBox.Left = PX2
           dBox.Width = (PX1 - PX2) + 1
         End If
         If PY2 > PY1 Then
           dBox.Top = PY1
           dBox.Height = (PY2 - PY1) + 1
         End If
         If PY2 < PY1 Then
           dBox.Top = PY2
           dBox.Height = (PY1 - PY2) + 1
         End If
        Case 10
         dLn(1).X1 = X: dLn(1).X2 = X
         dLn(2).Y1 = Y: dLn(2).Y2 = Y
         
         
  End Select
End If
  
  
End Sub

Private Sub pimage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  Select Case SelectedTool
        Case 1
         pimage.DrawWidth = dWidth
        Case 3
         If SelDone = True Then
           pregion.Left = dBox.Left: pregion.Top = dBox.Top
           pregion.Width = dBox.Width: pregion.Height = dBox.Height
           pregion.Visible = True
           dBox.Left = pregion.Left - 1
           dBox.Top = pregion.Top - 1
           dBox.Width = pregion.Width + 2
           dBox.Height = pregion.Height + 2
           prs.Left = dBox.Left + dBox.Width + 1
           prs.Top = dBox.Top + dBox.Height + 1
           prs.Visible = True
           BitBlt pregion.hdc, _
                  0, 0, pregion.Width, pregion.Height, _
                  pimage.hdc, pregion.Left, pregion.Top, SRCCOPY
           pregion.Refresh
         End If
        Case 5 'line
         dLn(1).Visible = False
         pimage.Line (dLn(1).X1, dLn(1).Y1)-(dLn(1).X2, dLn(1).Y2), dColor
        Case 6 'pollyline
         dLn(1).Visible = False
         dLn(1).X2 = X: dLn(1).Y2 = Y
         pimage.Line (dLn(1).X1, dLn(1).Y1)-(dLn(1).X2, dLn(1).Y2), dColor
        Case 7
         dBox.Visible = False
         pimage.Line (PX1, PY1)-(PX2, PY2), dColor, B
         If dbFill = True Then pimage.FillStyle = 1
         If dbBorder = True Then
           pimage.Line (PX1, PY1)-(PX2, PY2), bgColor, B
         End If
        Case 8
         dBox.Visible = False
         If dBox.Height > dBox.Width Then
           pimage.Circle (dBox.Left + (dBox.Width / 2), dBox.Top + (dBox.Height / 2)), (dBox.Height / 2), dColor, , , (dBox.Height / dBox.Width)
           Else
           pimage.Circle (dBox.Left + (dBox.Width / 2), dBox.Top + (dBox.Height / 2)), (dBox.Width / 2), dColor, , , (dBox.Height / dBox.Width)
         End If
         If dbFill = True Then pimage.FillStyle = 1
         If dbBorder = True Then
           If dBox.Height > dBox.Width Then
             pimage.Circle (dBox.Left + (dBox.Width / 2), dBox.Top + (dBox.Height / 2)), (dBox.Height / 2), bgColor, , , (dBox.Height / dBox.Width)
             Else
             pimage.Circle (dBox.Left + (dBox.Width / 2), dBox.Top + (dBox.Height / 2)), (dBox.Width / 2), bgColor, , , (dBox.Height / dBox.Width)
           End If
         End If
        Case 10
         fText.Show 1
         dLn(1).BorderStyle = 0: dLn(2).BorderStyle = 0
         dLn(1).Visible = False: dLn(2).Visible = False
         
  End Select
End If

End Sub

Private Sub pimage_Resize()

'sets the tiny resizng boxes position

psizing(1).Top = 1 + vs.Value: psizing(1).Left = 1 + hs.Value

pimage.Top = psizing(1).Top + 9
pimage.Left = psizing(1).Left + 9

psizing(2).Top = 1: psizing(2).Left = (pimage.Width / 2) + 5
psizing(3).Top = 1: psizing(3).Left = pimage.Width + 11
psizing(4).Top = (pimage.Height / 2) + 5: psizing(4).Left = 1
psizing(5).Top = (pimage.Height / 2) + 5: psizing(5).Left = pimage.Width + 11
psizing(6).Top = pimage.Height + 11: psizing(6).Left = 1
psizing(7).Top = pimage.Height + 11: psizing(7).Left = (pimage.Width / 2) + 5
psizing(8).Top = pimage.Height + 11: psizing(8).Left = pimage.Width + 11

If psizing(5).Left + psizing(5).Width > pcontainer.Width - vs.Width Then
  hs.Max = (psizing(5).Left + psizing(5).Width) - (pcontainer.Width - vs.Width)
  hs.Visible = True
  cdcenter.Visible = True
  Else
  hs.Value = 0
  hs.Visible = False
  cdcenter.Visible = False
End If
If psizing(7).Top + psizing(7).Height > pcontainer.Height - hs.Height Then
  vs.Max = (psizing(7).Top + psizing(7).Height) - (pcontainer.Height - hs.Height)
  vs.Visible = True
  cdcenter.Visible = True
  Else
  vs.Value = 0
  vs.Visible = False
  cdcenter.Visible = False
End If

End Sub


Private Sub pregion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  Select Case SelectedTool
        Case 3, 11
         PX1 = X: PY1 = Y




  End Select
End If

End Sub

Private Sub pregion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
 
If Button = 1 Then
  Select Case SelectedTool
        Case 3
         If Moved = False Then
           CreatUndo
           pimage.FillStyle = 0
           pimage.Line (pregion.Left, pregion.Top)- _
                      (pregion.Left + pregion.Width, _
                       pregion.Top + pregion.Height), bgColor, BF
           pimage.FillStyle = 1
           Moved = True
         End If
         pregion.Left = pregion.Left + (X - PX1)
         pregion.Top = pregion.Top + (Y - PY1)
         dBox.Top = pregion.Top - 1
         dBox.Left = pregion.Left - 1
         dBox.Width = pregion.Width + 2
         dBox.Height = pregion.Height + 2
         prs.Left = dBox.Left + dBox.Width + 1
         prs.Top = dBox.Top + dBox.Height + 1
        Case 11
         pregion.Left = pregion.Left + (X - PX1)
         pregion.Top = pregion.Top + (Y - PY1)
         dBox.Top = pregion.Top - 1
         dBox.Left = pregion.Left - 1
         dBox.Width = pregion.Width + 2
         dBox.Height = pregion.Height + 2
         prs.Left = dBox.Left + dBox.Width + 1
         prs.Top = dBox.Top + dBox.Height + 1

  End Select
End If

End Sub

Private Sub pregion_Resize()

'tb.Buttons(6).Enabled = True
tb.Buttons(7).Enabled = True
       
End Sub

Private Sub prs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  Select Case SelectedTool
        Case 3, 11
         CopyReg
         pregion.Picture = LoadPicture("")
         PX2 = X: PY2 = Y
         
         
  End Select
End If

End Sub

Private Sub prs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  Select Case SelectedTool
        Case 3, 11
         prs.Left = prs.Left + (X - PX2)
         prs.Top = prs.Top + (Y - PY2)
         RegWidth = (prs.Left - pregion.Left) - 3
         RegHeight = (prs.Top - pregion.Top) - 3
         If RegWidth < 1 Then RegWidth = 1
         If RegHeight < 1 Then RegHeight = 1
         pregion.Width = RegWidth
         pregion.Height = RegHeight
         dBox.Width = pregion.Width + 2
         dBox.Height = pregion.Height + 2
         
         
  End Select
End If

End Sub

Private Sub prs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  Select Case SelectedTool
        Case 3, 11
         pregion.Cls
         Call StretchBlt(pregion.hdc, 0, 0, pregion.Width, pregion.Height, ReghDC, 0, 0, RegOldWidth, RegOldHeight, SRCCOPY)
         pregion.Refresh
         prs.Left = dBox.Left + dBox.Width + 1
         prs.Top = dBox.Top + dBox.Height + 1
         
  End Select
End If

End Sub

Private Sub psizing_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  If Index = 5 Or Index = 7 Or Index = 8 Then boxSizing.Visible = True
  PX = X: PY = Y
  IgWidth = pimage.Width
  IgHeight = pimage.Height
End If

End Sub

Private Sub psizing_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo er

If Button = 1 Then
  Select Case Index
        Case 5
         psizing(5).Left = psizing(Index).Left - (PX - X)
         IgWidth = psizing(5).Left - 11
        Case 7
         psizing(7).Top = psizing(Index).Top - (PY - Y)
         IgHeight = psizing(7).Top - 11
        Case 8
         psizing(8).Left = psizing(Index).Left - (PX - X)
         psizing(8).Top = psizing(Index).Top - (PY - Y)
         IgWidth = psizing(8).Left - 11
         IgHeight = psizing(8).Top - 11
         
  End Select
  If IgWidth < 1 Then IgWidth = 1
  If IgHeight < 1 Then IgHeight = 1
  pimage.Width = IgWidth
  pimage.Height = IgHeight
  
  boxSizing.Top = psizing(1).Top + 8
  boxSizing.Left = psizing(1).Left + 8
  boxSizing.Width = IgWidth + 2
  boxSizing.Height = IgHeight + 2
  
End If

er:
 Exit Sub

End Sub

Private Sub psizing_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  If Index = 5 Or Index = 7 Or Index = 8 Then boxSizing.Visible = False
  pimage_Resize
  'draw grids after resizing
  DrawGrid Grids
End If

End Sub

Private Sub DrawGrid(Vis As Boolean)

Grids = Vis

If gLn.Count > 1 Then
  For Lp = 1 To gLn.Count - 1
     Unload gLn(Lp)
     DoEvents
  Next Lp
End If

If Vis = True Then
  Lp = 0
  gCtr = 0
  Do Until Lp >= IgWidth
    gCtr = gCtr + 1
    Load gLn(gCtr)
    gLn(gCtr).X1 = Lp
    gLn(gCtr).X2 = Lp
    gLn(gCtr).Y1 = 0
    gLn(gCtr).Y2 = pimage.ScaleHeight
    gLn(gCtr).Visible = True
    gLn(gCtr).BorderColor = &H0 'gColor ... you can do that
    Lp = Lp + gSpace
    DoEvents
  Loop
  Lp = 0
  Do Until Lp >= IgHeight
    gCtr = gCtr + 1
    Load gLn(gCtr)
    gLn(gCtr).X1 = 0
    gLn(gCtr).X2 = pimage.ScaleWidth
    gLn(gCtr).Y1 = Lp
    gLn(gCtr).Y2 = Lp
    gLn(gCtr).Visible = True
    gLn(gCtr).BorderColor = &H0 'gColor...you can do that
    Lp = Lp + gSpace
    DoEvents
  Loop
End If

End Sub



Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 7
       'tb.Buttons(6).Enabled = True
       tb.Buttons(7).Enabled = False
       Clipboard.Clear
       Clipboard.SetData pregion.Image
      Case 8
       CreatUndo
       pregion.AutoSize = True
       pregion.Picture = Clipboard.GetData
       pregion.AutoSize = False
       pregion.Top = 2: pregion.Left = 2
       pregion.Visible = True
       dBox.BorderStyle = 3
       dBox.Shape = 0
       dBox.Top = pregion.Top - 1
       dBox.Left = pregion.Left - 1
       dBox.Width = pregion.Width + 2
       dBox.Height = pregion.Height + 2
       dBox.Visible = True
       prs.Left = dBox.Left + dBox.Width + 1
       prs.Top = dBox.Top + dBox.Height + 1
       prs.Visible = True
       SelectedTool = 3
       Moved = True
       SelDone = True
       'tb.Buttons(6).Enabled = True
       tb.Buttons(7).Enabled = True
      Case 10
       GetDCBmp BmpCtr
       BmpCtr = BmpCtr - 1
       If SelectedTool = 6 Then plStarted = False
       If SelectedTool = 3 Then Moved = False
       If BmpCtr < 1 Then
         tb.Buttons(10).Enabled = False
       End If
       
      
End Select
       
End Sub

Private Sub tbesize_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 1
       eWidth = 2
      Case 2
       eWidth = 4
      Case 3
       eWidth = 8
      Case 4
       eWidth = 10
      Case 5
       eWidth = 12
      Case 6
       eWidth = 15
      
End Select
      
End Sub

Private Sub tbfs_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 1
       dbFill = False
       dbBorder = True
      Case 2
       dbFill = True
       dbBorder = True
      Case 3
       dbFill = True
       dbBorder = False

End Select

End Sub

Private Sub tbsize_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 1
       dWidth = 1
      Case 2
       dWidth = 2
      Case 3
       dWidth = 4
      Case 4
       dWidth = 6
      Case 5
       dWidth = 8
   
End Select

pimage.DrawWidth = dWidth

End Sub

Private Sub tbtools_ButtonClick(ByVal Button As MSComctlLib.Button)

If SelectedTool = 3 Or SelectedTool = 11 Then
  pregion.Visible = False
  dBox.BorderStyle = 2
  dBox.Visible = False
  prs.Visible = False
  Call DeleteDC(ReghDC)
  Call DeleteObject(ReghWnd)
  tbtools.Buttons(3).Value = tbrUnpressed
  tbtools.Buttons(11).Value = tbrUnpressed
  SelectedTool = 0
  SelDone = False
  Moved = False
End If

Select Case Button.Index
      Case 1
       pimage.MouseIcon = LoadResPicture(160, 2)
       SelectedTool = 1
      Case 2
       pimage.MouseIcon = LoadResPicture(159, 2)
       SelectedTool = 2
      Case 3
       pimage.MouseIcon = LoadResPicture(241, 2)
       SelectedTool = 3
       Moved = False
       SelDone = False
      Case 4
       pimage.MouseIcon = LoadResPicture(161, 2)
       SelectedTool = 4
      Case 5
       pimage.MouseIcon = LoadResPicture(177, 2)
       SelectedTool = 5
      Case 6
       pimage.MouseIcon = LoadResPicture(232, 2)
       SelectedTool = 6
       plStarted = False
      Case 7
       pimage.MouseIcon = LoadResPicture(167, 2)
       SelectedTool = 7
       'tbsize
      Case 8
       pimage.MouseIcon = LoadResPicture(173, 2)
       SelectedTool = 8
      Case 9
       pimage.MouseIcon = LoadResPicture(166, 2)
       SelectedTool = 9
      Case 10
       pimage.MouseIcon = LoadResPicture(156, 2)
       SelectedTool = 10
      Case 11
       pregion.AutoSize = True
       pregion.Picture = LoadPicture(GetFile)
       pregion.AutoSize = False
       pregion.Top = 2: pregion.Left = 2
       pregion.Visible = True
       dBox.BorderStyle = 3
       dBox.Shape = 0
       dBox.Top = pregion.Top - 1
       dBox.Left = pregion.Left - 1
       dBox.Width = pregion.Width + 2
       dBox.Height = pregion.Height + 2
       dBox.Visible = True
       prs.Left = dBox.Left + dBox.Width + 1
       prs.Top = dBox.Top + dBox.Height + 1
       prs.Visible = True
       SelectedTool = 11
       Moved = False
       tb.Buttons(6).Enabled = True
       tb.Buttons(7).Enabled = True
      Case 12
       If tbtools.Buttons(12).Value = tbrPressed Then
         DrawGrid True
         Else
         DrawGrid False
       End If
      Case 13
       If tbtools.Buttons(13).Value = tbrPressed Then
         fzoom.Show
         Zooming = True
         cbzoom.Visible = True
         Else
         Unload fzoom
         Zooming = False
         cbzoom.Visible = False
       End If
       
 
End Select
 
Select Case SelectedTool
      Case 1
       tbesize.Visible = True
       tbfs.Visible = False
       tbsize.Visible = False
      Case 4, 5, 6
       tbesize.Visible = False
       tbfs.Visible = False
       tbsize.Visible = True
      Case 7, 8
       tbesize.Visible = False
       tbfs.Visible = Visible
       tbsize.Visible = True
      Case Else
       tbesize.Visible = False
       tbfs.Visible = False
       tbsize.Visible = False

End Select
 
End Sub

Private Sub CreatUndo()

BmpCtr = BmpCtr + 1
If BmpCtr > 3 Then
  BitBlt BmphDC(1), _
   0, 0, OldWidth(2), OldHeight(2), _
   BmphDC(2), 0, 0, SRCCOPY
  BitBlt BmphDC(2), _
   0, 0, OldWidth(3), OldHeight(3), _
   BmphDC(3), 0, 0, SRCCOPY
  BitBlt BmphDC(3), _
   0, 0, pimage.Width, pimage.Height, _
   pimage.hdc, 0, 0, SRCCOPY
  BmpCtr = 3
  Exit Sub
End If

Call DeleteDC(BmphDC(BmpCtr))
Call DeleteObject(BmphWnd(BmpCtr))
OldWidth(BmpCtr) = pimage.Width
OldHeight(BmpCtr) = pimage.Height
BmphDC(BmpCtr) = CreateCompatibleDC(pimage.hdc)
BmphWnd(BmpCtr) = CreateCompatibleBitmap(pimage.hdc, OldWidth(BmpCtr), OldHeight(BmpCtr))
Call SelectObject(BmphDC(BmpCtr), BmphWnd(BmpCtr))
BitBlt BmphDC(BmpCtr), _
       0, 0, OldWidth(BmpCtr), OldHeight(BmpCtr), _
       pimage.hdc, 0, 0, SRCCOPY


tb.Buttons(10).Enabled = True

End Sub

Private Sub GetDCBmp(ind As Integer)

pimage.Cls

BitBlt pimage.hdc, _
       0, 0, OldWidth(ind), OldHeight(ind), _
       BmphDC(ind), 0, 0, SRCCOPY
pimage.Refresh

End Sub

Private Sub CopyReg()

Call DeleteDC(ReghDC)
Call DeleteObject(ReghWnd)
RegOldWidth = pregion.Width
RegOldHeight = pregion.Height
ReghDC = CreateCompatibleDC(pregion.hdc)
ReghWnd = CreateCompatibleBitmap(pregion.hdc, RegOldWidth, RegOldHeight)
Call SelectObject(ReghDC, ReghWnd)
BitBlt ReghDC, _
       0, 0, RegOldWidth, RegOldHeight, _
       pregion.hdc, 0, 0, SRCCOPY
       
End Sub

Private Sub tm_Timer()

Clipboard.GetFormat cbFmt
If Clipboard.GetFormat(1) Or Clipboard.GetFormat(2) Or Clipboard.GetFormat(3) Or Clipboard.GetFormat(8) Then
  tb.Buttons(8).Enabled = True
  Else
  tb.Buttons(8).Enabled = False
End If

End Sub
