VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fpicture 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Picture"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7965
   ControlBox      =   0   'False
   HelpContextID   =   2120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   600
      WhatsThisHelpID =   2120
      Width           =   7935
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   3915
         TabIndex        =   5
         Top             =   0
         WhatsThisHelpID =   2120
         Width           =   3975
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   735
         Left            =   6120
         TabIndex        =   4
         Top             =   1800
         WhatsThisHelpID =   2120
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Top             =   2520
         WhatsThisHelpID =   2120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   6120
         TabIndex        =   2
         Top             =   2520
         WhatsThisHelpID =   2120
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   600
      Left            =   6240
      TabIndex        =   0
      Top             =   4080
      WhatsThisHelpID =   2120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      WhatsThisHelpID =   2120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "fpicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
