VERSION 5.00
Begin VB.Form newfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Form"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   HelpContextID   =   2160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fm 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   2160
      Width           =   3375
      Begin VB.VScrollBar vs 
         Height          =   300
         Index           =   2
         LargeChange     =   5
         Left            =   1800
         Max             =   0
         Min             =   -700
         SmallChange     =   5
         TabIndex        =   8
         Top             =   720
         WhatsThisHelpID =   2160
         Width           =   375
      End
      Begin VB.VScrollBar vs 
         Height          =   300
         Index           =   1
         LargeChange     =   5
         Left            =   1800
         Max             =   0
         Min             =   -700
         SmallChange     =   5
         TabIndex        =   7
         Top             =   360
         WhatsThisHelpID =   2160
         Width           =   375
      End
      Begin VB.TextBox tx 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   720
         WhatsThisHelpID =   2160
         Width           =   855
      End
      Begin VB.TextBox tx 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   360
         WhatsThisHelpID =   2160
         Width           =   855
      End
      Begin VB.CommandButton cd 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         WhatsThisHelpID =   2160
         Width           =   975
      End
      Begin VB.CommandButton cd 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   1
         Top             =   1200
         WhatsThisHelpID =   2160
         Width           =   735
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "< Pixels >"
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
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Top             =   750
         WhatsThisHelpID =   2160
         Width           =   840
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "< Pixels >"
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
         Index           =   2
         Left            =   2280
         TabIndex        =   9
         Top             =   390
         WhatsThisHelpID =   2160
         Width           =   840
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
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
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         WhatsThisHelpID =   2160
         Width           =   585
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         WhatsThisHelpID =   2160
         Width           =   510
      End
   End
End
Attribute VB_Name = "newfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cd_Click(Index As Integer)

Select Case Index
      Case 1
       Unload Me
      Case 2
       With paintmain
       .pimage.Visible = True
       .pimage.Width = Val(tx(1).Text)
       .pimage.Height = Val(tx(2).Text)
       IgWidth = Val(tx(1).Text)
       IgHeight = Val(tx(2).Text)
       End With
       Unload Me
       
       
End Select

End Sub

Private Sub Form_Load()

tx(1).Text = IgWidth
tx(2).Text = IgHeight

vs(1).Value = -IgWidth
vs(2).Value = -IgHeight

End Sub

Private Sub vs_Change(Index As Integer)

tx(Index) = -(vs(Index).Value)

End Sub
