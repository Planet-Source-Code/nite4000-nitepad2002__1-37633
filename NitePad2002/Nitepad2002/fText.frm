VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fText 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5670
   ControlBox      =   0   'False
   HelpContextID   =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tbokno 
      Height          =   435
      Left            =   4680
      TabIndex        =   6
      Top             =   2010
      WhatsThisHelpID =   2130
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      ButtonWidth     =   794
      ButtonHeight    =   767
      Style           =   1
      ImageList       =   "il"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   4800
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fText.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fText.frx":0730
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fText.frx":0E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fText.frx":1590
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fText.frx":1CC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fm 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      WhatsThisHelpID =   2130
      Width           =   5655
      Begin MSComctlLib.Toolbar tb 
         Height          =   435
         Left            =   4110
         TabIndex        =   5
         Top             =   135
         WhatsThisHelpID =   2130
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ButtonWidth     =   794
         ButtonHeight    =   767
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "il"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbsize 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Text            =   "cbsize"
         Top             =   210
         WhatsThisHelpID =   2130
         Width           =   1095
      End
      Begin VB.ComboBox cbfonts 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "cbfonts"
         Top             =   210
         WhatsThisHelpID =   2130
         Width           =   2655
      End
   End
   Begin VB.Frame fm 
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   495
      WhatsThisHelpID =   2130
      Width           =   5655
      Begin VB.TextBox tx 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         WhatsThisHelpID =   2130
         Width           =   5415
      End
   End
End
Attribute VB_Name = "fText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbfonts_Click()

tx.Font.Name = cbfonts.Text
paintpaintmain.pimage.Font.Name = tx.Font.Name

End Sub


Private Sub cbsize_Click()

tx.Font.Size = cbsize.Text
paintpaintmain.pimage.Font.Size = tx.Font.Size

End Sub

Private Sub Form_Load()

For Lp = 0 To Screen.FontCount - 1
   cbfonts.AddItem Screen.Fonts(Lp)
   DoEvents
Next Lp
For Lp = 2 To 100 Step 2
   cbsize.AddItem Lp
   DoEvents
Next Lp
cbfonts.Text = paintmain.pimage.Font.Name
cbsize.Text = paintmain.pimage.Font.Size
tx.Font.Name = paintmain.pimage.Font.Name
tx.Font.Size = paintmain.pimage.Font.Size

End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 1
       Select Case Button.Value
             Case tbrPressed
              tx.Font.Bold = True
              paintmain.pimage.Font.Bold = True
              Case Else
              tx.Font.Bold = False
              paintmain.pimage.Font.Bold = False
       End Select
      Case 2
       Select Case Button.Value
             Case tbrPressed
              tx.Font.Italic = True
              paintmain.pimage.Font.Italic = True
              Case Else
              tx.Font.Italic = False
              paintmain.pimage.Font.Italic = False
       End Select
      Case 3
       Select Case Button.Value
             Case tbrPressed
              tx.Font.Underline = True
              paintmain.pimage.Font.uderline = True
              Case Else
              tx.Font.Underline = False
              paintmain.pimage.Font.Underline = False
       End Select
       
End Select
      


End Sub

Private Sub tbokno_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
      Case 1
       With paintmain
        Call TextOut(.pimage.hDC, .dLn(1).X1, .dLn(2).Y1, tx.Text, Len(tx.Text))
        .pimage.Refresh
        End With
       Unload Me
      Case 2
       Unload Me
       
End Select

End Sub
