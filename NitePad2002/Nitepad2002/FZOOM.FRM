VERSION 5.00
Begin VB.Form fzoom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Zoom"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4365
   ControlBox      =   0   'False
   HelpContextID   =   2140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tm 
      Interval        =   1
      Left            =   3240
      Top             =   2640
   End
   Begin VB.PictureBox pzoom 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   -120
      WhatsThisHelpID =   2140
      Width           =   4335
   End
End
Attribute VB_Name = "fzoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Resize()

pzoom.Top = 0
pzoom.Left = 0
pzoom.Width = fzoom.ScaleWidth
pzoom.Height = fzoom.ScaleHeight

End Sub

Private Sub tm_Timer()

Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW)

End Sub
