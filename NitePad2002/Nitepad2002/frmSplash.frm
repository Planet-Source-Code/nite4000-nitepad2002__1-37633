VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   390
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   2580
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1680
      Top             =   1080
   End
   Begin VB.Line Line2 
      X1              =   4635
      X2              =   4635
      Y1              =   2160
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   630
      X2              =   4635
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   465
      Left            =   600
      Top             =   2160
      Width           =   4380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00FF0000&
      Height          =   2670
      Left            =   -45
      Top             =   -45
      Width           =   690
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Startup()
    ' Show form
    frmSplash.Show
    ' Refresh form
    frmSplash.Refresh
    ' Call Loading sub
    Loading
    ' Close form
    Unload Me
    ' Show main form
    frmMDI.Show
End Sub

Public Sub Loading()
    CStatus "Loading Fonts..."
    LoadFonts 'Load fonts and sizes

    CStatus "Setting all menus and buttons..."
    SetAll 'Set all menus and toolbars

    CStatus "Getting settings from registry..."
    GetSettings 'Load settings from registry
End Sub

Public Sub CStatus(Message As String)
   
    
End Sub
