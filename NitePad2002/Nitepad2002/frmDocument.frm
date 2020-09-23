VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Document"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   HelpContextID   =   310
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   8235
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   3840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   8580
      HelpContextID   =   610
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   310
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   15134
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public bChanged As Boolean


Private Sub Form_Activate()
    EnableAll 'Enable all menus and toolbars
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim myResponse As Integer

If rtfText.Text = "" Then
    Exit Sub
Else
    myResponse = MsgBox("Do you wish to save current project?", vbYesNo)
    If myResponse = "6" Then
        frmMDI.mnuFileSave_Click
    Else
        Exit Sub
    End If
End If

If myCancel = True Then
    Cancel = True
End If

End Sub

Private Sub rtfText_Change()
    bChanged = True
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then 'If right mouse button is clicked
        PopupMenu frmMDI.mnuPop 'show popup menu
    End If
    SetAll
End Sub

Private Sub rtfText_SelChange()
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub
