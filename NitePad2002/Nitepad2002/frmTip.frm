VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   2715
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   HelpContextID   =   2110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5415
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   4080
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   1560
      WhatsThisHelpID =   2110
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      WhatsThisHelpID =   2110
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<= Previous"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      WhatsThisHelpID =   2110
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next =>"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      WhatsThisHelpID =   2110
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2355
      Left            =   120
      Picture         =   "frmTip.frx":2D42
      ScaleHeight     =   2295
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   2110
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   1815
         Left            =   540
         TabIndex        =   2
         Top             =   240
         WhatsThisHelpID =   2110
         Width           =   2895
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   1
         Top             =   840
         WhatsThisHelpID =   2110
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub cmdNext_Click()
If CurrentTip = 11 Then
CurrentTip = 1
Label1.Caption = Tips.Item(CurrentTip)
Else
    CurrentTip = CurrentTip + 1
    Label1.Caption = Tips.Item(CurrentTip)
End If

End Sub

Private Sub cmdPrevious_Click()

If CurrentTip = 0 Then
CurrentTip = 11
Label1.Caption = Tips.Item(CurrentTip)
ElseIf CurrentTip = 1 Then
CurrentTip = 11
Label1.Caption = Tips.Item(CurrentTip)
Else
    CurrentTip = CurrentTip - 1
    Label1.Caption = Tips.Item(CurrentTip)
End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
 
CurrentTip = 0

'Tips added
Tips.Add "That you can create a new document by click NEW on the File Menu?"
Tips.Add "That you can select all text by pressing CTRL + A"
Tips.Add "That You can design web pages by starting the HTML editor from the Tools menu by click HTML EDITOR?"
Tips.Add "That you can display the clock by clicking the tools menu and then Clock?"
Tips.Add "That You can Insert the Date into your document?"
Tips.Add "That You Can insert a Image File by Click on Picture on the Insert Menu"
Tips.Add "That You can view the internet by clicking web browser on the View menu"
Tips.Add "That You can search your document for a word"
Tips.Add "That you can draw with Nitepads paint program"
Tips.Add "That you can indent you paragraphs"
End Sub

