VERSION 5.00
Begin VB.Form Help 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nitepad2002 Help"
   ClientHeight    =   5895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000009&
   HelpContextID   =   2100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdtips 
      Caption         =   "&Show Tips"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox lstHelp 
      Height          =   2010
      ItemData        =   "Help.frx":0000
      Left            =   120
      List            =   "Help.frx":003D
      TabIndex        =   2
      Top             =   840
      WhatsThisHelpID =   2100
      Width           =   3975
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      WhatsThisHelpID =   2100
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Nitepad2002 Help Contents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   2100
      Width           =   4665
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdtips_Click()
frmTip.Show
Unload Me
End Sub

Private Sub lstHelp_Click()


If lstHelp.ListIndex = 0 Then
    lblHelp.Caption = "To Create a New Document then on the File Menu Click New"
ElseIf lstHelp.ListIndex = 1 Then
    lblHelp.Caption = "On the File Menu Click Open. The File Open Dialog Box shows. Select the Drive and folder where file is stored. In the Filter box change it to the file type you want and Double-Click the selected file and Click Open"
ElseIf lstHelp.ListIndex = 2 Then
    lblHelp.Caption = "To Close a Document Click CLOSE on the File Menu"
ElseIf lstHelp.ListIndex = 3 Then
    lblHelp.Caption = "On the File menu Click Save.If your document Dont exist the Save As dialog box will appear. Select the drive and folder to save your document at. In the File Name box give your document a file name. In the filter box you can select which format to save file and the Click OK"
ElseIf lstHelp.ListIndex = 4 Then
    lblHelp.Caption = "On the File menu Click Save As.. Select the drive and folder to save your document at. In the File Name box give your document a file name. In the filter box you can select which format to save the file in and then Click OK"
ElseIf lstHelp.ListIndex = 5 Then
   lblHelp.Caption = "On the File menu Click Page setup then select your printer"
ElseIf lstHelp.ListIndex = 6 Then
    lblHelp.Caption = "To print a document Click File and then Click Print"
ElseIf lstHelp.ListIndex = 7 Then
   lblHelp.Caption = "Click Search then click Find, Findnext or replace"
ElseIf lstHelp.ListIndex = 8 Then
    lblHelp.Caption = "Click the View menu and click web browser"
ElseIf lstHelp.ListIndex = 9 Then
    lblHelp.Caption = "Click View and then on Toolbars and click the toolbar you would like to view"
ElseIf lstHelp.ListIndex = 10 Then
    lblHelp.Caption = "Click View and then on Document Properties"
ElseIf lstHelp.ListIndex = 11 Then
    lblHelp.Caption = "Click Insert then on picture"
ElseIf lstHelp.ListIndex = 12 Then
    lblHelp.Caption = "Click Insert then click Date and time"
ElseIf lstHelp.ListIndex = 13 Then
    lblHelp.Caption = "Click Insert and then on symbols"
ElseIf lstHelp.ListIndex = 14 Then
    lblHelp.Caption = "Click Format then click Font"
ElseIf lstHelp.ListIndex = 15 Then
    lblHelp.Caption = "Click format then click paragraph"
ElseIf lstHelp.ListIndex = 16 Then
    lblHelp.Caption = "Click Format then choose increase or decrease indent"
ElseIf lstHelp.ListIndex = 17 Then
    lblHelp.Caption = "Click tools then Html Editor"
ElseIf lstHelp.ListIndex = 18 Then
    lblHelp.Caption = "Click Tools then click paint"
ElseIf lstHelp.ListIndex = 19 Then
End If

End Sub
