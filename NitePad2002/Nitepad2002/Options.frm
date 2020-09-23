VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1470
         TabIndex        =   10
         Top             =   1170
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                                                                            Auto Save Options for XtremePad."
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Auto Save every"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "IMPORTANT  Auto Save will work un Pre-Saved files. If it's a new document that hasn't been saved yet, the save box will be opened."
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   4455
      End
   End
   Begin VB.CommandButton apply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "Restore Original Values"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show Clock"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1035
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                                                                       General Options for XtremePad."
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.ToolTipText     =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AutoSave"
            Key             =   "Autosave"
            Object.ToolTipText     =   "AutoSave Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub apply_Click()
apply.Enabled = False
If Check1 = False Then
frmmdi.Timer1.Enabled = False
frmmdi.SB.Panels(2).Text = ""
Else
frmmdi.Timer1.Enabled = True
End If

If Check2 = False Then
frmmdi.Timer2.Enabled = False
frmmdi.SB.Panels(3).Text = ""
Else
frmmdi.Timer2.Enabled = True
End If

Dim UserInput
UserInput = Text1.Text
If UserInput = "" Then
Text1.Text = 0
End If
If Not IsNumeric(UserInput) Then
Exit Sub
MsgBox "Invalid number!"
MDIForm1.txtuserinput = "0"
MDIForm1.txtrepitition = "0"
MDIForm1.tmrautosave.Enabled = False
Else
frmmdi.txtuserinput = UserInput
frmmdi.tmrautosave.Enabled = True
End If

End Sub

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Check1_Click()
apply.Enabled = True
End Sub

Private Sub Check2_Click()
apply.Enabled = True
End Sub

Private Sub Command3_Click()
Check1.Value = 1
Check2.Value = 1
apply.Enabled = True
End Sub

Private Sub Ok_Click()
If Check1 = False Then
frmmdi.Timer1.Enabled = False
frmmdi.SB.Panels(2).Text = ""
Unload Me
Else
frmmdi.Timer1.Enabled = True
Unload Me
End If

If Check2 = False Then
MDIForm1.Timer2.Enabled = False
MDIForm1.SB.Panels(3).Text = ""
Unload Me
Else
frmmdi.Timer2.Enabled = True
Unload Me
End If

End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs.Item(1).Selected = True Then
Frame1.Visible = True
Frame2.Visible = False
End If
If TabStrip1.Tabs.Item(2).Selected = True Then
Frame2.Visible = True
Frame1.Visible = False
End If
End Sub

Private Sub Text1_Change()
apply.Enabled = True
End Sub
