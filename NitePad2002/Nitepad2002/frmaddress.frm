VERSION 5.00
Begin VB.Form frmaddress 
   Caption         =   "Address Book"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   HelpContextID   =   2080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   2760
      WhatsThisHelpID =   2080
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   2760
      WhatsThisHelpID =   2080
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      WhatsThisHelpID =   2080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      WhatsThisHelpID =   2080
      Width           =   1215
   End
   Begin VB.Data Databar 
      Connect         =   "Access"
      DatabaseName    =   "C:\vbcourse\Programs i made\Nitepad2002\Addbook_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AddBook"
      Top             =   3480
      WhatsThisHelpID =   2080
      Width           =   6900
   End
   Begin VB.TextBox txtcomments 
      DataField       =   "Comments"
      DataSource      =   "Databar"
      Height          =   1335
      Left            =   5040
      TabIndex        =   14
      Top             =   2040
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.TextBox txtemail 
      DataField       =   "Email"
      DataSource      =   "Databar"
      Height          =   405
      Left            =   1320
      TabIndex        =   13
      Top             =   2040
      WhatsThisHelpID =   2080
      Width           =   2535
   End
   Begin VB.TextBox txtzip 
      DataField       =   "Zip"
      DataSource      =   "Databar"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.TextBox txtstate 
      DataField       =   "State"
      DataSource      =   "Databar"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      WhatsThisHelpID =   2080
      Width           =   615
   End
   Begin VB.TextBox txtcity 
      DataField       =   "City"
      DataSource      =   "Databar"
      Height          =   405
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "Address"
      DataSource      =   "Databar"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.TextBox txtlastname 
      DataField       =   "Last"
      DataSource      =   "Databar"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   240
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      DataField       =   "First"
      DataSource      =   "Databar"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      WhatsThisHelpID =   2080
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Comments"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   2040
      WhatsThisHelpID =   2080
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Email"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      WhatsThisHelpID =   2080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Zip"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      WhatsThisHelpID =   2080
      Width           =   495
   End
   Begin VB.Label label5 
      Caption         =   "State"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      WhatsThisHelpID =   2080
      Width           =   1215
   End
   Begin VB.Label label4 
      Caption         =   "City"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      WhatsThisHelpID =   2080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      WhatsThisHelpID =   2080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      WhatsThisHelpID =   2080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   2080
      Width           =   1215
   End
End
Attribute VB_Name = "frmaddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Databar.Recordset.AddNew
End Sub

Private Sub cmddelete_Click()
Databar.Recordset.Delete
Databar.Recordset.MovePrevious
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
Databar.Recordset.Edit
End Sub
