VERSION 5.00
Begin VB.Form frmClock 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time Clock"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   HelpContextID   =   460
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdt2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5520
      WhatsThisHelpID =   460
      Width           =   1335
   End
   Begin VB.CommandButton Cmdt1 
      Caption         =   "Calendar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4920
      WhatsThisHelpID =   460
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   4560
      WhatsThisHelpID =   460
      Width           =   495
   End
   Begin VB.TextBox txtFormula 
      Height          =   1455
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmClock.frx":0742
      Top             =   3000
      WhatsThisHelpID =   460
      Width           =   2700
   End
   Begin VB.Timer tmrQuartz 
      Interval        =   1000
      Left            =   6000
      Top             =   3480
   End
   Begin VB.Image imgAuthor 
      BorderStyle     =   1  'Fixed Single
      Height          =   2790
      Left            =   4560
      Picture         =   "frmClock.frx":0977
      Top             =   120
      WhatsThisHelpID =   460
      Width           =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   120
      X2              =   4080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   10
      Height          =   6135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   4440
      WhatsThisHelpID =   460
      Width           =   2535
   End
   Begin VB.Line LineSecond 
      BorderColor     =   &H00FFFFFF&
      X1              =   2197
      X2              =   1000
      Y1              =   2160
      Y2              =   3240
   End
   Begin VB.Line LineMinute 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      X1              =   2190
      X2              =   1320
      Y1              =   2160
      Y2              =   1440
   End
   Begin VB.Line LineHour 
      BorderColor     =   &H8000000E&
      BorderWidth     =   5
      X1              =   2190
      X2              =   3600
      Y1              =   2160
      Y2              =   1200
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const PI = 3.14159

Private Sub Command1_Click()
 frmClock.Width = 7680
End Sub

Private Sub Cmdt1_Click()
frmcalendar.Show

End Sub

Private Sub Cmdt2_Click()
Unload Me
End Sub



Private Sub Form_Load()
Call tmrQuartz_Timer
End Sub



Private Sub tmrQuartz_Timer()
Dim Hours As Single, Minutes As Single, Seconds As Single
Dim TrueHours As Single

lblTime.Caption = Time
'Beep
Hours = Hour(Time)
Minutes = Minute(Time)
Seconds = Second(Time)
TrueHours = Hours + Minutes / 60

' I made all the X1 and Y1 equal in the form
LineHour.X2 = 1000 * Cos(PI / 180 * (30 * TrueHours - 90)) + LineHour.X1
LineHour.Y2 = 1000 * Sin(PI / 180 * (30 * TrueHours - 90)) + LineHour.Y1
    
LineMinute.X2 = 1500 * Cos(PI / 180 * (6 * Minutes - 90)) + LineHour.X1
LineMinute.Y2 = 1500 * Sin(PI / 180 * (6 * Minutes - 90)) + LineHour.Y1

LineSecond.X2 = 1600 * Cos(PI / 180 * (6 * Seconds - 90)) + LineHour.X1
LineSecond.Y2 = 1600 * Sin(PI / 180 * (6 * Seconds - 90)) + LineHour.Y1
    
End Sub
