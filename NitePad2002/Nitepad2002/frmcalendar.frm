VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmcalendar 
   Caption         =   "Calendar"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Time Clock"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   2
      Day             =   16
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmcalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' In the name of The Almighty, the Beneficent, the Merciful
'
'****************************************************************
'*     Author : K. O. Thaha Hussain MCA (thaha_ko@yahoo.com)     *
'*             URL : www.bcity.com/thahahussain                  *
'*  Copyright (c) K. O. Thaha Hussain.  All rights reserved      *
'*  Company : Indusware Solutions (www.induswareonline.com)      *
'*                Date : Thursday March 15 2001                  *
'*****************************************************************
' *** LICENSE AGREEMENT ***
' Get permission from the author to use the formulae commercially
' Feel free to make use of the Formulae for Non-Commercial Purposes,
' but the name of the Author should be accompanied along with the formulae.
'
' In most of the computer languages, X-axis is taken correctly.
' But Y-axis in the reverse of the normal Cartisian axis (Y-axis is
' incremented downword not upword) . So usual analytic manipulations
' such as, shifting the origin, calculation  of polar co-ordinates etc
' become difficult. Ofcourse, VB has techniques  to correct this problem.
' But for many lanuages, it is not available!
' Here is a solution to that problem.
'
' These are the difficulties faced while deriving the formulae.
'  -----------------------------------------------------------
' 1) The Y axis Problem.
' 2) Polar angles are measured in anti-clockwise direction, while the
'                     clock hands are moved in clockwise direction.
' 3) 'Zero' of polar angle and 'Zero' of Clock Hands causes a difference
'                                      of 90 degrees.
' The following are the Formulae obtained.

' May I call it "Thaha Hussain's Clock-Work formulae" ?
'
'
    'Hour Hand :
    'hour_x2 = LengthOfHourHand * Cos(PI/180*(30 * hour - 90)) + x1
    'hour_y2 = LengthOfHourHand * Sin(PI/180*(30 * hour - 90)) + y1
 '
    'Minute Hand:
    'minute_x2 = LengthOfMinuteHand * Cos(PI/180*(6 * minute - 90)) + x1
    'minute_y2 = LengthOfMinuteHand * Sin(PI/180*(6 * minute - 90)) + y1
'
    'Soconds Hand:
    'second_x2 = LengthOfSecondsHand * Cos(PI/180*(6 * second - 90)) + x1
    'second_y2 = LengthOfSecondsHand * Sin(PI/180*(6 * second - 90)) + y1
'
 '1) You can use the formulae in any Programming Language
 '                   without the Co-ordinate adjustment!
 '2) No problem for hour even in 24 Hr format!
 '3) You can adjust the Length of Clock Hands!
 '4) Shift the clock to anywhere by changing (midx,midy)
 '5) A variety of other uses in graphics!
 
'Simply Excellent! Right?
'
' a C Language Version is published in Dec 2000 issue of 'Electronics for you'
'           - Asia's most popular Electronis magazine (www.electronicsforu.com)
'
Const PI = 3.14159

Private Sub Command1_Click()
 Unload Me
 
End Sub

Private Sub Command2_Click()
  frmClock.Show
  
End Sub

Private Sub Form_Load()
Call tmrQuartz_Timer
End Sub



Private Sub tmrQuartz_Timer()
Dim Hours As Single, Minutes As Single, Seconds As Single
Dim TrueHours As Single


'Beep
Hours = Hour(Time)
Minutes = Minute(Time)
Seconds = Second(Time)
TrueHours = Hours + Minutes / 60

' I made all the X1 and Y1 equal in the form

End Sub
