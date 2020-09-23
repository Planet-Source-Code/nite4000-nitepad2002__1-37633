VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form mp3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                       Niteplayer2002"
   ClientHeight    =   3660
   ClientLeft      =   2670
   ClientTop       =   1080
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdopen 
      Caption         =   "Command1"
      Height          =   675
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      WhatsThisHelpID =   2060
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      WhatsThisHelpID =   2060
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   3960
      Top             =   2040
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   2160
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   480
      WhatsThisHelpID =   2060
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   2060
      Width           =   2175
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      WhatsThisHelpID =   2060
      Width           =   4455
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "mp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SwitchVal As Boolean
Dim GenresTypes
Dim Min As Integer
Dim Sec As Integer

Dim FileName As String
Dim FileOpen As Boolean
Dim CurrentTag As TagInfo
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
'Declare all needed variables
Dim BitD As Boolean, StartA%, StartB%, StartC%, StartD%, StartE%, StartF%, abc%, Abort As Boolean
Private Type TagInfo
    Tag As String * 3
    Songname As String * 30
    artist As String * 30
    album As String * 30
    year As String * 4
    comment As String * 30
    genre As String * 1
End Type



Private Sub Command3_Click()

File1.ListIndex = File1.ListIndex + 1
On Error GoTo er
With MediaPlayer1

        .FileName = FileName
        .AutoStart = True
        cmdopen.Caption = "Close"

End With
er:
If Err.Number = 381 Then File1.ListIndex = 0
End Sub

Private Sub Command4_Click()
File1.ListIndex = File1.ListIndex - 1
On Error GoTo er
With MediaPlayer1

        .FileName = FileName
        .AutoStart = True
        cmdopen.Caption = "Close"

End With
er:
If Err.Number = 381 Then File1.ListIndex = 0
End Sub

Private Sub Drive1_Change()
'this will change the path of the folders
On Error GoTo errDrive
Dir1.Path = Drive1.Drive
errDrive:
    If Err.Number = 68 Then 'this means that the drive isnt accessable!
        MsgBox "Drive Not Ready!", vbOKOnly
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

Private Sub Timer3_Timer()
    If Text2See(SwitchVal) = True Then
       Label7.FontSize = 13
       Label7.Caption = " SAMIKARAEEN@YAHOO.COM"
    Else
        Label7.Caption = ""
    End If

End Sub

Public Function Text2See(Switch As Boolean) As Boolean
    Select Case Switch
        Case True
        Text2See = True
        SwitchVal = False
        Case False
        Text2See = False
        SwitchVal = True
    End Select
End Function






Private Sub cmdWriteTage_Click()

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub File1_Click()
Dim temp As String
On Error Resume Next
EraseTXTBoxes

If Right(Dir1.Path, 1) = "\" Then
    FileName = Dir1.Path & File1.FileName
Else
    FileName = Dir1.Path & "\" & File1.FileName
End If
    
Open FileName For Binary As #1
With CurrentTag
    Get #1, FileLen(FileName) - 127, .Tag
    If Not .Tag = "TAG" Then
        lblMsg.Caption = "No tag"
        Close #1
        Exit Sub
    End If
    Get #1, , .Songname
    Get #1, , .artist
    Get #1, , .album
    Get #1, , .year
    Get #1, , .comment
    Get #1, , .genre
    Close #1

    txtTitle = RTrim(.Songname)
    txtArtist = RTrim(.artist)
    txtAlbum = RTrim(.album)
    txtYear = RTrim(.year)
    txtComment = RTrim(.comment)
    
    temp = RTrim(.genre)
    txtGenreCode = Asc(temp)
    
End With
End Sub
Private Sub Dir1_change()
File1.FileName = Dir1.Path
End Sub

Private Sub cmdOpen_Click()
With MediaPlayer1
    If Not FileOpen Then
        .FileName = FileName
        .AutoStart = False
        cmdopen.Caption = "Close"
    Timer1.Enabled = True
    Else
        .FileName = ""
        cmdopen.Caption = "Open"
   Timer1.Enabled = False
    
  StartA = 0
    End If
End With
End Sub

Private Sub File1_DblClick()
With MediaPlayer1

        .FileName = FileName
        .AutoStart = True
        cmdopen.Caption = "Close"

End With
End Sub

Private Sub MediaPlayer1_OpenStateChange(ByVal _
    OldState As Long, ByVal NewState As Long)

Min = MediaPlayer1.Duration \ 60
Sec = MediaPlayer1.Duration - (Min * 60)
lblTotalTime = "Total Time: " & Format(Min, "0#") _
    & ":" & Format(Sec, "0#") 'format time to 00:00
    
FileOpen = CBool(NewState)
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
Min = MediaPlayer1.CurrentPosition \ 60
Sec = MediaPlayer1.CurrentPosition - (Min * 60)
If Min > 0 Or Sec > 0 Then
    lblElapsedTime = "Elapsed Time: " & Format(Min, "0#") _
        & ":" & Format(Sec, "0#")
Else
    lblElapsedTime = "Elapsed Time: 00:00"
End If
End Sub
Private Sub EraseTXTBoxes()
lblMsg.Caption = ""
txtTitle = ""
txtArtist = ""
txtAlbum = ""
txtYear = ""
txtComment = ""
txtGenreCode = ""

End Sub



Sub Form_Paint()
Dim i As Integer, X As Integer, Y As Integer
Dim C As String
Cls

X = CurrentX
Y = CurrentY
C = Chr(i)
'Line -(X + TextWidth(C), Y = TextHeight(C)), _
QBColor(Rnd * 16), BF
CurrentX = X
CurrentY = Y







End Sub



Private Sub Timer2_Timer()

 Form_Paint
End Sub




Private Sub Timer4_Timer()

End Sub

Private Sub Timer6_Timer()
'********************************************************'
'Progress with no end. With nice gradients and looks cool'
'Idea taken from Netscape Communicator 5 BETA Installer  '
'********************************************************'
'This is the never-ending progress. On some computers it
'produces some little line in the PictureBox (but looks kewl too :))
On Error Resume Next 'Errors suck :]
If Abort Then End 'If user abort operation exit the program
If Not BitD Then 'If painting-mode is left->right
  For a = 0 To 250 Step 2
    'Start painting the gradient on the left
    picProgress1.Line (StartA + a * 2, 0)-(StartA + a * 2 + 2, picProgress1.Height), RGB(0, a, 0), BF
  Next a
  'Paints the inner box of progress-mark
  picProgress1.Line (StartA + 500, 0)-(StartA + 1500, picProgress1.Height), RGB(0, 255, 0), BF
  For a = 0 To 250 Step 2
    'Start painting the gradient on the right
    picProgress1.Line ((StartA + 1500) + a * 2, 0)-((StartA + 1500) + a * 2 + 2, picProgress1.Height), RGB(0, 255 - a, 0), BF
    Next a
  'Increase marks position by 45
  StartA = StartA + 45
End If
'If painting-mode is right->left
If StartA + 2000 >= picProgress1.Width Or BitD = True Then
  BitD = True 'Must be set to reenter this sub
  For a = 0 To 250 Step 2 'The gradient again
    picProgress1.Line (StartA + a * 2, 0)-(StartA + a * 2 + 2, picProgress1.Height), RGB(0, a, 0), BF
  Next a 'And the block...
  picProgress1.Line (StartA + 500, 0)-(StartA + 1500, picProgress1.Height), RGB(0, 255, 0), BF
  For a = 0 To 250 Step 2 'Next!
    picProgress1.Line ((StartA + 1500) + a * 2, 0)-((StartA + 1500) + a * 2 + 2, picProgress1.Height), RGB(0, 255 - a, 0), BF
  Next a
  StartA = StartA - 45 'Decrease marks position by 45
  If StartA <= 0 Then BitD = False 'Set position back if way's finished
End If
'That's it. Nice one :)
End Sub
