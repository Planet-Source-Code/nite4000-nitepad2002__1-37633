VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   2805
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   3480
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   430
   Icon            =   "Calculator.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   3480
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton MemoKey 
      Caption         =   "MC"
      Height          =   360
      Index           =   3
      Left            =   2940
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2280
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "MR"
      Height          =   360
      Index           =   2
      Left            =   2940
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M-"
      Height          =   360
      Index           =   1
      Left            =   2940
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M+"
      Height          =   360
      Index           =   0
      Left            =   2940
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   660
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   1200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00808080&
      Caption         =   "C"
      Height          =   360
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00808080&
      Caption         =   "CE"
      Height          =   360
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   660
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   660
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      WhatsThisHelpID =   430
      Width           =   960
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      WhatsThisHelpID =   430
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      WhatsThisHelpID =   430
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   810
      TabIndex        =   1
      Top             =   60
      WhatsThisHelpID =   430
      Width           =   2595
      Begin VB.Label Readout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   180
         WhatsThisHelpID =   430
         Width           =   2475
      End
   End
   Begin VB.CommandButton CopyButton 
      BackColor       =   &H00808080&
      Caption         =   "<"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Copy to clipboard"
      Top             =   240
      WhatsThisHelpID =   430
      Width           =   315
   End
   Begin VB.Label lblMemoFlag 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   540
      TabIndex        =   25
      ToolTipText     =   "If M, memory  not zero"
      Top             =   300
      WhatsThisHelpID =   430
      Width           =   225
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const Maxdigits = 16        ' After this, scientific notation
Dim Op1 As Variant          ' Prev input operand
Dim Op2 As Variant          ' Further prev input operand
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Numkey of operands, 0 to 2
Dim LastInput As String     ' Indicate type of last keypress event.
Dim OpFlag As String        ' Indicate pending operation.
Dim PrevReadout As String   ' For restore if "CE"
Dim MemoResult              ' Store result for memo keys
Dim XReadout As String
Dim XOp1 As Variant
Dim XOp2 As Variant
Dim XDecimalFlag As Integer
Dim XNumOps As Integer
Dim XLastInput As String
Dim XOpFlag As String
Dim XCaption As String
Dim XMemoResult



Private Sub Form_Load()
    ResetStatus
End Sub


Sub ResetStatus()
    Readout = Format(0, "0.")
    PrevReadout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    lblMemoFlag.Caption = " "
    MemoResult = 0
End Sub


Sub RestoreStatus()
    Readout = XReadout
    Op1 = XOp1
    Op2 = XOp2
    DecimalFlag = XDecimalFlag
    NumOps = XNumOps
    LastInput = XLastInput
    OpFlag = XOpFlag
    lblMemoFlag.Caption = XCaption
    MemoResult = XMemoResult
End Sub


Sub MarkStatus()
    XReadout = Readout
    XOp1 = Op1
    XOp2 = Op2
    XDecimalFlag = DecimalFlag
    XNumOps = NumOps
    XLastInput = LastInput
    XOpFlag = OpFlag
    XCaption = lblMemoFlag.Caption
    XMemoResult = MemoResult
End Sub


Private Function MaxReached()
    MaxReached = False
    If Len(Readout) >= Maxdigits Then       ' Not allow further Numkey
         MaxReached = True
    End If
End Function


Function HasDecimal(strToRead As String)
    HasDecimal = False
    Dim i As Integer
    For i = Len(strToRead) To 1 Step -1
         If InStr(i, strToRead, ".") Then
             HasDecimal = True
             Exit For
         End If
    Next
End Function


' Copy the "Label" Caption onto the Clipboard.
Private Sub CopyButton_Click()
    Clipboard.SetText Readout
End Sub


Private Sub Cancel_Click()
    ResetStatus
    Operator(4).SetFocus
End Sub


Private Sub CancelEntry_Click()
    RestoreStatus
    LastInput = "CE"
    Operator(4).SetFocus
End Sub




Private Sub Decimal_Click()
    If HasDecimal(Readout) Then             ' One is enough
        Exit Sub
    End If
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If Len(Readout) = Maxdigits Then
            MsgBox "Maximum digits " & Str(Maxdigits - 1) + _
                vbCrLf & "Cannot add another digit"
                Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
    Me.Decimal.SetFocus
    MarkStatus
    
    If LastInput = "NEG" Then
        If Abs(Val(Readout)) <> 0 Then
            Readout = Format(0, "-0.")
        End If
    ElseIf LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, "0.")
    End If
    
    DecimalFlag = True
    LastInput = "DIGI"
    
    If MaxReached Then
        MsgBox "Maximum digits " & Str(Maxdigits - 1) + _
           vbCrLf & "Result overflowed"
        RestoreStatus
        Exit Sub
    End If
    Operator(4).SetFocus
End Sub



Private Sub Numkey_Click(Index As Integer)
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If MaxReached Then
            MsgBox "Maximum digits " & Str(Maxdigits - 1) + _
               vbCrLf & "Cannot add another digit"
            Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
    Me.NumKey(Index).SetFocus
    MarkStatus
    If LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + NumKey(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + NumKey(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then
        Readout = "-" & Readout
    End If
    LastInput = "NUMS"
  
    Operator(4).SetFocus
End Sub



Private Sub Operator_Click(Index As Integer)
    Me.Operator(Index).SetFocus
    MarkStatus
    
    Dim strTempreadout As String
    strTempreadout = Readout
    
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        NumOps = NumOps + 1
    End If
    
    Select Case NumOps
        Case 0
            If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-" & Readout
                    LastInput = "NEG"
                End If
            End If
        Case 1
            Op1 = Readout
            If Operator(Index).Caption = "-" And (LastInput <> "NUMS" _
                    And LastInput <> "DIGI") And OpFlag <> "=" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-"
                    LastInput = "NEG"
                End If
            End If
        Case 2
            Op2 = strTempreadout
            Select Case OpFlag
                Case "+"
                    Op1 = CDbl(Op1) + CDbl(Op2)
                Case "-"
                    Op1 = CDbl(Op1) - CDbl(Op2)
                Case "*"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                Case "/"
                    If Op2 = 0 Then
                       MsgBox "Can't divide by zero", 48, "Calculator"
                       RestoreStatus
                       Exit Sub
                    Else
                       Op1 = CDbl(Op1) / CDbl(Op2)
                    End If
               Case "="
                    Op1 = CDbl(Op2)
             End Select
             Readout = Op1
             NumOps = 1
             
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
    
     ' Be consistent, since we always show a decimal point
    If Not HasDecimal(Readout) Then
        If Abs(Val(Readout)) = 0 Then
           Readout = "0."
        Else
           Readout = Readout + "."
        End If
    End If
    
    Operator(4).SetFocus
End Sub




Private Sub MemoKey_Click(Index As Integer)
    MarkStatus
    Select Case Index
       Case 0                    ' Memory Plus
            MemoResult = MemoResult + Val(Readout)
       Case 1                    ' Memory Minus
            MemoResult = MemoResult - Val(Readout)
       Case 2                    ' Memory Recall
            Dim s As String
            s = Str(MemoResult)
            If Not HasDecimal(Str(s)) Then
                s = s + "."
            End If
            Readout = s
       Case 3                    ' Memory Clear
            MemoResult = 0
    End Select
     ' Our system is, if MemoResult is not cleared, show "M"
    If MemoResult <> 0 Then
         lblMemoFlag.Caption = "M"
    Else
         lblMemoFlag.Caption = " "
    End If
    
    LastInput = "OPS"
    NumOps = 1
    Op1 = Readout
    Op2 = 0
    Operator(4).SetFocus
End Sub



' Detect keyboard key
Private Sub Form_KeyPress(KeyAscii As Integer)
    MarkStatus
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 46 And KeyAscii <> 43 And _
           KeyAscii <> 45 And KeyAscii <> 42 And _
           KeyAscii <> 47 And KeyAscii <> 61 And _
           KeyAscii <> 13 Then
               KeyAscii = 0
        Else
           Select Case KeyAscii
             Case 46                   ' "."
               Decimal_Click
             Case 43
               Operator_Click (0)      ' re Property "+"
             Case 45                   ' "-"
               Operator_Click (1)
             Case 42                   ' "*"
               Operator_Click (2)
             Case 47                   ' "/"
               Operator_Click (3)
             Case 61                   ' "="
               Operator_Click (4)
             Case 13                   ' As "=" (if Windows allows Enter)
               Operator_Click (4)
           End Select
        End If
    Else
        Numkey_Click (Val(Chr(KeyAscii)))
    End If
End Sub



