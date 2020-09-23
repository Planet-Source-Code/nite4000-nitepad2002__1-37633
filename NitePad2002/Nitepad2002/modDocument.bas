Attribute VB_Name = "modDocument"
Option Explicit

' Filter
Public Const epFilter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|Log Files (*.log)|*.log|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|All Files (*.*)|*.*|"








Public Function CreateNewDocument()
    '** Description:
    '** Create a new document
    On Error GoTo NewError
    Static DocCount As Long
    Dim frmDoc As frmDocument
    
    Set frmDoc = New frmDocument 'Create new form
    DocCount = DocCount + 1 'Increase document counter
    frmDoc.Caption = "Document " & DocCount 'Set document caption
    frmDoc.Show 'Show document
NewError:
    ErrorLog "modDocument/CreateNewDocument"
End Function


Public Function AlignLeft()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfLeft 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrPressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function AlignCenter()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfCenter 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrPressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function AlignRight()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfRight 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrPressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function Bullet()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm.rtfText
        'If there is not bullet
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            .SelBullet = True 'Put it
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrPressed
        ElseIf .SelBullet = True Then 'If there is bullet
            .SelBullet = False 'Remove it
            .SelHangingIndent = False
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrUnpressed
        End If
    End With
End Function



Public Function GetFTitle(strFilename As String)
    '** Description:
    '** Get file title from file name
    On Error GoTo GFTError
    Dim cbBuf As String
    
    cbBuf = String(250, vbNullChar) 'Fill buffer with null chars
    
    GetFTitle = Left(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer
GFTError:
    ErrorLog "modFileOperations/GetFTitle"
End Function





Public Function CommLineFile()
    Dim sFile As String
    Dim fType As String
    
    If Command$ <> "" Then
        sFile = Command$ 'Get command line filename
    End If
  
    'Get file extension
    Select Case UCase(Right(sFile, 3))
        Case "RTF"
            fType = rtfRTF
        Case Else
            fType = rtfText
    End Select
    
    On Error GoTo CommResume
    frmDocument.rtfText.LoadFile sFile, fType 'Load file
    frmDocument.Caption = sFile 'Set caption
    frmDocument.bChanged = False 'Set bChanged flag to false
CommResume:
    If Left(sFile, 1) = Chr(34) Then 'Remove "" from filename
        sFile = Right(sFile, Len(sFile) - 1)
        sFile = Left(sFile, Len(sFile) - 1)
    End If
    frmDocument.rtfText.LoadFile sFile, fType 'Load file
    frmDocument.Caption = sFile 'Set caption
    frmDocument.bChanged = False 'Set bChanged flag to false
End Function

Public Function SaveDocument()
    '** Description:
    '** Save the active document
    On Error GoTo SaveError
    Dim fType As String
    
    With frmMDI.ActiveForm
        If Left(.Caption, 8) = "Document" Then
            'If it is not saved then call SaveDocumentAs function
            SaveDocument
        Else
            ' Get file extension
            If UCase(Right(.Caption, 3)) = "RTF" Then
                fType = rtfText
            Else
                fType = rtfText
            End If
            .rtfText.SaveFile .Caption, fType 'Save document
            .bChanged = False 'Set bChanged flag to false
        End If
    End With
SaveError:
    ErrorLog "modDocument/SaveDocument"
End Function


