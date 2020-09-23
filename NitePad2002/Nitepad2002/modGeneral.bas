Attribute VB_Name = "modGeneral"

Option Explicit

Public Const MRUPath = "Software\Nitepad2002\MRUList"
Public Const ViewPath = "Software\Nitepad2002\View"
Public Const SettingsPath = "Software\Nitepad2002\Settings"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long


Sub Main()
    Load frmSplash 'Load Splash Screen form
    frmSplash.Startup 'Call StartUp sub
End Sub

Public Function GetTotalLines(RichTextBox As RichTextBox)
    Dim TotalLines As Long
    TotalLines = SendMessage(RichTextBox.hWnd, EM_GETLINECOUNT, 0, 0&)
    frmMDI.SB.Panels(5).Text = Format(TotalLines, "###,###,###,###")
End Function

Public Function GetCurrentLine(RichTextBox As RichTextBox)
    Dim CurrentLine As Long
    CurrentLine = SendMessage(RichTextBox.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    frmMDI.SB.Panels(3).Text = Format(CurrentLine, "###,###,###,###")
End Function

Public Function GetSettings()
    On Error GoTo GetSettingsError
    Dim ViewMode As Integer
    With frmMDI
        ' Set the status bar
        .mnuViewStatusBar.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Status Bar", 1)
        .SB.Visible = .mnuViewStatusBar.Checked
        ' Set the ruler
        .mnuViewRuler.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Ruler", 1)
        .cbrRuler.Visible = .mnuViewRuler.Checked
        ' Set the toolbar:Standard
        .mnuViewToolbarStandard.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Standard Toolbar", 1)
        .cbrBar.Bands(1).Visible = .mnuViewToolbarStandard.Checked
        ' Set the toolbar:File
        .mnuViewToolbarFile.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "File Toolbar", 0)
        .cbrBar.Bands(4).Visible = .mnuViewToolbarFile.Checked
        ' Set the toolbar:Edit
        .mnuViewToolbarEdit.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Edit Toolbar", 0)
        .cbrBar.Bands(5).Visible = .mnuViewToolbarEdit.Checked
        ' Set the toolbar:Format
        .mnuViewToolbarFormat.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Format Toolbar", 1)
        .cbrBar.Bands(3).Visible = .mnuViewToolbarFormat.Checked
        ' Set the toolbar:Font
        .mnuViewToolbarFont.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Font Toolbar", 1)
        .cbrBar.Bands(2).Visible = .mnuViewToolbarFont.Checked
        ' Set the toolbar:Window
        .mnuViewToolbarWindow.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Window Toolbar", 0)
        .cbrBar.Bands(6).Visible = .mnuViewToolbarWindow.Checked
        ' Set Stay on Top
        .mnuViewStayonTop.Checked = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Stay On Top", 0)
        If .mnuViewStayonTop.Checked Then OnTop frmMDI
        ' Set View Mode
        ViewMode = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "ViewMode", 1)
        .mnuViewMode(ViewMode).Checked = True
        SetViewMode ViewMode
        ' Set font name and size
        .cboFontName.Text = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Font Name", "Tahoma")
        .ActiveForm.rtfText.Font.Name = .cboFontName.Text
        .cboFontSize.Text = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Font Size", 9)
        .ActiveForm.rtfText.Font.Size = .cboFontSize.Text
    End With
GetSettingsError:
    ErrorLog "modGeneral\GetSettings"
End Function

Public Sub EnableAll()
    On Error GoTo EnableAllError
    Dim i As Integer
    With frmMDI
        
        .mnuFileSave.Enabled = True
        .mnuFileSaveAs.Enabled = True
        .mnuFileSaveAll.Enabled = True
        .mnuFilePrint.Enabled = True
        '------------------------------
        .mnuEditSelectAll.Enabled = True
        '------------------------------
        .mnuSearchFind.Enabled = True
        .mnuSearchReplace.Enabled = True
        .mnuSearchGoTo.Enabled = True
        '------------------------------
        .mnuViewDocumentProperties.Enabled = True
        '------------------------------
        .mnuInsertTimeDate.Enabled = True
        .mnuInsertPicture.Enabled = True
        .mnuInsertTextFile.Enabled = True
        .mnuInsertPathandFile.Enabled = True
        .mnuInsertSymbols.Enabled = True
        '------------------------------
        .mnuFormatFont.Enabled = True
        .mnuFormatIndent.Enabled = True
        .mnuFormatOutdent.Enabled = True
        .mnuFormatLower.Enabled = True
        .mnuFormatUpper.Enabled = True
        .mnuFormatScript.Enabled = True
        '-----TOOLBARS-----'
        '------------------------------
        With .tbrStandard
            .Buttons("Save").Enabled = True
            .Buttons("Print").Enabled = True
            .Buttons("FullScreen").Enabled = True
            .Buttons("WordWrap").Enabled = True
            .Buttons("Indent").Enabled = True
            .Buttons("Outdent").Enabled = True
        End With
        '------------------------------
        With .tbrFormat
            .Buttons("Find").Enabled = True
            .Buttons("Bold").Enabled = True
            .Buttons("Italic").Enabled = True
            .Buttons("Underline").Enabled = True
            .Buttons("StrikeThru").Enabled = True
            .Buttons("Left").Enabled = True
            .Buttons("Center").Enabled = True
            .Buttons("Right").Enabled = True
            .Buttons("Bullet").Enabled = True
        End With
        '------------------------------
        With .tbrFile
            .Buttons("Close").Enabled = True
            .Buttons("Save").Enabled = True
            .Buttons("SaveAll").Enabled = True
            .Buttons("Print").Enabled = True
        End With
        '------------------------------
        With .tbrEdit
            .Buttons("Indent").Enabled = True
            .Buttons("Outdent").Enabled = True
        End With
    End With
EnableAllError:
    ErrorLog "modGeneral\EnableAll"
End Sub

Public Sub DisableAll()
    On Error GoTo DisableAllError
    Dim i As Integer
    
    With frmMDI
        
        .mnuFileSave.Enabled = True
        .mnuFileSaveAs.Enabled = True
        .mnuFileSaveAll.Enabled = True
        .mnuFilePrint.Enabled = True
        .mnufilesendemail.Enabled = True
        
        '------------------------------
        .mnuEditUndo.Enabled = True
        .mnuEditCut.Enabled = True
        .mnuEditCopy.Enabled = True
        .mnuEditPaste.Enabled = True
        .mnuEditDelete.Enabled = True
        .mnuEditSelectAll.Enabled = True
        
        '------------------------------
        .mnuSearchFind.Enabled = True
        .mnuSearchFindNext.Enabled = True
        .mnuSearchReplace.Enabled = True
        .mnuSearchGoTo.Enabled = True
        '------------------------------
        .mnuViewDocumentProperties.Enabled = True
        '------------------------------
        .mnuInsertTimeDate.Enabled = True
        .mnuInsertPicture.Enabled = True
        .mnuInsertTextFile.Enabled = True
        .mnuInsertPathandFile.Enabled = True
        .mnuInsertSymbols.Enabled = True
        '------------------------------
        .mnuFormatFont.Enabled = True
        .mnuFormatIndent.Enabled = True
        .mnuFormatOutdent.Enabled = True
        .mnuFormatLower.Enabled = True
        .mnuFormatUpper.Enabled = True
        .mnuFormatScript.Enabled = True
        '-----TOOLBARS-----'
        '------------------------------
        With .tbrStandard
            .Buttons("Save").Enabled = True
            .Buttons("Print").Enabled = True
            .Buttons("FullScreen").Enabled = True
            .Buttons("WordWrap").Enabled = True
            .Buttons("Cut").Enabled = True
            .Buttons("Copy").Enabled = True
            .Buttons("Paste").Enabled = True
            .Buttons("Undo").Enabled = True
            .Buttons("Indent").Enabled = True
            .Buttons("Outdent").Enabled = True
        End With
        '------------------------------
        With .tbrFormat
            .Buttons("Find").Enabled = True
            .Buttons("Bold").Enabled = True
            .Buttons("Italic").Enabled = True
            .Buttons("Underline").Enabled = True
            .Buttons("StrikeThru").Enabled = True
            .Buttons("Left").Enabled = True
            .Buttons("Center").Enabled = True
            .Buttons("Right").Enabled = True
            .Buttons("Bullet").Enabled = True
        End With
        '------------------------------
        With .tbrFile
            .Buttons("Close").Enabled = True
            .Buttons("Save").Enabled = True
            .Buttons("SaveAll").Enabled = True
            .Buttons("Print").Enabled = True
            .Buttons("Email").Enabled = True
            
        End With
        '------------------------------
        With .tbrEdit
            .Buttons("Cut").Enabled = True
            .Buttons("Copy").Enabled = True
            .Buttons("Paste").Enabled = True
            .Buttons("Undo").Enabled = True
            .Buttons("Indent").Enabled = True
            .Buttons("Outdent").Enabled = True
        End With
    End With
DisableAllError:
    ErrorLog "modGeneral\DisableAll"
End Sub

Public Function SetAll()
    On Error Resume Next
    With frmMDI
        '------MENUS----------------
 
        .mnuEditUndo.Enabled = SendMessage(.ActiveForm.rtfText.hWnd, EM_CANUNDO, 0, 0&)
        .mnuEditCut.Enabled = .ActiveForm.rtfText.SelLength
        .mnuEditCopy.Enabled = .ActiveForm.rtfText.SelLength
        .mnuEditPaste.Enabled = Clipboard.GetFormat(vbCFText)
        .mnuEditDelete.Enabled = .ActiveForm.rtfText.SelLength
        '------TOOLBARS-------------
        '---------------------------
        With .tbrStandard
            .Buttons("Undo").Enabled = SendMessage(frmMDI.ActiveForm.rtfText.hWnd, EM_CANUNDO, 0, 0&)
            .Buttons("Cut").Enabled = frmMDI.ActiveForm.rtfText.SelLength
            .Buttons("Copy").Enabled = frmMDI.ActiveForm.rtfText.SelLength
            .Buttons("Paste").Enabled = Clipboard.GetFormat(vbCFText)
        End With
        '---------------------------
        With .tbrEdit
            .Buttons("Cut").Enabled = frmMDI.ActiveForm.rtfText.SelLength
            .Buttons("Copy").Enabled = frmMDI.ActiveForm.rtfText.SelLength
            .Buttons("Paste").Enabled = Clipboard.GetFormat(vbCFText)
        End With
        '---------------------------
        ' Set Align Left,Center and Right buttons
        If .ActiveForm.rtfText.SelAlignment = rtfLeft Then
            frmMDI.tbrFormat.Buttons("Left").Value = tbrPressed
            frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
            frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        ElseIf .ActiveForm.rtfText.SelAlignment = rtfCenter Then
            frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
            frmMDI.tbrFormat.Buttons("Center").Value = tbrPressed
            frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        ElseIf .ActiveForm.rtfText.SelAlignment = rtfRight Then
            frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
            frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
            frmMDI.tbrFormat.Buttons("Right").Value = tbrPressed
        End If
        ' Set Bold button
        If .ActiveForm.rtfText.SelBold = True Then
            frmMDI.tbrFormat.Buttons("Bold").Value = tbrPressed
        Else
            frmMDI.tbrFormat.Buttons("Bold").Value = tbrUnpressed
        End If
        ' Set Italic button
        If .ActiveForm.rtfText.SelItalic = True Then
            frmMDI.tbrFormat.Buttons("Italic").Value = tbrPressed
        Else
            frmMDI.tbrFormat.Buttons("Italic").Value = tbrUnpressed
        End If
        ' Set Underline button
        If .ActiveForm.rtfText.SelUnderline = True Then
            frmMDI.tbrFormat.Buttons("Underline").Value = tbrPressed
        Else
            frmMDI.tbrFormat.Buttons("Underline").Value = tbrUnpressed
        End If
        ' Set Strikethru button
        If .ActiveForm.rtfText.SelStrikeThru = True Then
            frmMDI.tbrFormat.Buttons("StrikeThru").Value = tbrPressed
        Else
            frmMDI.tbrFormat.Buttons("StrikeThru").Value = tbrUnpressed
        End If
        ' Set Bullet button
        If .ActiveForm.rtfText.SelBullet = True Then
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrPressed
        Else
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrUnpressed
        End If
        ' Set Font
        .cboFontName.Text = .ActiveForm.rtfText.SelFontName
        .cboFontSize.Text = .ActiveForm.rtfText.SelFontSize
    End With
End Function
