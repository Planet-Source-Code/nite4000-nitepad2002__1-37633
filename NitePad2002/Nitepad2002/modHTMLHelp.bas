Attribute VB_Name = "modHTMLHelp"

Option Explicit

Public Const HH_HELP_CONTEXT = &HF
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const HH_DISPLAY_TEXT_POPUP = &HE

Private Type tagHH_FTS_QUERY
    cbStruct          As Long
    fUniCodeStrings   As Long
    pszSearchQuery    As String
    iProximity        As Long
    fStemmedSearch    As Long
    fTitleOnly        As Long
    fExecute          As Long
    pszWindow         As String
End Type

Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByRef dwData As tagHH_FTS_QUERY) As Long

'----------SHOW HTMLHELP SEARCH----------'
Public Function HHShowSearch(lhWnd As Long)
    Dim HHFQ As tagHH_FTS_QUERY
    With HHFQ
        .cbStruct = Len(HHFQ)
        .fUniCodeStrings = 0&
        .pszSearchQuery = ""
        .iProximity = 0&
        .fStemmedSearch = 0&
        .fTitleOnly = 0&
        .fExecute = 1&
        .pszWindow = ""
    End With
    HTMLHelpCallSearch lhWnd, App.Path & "\Nitepad2002.chm" & ">Main", HH_DISPLAY_SEARCH, HHFQ
End Function

'----------SHOW HTMLHELP CONTENTS----------'
Public Function HHShowContents(lhWnd As Long)
    HTMLHelp lhWnd, App.Path & "\Nitepad2002.chm" & ">Main", HH_DISPLAY_TOC, 0
End Function

'----------SHOW HTMLHELP INDEX----------'
Public Function HHShowIndex(lhWnd As Long)
    HTMLHelp lhWnd, App.Path & "\Nitepad2002.chm" & ">Main", HH_DISPLAY_INDEX, 0
End Function

'----------SHOW HTMLHELP TOPIC----------'
Public Function HHShowTopic(lhWnd As Long, lngTopicID As Long)
    HTMLHelp lhWnd, App.Path & "\Nitepad2002.chm" & ">Main", HH_HELP_CONTEXT, lngTopicID
End Function
