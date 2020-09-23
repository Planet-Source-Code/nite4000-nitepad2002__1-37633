Attribute VB_Name = "Module3"
Public lngMenu As Long, lngSubMenu As Long, lngMenuItemID As Long, lngRet As Long

Public Declare Function GetMenu Lib "user32" _
(ByVal hWnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As _
Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked _
As Long) As Long


 Public myCancel As Boolean
 
Public undoStack(99) As String
Public undoStage As Integer
Public Undoing As Boolean
Public CountUndo As Integer

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long







Public Sub CenterForm(F As Form)
    F.Left = (Screen.Width - F.Width) / 2
    F.Top = (Screen.Height - F.Height) / 2
End Sub


