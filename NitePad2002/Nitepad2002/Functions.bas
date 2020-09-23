Attribute VB_Name = "Functions"
Option Explicit

Private Lp, Ctr, Tmr

Public Function GetFolder() As String

Dim bi As browseinfo
Dim pidl As Long
Dim folder As String
    
folder = Space$(255)
With bi
 .howner = paintmain.hwnd
 .ulflags = BIF_RETURNONLYFSDIRS
 .pidlroot = 0
 .lpsztitle = "Select a Folder" & Chr$(0)
End With
    
pidl = shbrowseforfolder(bi)
If SHGetPathFromIDList(ByVal pidl, ByVal folder) Then
  GetFolder = Left(folder, InStr(folder, Chr$(0)) - 1)
End If
    
End Function

Public Function GetFile() As String

Dim rst As Long
Dim ofn As OPENFILENAME
Const mbl = 256

With ofn
 .hwndOwner = paintmain.hwnd
 .hInstance = App.hInstance
 .lpstrTitle = "Open Image File"
 '.lpstrInitialDir =
 .lpstrFilter = "All Files" & Chr(0) & "*.*"
 .nFilterIndex = 1
 .lpstrFile = String(mbl, 0)
 .nMaxFile = mbl - 1
 .lpstrFileTitle = .lpstrFile
 .nMaxFileTitle = mbl - 1
 .lStructSize = Len(ofn)
End With

rst = GetOpenFileName(ofn)

If rst Then
  GetFile = Left(ofn.lpstrFile, ofn.nMaxFile)
End If

End Function

Public Function GetColor() As Long

Dim rst As Long
Dim pcc As CHOOSECOLOR_TYPE
Dim cc() As Byte

With pcc
 .hwndOwner = paintmain.hwnd
 .hInstance = App.hInstance
 .lpCustColors = StrConv(cc, vbUnicode)
 .Flags = 0
 .lStructSize = Len(pcc)
End With

rst = ChooseColor(pcc)

If rst Then
  GetColor = pcc.rgbResult
End If

End Function

