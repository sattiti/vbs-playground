Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Private Function ccb(ByVal r As String)
  Dim hgm As Long
  Dim lgm As Long
  Dim hcm As Long
  Dim x As Long
  
  hgm = GlobalAlloc(GHND, LenB(r) + 1)
  lgm = GlobalLock(hgm)
  lgm = lstrcpy(lgm, r)
  
  ' Unlock the memory.
  If GlobalUnlock(hgm) <> 0 Then
    MsgBox "Copy aborted."
    GoTo CBError
  End If
 
   If OpenClipboard(0&) = 0 Then
    MsgBox "Copy aborted."
    Exit Function
   End If
   
   x = EmptyClipboard()
 
   ' Copy data to the Clipboard.
   hcm = SetClipboardData(CF_TEXT, hgm)
   MsgBox "コピーしました。"

CBError:
If CloseClipboard() = 0 Then
  MsgBox "Could not close Clipboard."
End If
End Function
