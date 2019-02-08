Option Explicit

' 0900
Private Const kStartTime = 32400

' 1700
Private Const kEndTime = 61200


#If VBA7 Then
  ' 64 Bit
  Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
#Else
  ' 32 Bit
  Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If


Sub test()
  Dim tt As Long
  
  Do While 1
    tt = Math.Round(Timer())
    If tt < kStartTime And tt > kEndTime Then
      Exit Do
    End If

    If (tt Mod 3600) Mod 60 = 0 Then
      Debug.Print tt
      Debug.Print Time
    End If

    Sleep (1000)
    DoEvents
  Loop

End Sub
