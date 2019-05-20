Private Sub export(ByVal result As String)
  Dim output As String
  Dim fso As Object
  Dim wsh As Variant
  Dim dt, ext As String
  Dim s As ADODB.Stream
  
  Set wsh = CreateObject("WScript.Shell")
  dt = wsh.SpecialFolders("Desktop") & "\"
  ext = ".txt"
  
  ' dest
  output = dt & "table-" & format(Date, "yyyymmdd") & format(Time, "hhmmss") & ext
  
  On Error GoTo SaveError:
  Set s = New ADODB.Stream
  With s
    .Type = adTypeText
    .Charset = "UTF-8"
    .LineSeparator = adLF
    .Open
    .WriteText result
    .Position = 0
    .Type = adTypeBinary
    .Position = 3
  End With
  
  Dim buf As Variant
  buf = s.Read()
  
  With s
    .Position = 0
    .Write buf
    .SetEOS
    .SaveToFile output, adSaveCreateOverWrite
    .Close
  End With
  
  Set s = Nothing
  MsgBox output & vbCrLf & "Saved!"
  
  Set wsh = Nothing
  Exit Sub
  
SaveError:
  MsgBox "Error."
  Set wsh = Nothing
  Exit Sub
End Sub
