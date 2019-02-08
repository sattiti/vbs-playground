Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  If Intersect(Target, Range("A:Z")) Is Nothing Then Exit sub
  Application.EnableEvents = False
  
  Dim lRow, lCol As Long
  Dim fold As Bool
  
  lRow = Cells(Rows.Count, 1).End(xlUp).Row
  lCol = Cells(1, Columns.Count).End(xlToLeft).Column
  fold = true
  
  Dim v As String
  For i = 1 To lRow
    If Target.Count = 1 then
      If c.Value Like "*" & Target.Value & "*" then
        v = v & c.Value & "*"
      End If
    End If
  Next
  
  If Len(v) > 0 then
    With Target.Validation
      .Delete
      .Add Type := xlValidateList, Formula1 := v
      .ShowError = False
      .ShowInput = true
      .InCellDropdown = True
    End With
    
    If fold then
      Target.Select
      SendKeys "{%DOWN}"
    End If
    
    fold = False
  End If
  
  Application.EnableEvents = True
Exit Sub