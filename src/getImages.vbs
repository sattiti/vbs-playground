Option Explicit

Sub a()
  Dim u As String
  Dim img As Variant
  Dim startRow, endRow, startCol, imgMargin As Long
  Dim i As Variant
  Dim c As Range
  
  imgMargin = 5
  startRow = 2
  endRow = ActiveSheet.Cells(Rows.Count, startRow).End(xlUp).Row
  startCol = 5
  
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
  
  For i = startRow To endRow
    Set c = ActiveSheet.Cells(i, startCol)
    If c.EntireRow.Hidden = False Then
      u = CStr(c.Value)
      
      If getHttpStatusCode(u) = 200 Then
        Set img = ActiveSheet.Pictures.Insert(u)

        If Not img Is Nothing Then
          img.Top = c.Offset(0, 5).Top + imgMargin
          img.Left = c.Offset(0, 5).Left + imgMargin
        End If
      End If

      Set img = Nothing
    End If
    Set c = Nothing
  Next
  
  Application.Calculation = xlCalculationAutomatic
  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating = True
End Sub


Sub d()
  Dim shp As Variant
  For Each shp In ActiveSheet.Shapes
    If shp.Type = msoLinkedPicture Then shp.Delete
  Next shp
End Sub




Function getHttpStatusCode(u As String) As Integer
  Dim WinHttp As Object
  Dim c As Object
  Set c = CreateObject("MSXML2.XMLHTTP")
  On Error GoTo error
  c.Open "GET", u, False
  c.send
  getHttpStatusCode = c.Status
  Set c = Nothing
  
  Exit Function

error:
  getHttpStatusCode = 0
  Set c = Nothing
End Function
