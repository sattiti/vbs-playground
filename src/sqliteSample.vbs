Option Explicit

Sub test()
  Dim cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  dim i as Variant
  dim dbPath as String

  dbPath              = "\PATH\TO\DATABASE.db"
  Set cn              = New ADODB.Connection
  cn.connectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & dbPath

  cn.Open

  Set rs = cn.Execute("SELECT id, uri FROM pages;")
  i      = 1
  Do Until rs.EOF = True
    Worksheets(1).Cells(i, 1).Value = rs.Fields("uri").Value
    rs.MoveNext
    i = i + 1
  Loop
  
  rs.Close: Set rs = Nothing
  cn.Close: Set cn = Nothing
End Sub

Sub test2()
  Dim cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  dim server, port, dbName, user, passwd as String

  Set cn = New ADODB.Connection
  server = "http://SERVER;"
  port   = "8080;"
  dbName = "ABCD;"
  user   = "USERNAME;"
  passwd = "PASSWD;"

  cn.connectionString = "DRIVER={MySQL ODBC 5.34 Driver};" _
  & "SERVER=" & server _
  & "PORT=" & port _
  & "DATABASE=" & dbName _
  & "USER=" & user _
  & "PASSWORD=" & passwd

  cn.Open connectionString

  rs.Close: Set rs = Nothing
  cn.Close: Set cn = Nothing
End Sub
