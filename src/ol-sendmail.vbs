Sub sendmail()
  Dim mSubject As String
  Dim mTo As String
  Dim mCC As String
  Dim mBody As String
  
  Dim ol As Outlook.Application
  Dim m As Outlook.MailItem
  
  mTo = ActiveSheet.Range("B1").Value()
  mCC = ActiveSheet.Range("B2").Value()
  mSubject = ActiveSheet.Range("B3").Value()
  mBody = ActiveSheet.Range("B4").Value()
  
  Set ol = New Outlook.Application
  Set m = ol.CreateItem(olMailItem)
    
  With m
    .To = mTo
    .CC = mCC
    .subject = mSubject
    .Body = mBody
    .BodyFormat = olFormatPlain
    .Display
  End With
  
  Set ol = Nothing
  Set m = Nothing
End Sub
