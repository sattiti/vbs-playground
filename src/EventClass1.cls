VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents eventObj As EventClass2
Attribute eventObj.VB_VarHelpID = -1


Private Sub Class_Initialize()
  Call init
End Sub

Private Sub init()
    Set eventObj = New EventClass2
    eventObj.testMethod
End Sub

Private Sub eventObj_callback(ByRef strMsg As String)
    MsgBox strMsg
End Sub
