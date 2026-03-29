Attribute VB_Name = "WebBridge"
Option Explicit

Public Function PingBridge() As String
    PingBridge = "pong"
End Function

Public Function PickTemplateRange(ByVal currentSheet As String) As String
    PickTemplateRange = WebHost.PickTemplateRange(currentSheet)
End Function

