Attribute VB_Name = "WebProgress"
Option Explicit

Public Sub ReportProgress(ByVal stage As String, ByVal current As Long, ByVal total As Long, ByVal percent As Long, ByVal message As String)
      Dim payload As String
      Dim js As String

      On Error Resume Next

      If total <= 0 Then
          percent = 0
      Else
          If percent < 0 Then percent = 0
          If percent > 100 Then percent = 100
      End If

      payload = "{""stage"":" & QuoteJson(stage) & ",""current"":" & CStr(current) & ",""total"":" & CStr(total) & ",""percent"":" & CStr(percent) & ",""message"":" & QuoteJson(message) & "}"

      If Not frmWebHost Is Nothing Then
          If Not frmWebHost.wbMain Is Nothing Then
              If Not frmWebHost.wbMain.Document Is Nothing Then
                  frmWebHost.wbMain.Document.getElementById("status").innerText = message
                  frmWebHost.wbMain.Document.getElementById("output").innerText = message

                  js = "window.hostBridge.onProgress(" & payload & ");"
                  frmWebHost.wbMain.Document.parentWindow.execScript js, "JavaScript"
              End If
          End If
      End If

      DoEvents
  End Sub

Private Function QuoteJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, Chr$(34), "\" & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    QuoteJson = Chr$(34) & value & Chr$(34)
End Function
