Attribute VB_Name = "WebHost"
Option Explicit

Public Sub 打开证件批量生成工具()
    Dim htmlPath As String
    htmlPath = ThisWorkbook.Path & "\src\webui\index.html"

    On Error GoTo HandleError
    frmWebHost.Show vbModeless
    frmWebHost.wbMain.Navigate htmlPath
    Exit Sub

HandleError:
    MsgBox "Cannot open web panel. Please confirm frmWebHost and wbMain exist." & vbCrLf & _
           "HTML path: " & htmlPath, vbExclamation
End Sub

Public Function PickTemplateRange(ByVal currentSheet As String) As String
    On Error GoTo HandleError
    PickTemplateRange = Application.InputBox("请框选模板区域", "选择模板区域", Type:=8).Address(False, False)
    Exit Function

HandleError:
    PickTemplateRange = ""
End Function

Public Function HandleBrowserCommand(ByVal fullUrl As String) As String
    Dim commandText As String
    commandText = ParseCommand(fullUrl)

    Select Case LCase$(commandText)
        Case "ping"
            HandleBrowserCommand = "pong"
        Case "pickrange"
            HandleBrowserCommand = PickTemplateRange(GetParamValue(fullUrl, "sheet"))
        Case "getinitdata"
            HandleBrowserCommand = WebConfig.GetInitData()
        Case "getsheetfields"
            HandleBrowserCommand = WebConfig.GetSheetFields(GetParamValue(fullUrl, "sheet"), GetParamValue(fullUrl, "range"))
        Case "validaterange"
            HandleBrowserCommand = WebConfig.ValidateRange(GetParamValue(fullUrl, "sheet"), GetParamValue(fullUrl, "range"))
        Case "generatetemplate"
      HandleBrowserCommand = WebActions.GenerateTemplate(GetParamValue(fullUrl, "templateSheet"), GetParamValue(fullUrl, "templateRange"), GetParamValue(fullUrl, "generateCount"), GetParamValue(fullUrl, "perRowCount"))
      ShowResultMessage HandleBrowserCommand
        Case "writedata"
      HandleBrowserCommand = WebActions.WriteData(GetParamValue(fullUrl, "dataSheet"), GetParamValue(fullUrl, "templateSheet"), GetParamValue(fullUrl, "templateRange"), GetParamValue(fullUrl, "generateCount"), GetParamValue(fullUrl, "perRowCount"), GetParamValue(fullUrl, "pageField"))
      ShowResultMessage HandleBrowserCommand
        Case Else
            HandleBrowserCommand = "unknown-command"
    End Select
End Function

Public Function ParseCommand(ByVal fullUrl As String) As String
    ParseCommand = GetParamValue(fullUrl, "cmd")
End Function

Public Function ParseRequestId(ByVal fullUrl As String) As String
    ParseRequestId = GetParamValue(fullUrl, "rid")
End Function

Public Function GetParamValue(ByVal fullUrl As String, ByVal keyName As String) As String
    Dim hashPos As Long
    Dim hashText As String
    Dim parts() As String
    Dim i As Long
    Dim prefix As String

    hashPos = InStr(1, fullUrl, "#", vbTextCompare)
    If hashPos = 0 Then Exit Function

    hashText = Mid$(fullUrl, hashPos + 1)
    parts = Split(hashText, "&")
    prefix = LCase$(keyName) & "="

    For i = LBound(parts) To UBound(parts)
        If LCase$(Left$(parts(i), Len(prefix))) = prefix Then
            GetParamValue = UrlDecode(Mid$(parts(i), Len(prefix) + 1))
            Exit Function
        End If
    Next i
End Function

Public Function UrlDecode(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim hexValue As String
    Dim byteCount As Long
    Dim buffer() As Byte

    If Len(value) = 0 Then Exit Function

    ReDim buffer(0 To 0)
    byteCount = 0
    i = 1

    Do While i <= Len(value)
        ch = Mid$(value, i, 1)

        Select Case ch
            Case "%"
                If i + 2 <= Len(value) Then
                    hexValue = Mid$(value, i + 1, 2)
                    ReDim Preserve buffer(0 To byteCount)
                    buffer(byteCount) = CByte(CLng("&H" & hexValue))
                    byteCount = byteCount + 1
                    i = i + 3
                Else
                    ReDim Preserve buffer(0 To byteCount)
                    buffer(byteCount) = Asc("%")
                    byteCount = byteCount + 1
                    i = i + 1
                End If
            Case "+"
                ReDim Preserve buffer(0 To byteCount)
                buffer(byteCount) = Asc(" ")
                byteCount = byteCount + 1
                i = i + 1
            Case Else
                ReDim Preserve buffer(0 To byteCount)
                buffer(byteCount) = Asc(ch)
                byteCount = byteCount + 1
                i = i + 1
        End Select
    Loop

    UrlDecode = Utf8BytesToString(buffer)
End Function

Private Function Utf8BytesToString(ByRef bytes() As Byte) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    Utf8BytesToString = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function

Public Sub PushResultToBrowser(ByVal requestId As String, ByVal resultText As String, ByVal statusText As String)
      Dim js As String

      On Error Resume Next

      frmWebHost.wbMain.Document.getElementById("status").innerText = statusText

      js = "window.hostBridge.resolveCommand('" & EscapeJs(requestId) & "','" & EscapeJs(resultText) & "');" & _
           "window.location.hash='';"

      frmWebHost.wbMain.Document.parentWindow.execScript js, "JavaScript"

  End Sub

Public Function EscapeJs(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, "'", "\'")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJs = value
End Function

Private Sub ShowResultMessage(ByVal resultText As String)
If InStr(1, resultText, """summaryMessage""", vbTextCompare) = 0 Then Exit Sub
      Dim success As Boolean
      Dim summaryMessage As String
      Dim affectedCount As Long
      Dim messageText As String

      success = InStr(1, resultText, """success"":true", vbTextCompare) > 0
      summaryMessage = ExtractJsonString(resultText, "summaryMessage")
      affectedCount = Val(ExtractJsonNumber(resultText, "affectedCount"))

      If summaryMessage = "" Then
          summaryMessage = resultText
      End If

      If affectedCount > 0 And success Then
          messageText = summaryMessage & vbCrLf & "处理数量：" & affectedCount
      Else
          messageText = summaryMessage
      End If

      If success Then
          MsgBox messageText, vbInformation, "执行完成"
      Else
          MsgBox messageText, vbExclamation, "执行失败"
      End If
  End Sub
  
  Private Function ExtractJsonString(ByVal jsonText As String, ByVal keyName As String) As String
      Dim searchText As String
      Dim startPos As Long
      Dim valueStart As Long
      Dim valueEnd As Long
      Dim rawValue As String

      searchText = """" & keyName & """:"""
      startPos = InStr(1, jsonText, searchText, vbTextCompare)
      If startPos = 0 Then Exit Function

      valueStart = startPos + Len(searchText)
      valueEnd = InStr(valueStart, jsonText, """")
      If valueEnd = 0 Then Exit Function

      rawValue = Mid$(jsonText, valueStart, valueEnd - valueStart)
      rawValue = Replace(rawValue, "\n", vbCrLf)
      rawValue = Replace(rawValue, "\\", "\")
      rawValue = Replace(rawValue, "\" & Chr$(34), Chr$(34))

      ExtractJsonString = rawValue
  End Function

Private Function ExtractJsonNumber(ByVal jsonText As String, ByVal keyName As String) As String
      Dim searchText As String
      Dim startPos As Long
      Dim valueStart As Long
      Dim valueEnd As Long
      Dim ch As String
      Dim i As Long

      searchText = """" & keyName & """:"
      startPos = InStr(1, jsonText, searchText, vbTextCompare)
      If startPos = 0 Then Exit Function

      valueStart = startPos + Len(searchText)
      valueEnd = valueStart

      For i = valueStart To Len(jsonText)
          ch = Mid$(jsonText, i, 1)
          If (ch >= "0" And ch <= "9") Or ch = "-" Then
              valueEnd = i
          Else
              Exit For
          End If
      Next i

      If valueEnd >= valueStart Then
          ExtractJsonNumber = Mid$(jsonText, valueStart, valueEnd - valueStart + 1)
      End If
  End Function
