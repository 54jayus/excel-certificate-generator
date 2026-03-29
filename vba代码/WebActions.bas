Attribute VB_Name = "WebActions"
Option Explicit

Private Function NormalizeAddress(ByVal addressText As String) As String
    Dim markPos As Long
    markPos = InStr(addressText, "!")
    If markPos > 0 Then
        NormalizeAddress = Mid$(addressText, markPos + 1)
    Else
        NormalizeAddress = addressText
    End If
End Function

Private Function ParsePositiveLong(ByVal textValue As String) As Long
    If Len(Trim$(textValue)) = 0 Then Exit Function
    If Not IsNumeric(textValue) Then Exit Function

    Dim valueNumber As Double
    valueNumber = Val(textValue)
    If valueNumber < 1 Or valueNumber > 100000 Then Exit Function

    ParsePositiveLong = CLng(Fix(valueNumber))
End Function

Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set TryGetWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function TryGetRange(ByVal ws As Worksheet, ByVal rangeText As String) As Range
    On Error Resume Next
    Set TryGetRange = ws.Range(NormalizeAddress(rangeText))
    On Error GoTo 0
End Function

Private Function JsonSuccess(ByVal messageText As String, ByVal affectedCount As Long) As String
    JsonSuccess = "{""success"":true,""summaryMessage"":""" & messageText & """,""affectedCount"":" & CStr(affectedCount) & "}"
End Function

Private Function JsonError(ByVal messageText As String) As String
    JsonError = "{""success"":false,""summaryMessage"":""" & messageText & """,""affectedCount"":0}"
End Function

Private Sub RunGenerateTemplate(ByVal targetSheet As Worksheet, ByVal sourceRange As Range, ByVal generateCount As Long, ByVal perRowCount As Long)
    Call 执行批量生成模板(targetSheet, sourceRange, generateCount, perRowCount)
End Sub

Private Sub RunWriteData(ByVal dataSheet As Worksheet, ByVal targetSheet As Worksheet, ByVal sourceRange As Range, ByVal generateCount As Long, ByVal perRowCount As Long, ByVal pageFieldText As String)
    Call 执行批量写入数据(dataSheet, targetSheet, sourceRange, generateCount, perRowCount, pageFieldText)
End Sub

Public Function GenerateTemplate(ByVal templateSheetName As String, ByVal templateRangeText As String, ByVal generateCountText As String, ByVal perRowCountText As String) As String
    On Error GoTo HandleError

    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim generateCount As Long
    Dim perRowCount As Long

    Set targetSheet = TryGetWorksheet(templateSheetName)
    generateCount = ParsePositiveLong(generateCountText)
    perRowCount = ParsePositiveLong(perRowCountText)

    If targetSheet Is Nothing Then
        GenerateTemplate = JsonError("不存在模板sheet: " & templateSheetName)
        Exit Function
    End If

    Set sourceRange = TryGetRange(targetSheet, templateRangeText)
    If sourceRange Is Nothing Then
        GenerateTemplate = JsonError("无法处理的模板区域: " & templateRangeText)
        Exit Function
    End If

    If generateCount = 0 Then
        GenerateTemplate = JsonError("生成数量有误")
        Exit Function
    End If

    If perRowCount = 0 Then
        GenerateTemplate = JsonError("每行个数有误")
        Exit Function
    End If

    If perRowCount < 1 Or perRowCount > 10 Then
        GenerateTemplate = JsonError("Per-row count must be between 1 and 10")
        Exit Function
    End If

    WebProgress.ReportProgress "generate", 0, generateCount, 0, "开始批量生成模板"
    RunGenerateTemplate targetSheet, sourceRange, generateCount, perRowCount
    WebProgress.ReportProgress "generate", generateCount, generateCount, 100, "模板生成完成"

    GenerateTemplate = JsonSuccess("批量生成模板完成", generateCount)
    Exit Function

HandleError:
    GenerateTemplate = JsonError("批量生成模板失败: " & Err.Description)
End Function

Public Function WriteData(ByVal dataSheetName As String, ByVal templateSheetName As String, ByVal templateRangeText As String, ByVal generateCountText As String, ByVal perRowCountText As String, ByVal pageFieldText As String) As String
    On Error GoTo HandleError

    Dim dataSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim generateCount As Long
    Dim perRowCount As Long

    Set dataSheet = TryGetWorksheet(dataSheetName)
    Set targetSheet = TryGetWorksheet(templateSheetName)
    generateCount = ParsePositiveLong(generateCountText)
    perRowCount = ParsePositiveLong(perRowCountText)

    If dataSheet Is Nothing Then
        WriteData = JsonError("不存在数据sheet: " & dataSheetName)
        Exit Function
    End If

    If targetSheet Is Nothing Then
        WriteData = JsonError("不存在模板sheet: " & templateSheetName)
        Exit Function
    End If

    Set sourceRange = TryGetRange(targetSheet, templateRangeText)
    If sourceRange Is Nothing Then
        WriteData = JsonError("无法处理的模板区域: " & templateRangeText)
        Exit Function
    End If

    If generateCount = 0 Then
        WriteData = JsonError("生成数量有误")
        Exit Function
    End If

    If perRowCount = 0 Then
        WriteData = JsonError("每行个数有误")
        Exit Function
    End If

    If perRowCount < 1 Or perRowCount > 10 Then
        WriteData = JsonError("每个行数需在1-10之间")
        Exit Function
    End If

    WebProgress.ReportProgress "write", 0, generateCount, 0, "开始批量写入数据"""
    RunWriteData dataSheet, targetSheet, sourceRange, generateCount, perRowCount, pageFieldText
    WebProgress.ReportProgress "write", generateCount, generateCount, 100, "数据写入完成"

    WriteData = JsonSuccess("批量写入数据完成", generateCount)
    Exit Function

HandleError:
    WriteData = JsonError("批量写入数据失败：" & Err.Description)
End Function
