Attribute VB_Name = "WebConfig"
Option Explicit

Public Function GetInitData() As String
    GetInitData = "{""templateSheet"":"""",""dataSheet"":"""",""templateRange"":"""",""pageField"":"""",""perRowCount"":3,""generateCount"":1,""sheets"":" & GetSheetListJson() & "}"
End Function

Public Function GetSheetListJson() As String
    Dim i As Long
    Dim parts() As String

    ReDim parts(1 To ActiveWorkbook.Worksheets.Count)
    For i = 1 To ActiveWorkbook.Worksheets.Count
        parts(i) = QuoteJson(CStr(ActiveWorkbook.Worksheets(i).Name))
    Next i

    GetSheetListJson = "[" & Join(parts, ",") & "]"
End Function

Public Function GetSheetFields(ByVal sheetName As String, ByVal rangeText As String) As String
    On Error GoTo HandleError

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim placeholders As Collection
    Dim parts() As String
    Dim i As Long

    Set ws = TryGetWorksheet(sheetName)
    If ws Is Nothing Then
        GetSheetFields = "[]"
        Exit Function
    End If

    On Error Resume Next
    Set targetRange = ws.Range(NormalizeAddress(rangeText))
    On Error GoTo HandleError
    If targetRange Is Nothing Then
        GetSheetFields = "[]"
        Exit Function
    End If

    Set placeholders = ExtractPlaceholders(targetRange)
    If placeholders.Count = 0 Then
        GetSheetFields = "[]"
        Exit Function
    End If

    ReDim parts(1 To placeholders.Count)
    For i = 1 To placeholders.Count
        parts(i) = QuoteJson(CStr(placeholders(i)))
    Next i

    GetSheetFields = "[" & Join(parts, ",") & "]"
    Exit Function

HandleError:
    GetSheetFields = "[]"
End Function

Public Function ValidateRange(ByVal sheetName As String, ByVal rangeText As String) As String
    On Error GoTo InvalidRange

    Dim ws As Worksheet
    Dim targetRange As Range

    Set ws = ActiveWorkbook.Worksheets(sheetName)
    Set targetRange = ws.Range(rangeText)

    ValidateRange = "{""valid"":true,""address"":" & QuoteJson(targetRange.Address(False, False)) & "}"
    Exit Function

InvalidRange:
    ValidateRange = "{""valid"":false,""message"":" & QuoteJson("模板区域无效") & "}"
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

Private Function NormalizeAddress(ByVal addressText As String) As String
    Dim markPos As Long

    markPos = InStr(addressText, "!")
    If markPos > 0 Then
        NormalizeAddress = Mid$(addressText, markPos + 1)
    Else
        NormalizeAddress = addressText
    End If
End Function

Private Function ExtractPlaceholders(ByVal targetRange As Range) As Collection
    Dim results As New Collection
    Dim cell As Range
    Dim textValue As String
    Dim openPos As Long
    Dim closePos As Long
    Dim startPos As Long
    Dim placeholder As String

    For Each cell In targetRange.Cells
        If Not IsEmpty(cell.value) Then
            textValue = CStr(cell.value)
            startPos = 1
            Do
                openPos = InStr(startPos, textValue, "<<")
                If openPos = 0 Then Exit Do
                closePos = InStr(openPos + 2, textValue, ">>")
                If closePos = 0 Then Exit Do

                placeholder = Trim$(Mid$(textValue, openPos + 2, closePos - openPos - 2))
                If Len(placeholder) > 0 Then
                    AddUniqueItem results, placeholder
                End If
                startPos = closePos + 2
            Loop
        End If
    Next cell

    Set ExtractPlaceholders = results
End Function

Private Sub AddUniqueItem(ByRef items As Collection, ByVal value As String)
    Dim existing As Variant

    For Each existing In items
        If StrComp(CStr(existing), value, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next existing

    items.Add value
End Sub

Private Function QuoteJson(ByVal value As String) As String
      value = Replace(value, "\", "\\")
      value = Replace(value, Chr$(34), "\" & Chr$(34))
      value = Replace(value, vbCrLf, "\n")
      value = Replace(value, vbCr, "\n")
      value = Replace(value, vbLf, "\n")
      QuoteJson = Chr$(34) & value & Chr$(34)
  End Function
