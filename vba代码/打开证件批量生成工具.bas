Attribute VB_Name = "打开证件批量生成工具"
Option Explicit


Public Function 查找列号_指定工作表(目标表 As Worksheet, 查找内容 As String, Optional 搜索行 As Long = 1) As Long
    On Error GoTo ErrH
    Dim 查找结果 As Range
    Set 查找结果 = 目标表.Rows(搜索行).Find(What:=查找内容, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not 查找结果 Is Nothing Then
        查找列号_指定工作表 = 查找结果.Column
    Else
        Set 查找结果 = 目标表.Rows(搜索行).Find(What:=查找内容, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not 查找结果 Is Nothing Then
            查找列号_指定工作表 = 查找结果.Column
        Else
            查找列号_指定工作表 = 0
        End If
    End If
    Exit Function
ErrH:
    查找列号_指定工作表 = -1
End Function

Public Sub 执行批量生成模板(目标表 As Worksheet, 源区域 As Range, 复制份数 As Long, 每行个数 As Long)
    Dim i As Long, 实际行 As Long, 实际列 As Long
    Dim 行偏移 As Long, 列偏移 As Long
    Dim 目标区域 As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo CleanUp
    Application.DisplayStatusBar = True
    目标表.Activate
    源区域.Copy
    Dim 开始时间 As Double
    开始时间 = Timer
    Dim c As Long
    Dim 列索引 As Long
    For 列索引 = 1 To 每行个数 - 1
        For c = 1 To 源区域.Columns.Count
            目标表.Columns(c + 列索引 * 源区域.Columns.Count).ColumnWidth = 目标表.Columns(c).ColumnWidth
        Next c
    Next 列索引
    Dim r As Long
    Dim 行索引 As Long
    Dim 总行组数 As Long
    总行组数 = Int((复制份数 + 每行个数 - 1) / 每行个数)
    For 行索引 = 1 To 总行组数
        For r = 1 To 源区域.Rows.Count
            目标表.Rows(r + 行索引 * 源区域.Rows.Count).RowHeight = 目标表.Rows(r).RowHeight
        Next r
    Next 行索引
    For i = 每行个数 + 1 To 复制份数 + 每行个数
        Dim 调整后索引 As Long
        调整后索引 = i - 1
        实际行 = Int(调整后索引 / 每行个数)
        实际列 = 调整后索引 Mod 每行个数
        行偏移 = 实际行 * 源区域.Rows.Count
        列偏移 = 实际列 * 源区域.Columns.Count
        Set 目标区域 = 源区域.Offset(行偏移, 列偏移)
        目标区域.PasteSpecial xlPasteAll
        If (i Mod 50) = 0 Or i = 复制份数 + 3 Then
            更新进度 "第1步：生成", i - 3, 复制份数, 开始时间
        End If
    Next i
    Application.CutCopyMode = False
CleanUp:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub 执行批量写入数据(数据源表 As Worksheet, 目标表 As Worksheet, 源区域 As Range, 复制份数 As Long, 每行个数 As Long, Optional 分页字段 As String = "")
    Dim 占位符列表 As String, 占位符数组() As String
    Dim i As Long, 实际行 As Long, 实际列 As Long
    Dim 行偏移 As Long, 列偏移 As Long
    Dim 目标区域 As Range
    Dim n As Integer
    Dim 上一个分页值 As String, 当前分页值 As String
    Dim 跳过计数 As Long
    Dim 调整后索引 As Long
    Dim 分页列号 As Long

    ' 第一步：提取模板中的占位符
    占位符列表 = 提取模板占位符(源区域)
    If 占位符列表 <> "" Then
        占位符数组 = Split(占位符列表, "|")
    End If

    ' 第二步：查找分页字段列号（如果指定了分页字段）
    分页列号 = 0
    If 分页字段 <> "" Then
        分页列号 = 查找列号_指定工作表(数据源表, 分页字段)
        If 分页列号 = 0 Then
            MsgBox "分页字段 '" & 分页字段 & "' 在数据源表中未找到，将不分页", vbInformation
            分页字段 = ""
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlInterrupt
    On Error GoTo CleanUp
    Application.DisplayStatusBar = True
    目标表.Activate
    Dim 开始时间 As Double
    开始时间 = Timer
    上一个分页值 = ""
    跳过计数 = 0
    调整后索引 = 每行个数
    For i = 每行个数 + 1 To 复制份数 + 每行个数

        ' 读取当前分页字段值（如果指定了分页字段）
        当前分页值 = ""
        If 分页列号 > 0 Then
            当前分页值 = 数据源表.Cells(i - 每行个数, 分页列号).Text
        End If

        ' 检测分页字段变化
        If 分页字段 <> "" And i > 每行个数 + 1 And 当前分页值 <> "" And 上一个分页值 <> "" And 当前分页值 <> 上一个分页值 Then
            ' 计算当前在第几列
            Dim 当前列位置 As Long
            当前列位置 = 调整后索引 Mod 每行个数

            ' 如果不在第0列，需要跳过剩余列
            If 当前列位置 <> 0 Then
                Dim 需要跳过的列数 As Long
                需要跳过的列数 = 每行个数 - 当前列位置

                ' 清空被跳过位置的占位符
                Dim skip_idx As Long
                For skip_idx = 1 To 需要跳过的列数
                    Dim skip_调整后索引 As Long
                    skip_调整后索引 = 调整后索引 + skip_idx - 1
                    Dim skip_实际行 As Long, skip_实际列 As Long
                    skip_实际行 = Int(skip_调整后索引 / 每行个数)
                    skip_实际列 = skip_调整后索引 Mod 每行个数
                    Dim skip_行偏移 As Long, skip_列偏移 As Long
                    skip_行偏移 = skip_实际行 * 源区域.Rows.Count
                    skip_列偏移 = skip_实际列 * 源区域.Columns.Count
                    Dim skip_目标区域 As Range
                    Set skip_目标区域 = 源区域.Offset(skip_行偏移, skip_列偏移)

                    ' 清空所有占位符
                    清空模板占位符 skip_目标区域, 占位符列表
                Next skip_idx

                ' 调整索引，跳到下一行开始
                调整后索引 = 调整后索引 + 需要跳过的列数
                跳过计数 = 跳过计数 + 需要跳过的列数
            End If

            ' 在新分页值开始前插入水平分页符
            Dim 分页行号 As Long
            分页行号 = Int(调整后索引 / 每行个数) * 源区域.Rows.Count + 1
            目标表.HPageBreaks.Add Before:=目标表.Rows(分页行号)
        End If

        ' 重新计算实际行列
        实际行 = Int(调整后索引 / 每行个数)
        实际列 = 调整后索引 Mod 每行个数
        行偏移 = 实际行 * 源区域.Rows.Count
        列偏移 = 实际列 * 源区域.Columns.Count
        Set 目标区域 = 源区域.Offset(行偏移, 列偏移)
        If (i Mod 50) = 0 Or i = 复制份数 + 3 Then
            更新进度 "第2步：写入", i - 3, 复制份数, 开始时间
        End If

        ' 第三步：动态替换占位符
        If 占位符列表 <> "" Then
            Dim j As Long
            For j = 0 To UBound(占位符数组)
                Dim 占位符名 As String
                占位符名 = 占位符数组(j)
                If 占位符名 <> "" Then
                    Dim 列号 As Long
                    列号 = 查找列号_指定工作表(数据源表, 占位符名)
                    If 列号 > 0 Then
                        目标区域.Replace What:="<<" & 占位符名 & ">>", Replacement:=数据源表.Cells(i - 每行个数, 列号).Text, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
                    End If
                End If
            Next j
        End If

        ' 更新上一个分页值
        上一个分页值 = 当前分页值

        ' 索引递增
        调整后索引 = 调整后索引 + 1
    Next i
CleanUp:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' 清空模板中的所有占位符
Private Sub 清空模板占位符(目标区域 As Range, 占位符列表 As String)
    If 占位符列表 = "" Then Exit Sub
    Dim 占位符数组() As String
    占位符数组 = Split(占位符列表, "|")
    Dim i As Long
    For i = 0 To UBound(占位符数组)
        If 占位符数组(i) <> "" Then
            目标区域.Replace What:="<<" & 占位符数组(i) & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        End If
    Next i
End Sub

Public Sub 更新进度(步骤名 As String, 当前 As Long, 总数 As Long, 开始时间 As Double)
      Dim pct As String
      Dim percentValue As Long

      If 总数 <= 0 Then
          pct = "0%"
          percentValue = 0
      Else
          pct = Format(当前 / 总数, "0%")
          percentValue = CLng(当前 * 100 / 总数)
          If percentValue < 0 Then percentValue = 0
          If percentValue > 100 Then percentValue = 100
      End If

      Application.StatusBar = 步骤名 & " " & 当前 & "/" & 总数 & " (" & pct & ")  耗时 " & Format(Timer - 开始时间, "0.0") & "s"

      On Error Resume Next
      Frm证件批量生成工具.更新进度UI 步骤名, 当前, 总数
      Frm证件批量生成工具.Repaint

      If InStr(步骤名, "第1步") > 0 Then
          WebProgress.ReportProgress "generate", 当前, 总数, percentValue, 步骤名
      ElseIf InStr(步骤名, "第2步") > 0 Then
          WebProgress.ReportProgress "write", 当前, 总数, percentValue, 步骤名
      End If

      DoEvents
  End Sub

Private Sub 清空准考证占位符(目标区域 As Range)
    ' 清空基本信息占位符
    目标区域.Replace What:="<<姓名>>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    目标区域.Replace What:="<<考号>>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    目标区域.Replace What:="<<班级>>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    目标区域.Replace What:="<<学号>>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    ' 清空科目相关占位符（1-8）
    Dim n As Integer
    For n = 1 To 8
        目标区域.Replace What:="<<科目" & n & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        目标区域.Replace What:="<<时间" & n & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        目标区域.Replace What:="<<考场" & n & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        目标区域.Replace What:="<<考场号" & n & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        目标区域.Replace What:="<<座位号" & n & ">>", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    Next n
End Sub

' 提取模板中的所有占位符名称（不含<>符号）
Public Function 提取模板占位符(模板区域 As Range) As String
    Dim cell As Range
    Dim result As String
    result = ""
    For Each cell In 模板区域
        If Not IsEmpty(cell.value) Then
            Dim txt As String
            txt = CStr(cell.value)
            Dim startPos As Long
            startPos = 1
            Do
                Dim openPos As Long, closePos As Long
                openPos = InStr(startPos, txt, "<<")
                If openPos = 0 Then Exit Do
                closePos = InStr(openPos, txt, ">>")
                If closePos = 0 Then Exit Do
                Dim placeholder As String
                placeholder = Mid(txt, openPos + 2, closePos - openPos - 2)
                If InStr(result, "|" & placeholder & "|") = 0 Then
                    If result = "" Then
                        result = placeholder
                    Else
                        result = result & "|" & placeholder
                    End If
                End If
                startPos = closePos + 1
            Loop
        End If
    Next cell
    提取模板占位符 = result
End Function
