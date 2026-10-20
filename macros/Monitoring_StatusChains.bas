Option Explicit

Private Const RESULT_SHEET As String = "Результат"
Private Const COPY_BUTTON As String = "Monitoring_Copy_Status_Times"
Private Const COPY_RANGE As String = "B5:B17"

Public Sub Monitoring_Цепочек_Статусов()
    Dim wb As Workbook
    Dim src As Worksheet
    Dim res As Worksheet
    Dim sourceStatuses As Variant
    Dim resultStatuses As Variant
    Dim values() As Double
    Dim statusColumn As Long
    Dim responsibilityColumn As Long
    Dim timeColumn As Long
    Dim headerRow As Long
    Dim i As Long
    Dim problem As String
    Dim oldScreenUpdating As Boolean
    Dim button As Shape
    Dim anchor As Range
    Dim callerValue As Variant

    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo ErrorHandler

    ' Та же процедура обрабатывает кнопку. Отдельный макрос не нужен.
    callerValue = Application.Caller
    If VarType(callerValue) = vbString Then
        If CStr(callerValue) = COPY_BUTTON Then
            If TypeName(ActiveSheet) <> "Worksheet" Then
                Err.Raise vbObjectError + 1, , "Выберите лист «Результат»."
            End If
            Set res = ActiveSheet
            If res.Name <> RESULT_SHEET Then
                Err.Raise vbObjectError + 2, , "Кнопка должна находиться на листе «Результат»."
            End If
            For i = 5 To 17
                If IsError(res.Cells(i, 2).Value2) Then
                    Err.Raise vbObjectError + 3, , "В диапазоне есть ошибка. Повторите обработку выгрузки."
                End If
                If IsEmpty(res.Cells(i, 2).Value2) Then
                    Err.Raise vbObjectError + 4, , "В диапазоне есть пустое значение. Повторите обработку выгрузки."
                End If
                If Not IsNumeric(res.Cells(i, 2).Value2) Then
                    Err.Raise vbObjectError + 5, , "В диапазоне есть некорректное время."
                End If
                If CDbl(res.Cells(i, 2).Value2) < 0 Then
                    Err.Raise vbObjectError + 6, , "В диапазоне есть отрицательное время."
                End If
            Next i
            res.Range(COPY_RANGE).Copy
            Exit Sub
        End If
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Err.Raise vbObjectError + 7, , "Нет открытой книги с выгрузкой."
    End If

    sourceStatuses = Array( _
        "Получено (M)", _
        "Получено (Фрод) (M)", _
        "Одобрено (M)", _
        "Создать новую транзакцию (M)", _
        "Зачислено на другой счет (M)", _
        "Требуется доработка (M)", _
        "Файл более высокого качества и разрешения (M)", _
        "Не получено (M)", _
        "Требуются детали возврата (M)", _
        "Запрос выписки по депозиту (M)", _
        "Запрос скриншота депозита (M)", _
        "Возвращено на счет отправителя (M)", _
        "Файл не соответствует тикету (M)" _
    )
    resultStatuses = Array( _
        "Received (M) - 189", _
        "Received (Fraud) (M) - 236", _
        "Approved (M) - 177", _
        "Create new transaction (M) - 251", _
        "Credited to another account (M) - 179", _
        "Revision needed (M) - 238", _
        "Higher quality and resolution file (M) - 183", _
        "Not Received (M) - 187", _
        "Refund details needed (M) - 193", _
        "Request a statement for deposit (M) - 195", _
        "Request deposit screenshot (M) - 197", _
        "Returned to sender's account (M) - 201", _
        "File does not match ticket (M) - 255" _
    )

    Set src = GetSourceSheet(wb, sourceStatuses, headerRow, _
        responsibilityColumn, statusColumn, timeColumn, problem)
    If src Is Nothing Then
        Err.Raise vbObjectError + 8, , problem
    End If

    ' Проверка всех значений выполняется до изменения результата.
    If Not ReadL2Values(src, sourceStatuses, headerRow, _
        responsibilityColumn, statusColumn, timeColumn, values, problem) Then
        Err.Raise vbObjectError + 9, , problem
    End If

    Application.ScreenUpdating = False
    Set res = GetResultSheet(wb)
    res.Range("A1:B17").Clear
    With res
        .Range("A1").Value = "Показатель"
        .Range("B1").Value = "Готовое значение"
        .Range("A3").Value = "% изменения"
        .Range("B3").NumberFormat = "0.0%"
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").Interior.Color = RGB(31, 78, 121)
        .Range("A1:B1").Font.Color = RGB(255, 255, 255)
        .Columns("A").ColumnWidth = 48
        .Columns("B").ColumnWidth = 20
        .Range(COPY_RANGE).NumberFormat = "[h]:mm:ss"
        For i = LBound(sourceStatuses) To UBound(sourceStatuses)
            .Cells(5 + i, 1).Value = resultStatuses(i)
            .Cells(5 + i, 2).Value2 = values(i)
        Next i
        With .Range("A1:B17")
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .VerticalAlignment = xlCenter
        End With
    End With

    ' Обновляется только собственная кнопка. Другие объекты сохраняются.
    On Error Resume Next
    Set button = res.Shapes(COPY_BUTTON)
    On Error GoTo ErrorHandler
    If Not button Is Nothing Then button.Delete
    Set anchor = res.Range("B4")
    Set button = res.Shapes.AddShape(msoShapeRoundedRectangle, _
        anchor.Left + 2, anchor.Top + 2, _
        Application.Max(20, anchor.Width - 4), _
        Application.Max(14, anchor.Height - 4))
    With button
        .Name = COPY_BUTTON
        .AlternativeText = COPY_RANGE
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
            "'!Monitoring_Цепочек_Статусов"
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Fill.Solid
        .Line.ForeColor.RGB = RGB(31, 78, 121)
        .Shadow.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Копировать"
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Font.Name = "Arial"
        .TextFrame2.TextRange.Font.Size = 7
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Placement = xlMoveAndSize
    End With

    res.Activate
    res.Range("B5").Select
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Готово. Перенесено время по 13 статусам L2." & vbCrLf & _
        "Кнопка «Копировать» копирует B5:B17.", vbInformation, "Медианное время"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Не удалось выполнить макрос:" & vbCrLf & Err.Description, _
        vbExclamation, "Проверка данных"
End Sub

Private Function GetSourceSheet(ByVal wb As Workbook, ByVal statuses As Variant, _
    ByRef headerRow As Long, ByRef groupColumn As Long, _
    ByRef statusColumn As Long, ByRef timeColumn As Long, _
    ByRef problem As String) As Worksheet

    Dim ws As Worksheet
    Dim found As Worksheet
    Dim r As Long, g As Long, s As Long, t As Long
    Dim count As Long

    If TypeName(wb.ActiveSheet) = "Worksheet" Then
        Set ws = wb.ActiveSheet
        If ws.Name <> RESULT_SHEET Then
            If FindColumns(ws, statuses, r, g, s, t) Then
                Set GetSourceSheet = ws
                headerRow = r: groupColumn = g: statusColumn = s: timeColumn = t
                Exit Function
            End If
        End If
    End If
    For Each ws In wb.Worksheets
        If ws.Name <> RESULT_SHEET Then
            If FindColumns(ws, statuses, r, g, s, t) Then
                count = count + 1
                Set found = ws
                headerRow = r: groupColumn = g: statusColumn = s: timeColumn = t
            End If
        End If
    Next ws
    If count = 1 Then
        Set GetSourceSheet = found
    ElseIf count > 1 Then
        problem = "Найдено несколько подходящих листов. Выберите лист с нужной выгрузкой."
    Else
        problem = "Не найдена подходящая выгрузка." & vbCrLf & _
            "Нужны группа L2, начальные статусы и колонка «Медианное время перехода»." & vbCrLf & _
            "Проверьте настройки отчета: Ответственность, Начальный статус, Конечный статус."
    End If
End Function

Private Function FindColumns(ByVal ws As Worksheet, ByVal statuses As Variant, _
    ByRef headerRow As Long, ByRef groupColumn As Long, _
    ByRef statusColumn As Long, ByRef timeColumn As Long) As Boolean

    Dim cell As Range
    Dim c As Long, r As Long, i As Long
    Dim firstColumn As Long, lastColumn As Long, lastRow As Long
    Dim groupCount As Long, statusCount As Long, pairCount As Long
    Dim g As Long, s As Long
    Dim hasL2 As Boolean, hasStatus As Boolean
    Dim valueText As String

    firstColumn = ws.UsedRange.Column
    lastColumn = firstColumn + ws.UsedRange.Columns.Count - 1
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For Each cell In ws.UsedRange.Cells
        If CellText(cell) = NormalizeText("Медианное время перехода") Then
            groupCount = 0: statusCount = 0
            For c = firstColumn To lastColumn
                If c <> cell.Column Then
                    hasL2 = False: hasStatus = False
                    For r = cell.Row + 1 To lastRow
                        valueText = CellText(ws.Cells(r, c))
                        If valueText = "l2" Then hasL2 = True
                        If Not hasStatus Then
                            For i = LBound(statuses) To UBound(statuses)
                                If valueText = NormalizeText(CStr(statuses(i))) Then
                                    hasStatus = True
                                    Exit For
                                End If
                            Next i
                        End If
                    Next r
                    If hasL2 Then groupCount = groupCount + 1: g = c
                    If hasStatus Then statusCount = statusCount + 1: s = c
                End If
            Next c
            If groupCount = 1 And statusCount = 1 And g <> s Then
                pairCount = pairCount + 1
                headerRow = cell.Row
                groupColumn = g: statusColumn = s: timeColumn = cell.Column
            End If
        End If
    Next cell
    FindColumns = (pairCount = 1)
End Function

Private Function ReadL2Values(ByVal ws As Worksheet, ByVal statuses As Variant, _
    ByVal headerRow As Long, ByVal groupColumn As Long, _
    ByVal statusColumn As Long, ByVal timeColumn As Long, _
    ByRef values() As Double, ByRef problem As String) As Boolean

    Dim counts() As Long
    Dim r As Long, i As Long, lastRow As Long
    Dim currentGroup As String, groupText As String, statusText As String
    Dim duration As Double

    ReDim values(LBound(statuses) To UBound(statuses))
    ReDim counts(LBound(statuses) To UBound(statuses))
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = headerRow + 1 To lastRow
        ' Tableau оставляет название группы только в первой строке блока.
        If IsError(ws.Cells(r, groupColumn).Value2) Then
            currentGroup = ""
        Else
            groupText = CellText(ws.Cells(r, groupColumn))
            If Len(groupText) > 0 Then currentGroup = groupText
        End If
        If currentGroup = "l2" Then
            statusText = CellText(ws.Cells(r, statusColumn))
            For i = LBound(statuses) To UBound(statuses)
                If statusText = NormalizeText(CStr(statuses(i))) Then
                    counts(i) = counts(i) + 1
                    If TryReadDuration(ws.Cells(r, timeColumn), duration) Then
                        values(i) = duration
                    Else
                        problem = problem & vbCrLf & "Некорректное время: " & CStr(statuses(i))
                    End If
                    Exit For
                End If
            Next i
        End If
    Next r
    For i = LBound(statuses) To UBound(statuses)
        If counts(i) = 0 Then
            problem = problem & vbCrLf & "Не найден статус L2: " & CStr(statuses(i))
        ElseIf counts(i) > 1 Then
            problem = problem & vbCrLf & "Статус L2 повторяется: " & CStr(statuses(i)) & _
                ". Нужна одна строка с готовым временем."
        End If
    Next i
    ReadL2Values = (Len(problem) = 0)
End Function

Private Function TryReadDuration(ByVal cell As Range, ByRef result As Double) As Boolean
    Dim raw As Variant
    Dim textValue As String, fmt As String
    Dim parts As Variant
    Dim expression As Object
    Dim number As Double

    On Error GoTo InvalidValue
    raw = cell.Value2
    If IsError(raw) Or IsEmpty(raw) Then Exit Function
    If VarType(raw) <> vbString And IsNumeric(raw) Then
        number = CDbl(raw)
        If number < 0 Then Exit Function
        fmt = LCase$(CStr(cell.NumberFormat))
        If InStr(fmt, "h") > 0 And InStr(fmt, "s") > 0 Then
            result = number
        Else
            ' Числовые значения без формата времени читаются как секунды.
            result = number / 86400#
        End If
        TryReadDuration = True
        Exit Function
    End If
    textValue = Trim$(CStr(raw))
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, ChrW(8239), "")
    textValue = Replace(textValue, " ", "")
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = "^\d+:[0-5]\d:[0-5]\d([.,]\d+)?$"
    If Not expression.Test(textValue) Then Exit Function
    parts = Split(textValue, ":")
    result = (CDbl(parts(0)) * 3600# + CDbl(parts(1)) * 60# + _
        Val(Replace(parts(2), ",", "."))) / 86400#
    TryReadDuration = True
InvalidValue:
End Function

Private Function GetResultSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetResultSheet = wb.Worksheets(RESULT_SHEET)
    On Error GoTo 0
    If GetResultSheet Is Nothing Then
        Set GetResultSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetResultSheet.Name = RESULT_SHEET
    End If
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value2) Then Exit Function
    CellText = NormalizeText(CStr(cell.Value2))
End Function

Private Function NormalizeText(ByVal text As String) As String
    text = LCase$(text)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, ChrW(8239), " ")
    text = Replace(text, vbTab, " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, "ё", "е")
    text = Replace(text, "(м)", "(m)")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    NormalizeText = Trim$(text)
End Function

