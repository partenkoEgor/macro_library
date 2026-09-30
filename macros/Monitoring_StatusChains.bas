Option Explicit

Private Const RESULT_SHEET As String = "Результат"

Public Sub Monitoring_Цепочек_Статусов()

    Dim wb As Workbook
    Dim src As Worksheet
    Dim res As Worksheet

    Dim statusHeader As Range
    Dim timeHeader As Range
    Dim statusCell As Range

    Dim sourceStatuses As Variant
    Dim resultStatuses As Variant

    Dim durationValue As Double
    Dim totalValue As Double

    Dim i As Long
    Dim outputRow As Long
    Dim errorsText As String
    Dim oldScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set wb = ActiveWorkbook
    Set src = GetSourceSheet(wb)

    If src Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "Не найден лист с исходными данными."
    End If

    Set statusHeader = FindExactText(src, "Начальный статус")
    Set timeHeader = FindExactText(src, "Медианное время перехода")

    If statusHeader Is Nothing Then
        Err.Raise vbObjectError + 2, , _
            "Не найдена колонка «Начальный статус»."
    End If

    If timeHeader Is Nothing Then
        Err.Raise vbObjectError + 3, , _
            "Не найдена колонка «Медианное время перехода»."
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

    Set res = GetOrCreateResultSheet(wb)
    res.Cells.Clear

    With res

        .Range("A1").Value = "Показатель"
        .Range("B1").Value = "Готовое значение"

        .Range("A2").Value = "Общее"
        .Range("A3").Value = "% изменения"
        .Range("A4").Value = ""

        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").Interior.Color = RGB(31, 78, 121)
        .Range("A1:B1").Font.Color = RGB(255, 255, 255)

        .Range("B2").NumberFormat = "[h]:mm:ss"
        .Range("B3").NumberFormat = "0.0%"
        .Range("B5:B17").NumberFormat = "[h]:mm:ss"

        .Columns("A").ColumnWidth = 48
        .Columns("B").ColumnWidth = 20

    End With

    If TryFindTotalDuration( _
        src, _
        statusHeader.Column, _
        timeHeader.Column, _
        statusHeader.Row + 1, _
        totalValue _
    ) Then

        res.Range("B2").Value2 = totalValue
        res.Range("B2").NumberFormat = "[h]:mm:ss"

    Else

        errorsText = errorsText & vbCrLf & _
            "Не удалось определить значение «Общее»."

        res.Range("B2").Interior.Color = RGB(255, 199, 206)

    End If

    For i = LBound(sourceStatuses) To UBound(sourceStatuses)

        outputRow = 5 + i

        res.Cells(outputRow, 1).Value = resultStatuses(i)

        Set statusCell = FindStatusInColumn( _
            src, _
            statusHeader.Column, _
            statusHeader.Row + 1, _
            CStr(sourceStatuses(i)) _
        )

        If statusCell Is Nothing Then

            errorsText = errorsText & vbCrLf & _
                "Не найден статус: " & CStr(sourceStatuses(i))

            res.Cells(outputRow, 2).Interior.Color = _
                RGB(255, 199, 206)

        ElseIf TryReadDuration( _
            src.Cells(statusCell.Row, timeHeader.Column), _
            durationValue _
        ) Then

            res.Cells(outputRow, 2).Value2 = durationValue
            res.Cells(outputRow, 2).NumberFormat = "[h]:mm:ss"

        Else

            errorsText = errorsText & vbCrLf & _
                "Некорректное время у статуса: " & _
                CStr(sourceStatuses(i))

            res.Cells(outputRow, 2).Interior.Color = _
                RGB(255, 199, 206)

        End If

    Next i

    With res.Range("A1:B17")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .VerticalAlignment = xlCenter
    End With

    res.Range("B2").NumberFormat = "[h]:mm:ss"
    res.Range("B5:B17").NumberFormat = "[h]:mm:ss"

    res.Activate
    res.Range("B2:B17").Select

    If Len(errorsText) = 0 Then

        res.Range("B2:B17").Copy

        Application.ScreenUpdating = oldScreenUpdating

        MsgBox _
            "Готово." & vbCrLf & vbCrLf & _
            "Диапазон B2:B17 скопирован." & vbCrLf & _
            "Вставляйте его начиная с ячейки «Общее»." & _
            vbCrLf & vbCrLf & _
            "Формат времени: [h]:mm:ss" & vbCrLf & _
            "Строка «% изменения» оставлена пустой.", _
            vbInformation, _
            "Медианное время"

    Else

        Application.ScreenUpdating = oldScreenUpdating

        MsgBox _
            "Результат сформирован, но найдены ошибки:" & _
            vbCrLf & errorsText & vbCrLf & vbCrLf & _
            "Проблемные ячейки выделены красным.", _
            vbExclamation, _
            "Проверка данных"

    End If

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = oldScreenUpdating

    MsgBox _
        "Не удалось выполнить макрос:" & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Ошибка"

End Sub

Private Function GetSourceSheet( _
    ByVal wb As Workbook _
) As Worksheet

    Dim ws As Worksheet

    If TypeName(wb.ActiveSheet) = "Worksheet" Then

        If wb.ActiveSheet.Name <> RESULT_SHEET Then
            Set GetSourceSheet = wb.ActiveSheet
            Exit Function
        End If

    End If

    For Each ws In wb.Worksheets

        If ws.Name <> RESULT_SHEET Then

            If Not FindExactText( _
                ws, _
                "Начальный статус" _
            ) Is Nothing Then

                Set GetSourceSheet = ws
                Exit Function

            End If

        End If

    Next ws

End Function

Private Function GetOrCreateResultSheet( _
    ByVal wb As Workbook _
) As Worksheet

    On Error Resume Next
    Set GetOrCreateResultSheet = wb.Worksheets(RESULT_SHEET)
    On Error GoTo 0

    If GetOrCreateResultSheet Is Nothing Then

        Set GetOrCreateResultSheet = _
            wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

        GetOrCreateResultSheet.Name = RESULT_SHEET

    End If

End Function

Private Function FindExactText( _
    ByVal ws As Worksheet, _
    ByVal wantedText As String _
) As Range

    Dim cell As Range
    Dim wantedNormalized As String

    wantedNormalized = NormalizeText(wantedText)

    For Each cell In ws.UsedRange.Cells

        If NormalizeText(CStr(cell.Value2)) = wantedNormalized Then
            Set FindExactText = cell
            Exit Function
        End If

    Next cell

End Function

Private Function FindStatusInColumn( _
    ByVal ws As Worksheet, _
    ByVal statusColumn As Long, _
    ByVal firstRow As Long, _
    ByVal wantedStatus As String _
) As Range

    Dim lastRow As Long
    Dim rowNumber As Long
    Dim wantedNormalized As String

    lastRow = ws.Cells(ws.Rows.Count, statusColumn).End(xlUp).Row
    wantedNormalized = NormalizeText(wantedStatus)

    For rowNumber = firstRow To lastRow

        If NormalizeText( _
            CStr(ws.Cells(rowNumber, statusColumn).Value2) _
        ) = wantedNormalized Then

            Set FindStatusInColumn = _
                ws.Cells(rowNumber, statusColumn)

            Exit Function

        End If

    Next rowNumber

End Function

Private Function TryReadDuration( _
    ByVal sourceCell As Range, _
    ByRef durationResult As Double _
) As Boolean

    Dim rawValue As Variant
    Dim numericValue As Double
    Dim textValue As String
    Dim parts As Variant

    Dim hoursValue As Double
    Dim minutesValue As Double
    Dim secondsValue As Double

    rawValue = sourceCell.Value2

    If IsError(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    If IsNumeric(rawValue) And VarType(rawValue) <> vbString Then

        numericValue = CDbl(rawValue)

        If numericValue < 0 Then
            Exit Function
        End If

        If IsDurationFormat(sourceCell.NumberFormat) Then

            durationResult = numericValue

        ElseIf numericValue < 1 Then

            durationResult = numericValue

        Else

            durationResult = numericValue / 86400#

        End If

        TryReadDuration = True
        Exit Function

    End If

    textValue = Trim$(CStr(rawValue))
    textValue = Replace(textValue, ChrW(160), "")
    textValue = Replace(textValue, " ", "")

    If Not IsValidDurationText(textValue) Then
        Exit Function
    End If

    parts = Split(textValue, ":")

    hoursValue = CDbl(parts(0))
    minutesValue = CDbl(parts(1))
    secondsValue = Val(Replace(parts(2), ",", "."))

    durationResult = _
        (hoursValue * 3600# + _
        minutesValue * 60# + _
        secondsValue) / 86400#

    TryReadDuration = True

End Function

Private Function TryFindTotalDuration( _
    ByVal ws As Worksheet, _
    ByVal statusColumn As Long, _
    ByVal timeColumn As Long, _
    ByVal firstRow As Long, _
    ByRef durationResult As Double _
) As Boolean

    Dim lastRow As Long
    Dim lastColumn As Long

    Dim totalRow As Long
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim searchRow As Long

    Dim currentText As String
    Dim numericValue As Double

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastColumn = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    For rowNumber = firstRow To lastRow

        currentText = NormalizeText( _
            CStr(ws.Cells(rowNumber, statusColumn).Value2) _
        )

        If currentText = "grand total" _
        Or currentText = "общий итог" _
        Or currentText = "итого" Then

            totalRow = rowNumber
            Exit For

        End If

    Next rowNumber

    If totalRow = 0 Then
        Exit Function
    End If

    If TryReadDuration( _
        ws.Cells(totalRow, timeColumn), _
        durationResult _
    ) Then

        TryFindTotalDuration = True
        Exit Function

    End If

    For searchRow = totalRow To _
        Application.Min(totalRow + 2, lastRow)

        For columnNumber = 1 To lastColumn

            currentText = NormalizeText( _
                CStr(ws.Cells(searchRow, columnNumber).Value2) _
            )

            If InStr(1, currentText, _
                "медианное время перехода", _
                vbTextCompare) > 0 _
            And InStr(1, currentText, _
                "сек", _
                vbTextCompare) > 0 Then

                If columnNumber < lastColumn Then

                    If IsNumeric( _
                        ws.Cells(searchRow, columnNumber + 1).Value2 _
                    ) Then

                        numericValue = CDbl( _
                            ws.Cells( _
                                searchRow, _
                                columnNumber + 1 _
                            ).Value2 _
                        )

                        If numericValue >= 0 Then

                            durationResult = numericValue / 86400#
                            TryFindTotalDuration = True
                            Exit Function

                        End If

                    End If

                End If

            End If

        Next columnNumber

    Next searchRow

    For searchRow = totalRow + 1 To _
        Application.Min(totalRow + 2, lastRow)

        For columnNumber = timeColumn To lastColumn

            If IsNumeric( _
                ws.Cells(searchRow, columnNumber).Value2 _
            ) Then

                numericValue = CDbl( _
                    ws.Cells(searchRow, columnNumber).Value2 _
                )

                If numericValue >= 0 Then

                    durationResult = numericValue / 86400#
                    TryFindTotalDuration = True
                    Exit Function

                End If

            End If

        Next columnNumber

    Next searchRow

End Function

Private Function IsValidDurationText( _
    ByVal durationText As String _
) As Boolean

    Dim expression As Object

    Set expression = CreateObject("VBScript.RegExp")

    With expression
        .Pattern = "^\d+:[0-5]\d:[0-5]\d([.,]\d+)?$"
        .Global = False
        .IgnoreCase = True
    End With

    IsValidDurationText = expression.Test(durationText)

End Function

Private Function IsDurationFormat( _
    ByVal formatText As String _
) As Boolean

    Dim normalizedFormat As String

    normalizedFormat = LCase$(formatText)
    normalizedFormat = Replace(normalizedFormat, " ", "")
    normalizedFormat = Replace(normalizedFormat, "\", "")

    IsDurationFormat = _
        InStr(1, normalizedFormat, "[h]:mm:ss", vbTextCompare) > 0 _
        Or InStr(1, normalizedFormat, "h:mm:ss", vbTextCompare) > 0

End Function

Private Function NormalizeText( _
    ByVal sourceText As String _
) As String

    Dim resultText As String

    resultText = LCase$(Trim$(sourceText))
    resultText = Replace(resultText, ChrW(160), " ")
    resultText = Replace(resultText, vbTab, " ")
    resultText = Replace(resultText, "ё", "е")

    Do While InStr(resultText, "  ") > 0
        resultText = Replace(resultText, "  ", " ")
    Loop

    NormalizeText = resultText

End Function