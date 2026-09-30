Sub Monitoring_Переключения_Сотрудников()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Dim r As Long
    Dim c As Long
    Dim dayNumber As Long
    Dim outRow As Long

    Dim currentDate As Variant
    Dim shiftType As String

    Dim employeeName As String
    Dim employeeKey As String
    Dim mapKey As String

    Dim sourceValue As Variant
    Dim totalValue As Double

    Dim dataMap As Object

    Set wsSrc = Worksheets("Sheet 1")
    Set dataMap = CreateObject("Scripting.Dictionary")

    '--------------------------------
    ' Последние строки/колонки
    ' Оставлено как в исходном макросе
    '--------------------------------
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).Row
    lastCol = wsSrc.Cells(5, wsSrc.Columns.Count).End(xlToLeft).Column

    '--------------------------------
    ' Первый проход
    ' Собираем общую сумму сотрудника
    ' за каждый день со всех столов.
    '
    ' C = 1 число, день
    ' D = 1 число, ночь
    ' День + ночь суммируются.
    '--------------------------------
    For r = 6 To lastRow

        employeeName = Trim(CStr(wsSrc.Cells(r, 2).Value))

        If employeeName <> "" Then

            If LCase(employeeName) <> "total" Then

                employeeKey = LCase(Trim(employeeName))

                If employeeKey <> "" Then

                    For c = 3 To lastCol

                        currentDate = wsSrc.Cells(4, c).MergeArea.Cells(1, 1).Value
                        shiftType = LCase(Trim(CStr(wsSrc.Cells(5, c).Value)))

                        If shiftType = "день" Or shiftType = "ночь" Then

                            If IsNumeric(currentDate) Then

                                dayNumber = CLng(currentDate)

                                If dayNumber >= 1 And dayNumber <= 31 Then

                                    sourceValue = wsSrc.Cells(r, c).Value

                                    If Len(Trim(CStr(sourceValue))) > 0 Then

                                        If IsNumeric(sourceValue) Then

                                            mapKey = employeeKey & "|" & CStr(dayNumber)

                                            If dataMap.Exists(mapKey) Then
                                                dataMap(mapKey) = CDbl(dataMap(mapKey)) + CDbl(sourceValue)
                                            Else
                                                dataMap.Add mapKey, CDbl(sourceValue)
                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    Next c

                End If

            End If

        End If

    Next r

    '--------------------------------
    ' Создание нового листа
    '--------------------------------
    Application.DisplayAlerts = False

    On Error Resume Next
    Worksheets("Преобразованные данные").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Set wsDst = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsDst.Name = "Преобразованные данные"

    '--------------------------------
    ' Заголовки
    '--------------------------------
    wsDst.Cells(1, 1).Value = "Стол специалиста (переключения)"
    wsDst.Cells(1, 2).Value = "Имя специалиста (переключения)"

    For c = 1 To 31
        wsDst.Cells(1, c + 2).Value = c
    Next c

    outRow = 2

    '--------------------------------
    ' Второй проход
    ' Структуру строк сохраняем как раньше,
    ' но для одинакового сотрудника на разных
    ' столах выводим одну и ту же общую сумму.
    '
    ' Это не даёт старому fillSwitchData()
    ' перезаписать сумму данными только
    ' с последнего стола.
    '--------------------------------
    For r = 6 To lastRow

        employeeName = Trim(CStr(wsSrc.Cells(r, 2).Value))

        If employeeName <> "" Then

            If LCase(employeeName) <> "total" Then

                wsDst.Cells(outRow, 1).Value = wsSrc.Cells(r, 1).Value
                wsDst.Cells(outRow, 2).Value = wsSrc.Cells(r, 2).Value

                employeeKey = LCase(Trim(employeeName))

                For dayNumber = 1 To 31

                    mapKey = employeeKey & "|" & CStr(dayNumber)

                    If dataMap.Exists(mapKey) Then

                        totalValue = CDbl(dataMap(mapKey))

                        ' Нулевой итог оставляем пустым.
                        ' Тогда старый скрипт очистит
                        ' прежнее значение сотрудника
                        ' за эту дату при синхронизации.
                        If totalValue <> 0 Then
                            wsDst.Cells(outRow, dayNumber + 2).Value = totalValue
                        Else
                            wsDst.Cells(outRow, dayNumber + 2).ClearContents
                        End If

                    Else
                        wsDst.Cells(outRow, dayNumber + 2).ClearContents
                    End If

                Next dayNumber

                outRow = outRow + 1

            End If

        End If

    Next r

    MsgBox "Преобразование завершено! Данные записаны на лист '" & wsDst.Name & "'."

End Sub