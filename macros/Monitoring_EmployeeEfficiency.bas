Option Explicit

Sub Monitoring_Эффективность_Сотрудников()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Dim r As Long
    Dim c As Long
    Dim outRow As Long

    Dim currentDate As Variant
    Dim shiftType As String
    Dim dstCol As Long
    Dim sourceValue As Variant
    Dim currentValue As Variant

    On Error GoTo CleanFail

    Set wsSrc = Worksheets("Sheet 1")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Удаляем старый результат, чтобы всегда работать с актуальной выгрузкой
    On Error Resume Next
    Worksheets("Преобразованные").Delete
    On Error GoTo CleanFail

    Application.DisplayAlerts = True

    Set wsDst = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsDst.Name = "Преобразованные"

    ' Заголовки
    wsDst.Cells(1, 1).Value = "Команда"
    wsDst.Cells(1, 2).Value = "Стол специалиста"
    wsDst.Cells(1, 3).Value = "Имя специалиста"

    For c = 1 To 31
        wsDst.Cells(1, c + 3).Value = c
    Next c

    ' Оставляем исходную логику определения последней строки и колонки
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).Row
    lastCol = wsSrc.Cells(5, wsSrc.Columns.Count).End(xlToLeft).Column

    outRow = 2

    ' Основной цикл по строкам
    For r = 6 To lastRow

        If Trim(CStr(wsSrc.Cells(r, 3).Value)) <> "" Then

            If LCase(Trim(CStr(wsSrc.Cells(r, 3).Value))) <> "total" Then

                ' Основные данные оставляем в том же формате
                wsDst.Cells(outRow, 1).Value = wsSrc.Cells(r, 1).Value
                wsDst.Cells(outRow, 2).Value = wsSrc.Cells(r, 2).Value
                wsDst.Cells(outRow, 3).Value = wsSrc.Cells(r, 3).Value

                ' В этой выгрузке колонка D является первым днём месяца, поэтому начинаем с 4
                For c = 4 To lastCol

                    currentDate = wsSrc.Cells(4, c).MergeArea.Cells(1, 1).Value
                    shiftType = LCase(Trim(CStr(wsSrc.Cells(5, c).Value)))

                    If shiftType = "день" Or shiftType = "ночь" Then

                        If IsNumeric(currentDate) Then

                            If CDbl(currentDate) >= 1 And CDbl(currentDate) <= 31 Then

                                dstCol = CLng(currentDate) + 3
                                sourceValue = wsSrc.Cells(r, c).Value

                                If Not IsError(sourceValue) Then

                                    If sourceValue <> "" And IsNumeric(sourceValue) Then

                                        currentValue = wsDst.Cells(outRow, dstCol).Value

                                        ' Если за одну дату есть и день, и ночь, складываем их,
                                        ' а не перезаписываем предыдущее значение
                                        If currentValue = "" Then
                                            wsDst.Cells(outRow, dstCol).Value = CDbl(sourceValue)
                                        ElseIf IsNumeric(currentValue) Then
                                            wsDst.Cells(outRow, dstCol).Value = _
                                                CDbl(currentValue) + CDbl(sourceValue)
                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                Next c

                outRow = outRow + 1

            End If

        End If

    Next r

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
