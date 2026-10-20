Option Explicit

Sub Monitoring_TTR_FTT()
    Dim wb As Workbook, wsSource As Worksheet, wsDest As Worksheet
    Dim countries As Variant, groups As Object, values As Collection
    Dim r As Long, c As Long, i As Long, j As Long
    Dim headerRow As Long, colCountry As Long, colValue As Long
    Dim candidateCountry As Long, candidateValue As Long, valueCount As Long
    Dim lastCol As Long, lastRow As Long, oldLastRow As Long
    Dim header As String, valueHeaderFound As String
    Dim country As String, currentCountry As String, key As String
    Dim cell As Range, countryCell As Range, lastCell As Range
    Dim rawValue As Variant, duration As Double, result() As Variant
    Dim medianValues() As Variant, copyAddress As String

    On Error GoTo ErrorHandler
    Set wb = ActiveWorkbook
    If wb.Worksheets.Count < 2 Then
        MsgBox "В книге нет второго листа с медианными значениями.", vbCritical, "Monitoring_TTR_FTT"
        Exit Sub
    End If
    Set wsSource = wb.Worksheets(2)
    If wsSource.Name = "Упорядоченный список" Then
        MsgBox "Второй лист является результатом, а не исходной выгрузкой.", vbCritical, "Monitoring_TTR_FTT"
        Exit Sub
    End If

    ' Обе колонки должны находиться в одной строке заголовков.
    For r = 1 To 15
        candidateCountry = 0: candidateValue = 0: valueCount = 0
        lastCol = wsSource.Cells(r, wsSource.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            If Not IsError(wsSource.Cells(r, c).Value2) Then
                header = Trim$(Replace(CStr(wsSource.Cells(r, c).Value2), ChrW(160), " "))
                Select Case LCase$(header)
                    Case "страна", "гранулярность численности тикетов"
                        candidateCountry = c
                    Case "first_touch_hms_med", "time_to_resolve_hms_med", "med_dt_prev_diff_emp_sec_hms"
                        candidateValue = c
                        valueCount = valueCount + 1
                        valueHeaderFound = header
                End Select
            End If
        Next c
        If candidateCountry > 0 And valueCount > 1 Then
            MsgBox "В строке заголовков найдено несколько метрик времени. Оставьте одну метрику: FTT или TTR.", vbCritical, "Monitoring_TTR_FTT"
            Exit Sub
        End If
        If candidateCountry > 0 And candidateValue > 0 Then
            headerRow = r: colCountry = candidateCountry: colValue = candidateValue
            Exit For
        End If
    Next r
    If headerRow = 0 Then
        MsgBox "На втором листе не найдены колонки страны и медианного времени в одной строке заголовков.", vbCritical, "Monitoring_TTR_FTT"
        Exit Sub
    End If

    ' Проверяем весь лист, включая рефералы последнего объединённого ГЕО.
    Set lastCell = wsSource.Cells.Find(What:="*", After:=wsSource.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If lastCell Is Nothing Then Exit Sub
    lastRow = lastCell.Row
    Set groups = CreateObject("Scripting.Dictionary")

    For r = headerRow + 1 To lastRow
        Set countryCell = wsSource.Cells(r, colCountry)
        If countryCell.MergeCells Then
            country = CStr(countryCell.MergeArea.Cells(1, 1).Value2)
            currentCountry = NormalizeCountry_TTR_FTT(country)
        ElseIf Not IsError(countryCell.Value2) Then
            country = Trim$(CStr(countryCell.Value2))
            If Len(country) > 0 Then currentCountry = NormalizeCountry_TTR_FTT(country)
        Else
            Err.Raise vbObjectError + 701, , "Ошибка в названии ГЕО: " & countryCell.Address(False, False)
        End If

        Set cell = wsSource.Cells(r, colValue)
        ' Объединённое время учитываем один раз, только в верхней строке.
        If cell.MergeCells Then
            If cell.Row <> cell.MergeArea.Row Then GoTo NextSourceRow
        End If
        rawValue = cell.Value2
        If IsError(rawValue) Then
            Err.Raise vbObjectError + 702, , "Ошибка в значении времени: " & cell.Address(False, False)
        End If
        If Len(Trim$(CStr(rawValue))) = 0 Then GoTo NextSourceRow
        If Len(currentCountry) = 0 Then
            Err.Raise vbObjectError + 703, , "Время без ГЕО: " & cell.Address(False, False)
        End If
        If Not TryDuration_TTR_FTT(rawValue, duration) Then
            Err.Raise vbObjectError + 704, , "Некорректное время в " & cell.Address(False, False) & ": " & CStr(rawValue)
        End If
        duration = PreciseDuration_TTR_FTT(wsSource, cell, valueHeaderFound, duration)
        If Not groups.Exists(currentCountry) Then
            Set values = New Collection
            groups.Add currentCountry, values
        End If
        Set values = groups(currentCountry)
        values.Add duration
NextSourceRow:
    Next r

    countries = Array( _
        "Азербайджан", "Алжир", "Афганистан", "Бахрейн", "Боливия", _
        "Гватемала", "Гондурас", "Джибути", "Египет", "Иордания", _
        "Ирак", "Иран", "Йемен", "Канада", "Катар", _
        "Кувейт", "Кыргызстан", "Ливан", "Ливия", "Мавритания", _
        "Марокко", "Никарагуа", "ОАЭ", "Оман", "Палестина", _
        "Панама", "Парагвай", "Саудовская Аравия", "Сирия", "Сомали", _
        "Тайвань", "Тунис", "Турция", "Южный Судан", "Судан", _
        "Ямайка" _
    )


    ReDim result(1 To UBound(countries) - LBound(countries) + 2, 1 To 2)
    result(1, 1) = "Страна": result(1, 2) = valueHeaderFound
    For i = LBound(countries) To UBound(countries)
        result(i - LBound(countries) + 2, 1) = countries(i)
        key = NormalizeCountry_TTR_FTT(CStr(countries(i)))
        If groups.Exists(key) Then
            Set values = groups(key)
            ReDim medianValues(1 To values.Count)
            For j = 1 To values.Count
                medianValues(j) = CDbl(values(j))
            Next j
            ' Каждый реферал имеет одинаковый вес; число тикетов не используется.
            result(i - LBound(countries) + 2, 2) = Application.WorksheetFunction.Median(medianValues)
        Else
            result(i - LBound(countries) + 2, 2) = ""
        End If
    Next i

    ' Исходные листы не меняем. Существующий результат обновляем без удаления листа.
    On Error Resume Next
    Set wsDest = wb.Worksheets("Упорядоченный список")
    On Error GoTo ErrorHandler
    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wsSource)
        wsDest.Name = "Упорядоченный список"
    End If
    oldLastRow = Application.Max(wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row, _
        wsDest.Cells(wsDest.Rows.Count, 2).End(xlUp).Row, UBound(result, 1))
    wsDest.Range("A1:B" & oldLastRow).ClearContents
    wsDest.Range("A1").Resize(UBound(result, 1), 2).Value = result
    wsDest.Columns("B").NumberFormat = "[h]:mm:ss"
    wsDest.Columns("A:B").AutoFit
    copyAddress = "B2:B" & UBound(result, 1)
    AddCopyButton_CMB wsDest, "btnMonitoringTtrFtt", wsDest.Range("B1"), copyAddress, RGB(31, 78, 121)
    MsgBox "Готово! Рассчитана медиана значений рефералов каждого ГЕО со второго листа. Метрика: " & valueHeaderFound, vbInformation, "Monitoring_TTR_FTT"
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка " & Err.Number & ":" & vbCrLf & Err.Description, vbCritical, "Monitoring_TTR_FTT"
End Sub

Private Function PreciseDuration_TTR_FTT(ByVal ws As Worksheet, ByVal timeCell As Range, _
    ByVal metric As String, ByVal fallback As Double) As Double
    Dim bottomRow As Long, r As Long, label As String, expected As String
    Dim rawSeconds As Variant
    PreciseDuration_TTR_FTT = fallback
    Select Case LCase$(metric)
        Case "first_touch_hms_med": expected = "медианное first touch time"
        Case "time_to_resolve_hms_med": expected = "медианное time to resolve"
        Case Else: Exit Function
    End Select
    If Not timeCell.MergeCells Then Exit Function
    bottomRow = timeCell.MergeArea.Row + timeCell.MergeArea.Rows.Count - 1
    For r = timeCell.Row To bottomRow
        If Not IsError(ws.Cells(r, timeCell.Column + 1).Value2) Then
            label = LCase$(Trim$(CStr(ws.Cells(r, timeCell.Column + 1).Value2)))
            If label = expected Then
                rawSeconds = ws.Cells(r, timeCell.Column + 2).Value2
                If Not IsError(rawSeconds) Then
                    If VarType(rawSeconds) <> vbBoolean And IsNumeric(rawSeconds) Then
                        If Len(Trim$(CStr(rawSeconds))) > 0 Then
                            If CDbl(rawSeconds) >= 0 Then PreciseDuration_TTR_FTT = CDbl(rawSeconds) / 86400#
                        End If
                    End If
                End If
                Exit Function
            End If
        End If
    Next r
End Function

Private Function TryDuration_TTR_FTT(ByVal rawValue As Variant, ByRef duration As Double) As Boolean
    Dim parts As Variant, text As String, secondsText As String
    Dim hours As Double, minutes As Double, seconds As Double
    On Error GoTo InvalidDuration
    If IsError(rawValue) Or IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    If VarType(rawValue) = vbBoolean Then Exit Function
    If VarType(rawValue) <> vbString And IsNumeric(rawValue) Then
        duration = CDbl(rawValue)
        If duration < 0 Then Exit Function
        TryDuration_TTR_FTT = True
        Exit Function
    End If
    text = Trim$(Replace(CStr(rawValue), ChrW(160), " "))
    parts = Split(text, ":")
    If UBound(parts) <> 2 Then Exit Function
    If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Then Exit Function
    If parts(0) Like "*[!0-9]*" Or parts(1) Like "*[!0-9]*" Then Exit Function
    secondsText = Replace(parts(2), ",", ".")
    If Len(secondsText) = 0 Then Exit Function
    If secondsText Like "*[!0-9.]*" Then Exit Function
    If Len(secondsText) - Len(Replace(secondsText, ".", "")) > 1 Then Exit Function
    If Not secondsText Like "*[0-9]*" Then Exit Function
    hours = CDbl(parts(0)): minutes = CDbl(parts(1)): seconds = Val(secondsText)
    If minutes >= 60 Or seconds >= 60 Then Exit Function
    duration = (hours * 3600# + minutes * 60# + seconds) / 86400#
    TryDuration_TTR_FTT = True
    Exit Function
InvalidDuration:
    TryDuration_TTR_FTT = False
End Function

Private Function NormalizeCountry_TTR_FTT( _
    ByVal txt As String) As String

    txt = LCase$(Trim$(txt))
    txt = Replace(txt, ChrW(160), " ")
    txt = Replace(txt, "–", "-")
    txt = Replace(txt, "—", "-")
    txt = Replace(txt, " - ", "-")

    txt = FixLatChars(txt)

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizeCountry_TTR_FTT = txt

End Function

Function FixLatChars(ByVal txt As String) As String

    txt = Replace(txt, "t", "т")
    txt = Replace(txt, "p", "р")
    txt = Replace(txt, "y", "у")
    txt = Replace(txt, "a", "а")
    txt = Replace(txt, "c", "с")
    txt = Replace(txt, "o", "о")
    txt = Replace(txt, "e", "е")
    txt = Replace(txt, "x", "х")

    FixLatChars = txt

End Function

Private Sub AddCopyButton_CMB( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal buttonCell As Range, _
    ByVal targetAddress As String, _
    ByVal buttonColor As Long)

    Dim buttonShape As Shape
    Dim macroWorkbookName As String
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0

    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    buttonWidth = buttonCell.Width - 4
    buttonHeight = buttonCell.Height - 4
    buttonLeft = buttonCell.Left + 2
    buttonTop = buttonCell.Top + 2

    If buttonWidth < 20 Then buttonWidth = 20
    If buttonHeight < 14 Then buttonHeight = 14

    Set buttonShape = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        buttonLeft, _
        buttonTop, _
        buttonWidth, _
        buttonHeight)

    With buttonShape

        .Name = buttonName
        .AlternativeText = targetAddress
        .OnAction = "'" & macroWorkbookName & "'!TTR_FTT_CopyBlock"

        .Fill.ForeColor.RGB = buttonColor
        .Fill.Solid
        .Line.ForeColor.RGB = buttonColor
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

End Sub

Public Function TTR_FTT_CopyBlock() As Variant

    Dim sourceSheet As Worksheet
    Dim buttonShape As Shape
    Dim targetRange As Range

    On Error GoTo CopyFail

    Set sourceSheet = ActiveSheet
    Set buttonShape = sourceSheet.Shapes(CStr(Application.Caller))
    Set targetRange = sourceSheet.Range(buttonShape.AlternativeText)

    targetRange.Copy

    Application.StatusBar = _
        "Данные скопированы: " & _
        sourceSheet.Name & "!" & _
        targetRange.Address(False, False)

    TTR_FTT_CopyBlock = True
    Exit Function

CopyFail:

    MsgBox _
        "Не удалось скопировать данные." & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Monitoring_TTR_FTT"

    TTR_FTT_CopyBlock = False

End Function
