Option Explicit

Public Sub Monitoring_Переключения_Агентов()

    Const RESULT_SHEET As String = "Результат"

    Dim wb As Workbook
    Dim src As Worksheet
    Dim res As Worksheet
    Dim headerCell As Range

    Dim headerRow As Long
    Dim tableCol As Long
    Dim countryCol As Long
    Dim referralCol As Long
    Dim metricCol As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Dim dayCols As Collection
    Dim dayKeys As Collection
    Dim targets As Collection

    Dim dataMap As Object
    Dim dayMap As Object

    Dim currentTable As String
    Dim currentCountry As String
    Dim currentReferral As String
    Dim combinationKey As String
    Dim dayKey As String
    Dim metricText As String

    Dim outputData() As Variant
    Dim target As Variant
    Dim cellValue As Variant

    Dim oldCalculation As XlCalculation

    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim d As Long
    Dim outputLastRow As Long
    Dim outputLastCol As Long

    Set src = ActiveSheet
    Set wb = src.Parent

    If src.Name = RESULT_SHEET Then
        MsgBox "Открой лист с исходной выгрузкой и запусти макрос ещё раз.", _
               vbExclamation, "Переключения"
        Exit Sub
    End If

    oldCalculation = Application.Calculation

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Новые названия первых трёх колонок Tableau.
    ' Логика подсчёта ниже НЕ ИЗМЕНЯЕТСЯ.
    Set headerCell = src.Cells.Find( _
        What:="Гранулярность переключений", _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    ' Поддержка старого названия на случай старой выгрузки.
    If headerCell Is Nothing Then
        Set headerCell = src.Cells.Find( _
            What:="Стол переключения", _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
    End If

    If headerCell Is Nothing Then
        Err.Raise vbObjectError + 1, , _
                  "Не найден заголовок стола переключений."
    End If

    headerRow = headerCell.Row
    tableCol = headerCell.Column

    countryCol = FindHeaderColumn( _
        src, headerRow, "Гранулярность переключений 2 слой")

    If countryCol = 0 Then
        countryCol = FindHeaderColumn( _
            src, headerRow, "Страна (переключения)")
    End If

    referralCol = FindHeaderColumn( _
        src, headerRow, "Гранулярность переключений 3 слой")

    If referralCol = 0 Then
        referralCol = FindHeaderColumn( _
            src, headerRow, "Реферал (переключения)")
    End If

    If countryCol = 0 Or referralCol = 0 Then
        Err.Raise vbObjectError + 2, , _
                  "Не найдены столбцы страны или реферала."
    End If

    ' Метрика идёт сразу после трёх колонок гранулярности.
    ' Max позволяет не зависеть от порядка отображения A/B/C.
    metricCol = Application.WorksheetFunction.Max( _
        tableCol, countryCol, referralCol) + 1

    lastRow = GetLastRow(src)
    lastCol = GetLastColumn(src)

    Set dayCols = New Collection
    Set dayKeys = New Collection

    For c = metricCol + 1 To lastCol

        cellValue = src.Cells(headerRow, c).Value2

        If IsNumeric(cellValue) Then
            If CDbl(cellValue) >= 1 And CDbl(cellValue) <= 31 Then
                dayCols.Add c
                dayKeys.Add CStr(CLng(cellValue))
            End If
        End If

    Next c

    If dayCols.Count = 0 Then
        Err.Raise vbObjectError + 3, , _
                  "В выгрузке не найдены дни месяца."
    End If

    Set dataMap = CreateObject("Scripting.Dictionary")
    dataMap.CompareMode = vbTextCompare

    For r = headerRow + 1 To lastRow

        cellValue = src.Cells(r, tableCol).Value2

        If Len(Trim$(CStr(cellValue))) > 0 Then

            If NormalizeText(cellValue) = "grand total" Then
                currentTable = vbNullString
            Else
                currentTable = Trim$(CStr(cellValue))
            End If

        End If

        cellValue = src.Cells(r, countryCol).Value2

        If Len(Trim$(CStr(cellValue))) > 0 Then
            If NormalizeText(cellValue) <> "total" Then
                currentCountry = Trim$(CStr(cellValue))
            End If
        End If

        cellValue = src.Cells(r, referralCol).Value2

        If Len(Trim$(CStr(cellValue))) > 0 Then
            If NormalizeText(cellValue) <> "total" Then
                currentReferral = Trim$(CStr(cellValue))
            End If
        End If

        metricText = NormalizeText( _
            src.Cells(r, metricCol).Value2)

        If InStr( _
            1, _
            metricText, _
            "count of view_switches", _
            vbTextCompare) = 1 Then

            If Len(currentTable) > 0 _
               And Len(currentCountry) > 0 _
               And Len(currentReferral) > 0 Then

                combinationKey = BuildKey( _
                    currentTable, _
                    currentCountry, _
                    currentReferral)

                If Not dataMap.Exists(combinationKey) Then

                    Set dayMap = CreateObject("Scripting.Dictionary")
                    dayMap.CompareMode = vbTextCompare
                    dataMap.Add combinationKey, dayMap

                Else

                    Set dayMap = dataMap(combinationKey)

                End If

                For d = 1 To dayCols.Count

                    c = CLng(dayCols(d))
                    dayKey = CStr(dayKeys(d))
                    cellValue = src.Cells(r, c).Value2

                    If Not IsError(cellValue) Then

                        If Len(CStr(cellValue)) > 0 _
                           And IsNumeric(cellValue) Then

                            If dayMap.Exists(dayKey) Then

                                dayMap(dayKey) = _
                                    CDbl(dayMap(dayKey)) + _
                                    CDbl(cellValue)

                            Else

                                dayMap.Add dayKey, CDbl(cellValue)

                            End If

                        End If

                    End If

                Next d

            End If

        End If

    Next r

    Set targets = BuildTargets()

    ReDim outputData( _
        1 To targets.Count + 1, _
        1 To dayCols.Count + 3)

    outputData(1, 1) = "Стол"
    outputData(1, 2) = "Гео"
    outputData(1, 3) = "Реферал"

    For d = 1 To dayKeys.Count
        outputData(1, d + 3) = dayKeys(d)
    Next d

    For i = 1 To targets.Count

        target = targets(i)

        outputData(i + 1, 1) = target(0)
        outputData(i + 1, 2) = target(1)
        outputData(i + 1, 3) = target(2)

        For d = 1 To dayKeys.Count

            outputData(i + 1, d + 3) = GetTargetValue( _
                dataMap, _
                CStr(target(3)), _
                CStr(target(4)), _
                CStr(target(2)), _
                CStr(dayKeys(d)))

        Next d

    Next i

    Set res = GetOrCreateSheet(wb, RESULT_SHEET)

    res.Cells.Clear

    res.Range("A1").Resize( _
        targets.Count + 1, _
        dayCols.Count + 3).Value = outputData

    outputLastRow = targets.Count + 1
    outputLastCol = dayCols.Count + 3

    With res.Range( _
        res.Cells(1, 1), _
        res.Cells(outputLastRow, outputLastCol))

        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = False
        .VerticalAlignment = xlCenter
        .RowHeight = 15

    End With

    With res.Range( _
        res.Cells(1, 1), _
        res.Cells(1, outputLastCol))

        .Interior.Color = RGB(226, 239, 218)
        .HorizontalAlignment = xlCenter

    End With

    With res.Range( _
        res.Cells(2, 1), _
        res.Cells(outputLastRow, 3))

        .Interior.Color = RGB(242, 242, 242)

    End With

    With res.Range( _
        res.Cells(2, 4), _
        res.Cells(outputLastRow, outputLastCol))

        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter

    End With

    With res.Range( _
        res.Cells(1, 1), _
        res.Cells(outputLastRow, outputLastCol)).Borders

        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlHairline

    End With

    res.Columns(1).ColumnWidth = 11
    res.Columns(2).ColumnWidth = 20
    res.Columns(3).ColumnWidth = 28

    res.Range( _
        res.Cells(1, 4), _
        res.Cells(outputLastRow, outputLastCol) _
    ).ColumnWidth = 5.5

    res.Activate

    ActiveWindow.FreezePanes = False
    res.Range("D2").Select
    ActiveWindow.FreezePanes = True

    res.Range( _
        res.Cells(2, 4), _
        res.Cells(outputLastRow, outputLastCol) _
    ).Select

    Selection.Copy

    MsgBox _
        "Готово." & vbCrLf & vbCrLf & _
        "Данные находятся на листе «Результат» " & _
        "и уже скопированы." & vbCrLf & _
        "Вставь их в Google Таблицу начиная с E11.", _
        vbInformation, _
        "Monitoring_Переключения_Агентов"

SafeExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = oldCalculation

    Exit Sub

ErrorHandler:

    Application.CutCopyMode = False

    MsgBox _
        "Не удалось подготовить данные:" & _
        vbCrLf & vbCrLf & Err.Description, _
        vbCritical, _
        "Monitoring_Переключения_Агентов"

    Resume SafeExit

End Sub

Private Function BuildTargets() As Collection

    Dim targets As New Collection

    targets.Add Array("EGP 1-5", "Египет", "1xCasino", "EGP", "Egypt")
    targets.Add Array("EGP 1-5", "Египет", "1xGames", "EGP", "Egypt")
    targets.Add Array("EGP 1-5", "Египет", "webDefault", "EGP", "Egypt")
    targets.Add Array("EGP 1-5", "Ливия", "webDefault", "EGP", "Libya")

    targets.Add Array("MT4-BG", "Алжир", "1xCasino", "MT4-BG", "Algeria")
    targets.Add Array("MT4-BG", "Бахрейн", "webDefault", "MT4-BG", "Bahrain")
    targets.Add Array("MT4-BG", "Египет", "Afropari", "MT4-BG", "Egypt")
    targets.Add Array("MT4-BG", "Египет", "Bizbet Africa (Ar)", "MT4-BG", "Egypt")
    targets.Add Array("MT4-BG", "Иордания", "webDefault", "MT4-BG", "Jordan")
    targets.Add Array("MT4-BG", "Кыргызстан", "1xCasino", "MT4-BG", "Kyrgyzstan")
    targets.Add Array("MT4-BG", "Мавритания", "1xCasino", "MT4-BG", "Mauritania")
    targets.Add Array("MT4-BG", "Марокко", "1xCasino", "MT4-BG", "Morocco")
    targets.Add Array("MT4-BG", "Марокко", "Afropari", "MT4-BG", "Morocco")
    targets.Add Array("MT4-BG", "Марокко", "Bizbet Africa (Ar)", "MT4-BG", "Morocco")
    targets.Add Array("MT4-BG", "Турция", "1xCasino", "MT4-BG", "Turkey")
    targets.Add Array("MT4-BG", "Турция", "BizBet", "MT4-BG", "Turkey")
    targets.Add Array("MT4-BG", "Турция", "RolsBet", "MT4-BG", "Turkey")

    targets.Add Array("MAR 1-3", "Боливия", "webDefault + bo.1xbet.com", "MAR", "Bolivia")
    targets.Add Array("MAR 1-3", "Гватемала", "webDefault", "MAR", "Guatemala")
    targets.Add Array("MAR 1-3", "Гондурас", "webDefault", "MAR", "Honduras")
    targets.Add Array("MAR 1-3", "Коста-Рика", "webDefault", "MAR", "Costa Rica")
    targets.Add Array("MAR 1-3", "Марокко", "webDefault", "MAR", "Morocco")
    targets.Add Array("MAR 1-3", "Никарагуа", "webDefault", "MAR", "Nicaragua")
    targets.Add Array("MAR 1-3", "ОАЭ", "webDefault", "MAR", "United Arab Emirates")
    targets.Add Array("MAR 1-3", "Панама", "webDefault + 1xbet.pa", "MAR", "Panama")
    targets.Add Array("MAR 1-3", "Парагвай", "webDefault", "MAR", "Paraguay")

    targets.Add Array("MT1-SPB", "Джибути", "webDefault", "MT1-SPB", "Djibouti")
    targets.Add Array("MT1-SPB", "Иран", "Onjabet", "MT1-SPB", "Iran")
    targets.Add Array("MT1-SPB", "Иран", "webDefault + 1xir.com", "MT1-SPB", "Iran")
    targets.Add Array("MT1-SPB", "Сомали", "webDefault", "MT1-SPB", "Somalia")
    targets.Add Array("MT1-SPB", "Судан", "webDefault", "MT1-SPB", "Sudan")
    targets.Add Array("MT1-SPB", "Тунис", "webDefault + 1xbet.tn", "MT1-SPB", "Tunisia")
    targets.Add Array("MT1-SPB", "Южный Судан", "webDefault", "MT1-SPB", "South Sudan")

    targets.Add Array("MRU", "Мавритания", "webDefault", "MRU", "Mauritania")
    targets.Add Array("ALG 1-2", "Алжир", "webDefault", "ALG", "Algeria")

    targets.Add Array("MT2-KZN", "Гаити", "webDefault", "MT2-KZN", "Haiti")
    targets.Add Array("MT2-KZN", "Катар", "webDefault", "MT2-KZN", "Qatar")
    targets.Add Array("MT2-KZN", "Кувейт", "webDefault", "MT2-KZN", "Kuwait")
    targets.Add Array("MT2-KZN", "Ливан", "webDefault", "MT2-KZN", "Lebanon")
    targets.Add Array("MT2-KZN", "Оман", "webDefault", "MT2-KZN", "Oman")
    targets.Add Array("MT2-KZN", "Палестина", "webDefault", "MT2-KZN", "Palestine")
    targets.Add Array("MT2-KZN", "Саудовская Аравия", "webDefault", "MT2-KZN", "Saudi Arabia")
    targets.Add Array("MT2-KZN", "Сирия", "webDefault", "MT2-KZN", "Syria")

    targets.Add Array("MT3-BG", "Афганистан", "webDefault", "MT3-BG", "Afghanistan")
    targets.Add Array("MT3-BG", "Ирак", "webDefault", "MT3-BG", "Iraq")
    targets.Add Array("MT3-BG", "Йемен", "webDefault", "MT3-BG", "Yemen")
    targets.Add Array("MT3-BG", "Папуа Новая Гвинея", "webDefault", "MT3-BG", "Papua New Guinea")

    targets.Add Array("MT5-BG", "Доминиканская республика", "webDefault", "MT5-BG", "Dominican Republic")
    targets.Add Array("MT5-BG", "Канада", "webDefault", "MT5-BG", "Canada")
    targets.Add Array("MT5-BG", "Кыргызстан", "webDefault", "MT5-BG", "Kyrgyzstan")
    targets.Add Array("MT5-BG", "Тайвань", "webDefault", "MT5-BG", "Taiwan")
    targets.Add Array("MT5-BG", "Ямайка", "webDefault", "MT5-BG", "Jamaica")

    targets.Add Array("TUR API+BT", "Азербайджан", "1xCasino", "TUR", "Azerbaijan")
    targets.Add Array("TUR API+BT", "Азербайджан", "Onjabet", "TUR", "Azerbaijan")
    targets.Add Array("TUR API+BT", "Азербайджан", "webDefault", "TUR", "Azerbaijan")
    targets.Add Array("TUR API+BT", "Турция", "melbet", "TUR", "Turkey")
    targets.Add Array("TUR API+BT", "Турция", "vippari", "TUR", "Turkey")
    targets.Add Array("TUR API+BT", "Турция", "webDefault + 1xbet22.com", "TUR", "Turkey")

    Set BuildTargets = targets

End Function

Private Function GetTargetValue( _
    ByVal dataMap As Object, _
    ByVal tableName As String, _
    ByVal countryName As String, _
    ByVal referralSpec As String, _
    ByVal dayKey As String) As Variant

    Dim referrals As Variant
    Dim referral As Variant
    Dim combinationKey As String
    Dim dayMap As Object

    Dim total As Double
    Dim combinationFound As Boolean
    Dim valueFound As Boolean

    referrals = Split(referralSpec, "+")

    For Each referral In referrals

        combinationKey = BuildKey( _
            tableName, _
            countryName, _
            Trim$(CStr(referral)))

        If dataMap.Exists(combinationKey) Then

            combinationFound = True
            Set dayMap = dataMap(combinationKey)

            If dayMap.Exists(dayKey) Then
                total = total + CDbl(dayMap(dayKey))
                valueFound = True
            End If

        End If

    Next referral

    If Not combinationFound And UBound(referrals) > 0 Then

        combinationKey = BuildKey( _
            tableName, _
            countryName, _
            referralSpec)

        If dataMap.Exists(combinationKey) Then

            Set dayMap = dataMap(combinationKey)

            If dayMap.Exists(dayKey) Then
                total = CDbl(dayMap(dayKey))
                valueFound = True
            End If

        End If

    End If

    If valueFound Then
        GetTargetValue = total
    Else
        GetTargetValue = 0
    End If

End Function

Private Function FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal headerText As String) As Long

    Dim foundCell As Range

    Set foundCell = ws.Rows(headerRow).Find( _
        What:=headerText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If Not foundCell Is Nothing Then
        FindHeaderColumn = foundCell.Column
    Else
        FindHeaderColumn = 0
    End If

End Function

Private Function BuildKey( _
    ByVal tableName As String, _
    ByVal countryName As String, _
    ByVal referralName As String) As String

    BuildKey = _
        NormalizeText(tableName) & "|" & _
        NormalizeText(countryName) & "|" & _
        NormalizeText(referralName)

End Function

Private Function NormalizeText( _
    ByVal value As Variant) As String

    Dim text As String

    If IsError(value) Then
        NormalizeText = vbNullString
        Exit Function
    End If

    text = CStr(value)
    text = Replace(text, ChrW(160), " ")
    text = Replace(text, vbTab, " ")
    text = Trim$(text)

    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    NormalizeText = LCase$(text)

End Function

Private Function GetLastRow( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If Not foundCell Is Nothing Then
        GetLastRow = foundCell.Row
    End If

End Function

Private Function GetLastColumn( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious)

    If Not foundCell Is Nothing Then
        GetLastColumn = foundCell.Column
    End If

End Function

Private Function GetOrCreateSheet( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then

        Set GetOrCreateSheet = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.Count))

        GetOrCreateSheet.Name = sheetName

    End If

End Function