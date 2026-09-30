Option Explicit

' Значение колонки "From", означающее, что тикет завёл сам пользователь.
' Сравнение строгое: "by customer" является подстрокой "by customer support",
' поэтому InStr здесь использовать нельзя.
Private Const FROM_USER As String = "by customer"

Sub Monitoring_API_BT_M_PSP_SMP()

    Dim wb As Workbook
    Dim srcWs As Worksheet

    Set srcWs = ActiveSheet
    Set wb = srcWs.Parent

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Worksheets("API").Delete
    wb.Worksheets("BT M").Delete
    wb.Worksheets("PSP").Delete
    wb.Worksheets("SMP").Delete
    On Error GoTo CleanFail

    Application.DisplayAlerts = True

    ' Листы создаются в нужном порядке слева направо.
    BuildReport srcWs, "API", "API", "", "API"
    BuildReport srcWs, "BT M", "BT M", "", "BTM"
    BuildReport srcWs, "PSP", "PSP", "", "PSP"
    BuildReport srcWs, "SMP", "SMP M", "SendMePay", "SMP"

    wb.Worksheets("API").Activate

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Готово! Листы расположены по порядку: API, BT M, PSP, SMP.", vbInformation, "Мониторинг"
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Не удалось сформировать отчёт: " & Err.Description, vbCritical, "Мониторинг"

End Sub


Private Sub BuildReport(ByVal srcWs As Worksheet, _
                        ByVal resultSheetName As String, _
                        ByVal ticketTypeNeed As String, _
                        ByVal agentNeed As String, _
                        ByVal reportMode As String)

    Dim dstWs As Worksheet
    Dim lastRow As Long, i As Long

    reportMode = UCase(Trim(reportMode))

    Dim topicCol As Long, countryCol As Long, referalCol As Long
    Dim fromCol As Long, deptCol As Long, ticketTypeCol As Long, agentCol As Long
    Dim typeCol As Long, paymentMethodCol As Long

    topicCol = GetHeaderCol(srcWs, "Topic")
    countryCol = GetHeaderCol(srcWs, "Country")
    referalCol = GetHeaderCol(srcWs, "Referal")
    fromCol = GetHeaderCol(srcWs, "From")
    deptCol = GetHeaderCol(srcWs, "Department")
    ticketTypeCol = GetHeaderCol(srcWs, "Ticket type")
    typeCol = GetHeaderCol(srcWs, "Type")

    If reportMode = "API" Then
        paymentMethodCol = GetHeaderCol(srcWs, "Payment method")
    End If

    If agentNeed <> "" Then
        agentCol = GetHeaderCol(srcWs, "Agent")
    End If

    If topicCol = 0 Or countryCol = 0 Or referalCol = 0 Or fromCol = 0 Or deptCol = 0 Or ticketTypeCol = 0 Or typeCol = 0 Then
        MsgBox resultSheetName & ": не найдены нужные колонки Topic / Country / Referal / From / Department / Ticket type / Type.", vbCritical
        Exit Sub
    End If

    If reportMode = "API" And paymentMethodCol = 0 Then
        MsgBox resultSheetName & ": не найдена колонка Payment method для проверки новых API-рефералов.", vbCritical
        Exit Sub
    End If

    If agentNeed <> "" And agentCol = 0 Then
        MsgBox resultSheetName & ": не найдена колонка Agent для фильтрации SendMePay.", vbCritical
        Exit Sub
    End If

    lastRow = GetLastDataRow(srcWs, Array(ticketTypeCol, deptCol, countryCol, topicCol))

    Dim countryList As Variant, refList As Variant

    Select Case reportMode

        Case "API"
            countryList = Array("Azerbaijan", "Algeria", "Afghanistan", "Bahrain", "Bolivia", "Haiti", _
                                "Guatemala", "Honduras", "Djibouti", "Dominican Republic", "Egypt", "Jordan", "Iraq", _
                                "Iran", "Yemen", "Canada", "Qatar", "Kuwait", "Kyrgyzstan", "Lebanon", "Libya", _
                                "Mauritania", "Morocco", "Nicaragua", "United Arab Emirates", "Oman", "Palestine", _
                                "Panama", "Papua New Guinea", "Paraguay", "Saudi Arabia", "Syria", "Somalia", _
                                "Sudan", "Taiwan", "Tunisia", "Turkey", "South Sudan", "Jamaica")

            refList = Array("webdefault", "1xbet22.com", "melbet", "1xir.com", "1xgames", _
                            "bo.1xbet.com", "1xbet.tn", "1xbet.et", "bizbet", "1xcasino", _
                            "afropari", "onjabet", "bizbet africa (ar)", "1xbet.pa", _
                            "in.1xbet.com", "rolsbet", "1xbet.latam", "vippari")

        Case "BTM"
            ' Papua New Guinea и Paraguay остаются в выводе как технические
            ' позиции, чтобы не сдвигать строки в таблице за август.
            countryList = Array("Azerbaijan", "Algeria", "Afghanistan", "Bahrain", "Bolivia", "Haiti", _
                                "Guatemala", "Honduras", "Djibouti", "Dominican Republic", "Egypt", "Jordan", "Iraq", _
                                "Iran", "Yemen", "Canada", "Qatar", "Kuwait", "Kyrgyzstan", "Lebanon", "Libya", _
                                "Mauritania", "Morocco", "Nicaragua", "United Arab Emirates", "Oman", "Palestine", _
                                "Panama", "Papua New Guinea", "Paraguay", "Saudi Arabia", "Syria", "Somalia", _
                                "Taiwan", "Tunisia", "Turkey", "South Sudan")

            refList = Array("webdefault", "1xbet22.com", "melbet", "1xir.com", "1xgames", _
                            "bo.1xbet.com", "1xbet.tn", "1xbet.et", "bizbet", "1xcasino", _
                            "afropari", "onjabet", "bizbet africa (ar)", "1xbet.pa", _
                            "vippari")

        Case "PSP"
            countryList = Array("Azerbaijan", "Turkey", "Iran", "Somalia", "Kyrgyzstan")

            refList = Array("webdefault", "1xbet22.com", "melbet", "1xir.com", "1xgames", _
                            "bo.1xbet.com", "1xbet.tn", "1xbet.et", "bizbet", "1xcasino", _
                            "afropari", "onjabet", "bizbet africa (ar)", "1xbet.pa")

        Case "SMP"
            countryList = Array("Egypt", "Mauritania", "Sudan", _
                                "Dominican Republic", "Honduras", "Nicaragua", "South Sudan", "Paraguay", "Qatar", "United Arab Emirates", "Canada", "Panama", "Guatemala")

            refList = Array("webdefault", "1xbet22.com", "melbet", "1xir.com", "1xgames", _
                            "bo.1xbet.com", "1xbet.tn", "1xbet.et", "bizbet", "1xcasino", _
                            "afropari", "onjabet", "bizbet africa (ar)", "1xbet.pa")

    End Select

    Dim dM1 As Long, dUM1 As Long, dPM1 As Long, wM1 As Long, wUM1 As Long, wPM1 As Long
    Dim dML As Long, dUML As Long, dPML As Long, wML As Long, wUML As Long, wPML As Long
    Dim dBf As Long, dUBf As Long, dPBf As Long, wBf As Long, wUBf As Long, wPBf As Long
    
    Dim vipM1 As Long, vipML As Long, vipBf As Long

    Dim countryM1 As Object, countryML As Object, countryBf As Object
    Dim refM1 As Object, refML As Object, refBf As Object

    ' vbTextCompare: ключи стран кладутся в исходном регистре выгрузки, а ищутся
    ' по countryList. Без него страна с другим регистром пройдёт IsCountryInList,
    ' попадёт в ИТОГО и покажет 0 в блоке GEO.
    Set countryM1 = CreateObject("Scripting.Dictionary")
    countryM1.CompareMode = vbTextCompare
    Set countryML = CreateObject("Scripting.Dictionary")
    countryML.CompareMode = vbTextCompare
    Set countryBf = CreateObject("Scripting.Dictionary")
    countryBf.CompareMode = vbTextCompare
    Set refM1 = CreateObject("Scripting.Dictionary")
    refM1.CompareMode = vbTextCompare
    Set refML = CreateObject("Scripting.Dictionary")
    refML.CompareMode = vbTextCompare
    Set refBf = CreateObject("Scripting.Dictionary")
    refBf.CompareMode = vbTextCompare

    Dim vDept As String, vDeptKey As String, vTopic As String, vFrom As String
    Dim vCountry As String, vRef As String, refName As String
    Dim vTicketType As String, vAgent As String
    Dim vType As String, vPaymentMethod As String

    For i = 2 To lastRow

        vTicketType = Trim(CStr(srcWs.Cells(i, ticketTypeCol).Value))
        
        If UCase(vTicketType) <> UCase(ticketTypeNeed) Then
            GoTo NextRow
        End If

        If agentNeed <> "" Then
            vAgent = LCase(Trim(CStr(srcWs.Cells(i, agentCol).Value)))
            ' Пустой Agent не повод отбрасывать строку: принадлежность к
            ' SendMePay уже гарантирована значением Ticket type. Отбрасываем
            ' только строку, у которой явно проставлен чужой агент.
            If vAgent <> "" Then
                If InStr(1, vAgent, LCase(agentNeed), vbTextCompare) = 0 Then
                    GoTo NextRow
                End If
            End If
        End If

        vType = Trim(CStr(srcWs.Cells(i, typeCol).Value))
        
        Dim isVIP As Boolean
        isVIP = (UCase(vType) = "VIP" Or UCase(vType) = "VIP-CASINO" Or UCase(vType) = "VIP-HYBRID")

        vDept = Trim(CStr(srcWs.Cells(i, deptCol).Value))
        vDeptKey = GetDepartmentKey(vDept)
        vTopic = LCase(Trim(CStr(srcWs.Cells(i, topicCol).Value)))
        vFrom = LCase(Trim(CStr(srcWs.Cells(i, fromCol).Value)))
        vCountry = Trim(CStr(srcWs.Cells(i, countryCol).Value))
        vRef = LCase(Trim(CStr(srcWs.Cells(i, referalCol).Value)))

        ' Блок Papua New Guinea остаётся в отчёте для сохранения структуры,
        ' но строки этого GEO полностью исключаются из всех подсчётов.
        If LCase(vCountry) = "papua new guinea" Then
            GoTo NextRow
        End If

        ' Paraguay перешёл в SMP. В BT M строка сохраняется
        ' только для совместимости с августовской таблицей.
        If UCase(reportMode) = "BTM" And LCase(vCountry) = "paraguay" Then
            GoTo NextRow
        End If

        ' В BT M и SMP общие итоги, GEO и рефералы считаются
        ' только по согласованным спискам стран.
        If UCase(reportMode) = "BTM" Or UCase(reportMode) = "SMP" Then
            If Not IsCountryInList(vCountry, countryList) Then
                GoTo NextRow
            End If
        End If

        If InStr(vRef, "#") > 0 Then
            refName = Trim(Left(vRef, InStr(vRef, "#") - 1))
        Else
            refName = vRef
        End If

        ' Rolsbet полностью исключён из BT M.
        ' В API не учитывается только связка Египет + Rolsbet.
        If IsExcludedReferralCountry(reportMode, refName, vCountry) Then
            GoTo NextRow
        End If

        If reportMode = "API" Then
            vPaymentMethod = Trim(CStr(srcWs.Cells(i, paymentMethodCol).Value))
            If Not IsAllowedApiReferral(refName, vCountry, vPaymentMethod) Then
                GoTo NextRow
            End If
        End If

        If reportMode = "PSP" Then
            If LCase(Trim(vCountry)) = "tunisia" Then
                GoTo NextRow
            End If
        End If

        If vDeptKey = "M1" Then

            If isVIP Then
                vipM1 = vipM1 + 1
            Else
                If InStr(vTopic, "deposit") > 0 Then
                    dM1 = dM1 + 1
                    If vFrom = FROM_USER Then dUM1 = dUM1 + 1 Else dPM1 = dPM1 + 1
                ElseIf InStr(vTopic, "withdraw") > 0 Then
                    wM1 = wM1 + 1
                    If vFrom = FROM_USER Then wUM1 = wUM1 + 1 Else wPM1 = wPM1 + 1
                End If
            End If

            countryM1(vCountry) = countryM1(vCountry) + 1
            refM1(refName) = refM1(refName) + 1

        ElseIf vDeptKey = "ML" Then

            If isVIP Then
                vipML = vipML + 1
            Else
                If InStr(vTopic, "deposit") > 0 Then
                    dML = dML + 1
                    If vFrom = FROM_USER Then dUML = dUML + 1 Else dPML = dPML + 1
                ElseIf InStr(vTopic, "withdraw") > 0 Then
                    wML = wML + 1
                    If vFrom = FROM_USER Then wUML = wUML + 1 Else wPML = wPML + 1
                End If
            End If

            If reportMode = "API" Or reportMode = "BTM" Then
                countryML(vCountry) = countryML(vCountry) + 1
                refML(refName) = refML(refName) + 1
            End If

        ElseIf vDeptKey = "BF" Then

            ' Buffer выводится только на листе API. Для BT M / PSP / SMP строки
            ' департамента Buffer в отчёт не попадают — так согласовано.
            If reportMode = "API" Then
                If isVIP Then
                    vipBf = vipBf + 1
                Else
                    If InStr(vTopic, "deposit") > 0 Then
                        dBf = dBf + 1
                        If vFrom = FROM_USER Then dUBf = dUBf + 1 Else dPBf = dPBf + 1
                    ElseIf InStr(vTopic, "withdraw") > 0 Then
                        wBf = wBf + 1
                        If vFrom = FROM_USER Then wUBf = wUBf + 1 Else wPBf = wPBf + 1
                    End If
                End If

                countryBf(vCountry) = countryBf(vCountry) + 1
                refBf(refName) = refBf(refName) + 1
            End If

        End If

NextRow:
    Next i

    Set dstWs = srcWs.Parent.Worksheets.Add(After:=srcWs.Parent.Worksheets(srcWs.Parent.Worksheets.Count))
    dstWs.Name = resultSheetName

    With dstWs

        If reportMode = "API" Then

            .Range("A1:D1").Value = Array("Категория", "MENA 1x", "Mena Leads 1X", "Buffer")
            .Range("A1:D1").Font.Bold = True
            .Range("A2:A9").Value = Application.Transpose(Array("Депозиты", "  User", "  PS", "Выводы", "  User", "  PS", "VIP", "ИТОГО"))

            .Range("B2:B9").Value = Application.Transpose(Array(dM1, dUM1, dPM1, wM1, wUM1, wPM1, vipM1, dM1 + wM1 + vipM1))
            .Range("C2:C9").Value = Application.Transpose(Array(dML, dUML, dPML, wML, wUML, wPML, vipML, dML + wML + vipML))
            .Range("D2:D9").Value = Application.Transpose(Array(dBf, dUBf, dPBf, wBf, wUBf, wPBf, vipBf, dBf + wBf + vipBf))

            PrintGeoAPI dstWs, countryList, countryM1, countryML, countryBf
            PrintRefAPI dstWs, refList, refM1, refML, refBf

            .Columns("A:J").AutoFit

        ElseIf reportMode = "BTM" Then

            .Range("A1:C1").Value = Array("Категория", "MENA 1x", "Mena Leads 1X")
            .Range("A1:C1").Font.Bold = True
            .Range("A2:A9").Value = Application.Transpose(Array("Депозиты", "  User", "  PS", "Выводы", "  User", "  PS", "VIP", "ИТОГО"))

            .Range("B2:B9").Value = Application.Transpose(Array(dM1, dUM1, dPM1, wM1, wUM1, wPM1, vipM1, dM1 + wM1 + vipM1))
            .Range("C2:C9").Value = Application.Transpose(Array(dML, dUML, dPML, wML, wUML, wPML, vipML, dML + wML + vipML))

            PrintGeoBTM dstWs, countryList, countryM1, countryML
            PrintRefBTM dstWs, refList, refM1, refML

            .Columns("A:I").AutoFit

        ElseIf reportMode = "PSP" Then

            .Range("A1:B1").Value = Array("Категория", "MENA 1x")
            .Range("A1:B1").Font.Bold = True
            .Range("A2:A9").Value = Application.Transpose(Array("Депозиты", "  User", "  PS", "Выводы", "  User", "  PS", "VIP", "ИТОГО"))

            .Range("B2:B9").Value = Application.Transpose(Array(dM1, dUM1, dPM1, wM1, wUM1, wPM1, vipM1, dM1 + wM1 + vipM1))

            PrintGeoSingleWithGap dstWs, countryList, countryM1
            PrintRefSingleWithGap dstWs, refList, refM1

            .Columns("A:I").AutoFit

        ElseIf reportMode = "SMP" Then

            .Range("A1:B1").Value = Array("Категория", "MENA 1x")
            .Range("A1:B1").Font.Bold = True
            .Range("A2:A9").Value = Application.Transpose(Array("Депозиты", "  User", "  PS", "Выводы", "  User", "  PS", "VIP", "ИТОГО"))

            .Range("B2:B9").Value = Application.Transpose(Array(dM1, dUM1, dPM1, wM1, wUM1, wPM1, vipM1, dM1 + wM1 + vipM1))

            PrintGeoSMP dstWs, countryList, countryM1
            PrintRefSingleWithGap dstWs, refList, refM1

            .Columns("A:I").AutoFit

        End If

    End With

    FormatReportSheet dstWs, reportMode, countryList, refList

End Sub


Private Sub FormatReportSheet(ByVal ws As Worksheet, _
                              ByVal reportMode As String, _
                              ByVal countryList As Variant, _
                              ByVal refList As Variant)

    Dim summaryLastCol As Long
    Dim geoLabelCol As Long, refLabelCol As Long
    Dim geoStep As Long, refStep As Long
    Dim geoLastRow As Long, refLastRow As Long
    Dim lastReportCol As Long

    Dim headerColor As Long, headerTextColor As Long
    Dim groupFill As Long, subFill As Long
    Dim totalFill As Long, accentTextColor As Long, borderColor As Long

    headerColor = RGB(31, 78, 121)
    headerTextColor = RGB(255, 255, 255)
    groupFill = RGB(226, 232, 240)
    subFill = RGB(248, 250, 252)
    totalFill = RGB(226, 232, 240)
    accentTextColor = RGB(31, 78, 121)
    borderColor = RGB(203, 213, 225)

    Select Case UCase(reportMode)
        Case "API"
            summaryLastCol = 4
            geoLabelCol = 6
            refLabelCol = 9
            geoStep = 4
            refStep = 4

        Case "BTM"
            summaryLastCol = 3
            geoLabelCol = 5
            refLabelCol = 8
            geoStep = 3
            refStep = 3

        Case "PSP"
            summaryLastCol = 2
            geoLabelCol = 5
            refLabelCol = 8
            geoStep = 2
            refStep = 2

        Case Else
            summaryLastCol = 2
            geoLabelCol = 5
            refLabelCol = 8
            geoStep = 2
            refStep = 2
    End Select

    ws.Tab.Color = headerColor

    ' Отдельная верхняя строка для полноразмерных кнопок копирования.
    ws.Rows(1).Insert Shift:=xlDown

    If UCase(reportMode) = "SMP" Then
        ' В SMP между всеми GEO сохраняется пустая строка. Берём
        ' фактическую последнюю строку, чтобы кнопка копировала
        ' диапазон F3:F27 (13 стран, включая Катар, ОАЭ, Канаду, Панаму и Гватемалу).
        geoLastRow = ws.Cells(ws.Rows.Count, geoLabelCol + 1).End(xlUp).Row
    Else
        geoLastRow = 2 + (UBound(countryList) - LBound(countryList) + 1) * geoStep
    End If
    refLastRow = 2 + (UBound(refList) - LBound(refList) + 1) * refStep

    lastReportCol = refLabelCol + 1

    ' Новая строка остаётся обычной строкой Excel без оформления заголовков.
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastReportCol))
        .ClearFormats
        .Font.Name = "Arial"
        .Font.Size = 10
    End With

    StyleSummaryBlock ws, summaryLastCol, headerColor, headerTextColor, _
                      groupFill, subFill, totalFill, accentTextColor, borderColor

    StyleDetailSection ws, geoLabelCol, geoLastRow, geoStep, _
                       headerColor, headerTextColor, groupFill, subFill, _
                       accentTextColor, borderColor

    StyleDetailSection ws, refLabelCol, refLastRow, refStep, _
                       headerColor, headerTextColor, groupFill, subFill, _
                       accentTextColor, borderColor

    ws.Columns(1).ColumnWidth = 22
    ws.Range(ws.Cells(1, 2), ws.Cells(1, summaryLastCol)).EntireColumn.ColumnWidth = 14
    ws.Columns(geoLabelCol - 1).ColumnWidth = 3
    ws.Columns(geoLabelCol).ColumnWidth = 23
    ws.Columns(geoLabelCol + 1).ColumnWidth = 12

    If UCase(reportMode) = "SMP" Then
        ws.Columns(geoLabelCol).ColumnWidth = 29
    End If

    ws.Columns(refLabelCol - 1).ColumnWidth = 3
    ws.Columns(refLabelCol).ColumnWidth = 23
    ws.Columns(refLabelCol + 1).ColumnWidth = 12

    ws.Rows(1).RowHeight = 24
    ws.Rows(2).RowHeight = 24
    ws.Range("A3:A10").EntireRow.RowHeight = 20

    AddColumnCopyButtons ws, reportMode, summaryLastCol, geoLabelCol, geoLastRow, geoStep, _
                         refLabelCol, refLastRow, refStep, headerColor

    ws.Activate
    With ActiveWindow
        .DisplayGridlines = True
        .Zoom = 90
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 2
        .FreezePanes = True
    End With

End Sub


Private Sub StyleSummaryBlock(ByVal ws As Worksheet, _
                              ByVal lastCol As Long, _
                              ByVal headerColor As Long, _
                              ByVal headerTextColor As Long, _
                              ByVal groupFill As Long, _
                              ByVal subFill As Long, _
                              ByVal totalFill As Long, _
                              ByVal accentTextColor As Long, _
                              ByVal borderColor As Long)

    With ws.Range(ws.Cells(2, 1), ws.Cells(10, lastCol))
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
        .Interior.Color = headerColor
        .Font.Color = headerTextColor
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = headerColor
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(3, 1), ws.Cells(10, lastCol))
        .Interior.Color = subFill
        .Borders.LineStyle = xlContinuous
        .Borders.Color = borderColor
        .Borders.Weight = xlThin
    End With

    ws.Range("A4:A5").IndentLevel = 1
    ws.Range("A7:A8").IndentLevel = 1

    StyleSummaryCategoryRow ws, 3, lastCol, groupFill
    StyleSummaryCategoryRow ws, 6, lastCol, groupFill
    StyleSummaryCategoryRow ws, 9, lastCol, groupFill

    With ws.Range(ws.Cells(10, 1), ws.Cells(10, lastCol))
        .Interior.Color = totalFill
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = headerColor
    End With

    With ws.Range(ws.Cells(3, 2), ws.Cells(10, lastCol))
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    HighlightPositiveCounts ws.Range(ws.Cells(3, 2), ws.Cells(9, lastCol)), accentTextColor

End Sub


Private Sub StyleSummaryCategoryRow(ByVal ws As Worksheet, _
                                    ByVal rowNumber As Long, _
                                    ByVal lastCol As Long, _
                                    ByVal fillColor As Long)

    With ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, lastCol))
        .Interior.Color = fillColor
        .Font.Bold = True
    End With

End Sub


Private Sub StyleDetailSection(ByVal ws As Worksheet, _
                               ByVal labelCol As Long, _
                               ByVal lastRow As Long, _
                               ByVal stepSize As Long, _
                               ByVal headerColor As Long, _
                               ByVal headerTextColor As Long, _
                               ByVal groupFill As Long, _
                               ByVal subFill As Long, _
                               ByVal accentTextColor As Long, _
                               ByVal borderColor As Long)

    Dim r As Long, groupEndRow As Long

    With ws.Range(ws.Cells(2, labelCol), ws.Cells(lastRow, labelCol + 1))
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    ws.Cells(2, labelCol + 1).Value = "Количество"

    With ws.Range(ws.Cells(2, labelCol), ws.Cells(2, labelCol + 1))
        .Interior.Color = headerColor
        .Font.Color = headerTextColor
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = headerColor
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(3, labelCol + 1), ws.Cells(lastRow, labelCol + 1))
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    ' Возвращаем видимые ячейки во все блоки GEO и рефералов.
    With ws.Range(ws.Cells(3, labelCol), ws.Cells(lastRow, labelCol + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = borderColor
        .Borders.Weight = xlThin
    End With

    For r = 3 To lastRow Step stepSize
        groupEndRow = r + stepSize - 1
        If groupEndRow > lastRow Then groupEndRow = lastRow

        With ws.Range(ws.Cells(r, labelCol), ws.Cells(r, labelCol + 1))
            .Interior.Color = groupFill
            .Font.Bold = True
        End With

        If groupEndRow >= r + 1 Then
            With ws.Range(ws.Cells(r + 1, labelCol), ws.Cells(groupEndRow, labelCol + 1))
                .Interior.Color = subFill
            End With

            ws.Range(ws.Cells(r + 1, labelCol), ws.Cells(groupEndRow, labelCol)).IndentLevel = 1
        End If

        With ws.Range(ws.Cells(groupEndRow, labelCol), ws.Cells(groupEndRow, labelCol + 1))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = borderColor
            .Borders(xlEdgeBottom).Weight = xlThin
        End With

        ws.Range(ws.Cells(r, labelCol), ws.Cells(groupEndRow, labelCol + 1)).EntireRow.RowHeight = 20
    Next r

    HighlightPositiveCounts ws.Range(ws.Cells(3, labelCol + 1), ws.Cells(lastRow, labelCol + 1)), accentTextColor

End Sub


Private Sub AddColumnCopyButtons(ByVal ws As Worksheet, _
                                 ByVal reportMode As String, _
                                 ByVal summaryLastCol As Long, _
                                 ByVal geoLabelCol As Long, _
                                 ByVal geoLastRow As Long, _
                                 ByVal geoStep As Long, _
                                 ByVal refLabelCol As Long, _
                                 ByVal refLastRow As Long, _
                                 ByVal refStep As Long, _
                                 ByVal buttonColor As Long)

    Dim c As Long
    Dim geoCopyFirstRow As Long, refCopyFirstRow As Long
    Dim geoCopyLastRow As Long, refCopyLastRow As Long

    ' Сводка копируется по одной колонке, без строки ИТОГО.
    For c = 2 To summaryLastCol
        AddColumnCopyButton ws, "btnCopySummary" & CStr(c), c, 3, 9, buttonColor
    Next c

    geoCopyFirstRow = 3
    refCopyFirstRow = 3
    geoCopyLastRow = geoLastRow
    refCopyLastRow = refLastRow

    ' В API и BT M первая строка группы содержит название, поэтому начинаем с первой цифры.
    If geoStep > 2 Then geoCopyFirstRow = 4
    If refStep > 2 Then refCopyFirstRow = 4

    ' В PSP последняя строка является техническим пустым отступом.
    ' В SMP geoLastRow указывает на последнюю страну (Гватемала, строка 27),
    ' диапазон F3:F27 сохраняет пустые строки между всеми GEO.
    If geoStep = 2 And UCase(reportMode) <> "SMP" Then geoCopyLastRow = geoCopyLastRow - 1
    If refStep = 2 Then refCopyLastRow = refCopyLastRow - 1

    AddColumnCopyButton ws, "btnCopyGeo", geoLabelCol + 1, geoCopyFirstRow, geoCopyLastRow, buttonColor
    AddColumnCopyButton ws, "btnCopyRef", refLabelCol + 1, refCopyFirstRow, refCopyLastRow, buttonColor

End Sub


Private Sub AddColumnCopyButton(ByVal ws As Worksheet, _
                                ByVal buttonName As String, _
                                ByVal targetColumn As Long, _
                                ByVal firstRow As Long, _
                                ByVal lastRow As Long, _
                                ByVal buttonColor As Long)

    Dim buttonShape As Shape
    Dim buttonCell As Range
    Dim targetAddress As String
    Dim macroWorkbookName As String
    Dim buttonLeft As Double, buttonTop As Double
    Dim buttonWidth As Double, buttonHeight As Double

    Set buttonCell = ws.Cells(1, targetColumn)
    targetAddress = ws.Range(ws.Cells(firstRow, targetColumn), _
                             ws.Cells(lastRow, targetColumn)).Address(False, False)
    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    buttonWidth = buttonCell.Width - 4
    buttonHeight = buttonCell.Height - 4
    buttonLeft = buttonCell.Left + 2
    buttonTop = buttonCell.Top + 2

    Set buttonShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                         buttonLeft, buttonTop, _
                                         buttonWidth, buttonHeight)

    With buttonShape
        .Name = buttonName
        .AlternativeText = targetAddress
        .OnAction = "'" & macroWorkbookName & "'!CopyReportBlock"
        .Fill.ForeColor.RGB = buttonColor
        .Line.ForeColor.RGB = buttonColor
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


Public Function CopyReportBlock() As Variant
    Dim sourceSheet As Worksheet
    Dim buttonShape As Shape
    Dim targetRange As Range

    On Error GoTo CopyFail

    Set sourceSheet = ActiveSheet
    Set buttonShape = sourceSheet.Shapes(CStr(Application.Caller))
    Set targetRange = sourceSheet.Range(buttonShape.AlternativeText)

    targetRange.Copy
    Application.StatusBar = "Данные скопированы: " & sourceSheet.Name & "!" & targetRange.Address(False, False)
    Exit Function

CopyFail:
    Application.StatusBar = False
    MsgBox "Не удалось скопировать данные: " & Err.Description, vbExclamation, "Мониторинг"

End Function


Private Sub HighlightPositiveCounts(ByVal targetRange As Range, ByVal accentTextColor As Long)

    Dim cell As Range

    For Each cell In targetRange.Cells
        If Len(CStr(cell.Value)) > 0 Then
            If IsNumeric(cell.Value) Then
                If CDbl(cell.Value) > 0 Then
                    cell.Font.Bold = True
                    cell.Font.Color = accentTextColor
                End If
            End If
        End If
    Next cell

End Sub


Private Function IsExcludedReferralCountry(ByVal reportMode As String, _
                                           ByVal refName As String, _
                                           ByVal countryName As String) As Boolean

    Dim r As String, c As String, mode As String

    r = LCase(Trim(refName))
    c = LCase(Trim(countryName))
    mode = UCase(Trim(reportMode))

    If mode = "BTM" And r = "rolsbet" Then
        IsExcludedReferralCountry = True
    ElseIf mode = "API" And c = "egypt" And r = "rolsbet" Then
        IsExcludedReferralCountry = True
    Else
        IsExcludedReferralCountry = False
    End If

End Function


Private Function GetDepartmentKey(ByVal departmentName As String) As String

    ' Сравнение департамента должно быть регистронезависимым: при Option Compare
    ' Binary смена регистра в выгрузке молча обнулила бы весь отчёт.
    Select Case LCase(Trim(departmentName))
        Case "mena 1x"
            GetDepartmentKey = "M1"
        Case "mena leads 1x"
            GetDepartmentKey = "ML"
        Case "buffer"
            GetDepartmentKey = "BF"
        Case Else
            GetDepartmentKey = ""
    End Select

End Function


Private Function GetLastDataRow(ByVal ws As Worksheet, ByVal keyCols As Variant) As Long

    ' Последняя строка считается по максимуму из нескольких обязательных колонок.
    ' По одной колонке (раньше — Department) одна пустая ячейка обрезала хвост
    ' выгрузки. UsedRange намеренно не используем: его легко раздуть
    ' форматированием на пустых строках.
    Dim c As Variant, r As Long, maxRow As Long

    For Each c In keyCols
        If CLng(c) > 0 Then
            r = ws.Cells(ws.Rows.Count, CLng(c)).End(xlUp).Row
            If r > maxRow Then maxRow = r
        End If
    Next c

    If maxRow < 1 Then maxRow = 1
    GetLastDataRow = maxRow

End Function


Private Function IsCountryInList(ByVal countryName As String, _
                                 ByVal countryList As Variant) As Boolean

    Dim item As Variant

    For Each item In countryList
        If StrComp(Trim(CStr(item)), Trim(countryName), vbTextCompare) = 0 Then
            IsCountryInList = True
            Exit Function
        End If
    Next item

    IsCountryInList = False

End Function


Private Function IsAllowedApiReferral(ByVal refName As String, ByVal countryName As String, ByVal paymentMethod As String) As Boolean

    Dim r As String, c As String, p As String

    r = LCase(Trim(refName))
    c = LCase(Trim(countryName))
    p = LCase(Trim(paymentMethod))

    Select Case r

        Case "in.1xbet.com"
            IsAllowedApiReferral = (c = "turkey")

        Case "rolsbet"
            IsAllowedApiReferral = (c = "turkey")

        Case "1xbet.latam"
            If c = "jamaica" Then
                IsAllowedApiReferral = (p = "lynk")
            ElseIf c = "guatemala" Then
                IsAllowedApiReferral = (p = "gt continental" Or _
                                        p = "banrural" Or _
                                        p = "banco industrial" Or _
                                        p = "bac" Or _
                                        p = "zigi")
            Else
                IsAllowedApiReferral = False
            End If

        Case Else
            IsAllowedApiReferral = True

    End Select

End Function


Private Sub PrintGeoAPI(ByVal ws As Worksheet, ByVal countryList As Variant, _
                        ByVal countryM1 As Object, ByVal countryML As Object, ByVal countryBf As Object)

    Dim r As Long, cName As Variant

    r = 2
    ws.Cells(1, 6).Value = "Страна"
    ws.Range("F1:G1").Font.Bold = True

    For Each cName In countryList
        ws.Cells(r, 6).Value = cName
        ws.Cells(r, 6).Font.Bold = True
        ws.Cells(r + 1, 6).Value = "M1x"
        ws.Cells(r + 1, 7).Value = IIf(countryM1.Exists(cName), countryM1(cName), 0)
        ws.Cells(r + 2, 6).Value = "ML1x"
        ws.Cells(r + 2, 7).Value = IIf(countryML.Exists(cName), countryML(cName), 0)
        ws.Cells(r + 3, 6).Value = "Buffer"
        ws.Cells(r + 3, 7).Value = IIf(countryBf.Exists(cName), countryBf(cName), 0)
        r = r + 4
    Next cName

End Sub


Private Sub PrintRefAPI(ByVal ws As Worksheet, ByVal refList As Variant, _
                        ByVal refM1 As Object, ByVal refML As Object, ByVal refBf As Object)

    Dim r As Long, rName As Variant

    r = 2
    ws.Cells(1, 9).Value = "Реферал"
    ws.Range("I1:J1").Font.Bold = True

    For Each rName In refList
        ws.Cells(r, 9).Value = rName
        ws.Cells(r, 9).Font.Bold = True
        ws.Cells(r + 1, 9).Value = "M1x"
        ws.Cells(r + 1, 10).Value = IIf(refM1.Exists(rName), refM1(rName), 0)
        ws.Cells(r + 2, 9).Value = "ML1x"
        ws.Cells(r + 2, 10).Value = IIf(refML.Exists(rName), refML(rName), 0)
        ws.Cells(r + 3, 9).Value = "Buffer"
        ws.Cells(r + 3, 10).Value = IIf(refBf.Exists(rName), refBf(rName), 0)
        r = r + 4
    Next rName

End Sub


Private Sub PrintGeoBTM(ByVal ws As Worksheet, ByVal countryList As Variant, _
                        ByVal countryM1 As Object, ByVal countryML As Object)

    Dim r As Long, cName As Variant

    r = 2
    ws.Cells(1, 5).Value = "Страна"
    ws.Range("E1:F1").Font.Bold = True

    For Each cName In countryList
        ws.Cells(r, 5).Value = cName
        ws.Cells(r, 5).Font.Bold = True
        ws.Cells(r + 1, 5).Value = "M1x"
        ws.Cells(r + 1, 6).Value = IIf(countryM1.Exists(cName), countryM1(cName), 0)
        ws.Cells(r + 2, 5).Value = "ML1x"
        ws.Cells(r + 2, 6).Value = IIf(countryML.Exists(cName), countryML(cName), 0)
        r = r + 3
    Next cName

End Sub


Private Sub PrintRefBTM(ByVal ws As Worksheet, ByVal refList As Variant, _
                        ByVal refM1 As Object, ByVal refML As Object)

    Dim r As Long, rName As Variant

    r = 2
    ws.Cells(1, 8).Value = "Реферал"
    ws.Range("H1:I1").Font.Bold = True

    For Each rName In refList
        ws.Cells(r, 8).Value = rName
        ws.Cells(r, 8).Font.Bold = True
        ws.Cells(r + 1, 8).Value = "M1x"
        ws.Cells(r + 1, 9).Value = IIf(refM1.Exists(rName), refM1(rName), 0)
        ws.Cells(r + 2, 8).Value = "ML1x"
        ws.Cells(r + 2, 9).Value = IIf(refML.Exists(rName), refML(rName), 0)
        r = r + 3
    Next rName

End Sub


Private Sub PrintGeoSingleWithGap(ByVal ws As Worksheet, ByVal countryList As Variant, ByVal countryM1 As Object)

    Dim r As Long, cName As Variant

    r = 2
    ws.Cells(1, 5).Value = "Страна"
    ws.Range("E1:F1").Font.Bold = True

    For Each cName In countryList
        ws.Cells(r, 5).Value = cName
        ws.Cells(r, 6).Value = IIf(countryM1.Exists(cName), countryM1(cName), 0)
        r = r + 2
    Next cName

End Sub


Private Sub PrintGeoSMP(ByVal ws As Worksheet, ByVal countryList As Variant, ByVal countryM1 As Object)

    Dim r As Long, cName As Variant

    r = 2
    ws.Cells(1, 5).Value = "Страна"
    ws.Range("E1:F1").Font.Bold = True

    For Each cName In countryList
        ws.Cells(r, 5).Value = GetSmpCountryDisplayName(CStr(cName))
        ws.Cells(r, 6).Value = IIf(countryM1.Exists(cName), countryM1(cName), 0)
        r = r + 2
    Next cName

End Sub


Private Function GetSmpCountryDisplayName(ByVal countryName As String) As String

    Select Case LCase(Trim(countryName))
        Case "egypt"
            GetSmpCountryDisplayName = "Египет"
        Case "mauritania"
            GetSmpCountryDisplayName = "Мавритания"
        Case "sudan"
            GetSmpCountryDisplayName = "Судан"
        Case "dominican republic"
            GetSmpCountryDisplayName = "Доминиканская Республика"
        Case "honduras"
            GetSmpCountryDisplayName = "Гондурас"
        Case "nicaragua"
            GetSmpCountryDisplayName = "Никарагуа"
        Case "south sudan"
            GetSmpCountryDisplayName = "Южный Судан"
        Case "paraguay"
            GetSmpCountryDisplayName = "Парагвай"
        Case "qatar"
            GetSmpCountryDisplayName = "Катар"
        Case "united arab emirates"
            GetSmpCountryDisplayName = "ОАЭ"
        Case "panama"
            GetSmpCountryDisplayName = "Панама"
        Case "canada"
            GetSmpCountryDisplayName = "Канада"
        Case "guatemala"
            GetSmpCountryDisplayName = "Гватемала"
        Case Else
            GetSmpCountryDisplayName = countryName
    End Select

End Function


Private Sub PrintRefSingleWithGap(ByVal ws As Worksheet, ByVal refList As Variant, ByVal refM1 As Object)

    Dim r As Long, rName As Variant

    r = 2
    ws.Cells(1, 8).Value = "Реферал"
    ws.Range("H1:I1").Font.Bold = True

    For Each rName In refList
        ws.Cells(r, 8).Value = rName
        ws.Cells(r, 9).Value = IIf(refM1.Exists(rName), refM1(rName), 0)
        r = r + 2
    Next rName

End Sub


Private Function GetHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long

    Dim lastCol As Long, c As Long
    Dim currentHeader As String

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol

        currentHeader = Trim(CStr(ws.Cells(1, c).Value))
        currentHeader = Replace(currentHeader, Chr(160), " ")
        currentHeader = Replace(currentHeader, vbCr, "")
        currentHeader = Replace(currentHeader, vbLf, "")

        If LCase(currentHeader) = LCase(headerName) Then
            GetHeaderCol = c
            Exit Function
        End If

    Next c

    GetHeaderCol = 0

End Function