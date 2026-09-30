Option Explicit

Sub Monitoring_Выводы()

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsSource As Worksheet: Set wsSource = ActiveSheet

    If wsSource.name = "Результат" Or wsSource.name = "Фильтр" Then
        MsgBox "Запусти макрос с исходного листа выгрузки, а не с итогового листа.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Worksheets("Результат").Delete
    wb.Worksheets("Фильтр").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Dim colExtStatus As Long: colExtStatus = GetHeaderCol(wsSource, "External Status")
    Dim colTicketType As Long: colTicketType = GetHeaderCol(wsSource, "Ticket type")
    Dim colAgent As Long: colAgent = GetHeaderCol(wsSource, "Agent")
    Dim colCountry As Long: colCountry = GetHeaderCol(wsSource, "Country")
    Dim colDept As Long: colDept = GetHeaderCol(wsSource, "Department")
    Dim colTicketID As Long: colTicketID = GetHeaderCol(wsSource, "Ticket ID")

    If colExtStatus = 0 Or _
       colTicketType = 0 Or _
       colAgent = 0 Or _
       colCountry = 0 Or _
       colDept = 0 Or _
       colTicketID = 0 Then

        MsgBox "Не найдены нужные колонки: External Status / Ticket type / Agent / Country / Department / Ticket ID.", vbCritical

        Application.ScreenUpdating = True

        Exit Sub

    End If


    Dim wsResult As Worksheet

    Set wsResult = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count) _
    )

    wsResult.name = "Результат"


    Dim wsFilter As Worksheet

    Set wsFilter = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count) _
    )

    wsFilter.name = "Фильтр"


    Dim headers As Variant

    headers = Array( _
        "Department", _
        "Ticket type", _
        "External Status", _
        "From", _
        "Transaction Status", _
        "Transaction ID", _
        "Check amount", _
        "Subagent", _
        "Payment date", _
        "", _
        "Ticket ID", _
        "Country", _
        "Processing time", _
        "Internal comment", _
        "Agent" _
    )


    Dim h As Long
    Dim colTxStatusOut As Long


    For h = LBound(headers) To UBound(headers)

        wsFilter.Cells(1, h + 1).value = headers(h)

        If headers(h) = "Transaction Status" Then
            colTxStatusOut = h + 1
        End If

    Next h


    Dim bt238 As Long
    Dim bt205 As Long
    Dim bt175 As Long
    Dim bt185 As Long
    Dim bt191 As Long
    Dim bt199 As Long
    Dim bt203 As Long

    Dim smp81 As Long
    Dim smp53 As Long
    Dim smp38 As Long


    Dim lastRow As Long

    lastRow = wsSource.Cells( _
        wsSource.Rows.Count, _
        1 _
    ).End(xlUp).Row


    Dim i As Long
    Dim outRow As Long

    outRow = 2


    For i = 2 To lastRow


        Dim dept As String

        dept = NormalizeText( _
            wsSource.Cells(i, colDept).value _
        )


        If dept = "mena 1x" Or _
           dept = "mena leads 1x" Then


            Dim tType As String
            Dim agent As String
            Dim status As String
            Dim country As String


            tType = NormalizeText( _
                wsSource.Cells(i, colTicketType).value _
            )


            agent = NormalizeText( _
                wsSource.Cells(i, colAgent).value _
            )


            status = NormalizeText( _
                wsSource.Cells(i, colExtStatus).value _
            )


            country = NormalizeText( _
                wsSource.Cells(i, colCountry).value _
            )


            Dim isBTM As Boolean
            Dim isSMP As Boolean


            isBTM = ( _
                tType = "bt m" And _
                InStr(agent, "banktransfer") > 0 _
            )


            isSMP = ( _
                tType = "smp m" And _
                InStr(agent, "sendmepay") > 0 _
            )


            Dim addToFilter As Boolean

            addToFilter = False


            ' =================================================
            ' BT M
            ' =================================================

            If isBTM Then


                ' Актуальное: Review required (M)
                ' Старое название оставлено для совместимости.
                If status = "review required (m)" Or _
                   status = "revision needed (m)" Then

                    bt238 = bt238 + 1
                    addToFilter = True

                End If


                ' Актуальное: Money not sent, cancel (M)
                ' Старое название оставлено для совместимости.
                If status = "money not sent, cancel (m)" Or _
                   status = "the money has not been sent, cancel it (m)" Then

                    bt205 = bt205 + 1
                    addToFilter = True

                End If


                ' Актуальное: Adjust payment amount (M)
                ' Старое название оставлено для совместимости.
                If status = "adjust payment amount (m)" Or _
                   status = "adjust the payout amount (m)" Then

                    bt175 = bt175 + 1
                    addToFilter = True

                End If


                If IsNeededGeo(country) Then


                    If status = _
                        "limit reached on the recipient side (m)" Then

                        bt185 = bt185 + 1
                        addToFilter = True

                    End If


                    ' Актуальное: Recipient details incorrect (M)
                    ' Старое название оставлено для совместимости.
                    If status = "recipient details incorrect (m)" Or _
                       status = "recipient's details are not correct (m)" Then

                        bt191 = bt191 + 1
                        addToFilter = True

                    End If


                    ' Актуальное: Request for payment statement (M)
                    ' Старое название оставлено для совместимости.
                    If status = "request for payment statement (m)" Or _
                       status = "request statement for payout (m)" Then

                        bt199 = bt199 + 1
                        addToFilter = True

                    End If


                    If status = "sent (m)" Then

                        bt203 = bt203 + 1
                        addToFilter = True

                    End If


                End If


            ' =================================================
            ' SMP M
            ' =================================================

            ElseIf isSMP Then


                If status = "revision needed" Then

                    smp81 = smp81 + 1
                    addToFilter = True

                End If


                ' Актуальное название:
                ' Money not sent, cancel
                '
                ' Старое оставлено для совместимости.

                If status = "money not sent, cancel" Or _
                   status = "the money has not been sent, cancel it" Then

                    smp53 = smp53 + 1
                    addToFilter = True

                End If


                If status = "adjust the payout amount" Then

                    smp38 = smp38 + 1
                    addToFilter = True

                End If


            End If


            ' =================================================
            ' ЛИСТ ФИЛЬТР
            ' =================================================

            If addToFilter Then


                Dim sourceCol As Long


                For h = LBound(headers) To UBound(headers)


                    If headers(h) <> "" Then


                        sourceCol = GetHeaderCol( _
                            wsSource, _
                            CStr(headers(h)) _
                        )


                        If sourceCol > 0 Then


                            wsFilter.Cells( _
                                outRow, _
                                h + 1 _
                            ).value = _
                                wsSource.Cells( _
                                    i, _
                                    sourceCol _
                                ).value


                        End If


                    End If


                Next h


                outRow = outRow + 1


            End If


        End If


    Next i


    ' ========================================================
    ' ОБРАБОТКА ЛИСТА ФИЛЬТР
    ' ========================================================

    If outRow > 2 Then


        NormalizeFromColumn _
            wsFilter, _
            "From"


        NormalizeDateColumn _
            wsFilter, _
            "Payment date"


    Dim helperCol As Long
    Dim rejectedCount As Long

        helperCol = UBound(headers) + 2


        wsFilter.Cells(1, helperCol).value = _
            "Sort_Rejected"


        For i = 2 To outRow - 1


            If LCase( _
                Trim( _
                    CStr( _
                        wsFilter.Cells( _
                            i, _
                            colTxStatusOut _
                        ).value _
                    ) _
                ) _
            ) = "rejected" Then


                wsFilter.Cells(i, helperCol).value = 1
                rejectedCount = rejectedCount + 1


            Else


                wsFilter.Cells(i, helperCol).value = 0


            End If


        Next i


        wsFilter.Sort.SortFields.Clear


        wsFilter.Sort.SortFields.Add _
            key:=wsFilter.Range( _
                wsFilter.Cells(2, helperCol), _
                wsFilter.Cells(outRow - 1, helperCol) _
            ), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal


        With wsFilter.Sort


            .SetRange wsFilter.Range( _
                wsFilter.Cells(1, 1), _
                wsFilter.Cells(outRow - 1, helperCol) _
            )


            .Header = xlYes

            .Apply


        End With


        wsFilter.Columns(helperCol).Delete


    End If


    With wsFilter.Rows(1)


        .Font.Bold = True

        .Interior.Color = _
            RGB(230, 230, 230)

        .AutoFilter


    End With


    wsFilter.Columns.AutoFit

    wsFilter.Columns(10).ColumnWidth = 3

    wsFilter.Columns(10).Interior.Color = _
        RGB(242, 242, 242)

    wsFilter.Columns(11).ColumnWidth = 24

    ' ========================================================
    ' ФОРМАТИРОВАНИЕ ДАННЫХ
    ' Повторяем эффект двойного переключения «Переносить текст».
    ' Простое присваивание False не пересчитывает высоту, если
    ' перенос уже был выключен, поэтому сначала включаем его.
    ' ========================================================

    If outRow > 2 Then

        With wsFilter.Range( _
            wsFilter.Cells(1, 1), _
            wsFilter.Cells(outRow - 1, UBound(headers) + 1) _
        )

            .WrapText = True
            .WrapText = False
            .VerticalAlignment = xlTop

        End With

        ' Сбрасываем высоту каждой строки отдельно. Это не даёт
        ' многострочным Internal comment растягивать строки.

        For i = 1 To outRow - 1
            wsFilter.Rows(i).RowHeight = 15
        Next i

    End If

    AddWithdrawalCopyButton _
        wsFilter, _
        "btnCopyRejectedTicketIDs", _
        wsFilter.Range("K1"), _
        IIf(rejectedCount > 0, "K2:K" & rejectedCount + 1, "K2"), _
        RGB(31, 78, 121)


    ' ========================================================
    ' ЛИСТ РЕЗУЛЬТАТ
    ' ========================================================

    With wsResult


        ' ====================================================
        ' BT M
        '
        ' Строка 1 только под кнопку копирования.
        ' ====================================================

        .Rows(1).RowHeight = 24


        .Range("A2:B2").Merge

        .Range("A2").value = "BT M"

        .Range("A2").Font.Bold = True

        .Range("A2").HorizontalAlignment = _
            xlCenter


        .Range("A3").value = "Статус"

        .Range("B3").value = "Кол-во"

        .Range("A3:B3").Font.Bold = True


        ' Общая сумма остаётся для контроля.
        ' Кнопкой НЕ копируется.

        .Range("A4").value = _
            "Суммарное кол-во Выводы"


        .Range("B4").value = _
            bt238 + _
            bt205 + _
            bt175 + _
            bt185 + _
            bt191 + _
            bt199 + _
            bt203


        .Range("A5").value = _
            "238 Review required (M)"

        .Range("B5").value = bt238


        .Range("A6").value = _
            "205 Money not sent, cancel (M)"

        .Range("B6").value = bt205


        .Range("A7").value = _
            "175 Adjust payment amount (M)"

        .Range("B7").value = bt175


        .Range("A8").value = _
            "185 Limit reached on the recipient side (M)"

        .Range("B8").value = bt185


        .Range("A9").value = _
            "191 Recipient details incorrect (M)"

        .Range("B9").value = bt191


        .Range("A10").value = _
            "199 Request for payment statement (M)"

        .Range("B10").value = bt199


        .Range("A11").value = _
            "203 Sent (M)"

        .Range("B11").value = bt203


        ' ====================================================
        ' SMP M
        '
        ' Строка 12 пустая.
        ' Строка 13 только под кнопку копирования.
        ' ====================================================

        .Rows(13).RowHeight = 24


        .Range("A14:B14").Merge

        .Range("A14").value = "SMP M"

        .Range("A14").Font.Bold = True

        .Range("A14").HorizontalAlignment = _
            xlCenter


        .Range("A15").value = "Статус"

        .Range("B15").value = "Кол-во"

        .Range("A15:B15").Font.Bold = True


        ' Общая сумма остаётся для контроля.
        ' Кнопкой НЕ копируется.

        .Range("A16").value = _
            "Суммарное кол-во Выводы"


        .Range("B16").value = _
            smp81 + _
            smp53 + _
            smp38


        .Range("A17").value = _
            "81 Revision needed"

        .Range("B17").value = smp81


        .Range("A18").value = _
            "53 Money not sent, cancel"

        .Range("B18").value = smp53


        .Range("A19").value = _
            "38 Adjust the payout amount"

        .Range("B19").value = smp38


        ' ====================================================
        ' ГРАНИЦЫ И ЦВЕТА
        ' ====================================================

        .Range("A2:B11").Borders.LineStyle = _
            xlContinuous


        .Range("A14:B19").Borders.LineStyle = _
            xlContinuous


        .Range("A2:B3").Interior.Color = _
            RGB(221, 235, 247)


        .Range("A14:B15").Interior.Color = _
            RGB(221, 235, 247)


        .Range("A4:B4").Interior.Color = _
            RGB(197, 217, 241)


        .Range("A16:B16").Interior.Color = _
            RGB(197, 217, 241)


        .Columns("A:B").AutoFit


        ' Чтобы кнопки были одинакового размера.

        .Columns("B").ColumnWidth = 14


    End With


    ' ========================================================
    ' КНОПКИ КОПИРОВАНИЯ
    ' ========================================================

    AddWithdrawalCopyButtons wsResult


    Application.ScreenUpdating = True


End Sub


' ============================================================
' КНОПКИ
'
' BT M:
' B5:B11
'
' SMP M:
' B17:B19
'
' B4 и B16 являются общими суммами и НЕ копируются.
' ============================================================

Private Sub AddWithdrawalCopyButtons( _
    ByVal ws As Worksheet _
)


    Dim buttonColor As Long

    buttonColor = RGB(31, 78, 121)


    On Error Resume Next


    ws.Shapes("btnCopyBTM").Delete

    ws.Shapes("btnCopySMPM").Delete
    ws.Shapes("btnCopyRejectedTicketIDs").Delete


    On Error GoTo 0


    ' ========================================================
    ' BT M
    ' ========================================================

    AddWithdrawalCopyButton _
        ws, _
        "btnCopyBTM", _
        ws.Range("B1"), _
        "B5:B11", _
        buttonColor


    ' ========================================================
    ' SMP M
    ' ========================================================

    AddWithdrawalCopyButton _
        ws, _
        "btnCopySMPM", _
        ws.Range("B13"), _
        "B17:B19", _
        buttonColor


End Sub


' ============================================================
' СОЗДАНИЕ ОДНОЙ КНОПКИ
'
' Использует ту же рабочую схему,
' что и твои остальные кнопки.
' ============================================================

Private Sub AddWithdrawalCopyButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal buttonCell As Range, _
    ByVal targetAddress As String, _
    ByVal buttonColor As Long _
)


    Dim buttonShape As Shape

    Dim macroWorkbookName As String

    Dim buttonLeft As Double
    Dim buttonTop As Double

    Dim buttonWidth As Double
    Dim buttonHeight As Double


    macroWorkbookName = _
        Replace( _
            ThisWorkbook.name, _
            "'", _
            "''" _
        )


    buttonWidth = _
        buttonCell.Width - 4


    buttonHeight = _
        buttonCell.Height - 4


    buttonLeft = _
        buttonCell.Left + 2


    buttonTop = _
        buttonCell.Top + 2


    Set buttonShape = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            buttonLeft, _
            buttonTop, _
            buttonWidth, _
            buttonHeight _
        )


    With buttonShape


        .name = buttonName


        ' Сохраняем адрес диапазона
        ' непосредственно внутри кнопки.

        .AlternativeText = _
            targetAddress


        ' Макрос вызывается именно из книги,
        ' где хранится этот VBA-код.
        '
        ' Для твоего случая это PERSONAL.XLSB.

        .OnAction = _
            "'" & _
            macroWorkbookName & _
            "'!Monitoring_Выводы_CopyBlock"


        ' ====================================================
        ' ЕДИНЫЙ СТИЛЬ
        ' ====================================================

        .Fill.ForeColor.RGB = _
            buttonColor


        .Line.ForeColor.RGB = _
            buttonColor


        .TextFrame2.TextRange.text = _
            "Копировать"


        .TextFrame2.MarginLeft = 0

        .TextFrame2.MarginRight = 0

        .TextFrame2.MarginTop = 0

        .TextFrame2.MarginBottom = 0


        .TextFrame2.TextRange.Font.name = _
            "Arial"


        .TextFrame2.TextRange.Font.Size = _
            7


        .TextFrame2.TextRange.Font.Bold = _
            msoTrue


        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)


        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle


        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter


        .Placement = _
            xlMoveAndSize


    End With


End Sub


' ============================================================
' ОБРАБОТЧИК ОБЕИХ КНОПОК
'
' Кнопка BT M копирует B5:B11.
' Кнопка SMP M копирует B17:B19.
'
' Общие суммы сюда попасть не могут.
' ============================================================

Public Function Monitoring_Выводы_CopyBlock() As Variant
    Dim sourceSheet As Worksheet

    Dim buttonShape As Shape

    Dim targetRange As Range


    On Error GoTo CopyFail


    Set sourceSheet = _
        ActiveSheet


    Set buttonShape = _
        sourceSheet.Shapes( _
            CStr(Application.Caller) _
        )


    Set targetRange = _
        sourceSheet.Range( _
            buttonShape.AlternativeText _
        )


    targetRange.Copy


    Application.StatusBar = _
        "Данные скопированы: " & _
        sourceSheet.name & _
        "!" & _
        targetRange.Address(False, False)


    Exit Function


CopyFail:


    Application.StatusBar = False


    MsgBox _
        "Не удалось скопировать данные: " & _
        Err.Description, _
        vbExclamation, _
        "Monitoring Выводы"


End Function


' ============================================================
' НУЖНЫЕ GEO
' ============================================================

Private Function IsNeededGeo( _
    ByVal countryValue As String _
) As Boolean


    Select Case countryValue


        Case _
            "azerbaijan", _
            "afghanistan", _
            "bolivia", _
            "guatemala", _
            "honduras", _
            "iran", _
            "canada", _
            "nicaragua", _
            "panama", _
            "paraguay", _
            "taiwan", _
            "kyrgyzstan", _
            "jamaica"


            IsNeededGeo = True


        Case Else


            IsNeededGeo = False


    End Select


End Function


' ============================================================
' НОРМАЛИЗАЦИЯ ТЕКСТА
' ============================================================

Private Function NormalizeText( _
    ByVal textValue As String _
) As String


    Dim v As String


    v = CStr(textValue)


    v = Replace( _
        v, _
        Chr(160), _
        " " _
    )


    v = Replace( _
        v, _
        vbCr, _
        " " _
    )


    v = Replace( _
        v, _
        vbLf, _
        " " _
    )


    v = Trim(v)


    Do While InStr(v, "  ") > 0


        v = Replace( _
            v, _
            "  ", _
            " " _
        )


    Loop


    NormalizeText = LCase(v)


End Function


' ============================================================
' FROM
' ============================================================

Private Sub NormalizeFromColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String _
)


    Dim col As Long


    col = GetHeaderCol( _
        ws, _
        headerName _
    )


    If col = 0 Then Exit Sub


    Dim lastRow As Long


    lastRow = ws.Cells( _
        ws.Rows.Count, _
        1 _
    ).End(xlUp).Row


    If lastRow < 2 Then Exit Sub


    Dim rng As Range


    Set rng = ws.Range( _
        ws.Cells(2, col), _
        ws.Cells(lastRow, col) _
    )


    rng.Replace _
        "Through Customer Support", _
        "L1", _
        xlWhole


    rng.Replace _
        "By yourself", _
        "User", _
        xlWhole


End Sub


' ============================================================
' PAYMENT DATE
' ============================================================

Private Sub NormalizeDateColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String _
)


    Dim col As Long


    col = GetHeaderCol( _
        ws, _
        headerName _
    )


    If col = 0 Then Exit Sub


    Dim lastRow As Long


    lastRow = ws.Cells( _
        ws.Rows.Count, _
        1 _
    ).End(xlUp).Row


    If lastRow < 2 Then Exit Sub


    Dim cell As Range


    For Each cell In ws.Range( _
        ws.Cells(2, col), _
        ws.Cells(lastRow, col) _
    )


        If Not IsEmpty(cell.value) Then


            On Error Resume Next


            If InStr( _
                CStr(cell.value), _
                "-" _
            ) > 0 Then


                cell.value = CDate( _
                    Replace( _
                        Left( _
                            CStr(cell.value), _
                            10 _
                        ), _
                        "-", _
                        "/" _
                    ) _
                )


            ElseIf IsDate(cell.value) Then


                cell.value = _
                    CDate(cell.value)


            End If


            On Error GoTo 0


        End If


    Next cell


    ws.Range( _
        ws.Cells(2, col), _
        ws.Cells(lastRow, col) _
    ).NumberFormat = _
        "dd.mm.yyyy"


End Sub


' ============================================================
' ПОИСК КОЛОНКИ
' ============================================================

Private Function GetHeaderCol( _
    ByVal ws As Worksheet, _
    ByVal headerName As String _
) As Long


    Dim lastCol As Long


    lastCol = ws.Cells( _
        1, _
        ws.Columns.Count _
    ).End(xlToLeft).Column


    Dim c As Long

    Dim currentHeader As String


    For c = 1 To lastCol


        currentHeader = _
            CStr(ws.Cells(1, c).value)


        currentHeader = Replace( _
            currentHeader, _
            Chr(160), _
            " " _
        )


        currentHeader = Replace( _
            currentHeader, _
            vbCr, _
            "" _
        )


        currentHeader = Replace( _
            currentHeader, _
            vbLf, _
            "" _
        )


        If LCase( _
            Trim(currentHeader) _
        ) = LCase(headerName) Then


            GetHeaderCol = c

            Exit Function


        End If


    Next c


    GetHeaderCol = 0


End Function
