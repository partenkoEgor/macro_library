Option Explicit

' В выгрузке колонка A "Ответственность" (Не L2 / L2 / Grand Total),
' из-за этого все остальные колонки сдвинуты на 1 вправо:
'   A=Ответственность, B=Нач.статус, C=Кон.статус, D=признак строки, E..=значения по дням
'
' Учитываются только строки блока "L2". Блок "Не L2" в подсчёт не идёт.
'
' Результат - сетка по дням месяца, а не один общий итог: колонка B = день 1,
' последняя колонка = вчерашний день. Сегодняшний день не считается, его данные
' в выгрузке ещё не полные. Ширина сетки не фиксирована и не зависит от того,
' сколько колонок реально есть в выгрузке - она считается от сегодняшней даты
' (MAX_DAYS = число сегодня минус 1), поэтому автоматически адаптируется под
' любой день месяца без лишних нулевых колонок.
' День берётся не по номеру колонки, а по значению в строке 4 исходной выгрузки
' (DAY_HEADER_ROW), на случай если пивот когда-нибудь пропустит день.
' Лист каждый раз пересобирается заново из активной выгрузки, поэтому его можно
' перезапускать в течение месяца на растущей month-to-date выгрузке.
'
' 3 кнопки "Копировать" (по одной на раздел) копируют весь блок дней своего раздела.
' Кнопка - это Shape с адресом диапазона в .AlternativeText, обработчик - Public Function
' (не Sub), поэтому в списке макросов (Alt+F8) он не отображается.

Private Const COL_RESP As Long = 1        ' Ответственность: Не L2 / L2 / Grand Total
Private Const COL_INIT As Long = 2        ' Начальный статус
Private Const COL_FINAL As Long = 3       ' Конечный статус
Private Const COL_ROWTYPE As Long = 4     ' "Тикеты или взаимодействия" / "Доля от общего количества..."
Private Const COL_VALUE_FIRST As Long = 5 ' первая колонка со значениями (по одной на день месяца)
Private Const DAY_HEADER_ROW As Long = 4  ' строка с номерами дней (1, 2, 3...) над колонками значений

' Не Const: считается один раз в начале Sub от текущей даты (сегодня минус 1),
' дальше используется во всех функциях листа как ширина сетки дней.
Private MAX_DAYS As Long

Sub Monitoring_Перемещение_тикетов()

    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim dict As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, c As Long, d As Long
    Dim curInit As String, curFinal As String, curResp As String
    Dim outRow As Long
    Dim dayNum As Variant
    Dim cellVal As Variant

    Dim teamBHeaderRow As Long, teamBSubHeaderRow As Long, teamBDataStart As Long, teamBDataEnd As Long
    Dim apiHeaderRow As Long, apiDataStart As Long, apiDataEnd As Long
    Dim pspHeaderRow As Long, pspDataStart As Long, pspDataEnd As Long

    Set wsSrc = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    ' Сегодняшний день неполный, считаем по вчерашний включительно.
    MAX_DAYS = Day(Date) - 1

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    Worksheets("Маппинг_Цепочек").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Set wsOut = Worksheets.Add(After:=wsSrc)
    wsOut.Name = "Маппинг_Цепочек"

    lastRow = wsSrc.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = wsSrc.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    curInit = ""
    curFinal = ""
    curResp = ""

    For i = 1 To lastRow

        ' Строка "Grand Total" (итог по всей выгрузке) не относится ни к одной цепочке -
        ' сбрасываем текущие значения, чтобы её данные не приплюсовались к последней цепочке
        If LCase(CleanText(wsSrc.Cells(i, COL_RESP).Value)) = "grand total" Then
            curInit = ""
            curFinal = ""
            curResp = ""
        End If

        If IsRealResp(wsSrc.Cells(i, COL_RESP).Value) Then
            curResp = CleanText(wsSrc.Cells(i, COL_RESP).Value)
        End If

        If IsRealStatus(wsSrc.Cells(i, COL_INIT).Value) Then
            curInit = NormalizeStatus(wsSrc.Cells(i, COL_INIT).Value)
        End If

        If IsRealStatus(wsSrc.Cells(i, COL_FINAL).Value) Then
            curFinal = NormalizeStatus(wsSrc.Cells(i, COL_FINAL).Value)
        End If

        If IsTicketRow(wsSrc.Cells(i, COL_ROWTYPE).Value) Then

            If curInit <> "" And curFinal <> "" And LCase(curResp) = "l2" Then

                For c = COL_VALUE_FIRST To lastCol

                    dayNum = wsSrc.Cells(DAY_HEADER_ROW, c).MergeArea.Cells(1, 1).Value

                    If IsNumeric(dayNum) Then
                        If CLng(dayNum) >= 1 And CLng(dayNum) <= MAX_DAYS Then

                            cellVal = wsSrc.Cells(i, c).Value

                            If IsNumeric(cellVal) Then
                                AddToDict dict, curInit, curFinal, CLng(dayNum), CLng(cellVal)
                            End If

                        End If
                    End If

                Next c

            End If

        End If

    Next i

    wsOut.Cells(1, 1).Value = "Цепочка"
    For d = 1 To MAX_DAYS
        wsOut.Cells(1, d + 1).Value = d
    Next d
    outRow = 2

    teamBHeaderRow = outRow
    AddHeader wsOut, outRow, "Обработка Team B (BT M)"
    teamBSubHeaderRow = outRow
    AddSubHeader wsOut, outRow, "Mena 1x, Mena Leads 1x"
    teamBDataStart = outRow

    AddChain dict, wsOut, outRow, "Received (M) - In Progress (M)", "Получено (M)", "В процессе (M)"
    AddChain dict, wsOut, outRow, "Received (Fraud) (M) - In Progress (M)", "Получено (Фрод) (M)", "В процессе (M)"
    AddChain dict, wsOut, outRow, "Approved (M) - In Progress (M)", "Одобрено (M)", "В процессе (M)"
    AddChain dict, wsOut, outRow, "CTAA (M) - In Progress (M)", "Зачислено на другой счет (M)", "В процессе (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Sent for processing (M)", "В процессе (M)", "Отправлено на обработку (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Credited (M)", "В процессе (M)", "Зачислено (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Credited (Fraud) (M)", "В процессе (M)", "Зачислено (Фрод) (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Create new transaction (M)", "В процессе (M)", "Создать новую транзакцию (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Approved by agent (M)", "В процессе (M)", "Одобрено агентом (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - CTTA by agent (M)", "В процессе (M)", "Зачислено на другой счет агентом (M)"
    AddChain dict, wsOut, outRow, "In Progress (M) - Adjusted the amount (d) (M)", "В процессе (M)", "Сумма скорректирована (Депозит) (M)"

    teamBDataEnd = outRow - 1

    apiHeaderRow = outRow
    AddHeader wsOut, outRow, "Обработка Team A,B (API)"
    apiDataStart = outRow

    AddChain dict, wsOut, outRow, "Ожидает ответ от ПС - Запрос информации", "Ожидание ответа от PS", "Запрос информации"
    AddChain dict, wsOut, outRow, "Ожидает ответ от ПС - Ответ игроку", "Ожидание ответа от PS", "Ответ клиенту"
    AddChain dict, wsOut, outRow, "Ожидает ответ от ПС - В работе", "Ожидание ответа от PS", "В процессе"
    AddChain dict, wsOut, outRow, "Ожидает ответ от ПС - Требуется помощь поддержки", "Ожидание ответа от PS", "Требуется помощь поддержки"

    apiDataEnd = outRow - 1

    pspHeaderRow = outRow
    AddHeader wsOut, outRow, "Обработка Team A,B (PSP)"
    pspDataStart = outRow

    AddChain dict, wsOut, outRow, "Проверка транзакции - Запрос информации", "Проверка транзакции", "Запрос информации"
    AddChain dict, wsOut, outRow, "Проверка транзакции - Ответ игроку", "Проверка транзакции", "Ответ клиенту"
    AddChain dict, wsOut, outRow, "Проверка транзакции - В работе", "Проверка транзакции", "В процессе"
    AddChain dict, wsOut, outRow, "Проверка транзакции - Требуется помощь поддержки", "Проверка транзакции", "Требуется помощь поддержки"

    pspDataEnd = outRow - 1

    FormatResult wsOut

    On Error Resume Next
    wsOut.Shapes("btnCopyTeamB").Delete
    wsOut.Shapes("btnCopyAPI").Delete
    wsOut.Shapes("btnCopyPSP").Delete
    On Error GoTo 0

    AddCopyButton wsOut, "btnCopyTeamB", teamBHeaderRow, teamBSubHeaderRow, teamBDataStart, teamBDataEnd
    AddCopyButton wsOut, "btnCopyAPI", apiHeaderRow, apiHeaderRow, apiDataStart, apiDataEnd
    AddCopyButton wsOut, "btnCopyPSP", pspHeaderRow, pspHeaderRow, pspDataStart, pspDataEnd

    Application.ScreenUpdating = True

    MsgBox "Информация по перемещениям тикетов сформирована." & vbCrLf & _
           "by Dilmurat", vbInformation, "Monitoring_Перемещение_тикетов"

End Sub

' ============================================================
' КНОПКА КОПИРОВАНИЯ (Shape + адрес в AlternativeText)
' Копирует весь блок B:AF (все дни) для строк своего раздела.
' ============================================================

Private Sub AddCopyButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal topRow As Long, ByVal bottomRow As Long, ByVal dataStartRow As Long, ByVal dataEndRow As Long)

    Dim buttonShape As Shape
    Dim anchorCell As Range
    Dim buttonLeft As Double, buttonTop As Double, buttonWidth As Double, buttonHeight As Double
    Dim targetAddress As String

    ws.Rows(topRow).RowHeight = 20
    If bottomRow <> topRow Then ws.Rows(bottomRow).RowHeight = 20

    Set anchorCell = ws.Cells(topRow, 2)

    ' Фиксированный размер кнопки - колонки с днями теперь узкие (8 единиц),
    ' и кнопка по их ширине/высоте секции выглядела растянутой и обрезанной.
    buttonLeft = anchorCell.Left + 2
    buttonTop = anchorCell.Top + 2
    buttonWidth = 55
    buttonHeight = 18

    targetAddress = ws.Range(ws.Cells(dataStartRow, 2), ws.Cells(dataEndRow, MAX_DAYS + 1)).Address(False, False)

    Set buttonShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)

    With buttonShape

        .Name = buttonName
        .AlternativeText = targetAddress

        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!Monitoring_Перемещение_тикетов_CopyBlock"

        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(31, 78, 121)

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

' ============================================================
' ОБРАБОТЧИК ВСЕХ КНОПОК КОПИРОВАНИЯ
' Function, а не Sub - не появляется в списке макросов (Alt+F8)
' ============================================================

Public Function Monitoring_Перемещение_тикетов_CopyBlock() As Variant

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
    MsgBox "Не удалось скопировать данные: " & Err.Description, vbExclamation, "Monitoring_Перемещение_тикетов"

End Function

Private Sub AddToDict(ByVal dict As Object, ByVal initStatus As String, ByVal finalStatus As String, ByVal dayNum As Long, ByVal valueTickets As Long)

    Dim key As String
    Dim dayDict As Object

    key = MakeKey(initStatus, finalStatus)

    If dict.Exists(key) Then
        Set dayDict = dict(key)
    Else
        Set dayDict = CreateObject("Scripting.Dictionary")
        dict.Add key, dayDict
    End If

    If dayDict.Exists(dayNum) Then
        dayDict(dayNum) = CLng(dayDict(dayNum)) + valueTickets
    Else
        dayDict.Add dayNum, valueTickets
    End If

End Sub

Private Sub AddChain(ByVal dict As Object, ByVal ws As Worksheet, ByRef r As Long, ByVal displayText As String, ByVal initStatus As String, ByVal finalStatus As String)

    Dim key As String
    Dim dayDict As Object
    Dim d As Long

    key = MakeKey(initStatus, finalStatus)

    ws.Cells(r, 1).Value = displayText

    If dict.Exists(key) Then

        Set dayDict = dict(key)

        For d = 1 To MAX_DAYS
            If dayDict.Exists(d) Then
                ws.Cells(r, d + 1).Value = CLng(dayDict(d))
            Else
                ws.Cells(r, d + 1).Value = 0
            End If
        Next d

    Else

        For d = 1 To MAX_DAYS
            ws.Cells(r, d + 1).Value = 0
        Next d

    End If

    r = r + 1

End Sub

Private Function MakeKey(ByVal initStatus As String, ByVal finalStatus As String) As String
    MakeKey = NormalizeKey(initStatus) & "||" & NormalizeKey(finalStatus)
End Function

Private Function NormalizeStatus(ByVal v As Variant) As String

    Dim s As String
    s = CleanText(v)

    Select Case s

        Case "В процессе (M)", "In Progress (M)"
            s = "В процессе (M)"

        Case "Получено (M)", "Received (M)"
            s = "Получено (M)"

        Case "Получено (Фрод) (M)", "Received (Fraud) (M)"
            s = "Получено (Фрод) (M)"

        Case "Одобрено (M)", "Approved (M)"
            s = "Одобрено (M)"

        Case "Зачислено на другой счет (M)", "CTAA (M)", "Credited to another account (M)"
            s = "Зачислено на другой счет (M)"

        Case "Отправлено на обработку (M)", "Sent for processing (M)"
            s = "Отправлено на обработку (M)"

        Case "Зачислено (M)", "Credited (M)"
            s = "Зачислено (M)"

        Case "Зачислено (Фрод) (M)", "Credited (Fraud) (M)"
            s = "Зачислено (Фрод) (M)"

        Case "Создать новую транзакцию (M)", "Create new transaction (M)"
            s = "Создать новую транзакцию (M)"

        Case "Одобрено агентом (M)", "Approved by agent (M)"
            s = "Одобрено агентом (M)"

        Case "Зачислено на другой счет агентом (M)", "CTTA by agent (M)", "CTAA by agent (M)", "Credited to another account by the agent (M)"
            s = "Зачислено на другой счет агентом (M)"

        Case "Сумма скорректирована (Депозит) (M)", "Adjusted the amount (d) (M)"
            s = "Сумма скорректирована (Депозит) (M)"

        Case "Ожидание ответа от PS", "Ожидает ответ от ПС", "Ожидание ответа от ПС", "Waiting for PS response"
            s = "Ожидание ответа от PS"

        Case "Проверка транзакции", "Transaction verification"
            s = "Проверка транзакции"

        Case "Запрос информации", "Request information"
            s = "Запрос информации"

        Case "Ответ клиенту", "Ответ игроку", "Reply to player", "Reply to client"
            s = "Ответ клиенту"

        Case "В процессе", "В работе", "In Progress"
            s = "В процессе"

        Case "Требуется помощь поддержки", "Support help required"
            s = "Требуется помощь поддержки"

    End Select

    NormalizeStatus = s

End Function

Private Function NormalizeKey(ByVal v As Variant) As String

    Dim s As String
    s = NormalizeStatus(v)

    s = LCase(Trim(s))
    s = Replace(s, "ё", "е")
    s = Replace(s, "'", "'")
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, "ctta", "ctaa")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeKey = s

End Function

Private Function CleanText(ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CleanText = ""
        Exit Function
    End If

    s = Trim(CStr(v))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, "ё", "е")
    s = Replace(s, "Ё", "Е")
    s = Replace(s, "(М)", "(M)")
    s = Replace(s, "(м)", "(M)")
    s = Replace(s, "'", "'")
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanText = Trim(s)

End Function

Private Function IsTicketRow(ByVal v As Variant) As Boolean

    Dim s As String
    s = LCase(CleanText(v))

    If InStr(1, s, "тикет", vbTextCompare) > 0 Or _
       InStr(1, s, "взаимодейств", vbTextCompare) > 0 Or _
       InStr(1, s, "ticket", vbTextCompare) > 0 Then
        IsTicketRow = True
    End If

End Function

Private Function IsRealStatus(ByVal v As Variant) As Boolean

    Dim s As String
    s = CleanText(v)

    If s = "" Then Exit Function
    If LCase(s) = "nan" Then Exit Function
    If IsNumeric(s) Then Exit Function
    If InStr(1, s, "Доля", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "Тикеты", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "взаимодействия", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "Начальный статус", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "Конечный статус", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "Grand Total", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "Дата", vbTextCompare) > 0 Then Exit Function

    IsRealStatus = True

End Function

Private Function IsRealResp(ByVal v As Variant) As Boolean

    Dim s As String
    s = LCase(CleanText(v))

    If s = "" Then Exit Function
    If s = "grand total" Then Exit Function

    IsRealResp = True

End Function

Private Sub AddHeader(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String)

    If r > 2 Then r = r + 1

    ws.Cells(r, 1).Value = title

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_DAYS + 1))
        .Font.Bold = True
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
    End With

    r = r + 1

End Sub

Private Sub AddSubHeader(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String)

    ws.Cells(r, 1).Value = title

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, MAX_DAYS + 1))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
    End With

    r = r + 1

End Sub

Private Sub FormatResult(ByVal ws As Worksheet)

    With ws.UsedRange
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 220, 220)
    End With

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, MAX_DAYS + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(27, 54, 93)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 75
    ws.Columns("A").WrapText = True

    With ws.Range(ws.Cells(1, 2), ws.Cells(1, MAX_DAYS + 1)).EntireColumn
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
        .NumberFormat = "#,##0"
    End With

End Sub
