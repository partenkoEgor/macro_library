Option Explicit

' ============================================================
' БУФЕР ОБМЕНА WINDOWS
' Нужен для копирования двух несмежных колонок одним текстовым блоком.
' Используем Unicode API Windows вместо Forms.DataObject: он может быть
' недоступен в конкретной установке Office.
' ============================================================
#If VBA7 Then
    Private Declare PtrSafe Function AP3_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function AP3_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
    Private Declare PtrSafe Function AP3_EmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
    Private Declare PtrSafe Function AP3_SetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function AP3_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function AP3_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function AP3_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function AP3_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub AP3_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function AP3_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
    Private Declare Function AP3_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
    Private Declare Function AP3_EmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
    Private Declare Function AP3_SetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function AP3_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function AP3_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
    Private Declare Function AP3_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
    Private Declare Function AP3_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
    Private Declare Sub AP3_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const AP3_CF_UNICODETEXT As Long = 13
Private Const AP3_GMEM_MOVEABLE_ZEROINIT As Long = &H42

' ============================================================
' Approved_Перенос2_v3
' Версия 3.1.0
'
' Пара к Approved_Фильтр1_v3.
'
' Отличие от Approved_Перенос2: после сверки итог раскладывается
' на два листа по колонке "Источник номера".
'
'   "Из выгрузки"    - номер стоял в колонке Transaction ID
'   "Из комментария" - номер подобран из Internal comment
'
' Сама логика переноса и сверки не менялась.
'
' Все процедуры имеют префикс AP3_, а обработчики кнопок
' заканчиваются на _v3, чтобы модуль мог лежать в одной книге
' со старым Approved_Перенос2 без конфликта имён.
' ============================================================

Private Const AP3_FILTER_HEADER_ROW As Long = 2
Private Const AP3_FILTER_DATA_ROW As Long = 3
Private Const AP3_SOURCE_HEADER_ROW As Long = 1
Private Const AP3_SOURCE_DATA_ROW As Long = 2

' Имена итоговых листов. Меняются здесь, в одном месте.
Private Const AP3_SHEET_EXPORT As String = "Из выгрузки"
Private Const AP3_SHEET_COMMENT As String = "Из комментария"

' Значение в колонке "Источник номера" для строк,
' где номер пришёл из самой выгрузки.
Private Const AP3_TAG_EXPORT As String = "Выгрузка"

' Размеры кнопок интерфейса.
Private Const AP3_BUTTON_PITCH_Y As Double = 20
Private Const AP3_BUTTON_HEIGHT As Double = 18


Public Sub Approved_Перенос2_v3()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet

    Dim wsExport As Worksheet
    Dim wsComment As Worksheet
    Dim wsOpen As Worksheet

    Dim lastRowSrc As Long
    Dim lastRowDst As Long

    Dim srcCols As Object
    Dim dstCols As Object

    Dim r As Long
    Dim c As Long
    Dim i As Long

    Dim foundRow As Long
    Dim rDst As Long

    Dim key As String

    Dim requiredSrc As Variant
    Dim requiredDst As Variant
    Dim colNames As Variant

    Dim lastColSrc As Long
    Dim lastColDst As Long

    Dim srcData As Variant
    Dim dstData As Variant

    Dim rowsSrc As Long
    Dim rowsDst As Long
    Dim pairCount As Long
    Dim keptCount As Long
    Dim keepRow As Boolean
    Dim statusText As String
    Dim alignWarning As String

    Dim didSplit As Boolean
    Dim countExport As Long
    Dim countComment As Long
    Dim buttonWarning As String

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim completedOK As Boolean
    Dim errorText As String

    Set wb = ActiveWorkbook

    On Error Resume Next

    Set wsSrc = wb.Worksheets("Источник")
    Set wsDst = wb.Worksheets("Фильтр")

    On Error GoTo 0

    If wsSrc Is Nothing Then

        MsgBox "Не найден лист 'Источник'.", vbCritical
        Exit Sub

    End If

    If wsDst Is Nothing Then

        MsgBox "Не найден лист 'Фильтр'.", vbCritical
        Exit Sub

    End If

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts

    On Error GoTo FailHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Старые листы для копирования больше не используются.
    AP3_DeleteSheetIfExists wb, "Для копирования"
    AP3_DeleteSheetIfExists wb, "Закрыть статусом 225"

    ' ========================================================
    ' 1. ПЕРЕНОС ДАННЫХ АДМИНКИ
    '
    ' Логика не менялась.
    ' Источник: заголовки в строке 1.
    ' Фильтр: заголовки в строке 2, строка 1 занята кнопками.
    ' ========================================================

    Set srcCols = CreateObject("Scripting.Dictionary")

    lastColSrc = wsSrc.Cells( _
        AP3_SOURCE_HEADER_ROW, _
        wsSrc.Columns.Count _
    ).End(xlToLeft).Column

    For c = 1 To lastColSrc

        srcCols( _
            Trim$(LCase$(CStr(wsSrc.Cells(AP3_SOURCE_HEADER_ROW, c).Value))) _
        ) = c

    Next c

    Set dstCols = CreateObject("Scripting.Dictionary")

    lastColDst = wsDst.Cells( _
        AP3_FILTER_HEADER_ROW, _
        wsDst.Columns.Count _
    ).End(xlToLeft).Column

    For c = 1 To lastColDst

        dstCols( _
            Trim$(LCase$(CStr(wsDst.Cells(AP3_FILTER_HEADER_ROW, c).Value))) _
        ) = c

    Next c

    requiredSrc = Array( _
        "номер заказа", _
        "сумма пополнения", _
        "статус", _
        "id пользователя", _
        "id субагента" _
    )

    For i = LBound(requiredSrc) To UBound(requiredSrc)

        If Not srcCols.Exists(requiredSrc(i)) Then

            MsgBox _
                "Не найдена колонка на листе Источник: " & _
                requiredSrc(i), _
                vbCritical

            GoTo CleanExit

        End If

    Next i

    requiredDst = Array( _
        "adm transaction id", _
        "adm check amount", _
        "transaction status", _
        "adm user id", _
        "adm subagent id" _
    )

    For i = LBound(requiredDst) To UBound(requiredDst)

        If Not dstCols.Exists(requiredDst(i)) Then

            lastColDst = lastColDst + 1

            wsDst.Cells( _
                AP3_FILTER_HEADER_ROW, _
                lastColDst _
            ).Value = requiredDst(i)

            dstCols(requiredDst(i)) = lastColDst

        End If

    Next i

    ' Колонки, приходящие от первого макроса. ADM-колонки выше
    ' уже гарантированы, эти надо проверить: без них обращения
    ' к массиву ниже упали бы на пустом индексе.
    colNames = Array( _
        "transaction status", _
        "adm transaction id", _
        "transaction id", _
        "adm user id", _
        "user id", _
        "adm check amount", _
        "check amount", _
        "adm subagent id", _
        "subagent id" _
    )

    For c = LBound(colNames) To UBound(colNames)

        If Not dstCols.Exists(colNames(c)) Then

            MsgBox _
                "Не найдена колонка на листе Фильтр: " & _
                colNames(c), _
                vbCritical

            GoTo CleanExit

        End If

    Next c

    ' Границу "Фильтра" берём по колонке 1: ADM Transaction ID
    ' до переноса пуста, и End(xlUp) по ней попадал бы в заголовок.
    lastRowSrc = wsSrc.Cells( _
        wsSrc.Rows.Count, _
        srcCols("номер заказа") _
    ).End(xlUp).Row

    lastRowDst = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row

    rowsSrc = lastRowSrc - AP3_SOURCE_DATA_ROW + 1
    rowsDst = lastRowDst - AP3_FILTER_DATA_ROW + 1

    If rowsSrc < 0 Then rowsSrc = 0
    If rowsDst < 0 Then rowsDst = 0

    ' Несовпадение числа строк означает, что админка вернула
    ' не столько, сколько ушло, и всё после пропуска съехало.
    ' Сверка ниже вычистит съехавшее, но пользователь должен
    ' понимать, почему результат неожиданно короткий.
    If rowsSrc <> rowsDst Then

        alignWarning = _
            "Строк в 'Источнике': " & CStr(rowsSrc) & _
            ", строк в 'Фильтре': " & CStr(rowsDst) & "."

    End If

    If rowsDst > 0 Then

        ' Оба листа читаются одним обращением, дальше всё в памяти.
        dstData = wsDst.Range( _
            wsDst.Cells(AP3_FILTER_DATA_ROW, 1), _
            wsDst.Cells(lastRowDst, lastColDst) _
        ).Value

        If rowsSrc > 0 Then

            srcData = wsSrc.Range( _
                wsSrc.Cells(AP3_SOURCE_DATA_ROW, 1), _
                wsSrc.Cells(lastRowSrc, lastColSrc) _
            ).Value

        End If

        ' ----------------------------------------------------
        ' Перенос: строка i "Источника" ложится на строку i
        ' "Фильтра". Так было и раньше, просто неявно - колонка
        ' ADM Transaction id пуста, поэтому поиск совпадения
        ' всегда промахивался и строка дописывалась по порядку.
        ' ----------------------------------------------------

        pairCount = rowsSrc

        If rowsDst < pairCount Then pairCount = rowsDst

        For i = 1 To pairCount

            dstData(i, dstCols("adm transaction id")) = _
                srcData(i, srcCols("номер заказа"))

            dstData(i, dstCols("adm check amount")) = _
                srcData(i, srcCols("сумма пополнения"))

            dstData(i, dstCols("transaction status")) = _
                srcData(i, srcCols("статус"))

            dstData(i, dstCols("adm user id")) = _
                srcData(i, srcCols("id пользователя"))

            dstData(i, dstCols("adm subagent id")) = _
                srcData(i, srcCols("id субагента"))

        Next i

        ' ----------------------------------------------------
        ' Сверка: те же условия, что и раньше, но одним проходом.
        ' Они независимы построчно, поэтому порядок не важен.
        ' Выжившие строки уплотняются к началу массива.
        ' ----------------------------------------------------

        keptCount = 0

        For i = 1 To rowsDst

            statusText = AP3_NormalizeStatusText( _
                CStr(dstData(i, dstCols("transaction status"))) _
            )

            keepRow = True

            If statusText = "отказ" Then
                keepRow = False
            ElseIf statusText = "в очереди" Then
                keepRow = False
            ElseIf Trim$(CStr(dstData(i, dstCols("adm transaction id")))) <> _
                   Trim$(CStr(dstData(i, dstCols("transaction id")))) Then
                keepRow = False
            ElseIf Trim$(CStr(dstData(i, dstCols("adm user id")))) <> _
                   Trim$(CStr(dstData(i, dstCols("user id")))) Then
                keepRow = False
            ElseIf Trim$(CStr(dstData(i, dstCols("adm check amount")))) <> _
                   Trim$(CStr(dstData(i, dstCols("check amount")))) Then
                keepRow = False
            ElseIf Trim$(CStr(dstData(i, dstCols("adm subagent id")))) <> _
                   Trim$(CStr(dstData(i, dstCols("subagent id")))) Then
                keepRow = False
            End If

            If keepRow Then

                keptCount = keptCount + 1

                If keptCount <> i Then

                    For c = 1 To lastColDst
                        dstData(keptCount, c) = dstData(i, c)
                    Next c

                End If

            End If

        Next i

        ' Запись одним обращением, затем один непрерывный Delete
        ' вместо тысяч поштучных.
        If keptCount > 0 Then

            wsDst.Range( _
                wsDst.Cells(AP3_FILTER_DATA_ROW, 1), _
                wsDst.Cells(AP3_FILTER_DATA_ROW + keptCount - 1, lastColDst) _
            ).Value = dstData

        End If

        If keptCount < rowsDst Then

            wsDst.Rows( _
                (AP3_FILTER_DATA_ROW + keptCount) & ":" & lastRowDst _
            ).Delete

        End If

    End If

    ' ========================================================
    ' 3. РАЗДЕЛЕНИЕ ПО ИСТОЧНИКУ НОМЕРА
    '
    ' Резать после сверки, а не до неё: в админку уходит
    ' один список, значит ручной шаг остаётся один,
    ' а позиционное сопоставление выше не ломается.
    ' ========================================================

    didSplit = AP3_SplitBySource( _
        wb, wsDst, wsExport, wsComment, countExport, countComment _
    )

    ' Перенос, сверка и разделение к этому моменту уже отработали,
    ' данные на листах верные. Сбой на кнопках - косметика,
    ' он не имеет права обесценить сделанную работу,
    ' поэтому собирается в предупреждение, а не в FailHandler.
    If didSplit Then

        buttonWarning = AP3_JoinWarnings( _
            AP3_TryAddButtons(wsExport), _
            AP3_TryAddButtons(wsComment) _
        )

        Set wsOpen = wsExport

    Else

        ' Книга собрана базовым Approved_Фильтр1:
        ' колонки "Источник номера" нет, делить нечем.
        buttonWarning = AP3_TryAddButtons(wsDst)

        Set wsOpen = wsDst

    End If

    completedOK = True

CleanExit:

    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating

    If wsOpen Is Nothing Then Set wsOpen = wsDst

    AP3_OpenSheet wsOpen

    If completedOK Then

        If didSplit Then

            MsgBox _
                "Готово. Перенос и сверка выполнены." & vbCrLf & vbCrLf & _
                "'" & AP3_SHEET_EXPORT & "': " & CStr(countExport) & vbCrLf & _
                "'" & AP3_SHEET_COMMENT & "': " & CStr(countComment) & vbCrLf & vbCrLf & _
                "Строки со статусом 'Отказ' и 'В очереди' удалены." & _
                AP3_AlignStatusText(alignWarning) & _
                AP3_ButtonStatusText(buttonWarning), _
                vbInformation

        Else

            MsgBox _
                "Готово. Перенос и сверка выполнены." & vbCrLf & vbCrLf & _
                "Колонка 'Источник номера' не найдена, " & _
                "поэтому лист 'Фильтр' не разделён." & vbCrLf & _
                "Так бывает, если выгрузка собрана старым Approved_Фильтр1." & _
                AP3_AlignStatusText(alignWarning) & _
                AP3_ButtonStatusText(buttonWarning), _
                vbInformation

        End If

    End If

    Exit Sub

FailHandler:

    errorText = Err.Description

    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating

    ' Если сбой случился уже после удаления "Фильтра",
    ' ссылка на него мертва - открываем то, что уцелело.
    On Error Resume Next

    If Not wsExport Is Nothing Then
        AP3_OpenSheet wsExport
    Else
        AP3_OpenSheet wsDst
    End If

    On Error GoTo 0

    MsgBox _
        "Не удалось выполнить Approved_Перенос2_v3: " & errorText, _
        vbCritical

End Sub


' ============================================================
' РАЗДЕЛЕНИЕ ЛИСТА
'
' Возвращает True, если листы созданы.
' False означает, что колонки "Источник номера" нет
' и делить нечего - "Фильтр" остаётся нетронутым.
' ============================================================

Private Function AP3_SplitBySource( _
    ByVal wb As Workbook, _
    ByVal wsDst As Worksheet, _
    ByRef wsExport As Worksheet, _
    ByRef wsComment As Worksheet, _
    ByRef countExport As Long, _
    ByRef countComment As Long _
) As Boolean

    Dim sourceCol As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Dim rowsData As Variant
    Dim headers As Variant

    Dim expData As Variant
    Dim comData As Variant

    Dim rowCount As Long
    Dim r As Long
    Dim c As Long

    AP3_SplitBySource = False

    countExport = 0
    countComment = 0

    sourceCol = AP3_FindHeaderCol( _
        wsDst, "Источник номера", AP3_FILTER_HEADER_ROW _
    )

    If sourceCol = 0 Then Exit Function

    lastCol = wsDst.Cells( _
        AP3_FILTER_HEADER_ROW, _
        wsDst.Columns.Count _
    ).End(xlToLeft).Column

    lastRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row

    headers = wsDst.Range( _
        wsDst.Cells(AP3_FILTER_HEADER_ROW, 1), _
        wsDst.Cells(AP3_FILTER_HEADER_ROW, lastCol) _
    ).Value

    rowCount = lastRow - AP3_FILTER_DATA_ROW + 1

    If rowCount > 0 Then

        rowsData = wsDst.Range( _
            wsDst.Cells(AP3_FILTER_DATA_ROW, 1), _
            wsDst.Cells(lastRow, lastCol) _
        ).Value

        ReDim expData(1 To rowCount, 1 To lastCol)
        ReDim comData(1 To rowCount, 1 To lastCol)

        For r = 1 To rowCount

            If Trim$(CStr(rowsData(r, sourceCol))) = AP3_TAG_EXPORT Then

                countExport = countExport + 1

                For c = 1 To lastCol
                    expData(countExport, c) = rowsData(r, c)
                Next c

            Else

                countComment = countComment + 1

                For c = 1 To lastCol
                    comData(countComment, c) = rowsData(r, c)
                Next c

            End If

        Next r

    End If

    ' Листы создаются до того, как "Фильтр" будет удалён:
    ' если что-то сорвётся, исходный лист останется на месте.
    AP3_DeleteSheetIfExists wb, AP3_SHEET_EXPORT
    AP3_DeleteSheetIfExists wb, AP3_SHEET_COMMENT

    Set wsExport = wb.Worksheets.Add(Before:=wsDst)
    wsExport.Name = AP3_SHEET_EXPORT

    Set wsComment = wb.Worksheets.Add(Before:=wsDst)
    wsComment.Name = AP3_SHEET_COMMENT

    AP3_WriteSheet wsExport, headers, expData, countExport, lastCol
    AP3_WriteSheet wsComment, headers, comData, countComment, lastCol

    Application.DisplayAlerts = False
    wsDst.Delete

    AP3_SplitBySource = True

End Function


' Пустая группа тоже получает свой лист, только с заголовками:
' состав выгрузки не должен менять набор листов.
Private Sub AP3_WriteSheet( _
    ByVal ws As Worksheet, _
    ByVal headers As Variant, _
    ByVal rowsData As Variant, _
    ByVal rowCount As Long, _
    ByVal lastCol As Long _
)

    Dim commentCol As Long
    Dim noteCol As Long
    Dim candCol As Long

    ' Строка 1 отдана под кнопки, как на листе "Фильтр".
    ws.Rows(1).RowHeight = 24

    ws.Range( _
        ws.Cells(AP3_FILTER_HEADER_ROW, 1), _
        ws.Cells(AP3_FILTER_HEADER_ROW, lastCol) _
    ).Value = headers

    If rowCount > 0 Then

        ws.Range( _
            ws.Cells(AP3_FILTER_DATA_ROW, 1), _
            ws.Cells(AP3_FILTER_DATA_ROW + rowCount - 1, lastCol) _
        ).Value = rowsData

    End If

    ws.Rows(AP3_FILTER_HEADER_ROW).Font.Bold = True
    ws.Columns.AutoFit

    ' Длинные текстовые колонки иначе растягивают лист.
    commentCol = AP3_FindHeaderCol(ws, "Internal comment", AP3_FILTER_HEADER_ROW)
    If commentCol > 0 Then ws.Columns(commentCol).ColumnWidth = 20

    candCol = AP3_FindHeaderCol(ws, "Кандидаты из комментария", AP3_FILTER_HEADER_ROW)
    If candCol > 0 Then ws.Columns(candCol).ColumnWidth = 22

    noteCol = AP3_FindHeaderCol(ws, "Примечание", AP3_FILTER_HEADER_ROW)
    If noteCol > 0 Then ws.Columns(noteCol).ColumnWidth = 44


    AP3_NormalizeRows ws, AP3_FILTER_DATA_ROW

End Sub


' ============================================================
' КНОПКИ
' ============================================================

Private Sub AP3_AddTicketCopyButtons(ByVal ws As Worksheet)

    Dim ticketCol As Long
    Dim lastRow As Long
    Dim totalTickets As Long
    Dim gridRight As Double

    Dim shp As Shape
    Dim i As Long

    ' Листы после разделения создаются с нуля, но при повторном
    ' прогоне без разделения кнопки уже могут быть на месте.
    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1

        Set shp = ws.Shapes(i)

        If Left$(shp.Name, 11) = "btnApproved" Then
            shp.Delete
        End If

    Next i

    On Error GoTo 0

    ' На листе "Из комментария" нужен один список для Bulk Approve:
    ' Ticket ID + Transaction ID целиком.
    If LCase$(Trim$(ws.Name)) = LCase$(AP3_SHEET_COMMENT) Then

        AP3_HideColumnByHeader ws, "Transaction status"
        AP3_AddCommentPairCopyButton ws
        Exit Sub

    End If

    ticketCol = AP3_FindHeaderCol(ws, "Ticket ID", AP3_FILTER_HEADER_ROW)

    If ticketCol = 0 Then

        MsgBox _
            "Не найдена колонка 'Ticket ID' для кнопки копирования " & _
            "на листе '" & ws.Name & "'.", _
            vbExclamation

        Exit Sub

    End If

    lastRow = ws.Cells(ws.Rows.Count, ticketCol).End(xlUp).Row

    If ws.Columns(ticketCol).ColumnWidth < 22 Then
        ws.Columns(ticketCol).ColumnWidth = 22
    End If

    ws.Rows(1).RowHeight = 24

    ' На листе "Из выгрузки" Ticket ID копируются одной кнопкой целиком.
    If lastRow >= AP3_FILTER_DATA_ROW Then

        totalTickets = lastRow - AP3_FILTER_DATA_ROW + 1

        AP3_CreateTicketCopyButton _
            ws, ticketCol, 1, 1, _
            AP3_FILTER_DATA_ROW, lastRow, 1, totalTickets

    End If

    ' Правая граница одной кнопки Ticket ID. Отдельная кнопка
    ' Transaction ID остаётся на своём месте и не перекрывается.
    gridRight = ws.Cells(1, ticketCol).Left + ws.Cells(1, ticketCol).Width

    AP3_AddTransactionCopyButton ws, gridRight

End Sub

Private Sub AP3_CreateTicketCopyButton( _
    ByVal ws As Worksheet, _
    ByVal ticketCol As Long, _
    ByVal buttonIndex As Long, _
    ByVal perColumn As Long, _
    ByVal startDataRow As Long, _
    ByVal endDataRow As Long, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long _
)

    Dim shp As Shape
    Dim baseCell As Range
    Dim macroWorkbookName As String

    Dim colIdx As Long
    Dim rowIdx As Long

    Set baseCell = ws.Cells(1, ticketCol)

    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    ' Заполняем столбец сверху вниз, затем уходим правее.
    colIdx = (buttonIndex - 1) \ perColumn
    rowIdx = (buttonIndex - 1) Mod perColumn

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        baseCell.Left + 2 + colIdx * baseCell.Width, _
        baseCell.Top + 2 + rowIdx * AP3_BUTTON_PITCH_Y, _
        baseCell.Width - 4, _
        AP3_BUTTON_HEIGHT _
    )

    With shp

        .Name = "btnApprovedCopyTickets_" & CStr(buttonIndex)

        .AlternativeText = _
            CStr(startDataRow) & "|" & _
            CStr(endDataRow) & "|" & _
            CStr(ticketCol)

        .OnAction = _
            "'" & macroWorkbookName & _
            "'!Approved_Reconcile_CopyTicketBlock_v3"

        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(31, 78, 121)

        .TextFrame2.TextRange.Text = _
            "Копировать " & CStr(endIndex - startIndex + 1)

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

        .Placement = xlFreeFloating

    End With

End Sub


Private Sub AP3_AddTransactionCopyButton( _
    ByVal ws As Worksheet, _
    ByVal gridRight As Double _
)

    Dim transactionCol As Long
    Dim buttonCell As Range
    Dim shp As Shape
    Dim macroWorkbookName As String
    Dim buttonLeft As Double

    transactionCol = AP3_FindHeaderCol( _
        ws, "Transaction ID", AP3_FILTER_HEADER_ROW _
    )

    If transactionCol = 0 Then

        MsgBox _
            "Не найдена колонка 'Transaction ID' для кнопки копирования " & _
            "на листе '" & ws.Name & "'.", _
            vbExclamation

        Exit Sub

    End If

    Set buttonCell = ws.Cells(1, transactionCol)

    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    ' Своё место над колонкой Transaction ID, но не под сеткой
    ' кнопок Ticket ID, если та дотянулась сюда справа.
    buttonLeft = buttonCell.Left + 2

    If gridRight + 4 > buttonLeft Then
        buttonLeft = gridRight + 4
    End If

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        buttonLeft, _
        buttonCell.Top + 2, _
        buttonCell.Width - 4, _
        AP3_BUTTON_HEIGHT _
    )

    With shp

        .Name = "btnApprovedCopyTransactions"
        .AlternativeText = "Transaction ID"

        .OnAction = _
            "'" & macroWorkbookName & _
            "'!Approved_Reconcile_CopyTransactions_v3"

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

        .Placement = xlFreeFloating

    End With

End Sub


Private Sub AP3_AddCommentPairCopyButton(ByVal ws As Worksheet)

    Dim ticketCol As Long
    Dim transactionCol As Long
    Dim lastRow As Long
    Dim totalRows As Long

    Dim shp As Shape
    Dim baseCell As Range
    Dim macroWorkbookName As String

    Dim ticketCenter As Double
    Dim transactionCenter As Double
    Dim buttonWidth As Double
    Dim buttonLeft As Double

    ticketCol = AP3_FindHeaderCol(ws, "Ticket ID", AP3_FILTER_HEADER_ROW)
    transactionCol = AP3_FindHeaderCol(ws, "Transaction ID", AP3_FILTER_HEADER_ROW)

    If ticketCol = 0 Or transactionCol = 0 Then

        MsgBox _
            "Не найдены Ticket ID или Transaction ID на листе '" & ws.Name & "'.", _
            vbExclamation

        Exit Sub

    End If

    lastRow = ws.Cells(ws.Rows.Count, ticketCol).End(xlUp).Row

    If lastRow < AP3_FILTER_DATA_ROW Then Exit Sub

    totalRows = lastRow - AP3_FILTER_DATA_ROW + 1

    Set baseCell = ws.Cells(1, ticketCol)

    If ws.Columns(ticketCol).ColumnWidth < 22 Then
        ws.Columns(ticketCol).ColumnWidth = 22
    End If

    ws.Rows(1).RowHeight = 24

    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    ' Кнопка копирует Ticket ID + Transaction ID, поэтому ставим её
    ' визуально по центру между этими двумя колонками.
    ticketCenter = ws.Cells(1, ticketCol).Left + ws.Cells(1, ticketCol).Width / 2
    transactionCenter = ws.Cells(1, transactionCol).Left + ws.Cells(1, transactionCol).Width / 2

    buttonWidth = baseCell.Width - 4
    buttonLeft = (ticketCenter + transactionCenter) / 2 - buttonWidth / 2

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        buttonLeft, _
        baseCell.Top + 2, _
        buttonWidth, _
        AP3_BUTTON_HEIGHT _
    )

    With shp

        .Name = "btnApprovedCopyCommentPairs"
        .AlternativeText = "Ticket ID|Transaction ID"

        .OnAction = _
            "'" & macroWorkbookName & "'!Approved_Reconcile_CopyCommentPairs_v3"

        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.ForeColor.RGB = RGB(31, 78, 121)

        .TextFrame2.TextRange.Text = "Копировать " & CStr(totalRows)

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

        .Placement = xlFreeFloating

    End With

End Sub


Private Sub AP3_HideColumnByHeader( _
    ByVal ws As Worksheet, _
    ByVal headerText As String _
)

    Dim col As Long

    col = AP3_FindHeaderCol(ws, headerText, AP3_FILTER_HEADER_ROW)

    If col > 0 Then
        ws.Columns(col).Hidden = True
    End If

End Sub


Private Sub AP3_NormalizeRows( _
    ByVal ws As Worksheet, _
    ByVal dataStartRow As Long _
)

    Dim lastCell As Range
    Dim lastRow As Long

    If ws Is Nothing Then Exit Sub

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If lastCell Is Nothing Then Exit Sub

    lastRow = lastCell.Row

    ws.UsedRange.WrapText = False

    If lastRow >= dataStartRow Then
        ws.Rows(dataStartRow & ":" & lastRow).RowHeight = ws.StandardHeight
    End If

End Sub


' Кнопки строятся под своим обработчиком ошибок: к этому моменту
' данные уже разложены по листам, и падение на косметике
' не должно превращаться в "Не удалось выполнить".
Private Function AP3_TryAddButtons(ByVal ws As Worksheet) As String

    On Error GoTo ButtonFail

    If ws Is Nothing Then Exit Function

    AP3_NormalizeRows ws, AP3_FILTER_DATA_ROW
    AP3_AddTicketCopyButtons ws

    AP3_TryAddButtons = ""

    Exit Function

ButtonFail:

    AP3_TryAddButtons = _
        "лист '" & ws.Name & "': " & Err.Description

End Function


Private Function AP3_JoinWarnings( _
    ByVal first As String, _
    ByVal second As String _
) As String

    If first = "" Then
        AP3_JoinWarnings = second
    ElseIf second = "" Then
        AP3_JoinWarnings = first
    Else
        AP3_JoinWarnings = first & vbCrLf & second
    End If

End Function


Private Function AP3_AlignStatusText( _
    ByVal warning As String _
) As String

    If warning = "" Then Exit Function

    AP3_AlignStatusText = _
        vbCrLf & vbCrLf & _
        "ВНИМАНИЕ: число строк не совпало." & vbCrLf & _
        warning & vbCrLf & _
        "Листы сопоставляются построчно, поэтому всё после пропуска" & _
        vbCrLf & _
        "съехало и было вычищено сверкой. Уцелела только та часть," & _
        vbCrLf & _
        "что шла до пропуска."

End Function


Private Function AP3_ButtonStatusText( _
    ByVal warning As String _
) As String

    If warning = "" Then

        AP3_ButtonStatusText = _
            vbCrLf & "Кнопки копирования добавлены."

    Else

        AP3_ButtonStatusText = _
            vbCrLf & vbCrLf & _
            "Данные разложены верно, но кнопки построить не удалось:" & _
            vbCrLf & warning & vbCrLf & vbCrLf & _
            "Встань на нужный лист и запусти Approved_Кнопки_v3."

    End If

End Function


' ============================================================
' ДОСТРОЙКА КНОПОК НА ГОТОВОМ ЛИСТЕ
'
' Запускается вручную на листе результата. Нужна, чтобы сбой
' на кнопках не стоил повторного похода в админ-панель:
' данные уже на месте, не хватает только кнопок.
' ============================================================

Public Function Approved_Кнопки_v3() As Variant

    Dim ws As Worksheet
    Dim warning As String

    Dim oldScreenUpdating As Boolean

    Set ws = ActiveSheet

    If Not AP3_IsResultSheet(ws.Name) Then

        MsgBox _
            "Макрос работает на листах '" & AP3_SHEET_EXPORT & _
            "' и '" & AP3_SHEET_COMMENT & "'." & vbCrLf & _
            "Сейчас активен лист '" & ws.Name & "'.", _
            vbExclamation

        Exit Function

    End If

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    warning = AP3_TryAddButtons(ws)

    Application.ScreenUpdating = oldScreenUpdating

    If warning = "" Then

        MsgBox _
            "Кнопки на листе '" & ws.Name & "' построены заново.", _
            vbInformation

    Else

        MsgBox _
            "Не удалось построить кнопки: " & warning, _
            vbExclamation

    End If

End Function


' ============================================================
' ОБРАБОТЧИКИ КНОПОК
'
' Function используется специально, чтобы служебный обработчик
' не отображался в Alt+F8.
' ============================================================

Public Function Approved_Reconcile_CopyCommentPairs_v3() As Variant

    Dim ws As Worksheet
    Dim ticketCol As Long
    Dim transactionCol As Long
    Dim lastRow As Long
    Dim r As Long

    Dim clipboardText As String

    On Error GoTo CopyFail

    Set ws = ActiveSheet

    If LCase$(Trim$(ws.Name)) <> LCase$(AP3_SHEET_COMMENT) Then

        MsgBox _
            "Кнопка работает только на листе '" & AP3_SHEET_COMMENT & "'.", _
            vbExclamation

        Exit Function

    End If

    ticketCol = AP3_FindHeaderCol(ws, "Ticket ID", AP3_FILTER_HEADER_ROW)
    transactionCol = AP3_FindHeaderCol(ws, "Transaction ID", AP3_FILTER_HEADER_ROW)

    If ticketCol = 0 Or transactionCol = 0 Then

        MsgBox _
            "Не найдены Ticket ID или Transaction ID.", _
            vbExclamation

        Exit Function

    End If

    lastRow = ws.Cells(ws.Rows.Count, ticketCol).End(xlUp).Row

    If lastRow < AP3_FILTER_DATA_ROW Then

        MsgBox "Нет данных для копирования.", vbExclamation
        Exit Function

    End If

    For r = AP3_FILTER_DATA_ROW To lastRow

        If r > AP3_FILTER_DATA_ROW Then
            clipboardText = clipboardText & vbCrLf
        End If

        clipboardText = clipboardText & _
            CStr(ws.Cells(r, ticketCol).Value2) & vbTab & _
            CStr(ws.Cells(r, transactionCol).Value2)

    Next r

    Application.CutCopyMode = False

    If Not AP3_CopyUnicodeTextToClipboard(clipboardText) Then
        Err.Raise vbObjectError + 3101, , _
            "Не удалось записать данные в буфер обмена Windows."
    End If

    Application.StatusBar = _
        "Скопировано Ticket ID + Transaction ID: " & _
        CStr(lastRow - AP3_FILTER_DATA_ROW + 1) & " строк"

    Exit Function

CopyFail:

    Application.StatusBar = False

    MsgBox _
        "Не удалось скопировать Ticket ID + Transaction ID: " & Err.Description, _
        vbExclamation

End Function


Public Function Approved_Reconcile_CopyTicketBlock_v3() As Variant

    Dim ws As Worksheet
    Dim shp As Shape
    Dim params As Variant

    Dim startDataRow As Long
    Dim endDataRow As Long
    Dim ticketCol As Long

    On Error GoTo CopyFail

    Set ws = ActiveSheet

    If Not AP3_IsResultSheet(ws.Name) Then

        MsgBox AP3_WrongSheetText, vbExclamation
        Exit Function

    End If

    Set shp = ws.Shapes(CStr(Application.Caller))

    params = Split(shp.AlternativeText, "|")

    If UBound(params) <> 2 Then

        MsgBox "Некорректные параметры кнопки копирования.", vbExclamation
        Exit Function

    End If

    startDataRow = CLng(params(0))
    endDataRow = CLng(params(1))
    ticketCol = CLng(params(2))

    ws.Range( _
        ws.Cells(startDataRow, ticketCol), _
        ws.Cells(endDataRow, ticketCol) _
    ).Copy

    Application.StatusBar = _
        "Скопировано Ticket ID: " & _
        CStr(startDataRow - AP3_FILTER_DATA_ROW + 1) & "-" & _
        CStr(endDataRow - AP3_FILTER_DATA_ROW + 1) & _
        " (лист '" & ws.Name & "')"

    Exit Function

CopyFail:

    Application.StatusBar = False

    MsgBox "Не удалось скопировать Ticket ID: " & Err.Description, vbExclamation

End Function


Public Function Approved_Reconcile_CopyTransactions_v3() As Variant

    Dim ws As Worksheet
    Dim transactionCol As Long
    Dim lastRow As Long

    On Error GoTo CopyFail

    Set ws = ActiveSheet

    If Not AP3_IsResultSheet(ws.Name) Then

        MsgBox AP3_WrongSheetText, vbExclamation
        Exit Function

    End If

    transactionCol = AP3_FindHeaderCol( _
        ws, "Transaction ID", AP3_FILTER_HEADER_ROW _
    )

    If transactionCol = 0 Then

        MsgBox "Не найдена колонка 'Transaction ID'.", vbExclamation
        Exit Function

    End If

    lastRow = ws.Cells(ws.Rows.Count, transactionCol).End(xlUp).Row

    If lastRow < AP3_FILTER_DATA_ROW Then

        MsgBox _
            "В колонке 'Transaction ID' нет данных для копирования.", _
            vbExclamation

        Exit Function

    End If

    ws.Range( _
        ws.Cells(AP3_FILTER_DATA_ROW, transactionCol), _
        ws.Cells(lastRow, transactionCol) _
    ).Copy

    Application.StatusBar = _
        "Скопирован список Transaction ID: " & _
        CStr(lastRow - AP3_FILTER_DATA_ROW + 1) & " строк" & _
        " (лист '" & ws.Name & "')"

    Exit Function

CopyFail:

    Application.StatusBar = False

    MsgBox _
        "Не удалось скопировать Transaction ID: " & Err.Description, _
        vbExclamation

End Function



' Записывает обычный Unicode-текст в системный буфер обмена Windows.
' Формат строк для Bulk Approve остаётся: Ticket ID<TAB>Transaction ID.
Private Function AP3_CopyUnicodeTextToClipboard( _
    ByVal textValue As String _
) As Boolean

#If VBA7 Then
    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim setResult As LongPtr
    Dim byteCount As LongPtr
#Else
    Dim hMem As Long
    Dim pMem As Long
    Dim setResult As Long
    Dim byteCount As Long
#End If

    Dim clipboardIsOpen As Boolean
    Dim tryNo As Long

    On Error GoTo ClipboardFail

    AP3_CopyUnicodeTextToClipboard = False

    If Len(textValue) = 0 Then Exit Function

    byteCount = LenB(textValue)

    ' +2 байта оставляют завершающий Unicode-ноль.
    hMem = AP3_GlobalAlloc(AP3_GMEM_MOVEABLE_ZEROINIT, byteCount + 2)
    If hMem = 0 Then Exit Function

    pMem = AP3_GlobalLock(hMem)
    If pMem = 0 Then GoTo ClipboardFail

    AP3_CopyMemory pMem, StrPtr(textValue), byteCount
    AP3_GlobalUnlock hMem
    pMem = 0

    ' Буфер иногда на долю секунды занят другим приложением.
    ' Несколько попыток + DoEvents делают копирование устойчивее.
    For tryNo = 1 To 10
        If AP3_OpenClipboard(0) <> 0 Then
            clipboardIsOpen = True
            Exit For
        End If
        DoEvents
    Next tryNo

    If Not clipboardIsOpen Then GoTo ClipboardFail

    If AP3_EmptyClipboard() = 0 Then GoTo ClipboardFail

    setResult = AP3_SetClipboardData(AP3_CF_UNICODETEXT, hMem)
    If setResult = 0 Then GoTo ClipboardFail

    ' После успешного SetClipboardData память принадлежит Windows.
    hMem = 0

    AP3_CloseClipboard
    clipboardIsOpen = False

    AP3_CopyUnicodeTextToClipboard = True
    Exit Function

ClipboardFail:

    On Error Resume Next

    If pMem <> 0 Then AP3_GlobalUnlock hMem

    If clipboardIsOpen Then AP3_CloseClipboard

    If hMem <> 0 Then AP3_GlobalFree hMem

    On Error GoTo 0

End Function


' ============================================================
' СЛУЖЕБНЫЕ ФУНКЦИИ
' ============================================================

' "Фильтр" остаётся в списке на случай, когда разделение
' было пропущено из-за отсутствия колонки "Источник номера".
Private Function AP3_IsResultSheet(ByVal sheetName As String) As Boolean

    Dim n As String

    n = LCase$(Trim$(sheetName))

    AP3_IsResultSheet = _
        (n = LCase$(AP3_SHEET_EXPORT)) Or _
        (n = LCase$(AP3_SHEET_COMMENT)) Or _
        (n = LCase$("Фильтр"))

End Function


Private Function AP3_WrongSheetText() As String

    AP3_WrongSheetText = _
        "Кнопка работает на листах '" & AP3_SHEET_EXPORT & _
        "' и '" & AP3_SHEET_COMMENT & "'."

End Function


Private Function AP3_FindHeaderCol( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal headerRow As Long _
) As Long

    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        If Trim$(LCase$(CStr(ws.Cells(headerRow, c).Value))) = _
           Trim$(LCase$(headerName)) Then

            AP3_FindHeaderCol = c
            Exit Function

        End If

    Next c

    AP3_FindHeaderCol = 0

End Function


Private Function AP3_NormalizeStatusText( _
    ByVal valueText As String _
) As String

    Dim s As String

    s = LCase$(Trim$(valueText))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, "ё", "е")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    AP3_NormalizeStatusText = s

End Function


Private Sub AP3_DeleteSheetIfExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
)

    On Error Resume Next
    wb.Worksheets(sheetName).Delete
    On Error GoTo 0

End Sub


Private Sub AP3_OpenSheet(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    ws.Activate
    ws.Range("A1").Select

    On Error Resume Next

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    On Error GoTo 0

End Sub
