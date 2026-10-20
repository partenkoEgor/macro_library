Option Explicit

' ============================================================
' Approved_Фильтр1_v3
' Версия 3.1.4
'
' Отличие от Approved_Фильтр1: если поле Transaction ID пустое,
' номер транзакции пытается вычлениться из поля Internal comment.
'
' Логика подбора намеренно консервативная: лучше отправить
' строку на ручную проверку, чем подставить чужой номер.
' Всё, что вызвало хоть малейшее сомнение, уходит
' на лист "На проверку" с указанием причины.
'
' Выходная раскладка листа "Фильтр" полностью совпадает
' с прошлой версией (колонки A..X), поэтому
' Approved_Перенос2 работает без изменений.
' Служебные колонки Y и Z добавлены в конец.
' ============================================================

Private Const AF3_FILTER_HEADER_ROW As Long = 2
Private Const AF3_FILTER_DATA_ROW As Long = 3

Private Const AF3_REVIEW_HEADER_ROW As Long = 1
Private Const AF3_REVIEW_DATA_ROW As Long = 2

' Лист с цифрами прогона в раскладке ежедневного отчёта.
' Строки 3-15 совпадают со строками отчёта, поэтому колонку B
' можно вставить в столбец нужного дня как есть.
' Имя должно совпадать с AP3_SHEET_REPORT в Approved_Перенос2_v3.
Private Const AF3_SHEET_REPORT As String = "Отчёт"

' Лист прошлой версии, его остатки удаляются при запуске.
Private Const AF3_SHEET_SUMMARY_OLD As String = "Итоги"

' Длина номера транзакции. Во всей выгрузке она ровно 11.
Private Const AF3_TRANS_LEN As Long = 11

' Максимум Transaction ID в одной кнопке для загрузки в админ-панель.
Private Const AF3_COPY_BLOCK_SIZE As Long = 5000

' ------------------------------------------------------------
' НАСТРОЙКИ БЕЗОПАСНОСТИ
'
' Чем меньше включено, тем строже отбор.
' Любая выключенная категория не теряется,
' а попадает на лист "На проверку".
' ------------------------------------------------------------

' Главный выключатель подбора номера из комментария.
' False -> макрос ведёт себя как прошлая версия.
Private Const AF3_ENABLE_COMMENT_EXTRACT As Boolean = True

' Категория B2: перед номером стоит фраза подтверждения агента,
' например "Sir, the order has been confirmed by 23806942713"
' или "the transaction number is 23658965287". Список фраз -
' в AF3_ConfirmPhrases. Номер жалобы из "confirmed in complaint N"
' сюда не попадает: прямо перед ним стоит слово "complaint".
Private Const AF3_ACCEPT_TIER_B2 As Boolean = True

' Категория D: короткий комментарий без слова-метки,
' например "23340367235 already confirmed".
Private Const AF3_ACCEPT_TIER_D As Boolean = True

' Категория E: номер стоит посреди свободного текста
' и ничем не помечен. Именно здесь живут опасные случаи
' вроде "confirmed in complaint 23484695109" - это номер
' жалобы, а не транзакции. По умолчанию такие строки
' идут на проверку.
Private Const AF3_ACCEPT_TIER_E As Boolean = False

' Один и тот же номер в нескольких тикетах ОДНОГО пользователя.
' Ответ у таких тикетов одинаковый и верный, опасности нет.
'
' Но Approved_Перенос2 сопоставляет "Источник" и "Фильтр"
' ПОЗИЦИОННО, строка к строке, а не по номеру. Значит список,
' который уходит в админку, обязан быть без повторов, иначе
' ответ админки сдвинется относительно тикетов.
'
' Поэтому по умолчанию в работе остаётся один тикет из группы,
' а остальные уходят на проверку со ссылкой на основной.
' Ставить True можно, только если админка гарантированно
' возвращает отдельную строку на каждый повтор номера.
Private Const AF3_ACCEPT_DUP_SAME_USER As Boolean = False

' Один и тот же номер у РАЗНЫХ пользователей всегда уходит
' на проверку и настройкой не отключается. Прав тут максимум
' один тикет, а значит остальные получили бы неверный ответ.
' Именно ради этого случая и написана вся проверка.

' Приведение поля From к значениям PS / User.
' В прошлой версии правила не срабатывали:
' выгрузка отдаёт "By customer" и "By Customer Support",
' а искались "by yourself" и "through customer support".
Private Const AF3_NORMALIZE_FROM As Boolean = True


Public Sub Approved_Фильтр1_v3()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim wsReview As Worksheet
    Dim wsOrigin As Worksheet

    Dim srcData As Variant
    Dim outData As Variant
    Dim revData As Variant

    Dim srcCols As Object

    Dim lastRow As Long
    Dim lastCol As Long

    Dim r As Long
    Dim outCount As Long
    Dim revCount As Long

    Dim statKept As Long
    Dim statWithTID As Long
    Dim statFromComment As Long
    Dim statReview As Long
    Dim statDropped As Long

    Dim transID As String
    Dim sourceTag As String
    Dim candList As String
    Dim reason As String
    Dim tier As String

    Dim verdict() As Long
    Dim note() As String
    Dim keepCount As Long
    Dim outRow As Long
    Dim dupLeft As Long
    Dim msg As String
    Dim wsReport As Worksheet

    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    Set wb = ActiveWorkbook
    Set wsSrc = ActiveSheet

    If AF3_IsServiceSheet(wsSrc.Name) Then

        MsgBox "Запусти макрос с исходного листа выгрузки.", vbExclamation
        Exit Sub

    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    On Error GoTo FailHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    AF3_DeleteSheetIfExists wb, "Фильтр"
    AF3_DeleteSheetIfExists wb, "Источник"
    AF3_DeleteSheetIfExists wb, "На проверку"
    AF3_DeleteSheetIfExists wb, "Для копирования"
    AF3_DeleteSheetIfExists wb, "Закрыть статусом 225"

    ' Итоговые листы прошлого прогона Approved_Перенос2_v3.
    AF3_DeleteSheetIfExists wb, "Из выгрузки"
    AF3_DeleteSheetIfExists wb, "Из комментария"

    ' Отчёт прошлого прогона.
    AF3_DeleteSheetIfExists wb, AF3_SHEET_REPORT
    AF3_DeleteSheetIfExists wb, AF3_SHEET_SUMMARY_OLD

    ' --------------------------------------------------------
    ' 1. ЧИТАЕМ ВЫГРУЗКУ В МАССИВ
    '
    ' Работа через массив, а не через ячейки:
    ' на 4800 строках это разница в десятки раз.
    ' --------------------------------------------------------

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then

        MsgBox "На листе выгрузки нет данных.", vbExclamation
        GoTo CleanExit

    End If

    srcData = wsSrc.Range( _
        wsSrc.Cells(1, 1), _
        wsSrc.Cells(lastRow, lastCol) _
    ).Value

    Set srcCols = AF3_BuildHeaderMap(srcData, lastCol)

    If Not srcCols.Exists("transaction id") Then

        MsgBox "На листе выгрузки не найдена колонка 'Transaction ID'.", vbCritical
        GoTo CleanExit

    End If

    ' --------------------------------------------------------
    ' 2. ОСНОВНОЙ ПРОХОД
    '
    ' Для каждой строки определяем номер транзакции
    ' и решаем, куда она пойдёт.
    ' --------------------------------------------------------

    ReDim outData(1 To lastRow, 1 To 27)
    ReDim revData(1 To lastRow, 1 To 10)

    outCount = 0
    revCount = 0

    For r = 2 To lastRow

        transID = AF3_CleanNumber( _
            AF3_CellText(srcData, r, srcCols, "transaction id") _
        )

        candList = ""
        reason = ""
        tier = ""

        If transID <> "" Then

            ' Номер пришёл из выгрузки. Ничего не выдумываем.
            sourceTag = "Выгрузка"
            statWithTID = statWithTID + 1

        ElseIf Not AF3_ENABLE_COMMENT_EXTRACT Then

            sourceTag = ""

        Else

            transID = AF3_ResolveFromComment( _
                srcData, r, srcCols, tier, candList, reason _
            )

            If transID <> "" Then
                sourceTag = "Комментарий " & tier
            Else
                sourceTag = ""
            End If

        End If

        If transID = "" And reason <> "" Then

            ' Кандидат был, но не прошёл проверку.
            revCount = revCount + 1
            AF3_FillReviewRow revData, revCount, srcData, r, srcCols, candList, reason
            statReview = statReview + 1

        ElseIf transID = "" Then

            ' Ни номера, ни кандидата. Строка отбрасывается,
            ' ровно как в прошлой версии.
            statDropped = statDropped + 1

        Else

            outCount = outCount + 1
            AF3_FillOutRow outData, outCount, srcData, r, srcCols, transID, sourceTag, candList

            If sourceTag = "Выгрузка" Then
                statKept = statKept + 1
            Else
                statFromComment = statFromComment + 1
            End If

        End If

    Next r

    ' --------------------------------------------------------
    ' 3. ПРОВЕРКА НА ДУБЛИКАТЫ
    '
    ' Один номер транзакции, приписанный разным пользователям,
    ' означает неверный ответ как минимум одному из них.
    ' Такие группы уходят на проверку целиком.
    ' --------------------------------------------------------

    ReDim verdict(1 To outCount + 1)
    ReDim note(1 To outCount + 1)

    AF3_ResolveDuplicates outData, outCount, verdict, note

    keepCount = 0

    For outRow = 1 To outCount

        outData(outRow, 27) = note(outRow)

        If verdict(outRow) = 1 Then

            revCount = revCount + 1

            AF3_FillReviewRowFromOut _
                revData, revCount, outData, outRow, note(outRow)

            statReview = statReview + 1

            If CStr(outData(outRow, 25)) = "Выгрузка" Then
                statKept = statKept - 1
            Else
                statFromComment = statFromComment - 1
            End If

        Else

            keepCount = keepCount + 1

            If keepCount <> outRow Then
                AF3_MoveOutRow outData, outRow, keepCount
            End If

        End If

    Next outRow

    outCount = keepCount

    ' --------------------------------------------------------
    ' 4. ЗАПИСЬ ЛИСТА "ФИЛЬТР"
    ' --------------------------------------------------------

    Set wsDst = wb.Worksheets.Add(Before:=wsSrc)
    wsDst.Name = "Фильтр"

    ' Строка 1 отдана под кнопки.
    wsDst.Rows(1).RowHeight = 24

    AF3_WriteHeaders wsDst

    If outCount > 0 Then

        wsDst.Range( _
            wsDst.Cells(AF3_FILTER_DATA_ROW, 1), _
            wsDst.Cells(AF3_FILTER_DATA_ROW + outCount - 1, 27) _
        ).Value = outData

        AF3_SortByTransaction wsDst, outCount

    End If

    AF3_FormatFilter wsDst
    AF3_NormalizeRows wsDst, AF3_FILTER_DATA_ROW

    ' --------------------------------------------------------
    ' 5. ЗАПИСЬ ЛИСТА "НА ПРОВЕРКУ"
    ' --------------------------------------------------------

    Set wsReview = wb.Worksheets.Add(After:=wsDst)
    wsReview.Name = "На проверку"

    AF3_WriteReviewHeaders wsReview

    If revCount > 0 Then

        wsReview.Range( _
            wsReview.Cells(AF3_REVIEW_DATA_ROW, 1), _
            wsReview.Cells(AF3_REVIEW_DATA_ROW + revCount - 1, 10) _
        ).Value = revData

    End If

    AF3_FormatReview wsReview
    AF3_NormalizeRows wsReview, AF3_REVIEW_DATA_ROW

    ' --------------------------------------------------------
    ' 5а. ЛИСТ "ОТЧЁТ"
    '
    ' Цифры прогона в раскладке ежедневного отчёта.
    ' Блок второго макроса заполнит Approved_Перенос2_v3.
    ' --------------------------------------------------------

    Set wsReport = wb.Worksheets.Add(After:=wsReview)
    wsReport.Name = AF3_SHEET_REPORT

    AF3_WriteReport wsReport, wsSrc.Name, lastRow - 1, statWithTID, _
        outCount, statKept, statFromComment, statDropped

    ' --------------------------------------------------------
    ' 6. ПУСТОЙ ЛИСТ "ИСТОЧНИК"
    ' --------------------------------------------------------

    Set wsOrigin = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count) _
    )

    wsOrigin.Name = "Источник"

    AF3_AddButtons wsDst

    ' Approved_Перенос2 сопоставляет листы позиционно,
    ' поэтому повтор номера в списке сдвинул бы ответ админки.
    ' Проверяем это явно, а не надеемся на логику выше.
    dupLeft = AF3_CountDuplicateIDs(outData, outCount)

    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    AF3_OpenFilter wsDst

    msg = "Готово." & vbCrLf & vbCrLf & _
        "На лист 'Фильтр' попало: " & CStr(outCount) & vbCrLf & _
        "  из колонки Transaction ID: " & CStr(statKept) & vbCrLf & _
        "  подобрано из комментария: " & CStr(statFromComment) & vbCrLf & vbCrLf & _
        "На лист 'На проверку': " & CStr(statReview) & vbCrLf & _
        "Отброшено без номера: " & CStr(statDropped) & vbCrLf & vbCrLf & _
        "Проверь лист 'На проверку' перед выгрузкой в админку." & vbCrLf & _
        "Цифры для отчёта - на листе '" & AF3_SHEET_REPORT & "'."

    If dupLeft > 0 Then

        msg = msg & vbCrLf & vbCrLf & _
            "ВНИМАНИЕ: в списке " & CStr(dupLeft) & _
            " повторяющихся номеров." & vbCrLf & _
            "Approved_Перенос2 сопоставляет строки позиционно." & vbCrLf & _
            "Убедись, что админка возвращает отдельную строку " & _
            "на каждый повтор, иначе поставь " & _
            "AF3_ACCEPT_DUP_SAME_USER = False."

        MsgBox msg, vbExclamation

    Else

        MsgBox msg, vbInformation

    End If

    Exit Sub

CleanExit:

    Application.CutCopyMode = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    Exit Sub

FailHandler:

    Application.CutCopyMode = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox _
        "Не удалось выполнить Approved_Фильтр1_v3: " & Err.Description, _
        vbCritical

End Sub


' Считает, сколько строк несут номер, встречающийся
' в списке более одного раза.
Private Function AF3_CountDuplicateIDs( _
    ByRef outData As Variant, _
    ByVal outCount As Long _
) As Long

    Dim seen As Object
    Dim i As Long
    Dim k As String
    Dim total As Long

    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To outCount

        k = CStr(outData(i, 8))

        If seen.Exists(k) Then
            seen(k) = seen(k) + 1
        Else
            seen(k) = 1
        End If

    Next i

    For i = 1 To outCount

        If seen(CStr(outData(i, 8))) > 1 Then
            total = total + 1
        End If

    Next i

    AF3_CountDuplicateIDs = total

End Function


' ============================================================
' РАЗБОР ДУБЛИКАТОВ
'
' verdict(i) = 0 -> строка остаётся на листе "Фильтр"
' verdict(i) = 1 -> строка уходит на лист "На проверку"
'
' Правила:
'   один номер у одного пользователя  -> оставляем, помечаем
'   один номер у разных пользователей -> вся группа на проверку,
'   но если ровно один из тикетов взял номер прямо из выгрузки,
'   он достоверен и остаётся, а домыслы из комментариев уходят.
' ============================================================

Private Sub AF3_ResolveDuplicates( _
    ByRef outData As Variant, _
    ByVal outCount As Long, _
    ByRef verdict() As Long, _
    ByRef note() As String _
)

    Dim groups As Object
    Dim i As Long
    Dim transID As String
    Dim groupKey As Variant

    Dim idxList As Variant
    Dim k As Long
    Dim idx As Long

    Dim users As Object
    Dim realIdx As Long
    Dim realCount As Long
    Dim groupSize As Long
    Dim mainIdx As Long

    Set groups = CreateObject("Scripting.Dictionary")

    ' Группируем номера строк по номеру транзакции.
    For i = 1 To outCount

        verdict(i) = 0
        note(i) = ""

        transID = CStr(outData(i, 8))

        If groups.Exists(transID) Then
            groups(transID) = groups(transID) & "|" & CStr(i)
        Else
            groups(transID) = CStr(i)
        End If

    Next i

    For Each groupKey In groups.keys

        transID = CStr(groupKey)

        idxList = Split(CStr(groups(transID)), "|")
        groupSize = UBound(idxList) - LBound(idxList) + 1

        If groupSize > 1 Then

            ' Сколько разных пользователей и сколько строк
            ' с номером прямо из выгрузки.
            Set users = CreateObject("Scripting.Dictionary")

            realIdx = 0
            realCount = 0

            For k = LBound(idxList) To UBound(idxList)

                idx = CLng(idxList(k))

                users(Trim$(CStr(outData(idx, 10)))) = 1

                If CStr(outData(idx, 25)) = "Выгрузка" Then
                    realCount = realCount + 1
                    realIdx = idx
                End If

            Next k

            If users.Count = 1 Then

                ' Один человек, несколько жалоб на один платёж.
                ' Опасности неверного ответа нет, вопрос только
                ' в уникальности списка для админки.
                If AF3_ACCEPT_DUP_SAME_USER Then

                    For k = LBound(idxList) To UBound(idxList)

                        idx = CLng(idxList(k))

                        note(idx) = "Дубликат: " & CStr(groupSize) & _
                            " тикета(ов) одного пользователя"

                    Next k

                Else

                    ' Основным считаем тикет с номером из выгрузки,
                    ' а если такого нет - первый в группе.
                    If realIdx > 0 Then
                        mainIdx = realIdx
                    Else
                        mainIdx = CLng(idxList(LBound(idxList)))
                    End If

                    For k = LBound(idxList) To UBound(idxList)

                        idx = CLng(idxList(k))

                        If idx = mainIdx Then

                            note(idx) = "Основной тикет группы, ещё " & _
                                CStr(groupSize - 1) & _
                                " тикет(ов) того же пользователя " & _
                                "с этим номером - на проверке"

                        Else

                            verdict(idx) = 1

                            note(idx) = "Тот же номер " & transID & _
                                " и тот же пользователь, " & _
                                "основной тикет " & _
                                CStr(outData(mainIdx, 5))

                        End If

                    Next k

                End If

            ElseIf realCount = 1 Then

                ' Номер достоверно принадлежит тикету из выгрузки.
                ' Остальные приписали его себе в комментарии.
                For k = LBound(idxList) To UBound(idxList)

                    idx = CLng(idxList(k))

                    If idx = realIdx Then

                        note(idx) = "Номер также упомянут в комментариях " & _
                            "других тикетов (" & CStr(groupSize - 1) & ")"

                    Else

                        verdict(idx) = 1

                        note(idx) = "Номер " & transID & _
                            " уже есть в выгрузке у тикета " & _
                            CStr(outData(realIdx, 5)) & _
                            ", пользователь другой"

                    End If

                Next k

            Else

                ' Разные пользователи и ни одного достоверного
                ' источника. Прав максимум один тикет.
                For k = LBound(idxList) To UBound(idxList)

                    idx = CLng(idxList(k))

                    verdict(idx) = 1

                    note(idx) = "Номер " & transID & " указан у " & _
                        CStr(users.Count) & " разных пользователей (" & _
                        CStr(groupSize) & " тикетов)"

                Next k

            End If

        End If

    Next groupKey

End Sub


' ============================================================
' ПОДБОР НОМЕРА ИЗ КОММЕНТАРИЯ
'
' Возвращает номер транзакции либо пустую строку.
' Если кандидат был, но забракован, причина попадает
' в reason и строка уйдёт на лист "На проверку".
' ============================================================

Private Function AF3_ResolveFromComment( _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByRef tier As String, _
    ByRef candList As String, _
    ByRef reason As String _
) As String

    Dim commentText As String

    Dim runs() As String
    Dim starts() As Long
    Dim runCount As Long

    Dim i As Long
    Dim cand As String

    Dim uniq As Object
    Dim rejected As Object
    Dim clash As String

    Dim keys As Variant
    Dim only As String
    Dim onlyPos As Long

    AF3_ResolveFromComment = ""
    tier = ""
    candList = ""
    reason = ""

    commentText = AF3_CellText(srcData, r, srcCols, "internal comment")

    If Trim$(commentText) = "" Then Exit Function

    AF3_ScanDigitRuns commentText, runs, starts, runCount

    If runCount = 0 Then Exit Function

    Set uniq = CreateObject("Scripting.Dictionary")
    Set rejected = CreateObject("Scripting.Dictionary")

    For i = 1 To runCount

        cand = runs(i)

        ' Правило 1. Ровно 11 цифр.
        If Len(cand) <> AF3_TRANS_LEN Then GoTo NextRun

        ' Правило 2. Ведущий ноль означает телефон
        ' или номер платежа, но не номер транзакции.
        If Left$(cand, 1) = "0" Then GoTo NextRun

        ' Правило 3. Хвост десятичной дроби.
        ' Например "correct the amount to 10.23414215485".
        If AF3_IsDecimalTail(commentText, starts(i)) Then

            rejected("Похоже на дробную часть числа: " & cand) = 1
            GoTo NextRun

        End If

        ' Правило 4. Номер совпал с другим полем этой же строки:
        ' кошелёк, номер перевода, ID пользователя и так далее.
        clash = AF3_ClashesWithRowField(srcData, r, srcCols, cand)

        If clash <> "" Then

            rejected(cand & " совпал с полем " & clash) = 1
            GoTo NextRun

        End If

        If Not uniq.Exists(cand) Then
            uniq(cand) = starts(i)
        End If

NextRun:

    Next i

    candList = Join(uniq.keys, ", ")

    If uniq.Count = 0 Then

        If rejected.Count > 0 Then
            reason = Join(rejected.keys, "; ")
        End If

        Exit Function

    End If

    ' Правило 5. Два и более разных кандидата -
    ' макрос не угадывает, какой из них нужный.
    If uniq.Count > 1 Then

        reason = "Несколько кандидатов, выбор вручную"
        Exit Function

    End If

    keys = uniq.keys
    only = CStr(keys(0))
    onlyPos = CLng(uniq(only))

    tier = AF3_ClassifyContext(commentText, only, onlyPos)

    ' Правило 6. Категории, отключённые в настройках,
    ' уходят на проверку, а не в работу.
    If tier = "B2" And Not AF3_ACCEPT_TIER_B2 Then

        reason = "Фраза подтверждения перед номером (категория B2)"
        Exit Function

    End If

    If tier = "D" And Not AF3_ACCEPT_TIER_D Then

        reason = "Короткий текст без метки (категория D)"
        Exit Function

    End If

    If tier = "E" And Not AF3_ACCEPT_TIER_E Then

        reason = "Номер внутри свободного текста (категория E)"
        Exit Function

    End If

    AF3_ResolveFromComment = only

End Function


' ============================================================
' РАЗБОР ТЕКСТА
' ============================================================

' Собирает все непрерывные группы цифр.
' Группа обязана быть окружена нецифровыми символами,
' поэтому "123456789012" не даст кандидата длиной 11.
Private Sub AF3_ScanDigitRuns( _
    ByVal t As String, _
    ByRef runs() As String, _
    ByRef starts() As Long, _
    ByRef runCount As Long _
)

    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim buf As String
    Dim bufStart As Long

    n = Len(t)

    ReDim runs(1 To n + 1)
    ReDim starts(1 To n + 1)

    runCount = 0
    buf = ""
    bufStart = 0

    For i = 1 To n

        ch = Mid$(t, i, 1)

        If ch >= "0" And ch <= "9" Then

            If buf = "" Then bufStart = i
            buf = buf & ch

        Else

            If buf <> "" Then

                runCount = runCount + 1
                runs(runCount) = buf
                starts(runCount) = bufStart
                buf = ""

            End If

        End If

    Next i

    If buf <> "" Then

        runCount = runCount + 1
        runs(runCount) = buf
        starts(runCount) = bufStart

    End If

End Sub


' Проверяет, не является ли группа цифр дробной частью.
' Дробная часть - это когда прямо перед группой стоит
' точка или запятая, а перед разделителем стоит цифра.
Private Function AF3_IsDecimalTail( _
    ByVal t As String, _
    ByVal startPos As Long _
) As Boolean

    Dim sep As String
    Dim prev As String

    AF3_IsDecimalTail = False

    If startPos < 3 Then Exit Function

    sep = Mid$(t, startPos - 1, 1)

    If sep <> "." And sep <> "," Then Exit Function

    prev = Mid$(t, startPos - 2, 1)

    If prev >= "0" And prev <= "9" Then
        AF3_IsDecimalTail = True
    End If

End Function


' Сравнивает кандидата с остальными числовыми полями строки.
' Если номер равен кошельку, номеру перевода или ID,
' значит это не номер транзакции.
Private Function AF3_ClashesWithRowField( _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByVal cand As String _
) As String

    Dim fields As Variant
    Dim i As Long
    Dim v As String

    fields = Array( _
        "user wallet", _
        "agent's wallet", _
        "unique transfer number", _
        "user id", _
        "ticket id", _
        "subagent id", _
        "agent id", _
        "admin id", _
        "provider id", _
        "check amount", _
        "transactions amount" _
    )

    AF3_ClashesWithRowField = ""

    For i = LBound(fields) To UBound(fields)

        If srcCols.Exists(fields(i)) Then

            v = AF3_CleanNumber( _
                AF3_CellText(srcData, r, srcCols, CStr(fields(i))) _
            )

            If v <> "" Then

                ' Сравниваем и как есть, и без ведущих нулей:
                ' кошелёк может храниться как 01556889173.
                If v = cand Or AF3_StripLeadingZeros(v) = cand Then

                    AF3_ClashesWithRowField = CStr(fields(i))
                    Exit Function

                End If

            End If

        End If

    Next i

End Function


' Определяет, насколько уверенно номер стоит в тексте.
'
' A  - весь комментарий состоит только из номера
' A2 - номер открывает комментарий либо занимает
'      отдельную строку целиком
' B  - перед номером стоит слово-метка
' C  - слово-метка стоит сразу после номера
' B2 - перед номером фраза подтверждения ("accepted by" и т. п.)
' D  - короткий текст без метки
' E  - номер посреди свободного текста
Private Function AF3_ClassifyContext( _
    ByVal t As String, _
    ByVal cand As String, _
    ByVal startPos As Long _
) As String

    Dim before As String
    Dim after As String
    Dim trimmed As String

    trimmed = Trim$(t)

    If trimmed = cand Then

        AF3_ClassifyContext = "A"
        Exit Function

    End If

    before = Left$(t, startPos - 1)
    after = Mid$(t, startPos + Len(cand))

    If AF3_EndsWithLabel(before) Then

        AF3_ClassifyContext = "B"
        Exit Function

    End If

    If AF3_StartsWithLabel(after) Then

        AF3_ClassifyContext = "C"
        Exit Function

    End If

    ' Позиция сама по себе сильный признак: агент почти всегда
    ' начинает комментарий с номера либо выносит его в свою строку.
    If AF3_TrimSeparators(before) = "" Then

        AF3_ClassifyContext = "A2"
        Exit Function

    End If

    If AF3_IsWholeLine(t, startPos, Len(cand)) Then

        AF3_ClassifyContext = "A2"
        Exit Function

    End If

    ' Фраза подтверждения прямо перед номером. Проверяется после
    ' меток и позиции: те надёжнее и сохраняют свою категорию.
    If AF3_EndsWithAny(before, AF3_ConfirmPhrases) Then

        AF3_ClassifyContext = "B2"
        Exit Function

    End If

    If Len(trimmed) <= 40 Then
        AF3_ClassifyContext = "D"
    Else
        AF3_ClassifyContext = "E"
    End If

End Function


' Проверяет, что строка комментария, в которой стоит номер,
' не содержит ничего кроме этого номера.
Private Function AF3_IsWholeLine( _
    ByVal t As String, _
    ByVal startPos As Long, _
    ByVal candLen As Long _
) As Boolean

    Dim lineStart As Long
    Dim lineEnd As Long
    Dim i As Long
    Dim ch As String

    lineStart = 1

    For i = startPos - 1 To 1 Step -1

        ch = Mid$(t, i, 1)

        If ch = vbCr Or ch = vbLf Then
            lineStart = i + 1
            Exit For
        End If

    Next i

    lineEnd = Len(t)

    For i = startPos + candLen To Len(t)

        ch = Mid$(t, i, 1)

        If ch = vbCr Or ch = vbLf Then
            lineEnd = i - 1
            Exit For
        End If

    Next i

    AF3_IsWholeLine = _
        (AF3_TrimSeparatorsBoth(Mid$(t, lineStart, lineEnd - lineStart + 1)) = _
         Mid$(t, startPos, candLen))

End Function


Private Function AF3_TrimSeparatorsBoth(ByVal s As String) As String

    Do While Len(s) > 0

        If AF3_IsSeparator(Left$(s, 1)) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If

    Loop

    AF3_TrimSeparatorsBoth = AF3_TrimSeparators(s)

End Function


' Слова, после которых идёт номер транзакции.
Private Function AF3_LabelWords() As Variant

    AF3_LabelWords = Array( _
        "transaction id", "transaction", "transactions", "trans", _
        "trx", "txn", "tx", "id", "ids", _
        "deposit request", "request", "order", "payment", _
        "no", "number", "num", _
        "номер", "транзакция", "транзакции", "транзакцию", _
        "заказ", "заказа", "платеж", "платёж" _
    )

End Function


' Проверяет, заканчивается ли текст перед номером
' одним из слов-меток. Разделители вида ":", "#", "№"
' и пробелы перед номером отбрасываются.
Private Function AF3_EndsWithLabel(ByVal before As String) As Boolean

    AF3_EndsWithLabel = AF3_EndsWithAny(before, AF3_LabelWords)

End Function


' Фразы, которыми агенты подтверждают платёж прямо перед номером.
' Взяты из комментариев, ушедших в категорию E, где номер
' при проверке оказался верным. Только латиница: модуль
' хранится в CP1251, где нет французских букв с акцентами.
Private Function AF3_ConfirmPhrases() As Variant

    AF3_ConfirmPhrases = Array( _
        "accepted by", "confirmed by", "approved by", _
        "accepted", "confirmed", "approved", "approuved", _
        "transaction number is", "transaction id is", _
        "order number is", "order id is", _
        "for sure", "identification", "numero", "i'd", "n" _
    )

End Function


' Проверяет, заканчивается ли текст перед номером одним
' из слов или фраз списка. Разделители вида ":", "#", "№"
' в хвосте перед проверкой отбрасываются.
Private Function AF3_EndsWithAny( _
    ByVal before As String, _
    ByVal words As Variant _
) As Boolean

    Dim s As String
    Dim i As Long
    Dim w As String

    AF3_EndsWithAny = False

    s = AF3_TrimSeparators(LCase$(before))

    If s = "" Then Exit Function

    For i = LBound(words) To UBound(words)

        w = CStr(words(i))

        If Len(s) >= Len(w) Then

            If Right$(s, Len(w)) = w Then

                ' Метка обязана быть целым словом,
                ' иначе "paid" превратится в "id".
                If Len(s) = Len(w) Then

                    AF3_EndsWithAny = True
                    Exit Function

                ElseIf Not AF3_IsWordChar(Mid$(s, Len(s) - Len(w), 1)) Then

                    AF3_EndsWithAny = True
                    Exit Function

                End If

            End If

        End If

    Next i

End Function


Private Function AF3_StartsWithLabel(ByVal after As String) As Boolean

    Dim s As String
    Dim words As Variant
    Dim i As Long
    Dim w As String

    AF3_StartsWithLabel = False

    s = LCase$(after)

    ' Снимаем разделители в начале.
    Do While Len(s) > 0

        If AF3_IsSeparator(Left$(s, 1)) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If

    Loop

    If s = "" Then Exit Function

    words = AF3_LabelWords

    For i = LBound(words) To UBound(words)

        w = CStr(words(i))

        If Len(s) >= Len(w) Then

            If Left$(s, Len(w)) = w Then

                If Len(s) = Len(w) Then

                    AF3_StartsWithLabel = True
                    Exit Function

                ElseIf Not AF3_IsWordChar(Mid$(s, Len(w) + 1, 1)) Then

                    AF3_StartsWithLabel = True
                    Exit Function

                End If

            End If

        End If

    Next i

End Function


Private Function AF3_TrimSeparators(ByVal s As String) As String

    Do While Len(s) > 0

        If AF3_IsSeparator(Right$(s, 1)) Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If

    Loop

    AF3_TrimSeparators = s

End Function


Private Function AF3_IsSeparator(ByVal ch As String) As Boolean

    Select Case ch

        Case " ", ":", ";", "#", "-", ".", ",", "№", "=", _
             vbTab, vbCr, vbLf, Chr$(160)

            AF3_IsSeparator = True

        Case Else

            AF3_IsSeparator = False

    End Select

End Function


Private Function AF3_IsWordChar(ByVal ch As String) As Boolean

    Dim c As String

    c = LCase$(ch)

    AF3_IsWordChar = _
        (c >= "a" And c <= "z") Or _
        (c >= "0" And c <= "9") Or _
        (c >= "а" And c <= "я") Or _
        (c = "ё")

End Function


' ============================================================
' ЗАПОЛНЕНИЕ СТРОК
' ============================================================

Private Sub AF3_FillOutRow( _
    ByRef outData As Variant, _
    ByVal i As Long, _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByVal transID As String, _
    ByVal sourceTag As String, _
    ByVal candList As String _
)

    outData(i, 1) = AF3_CellValue(srcData, r, srcCols, "date created")
    outData(i, 2) = AF3_NormalizeFrom( _
        AF3_CellText(srcData, r, srcCols, "from") _
    )
    outData(i, 3) = AF3_CellValue(srcData, r, srcCols, "external status")
    outData(i, 4) = AF3_CellValue(srcData, r, srcCols, "country")
    outData(i, 5) = AF3_CellValue(srcData, r, srcCols, "ticket id")

    ' 6 Transaction status и ADM колонки
    ' остаются пустыми - их заполняет Approved_Перенос2.
    outData(i, 6) = ""
    outData(i, 7) = ""

    outData(i, 8) = transID

    outData(i, 9) = ""
    outData(i, 10) = AF3_CellValue(srcData, r, srcCols, "user id")

    outData(i, 11) = ""
    outData(i, 12) = AF3_CellValue(srcData, r, srcCols, "check amount")

    outData(i, 13) = AF3_CellValue(srcData, r, srcCols, "ticket currency")

    outData(i, 14) = ""
    outData(i, 15) = AF3_CellValue(srcData, r, srcCols, "subagent id")

    outData(i, 16) = AF3_CellValue(srcData, r, srcCols, "subagent")
    outData(i, 17) = AF3_CellValue(srcData, r, srcCols, "agent's wallet")
    outData(i, 18) = AF3_CellValue(srcData, r, srcCols, "user wallet")
    outData(i, 19) = AF3_CellValue(srcData, r, srcCols, "unique transfer number")
    outData(i, 20) = AF3_CellValue(srcData, r, srcCols, "topic")
    outData(i, 21) = AF3_CellValue(srcData, r, srcCols, "ticket type")
    outData(i, 22) = AF3_CellValue(srcData, r, srcCols, "department")
    outData(i, 23) = AF3_CellValue(srcData, r, srcCols, "agent")
    outData(i, 24) = AF3_CellValue(srcData, r, srcCols, "internal comment")

    outData(i, 25) = sourceTag
    outData(i, 26) = candList
    outData(i, 27) = ""

End Sub


Private Sub AF3_MoveOutRow( _
    ByRef outData As Variant, _
    ByVal fromRow As Long, _
    ByVal toRow As Long _
)

    Dim c As Long

    For c = 1 To 27
        outData(toRow, c) = outData(fromRow, c)
    Next c

End Sub


Private Sub AF3_FillReviewRow( _
    ByRef revData As Variant, _
    ByVal i As Long, _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByVal candList As String, _
    ByVal reason As String _
)

    revData(i, 1) = AF3_CellValue(srcData, r, srcCols, "ticket id")
    revData(i, 2) = AF3_CellValue(srcData, r, srcCols, "date created")
    revData(i, 3) = AF3_CellValue(srcData, r, srcCols, "country")
    revData(i, 4) = AF3_CellValue(srcData, r, srcCols, "user id")
    revData(i, 5) = AF3_CellValue(srcData, r, srcCols, "check amount")
    revData(i, 6) = AF3_CellValue(srcData, r, srcCols, "ticket currency")
    revData(i, 7) = AF3_CellValue(srcData, r, srcCols, "subagent")
    revData(i, 8) = AF3_CellValue(srcData, r, srcCols, "internal comment")
    revData(i, 9) = candList
    revData(i, 10) = reason

End Sub


Private Sub AF3_FillReviewRowFromOut( _
    ByRef revData As Variant, _
    ByVal i As Long, _
    ByRef outData As Variant, _
    ByVal outRow As Long, _
    ByVal reason As String _
)

    revData(i, 1) = outData(outRow, 5)
    revData(i, 2) = outData(outRow, 1)
    revData(i, 3) = outData(outRow, 4)
    revData(i, 4) = outData(outRow, 10)
    revData(i, 5) = outData(outRow, 12)
    revData(i, 6) = outData(outRow, 13)
    revData(i, 7) = outData(outRow, 16)
    revData(i, 8) = outData(outRow, 24)
    revData(i, 9) = CStr(outData(outRow, 8)) & _
        " (" & CStr(outData(outRow, 25)) & ")"
    revData(i, 10) = reason

End Sub


' ============================================================
' ЛИСТЫ И ФОРМАТИРОВАНИЕ
' ============================================================

Private Sub AF3_WriteHeaders(ByVal ws As Worksheet)

    Dim h As Variant

    h = Array( _
        "Date created", "From", "External Status", "Country", "Ticket ID", _
        "Transaction status", "ADM Transaction ID", "Transaction ID", _
        "ADM User ID", "User ID", _
        "ADM Check amount", "Check amount", "Ticket currency", _
        "ADM Subagent ID", "Subagent ID", _
        "Subagent", "Agent's wallet", "User wallet", "Unique transfer number", _
        "Topic", "Ticket type", "Department", "Agent", "Internal comment", _
        "Источник номера", "Кандидаты из комментария", "Примечание" _
    )

    ws.Range( _
        ws.Cells(AF3_FILTER_HEADER_ROW, 1), _
        ws.Cells(AF3_FILTER_HEADER_ROW, 27) _
    ).Value = h

End Sub


Private Sub AF3_WriteReviewHeaders(ByVal ws As Worksheet)

    Dim h As Variant

    h = Array( _
        "Ticket ID", "Date created", "Country", "User ID", _
        "Check amount", "Ticket currency", "Subagent", _
        "Internal comment", "Кандидаты", "Причина" _
    )

    ws.Range( _
        ws.Cells(AF3_REVIEW_HEADER_ROW, 1), _
        ws.Cells(AF3_REVIEW_HEADER_ROW, 10) _
    ).Value = h

End Sub


Private Sub AF3_SortByTransaction( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long _
)

    Dim lastRow As Long

    lastRow = AF3_FILTER_DATA_ROW + rowCount - 1

    ws.Sort.SortFields.Clear

    ws.Sort.SortFields.Add _
        Key:=ws.Range( _
            ws.Cells(AF3_FILTER_DATA_ROW, 8), _
            ws.Cells(lastRow, 8) _
        ), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal

    With ws.Sort

        .SetRange ws.Range( _
            ws.Cells(AF3_FILTER_HEADER_ROW, 1), _
            ws.Cells(lastRow, 27) _
        )

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

End Sub


Private Sub AF3_FormatFilter(ByVal ws As Worksheet)

    ws.Rows(AF3_FILTER_HEADER_ROW).Font.Bold = True
    ws.Columns.AutoFit

    ' Комментарий бывает многострочным, поэтому ширина фиксируется.
    ws.Columns(24).ColumnWidth = 20
    ws.Columns(25).ColumnWidth = 16
    ws.Columns(26).ColumnWidth = 22
    ws.Columns(27).ColumnWidth = 44

    ' Номера транзакций показываем без экспоненты.
    ws.Columns(7).NumberFormat = "0"
    ws.Columns(8).NumberFormat = "0"

End Sub


' Раскладка повторяет ежедневный отчёт: подписи в колонке A,
' числа в B, строки 3-15 те же, что в отчёте. Блок скрипта
' (строки 17 и ниже) не пишется: его заполняет другой инструмент.
' Строки 14-15 заполнит Approved_Перенос2_v3.
Private Sub AF3_WriteReport( _
    ByVal ws As Worksheet, _
    ByVal sourceSheetName As String, _
    ByVal totalRows As Long, _
    ByVal withTID As Long, _
    ByVal outCount As Long, _
    ByVal statKept As Long, _
    ByVal statFromComment As Long, _
    ByVal statDropped As Long _
)

    Dim labels(1 To 15, 1 To 1) As Variant
    Dim values(1 To 15, 1 To 1) As Variant

    labels(1, 1) = "Дата прогона"
    values(1, 1) = Now

    labels(2, 1) = "Лист выгрузки"
    values(2, 1) = sourceSheetName

    labels(3, 1) = "Кол-во тикетов в статусе Approved (M)"
    values(3, 1) = totalRows

    labels(4, 1) = "С транзакциями"
    values(4, 1) = withTID

    labels(5, 1) = "Без транзакций"

    labels(7, 1) = "Блок результатов после первого макроса"

    labels(8, 1) = "Прошло на этап проверки с админкой"
    values(8, 1) = outCount

    labels(9, 1) = "          - из поля Transaction ID"
    values(9, 1) = statKept

    labels(10, 1) = "          - из поля Inside comment"
    values(10, 1) = statFromComment

    labels(11, 1) = "Нет транзакции в обоих полях"
    values(11, 1) = statDropped

    labels(13, 1) = "Блок результатов после второго макроса"
    labels(14, 1) = "Из поля Transaction ID"
    labels(15, 1) = "Из поля Inside comment"

    ws.Range("A1:A15").Value = labels
    ws.Range("B1:B15").Value = values

    ' Формула, а не число: при вставке в отчёт она
    ' сама перестроится на столбец нужного дня.
    ws.Range("B5").Formula = "=B3-B4"

    ws.Range("A1:B2").Font.Bold = True
    ws.Range("A7,A13").Font.Bold = True

    ws.Range("B1").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Range("B3:B15").NumberFormat = "0"
    ws.Range("B1:B15").HorizontalAlignment = xlRight

    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").AutoFit

End Sub


Private Sub AF3_FormatReview(ByVal ws As Worksheet)

    ws.Rows(AF3_REVIEW_HEADER_ROW).Font.Bold = True
    ws.Columns.AutoFit

    ws.Columns(8).ColumnWidth = 60
    ws.Columns(9).ColumnWidth = 26
    ws.Columns(10).ColumnWidth = 40

    ws.Columns(8).WrapText = False

End Sub


' ============================================================
' КНОПКА КОПИРОВАНИЯ
' ============================================================

Private Sub AF3_AddButtons(ByVal ws As Worksheet)

    Dim transactionCol As Long
    Dim lastRow As Long
    Dim totalRows As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim startDataRow As Long
    Dim endDataRow As Long
    Dim buttonIndex As Long
    Dim buttonCount As Long
    Dim buttonsPerStack As Long
    Dim stackCol As Long
    Dim stackRow As Long
    Dim visibleStackRows As Long

    Dim shp As Shape
    Dim baseCell As Range
    Dim macroWorkbookName As String
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim buttonGap As Double
    Dim buttonPitch As Double
    Dim requiredRowHeight As Double
    Dim i As Long

    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, 25) = "btnApprovedCopyTransBlock" Or _
           shp.Name = "btnApprovedCopyTransactions" Then
            shp.Delete
        End If
    Next i

    On Error GoTo 0

    transactionCol = AF3_FindHeaderCol(ws, "Transaction ID", AF3_FILTER_HEADER_ROW)

    If transactionCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, transactionCol).End(xlUp).Row

    If lastRow < AF3_FILTER_DATA_ROW Then Exit Sub

    totalRows = lastRow - AF3_FILTER_DATA_ROW + 1

    Set baseCell = ws.Cells(1, transactionCol)

    macroWorkbookName = Replace(ThisWorkbook.Name, "'", "''")

    buttonWidth = 112
    buttonHeight = 20
    buttonGap = 4
    buttonPitch = buttonHeight + buttonGap

    buttonCount = (totalRows + AF3_COPY_BLOCK_SIZE - 1) \ AF3_COPY_BLOCK_SIZE

    ' Кнопки идут стопкой вниз от колонки Transaction ID (обычно H).
    ' В одну стопку помещаем столько кнопок, сколько безопасно влезает
    ' в максимальную высоту строки Excel. При очень большом списке
    ' следующая стопка начнётся правее, чтобы не перекрывать данные.
    buttonsPerStack = Int((409 - 4) / buttonPitch)
    If buttonsPerStack < 1 Then buttonsPerStack = 1

    visibleStackRows = buttonCount
    If visibleStackRows > buttonsPerStack Then visibleStackRows = buttonsPerStack

    requiredRowHeight = 4 + visibleStackRows * buttonPitch
    If requiredRowHeight < 24 Then requiredRowHeight = 24
    If requiredRowHeight > 409 Then requiredRowHeight = 409

    ws.Rows(1).RowHeight = requiredRowHeight

    startIndex = 1
    buttonIndex = 1

    Do While startIndex <= totalRows

        endIndex = Application.Min( _
            startIndex + AF3_COPY_BLOCK_SIZE - 1, totalRows _
        )

        startDataRow = AF3_FILTER_DATA_ROW + startIndex - 1
        endDataRow = AF3_FILTER_DATA_ROW + endIndex - 1

        stackCol = (buttonIndex - 1) \ buttonsPerStack
        stackRow = (buttonIndex - 1) Mod buttonsPerStack

        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            baseCell.Left + 2 + stackCol * (buttonWidth + buttonGap), _
            baseCell.Top + 2 + stackRow * buttonPitch, _
            buttonWidth, _
            buttonHeight _
        )

        With shp

            .Name = "btnApprovedCopyTransBlock_" & CStr(buttonIndex)

            .AlternativeText = _
                CStr(startDataRow) & "|" & _
                CStr(endDataRow) & "|" & _
                CStr(transactionCol)

            .OnAction = _
                "'" & macroWorkbookName & "'!Approved_UI_CopyBlock_v3"

            .Fill.ForeColor.RGB = RGB(31, 78, 121)
            .Line.ForeColor.RGB = RGB(31, 78, 121)

            .TextFrame2.TextRange.Text = _
                "Копировать " & CStr(startIndex) & "-" & CStr(endIndex)

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

        startIndex = endIndex + 1
        buttonIndex = buttonIndex + 1

    Loop

End Sub


' Function используется специально, чтобы служебный обработчик
' не отображался в Alt+F8.
Public Function Approved_UI_CopyBlock_v3() As Variant

    Dim ws As Worksheet
    Dim shp As Shape
    Dim params As Variant

    Dim startDataRow As Long
    Dim endDataRow As Long
    Dim transactionCol As Long

    On Error GoTo CopyFail

    Set ws = ActiveSheet

    If LCase$(Trim$(ws.Name)) <> LCase$("Фильтр") Then

        MsgBox "Кнопка работает на листе 'Фильтр'.", vbExclamation
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
    transactionCol = CLng(params(2))

    ws.Range( _
        ws.Cells(startDataRow, transactionCol), _
        ws.Cells(endDataRow, transactionCol) _
    ).Copy

    Application.StatusBar = _
        "Скопировано Transaction ID: " & _
        CStr(startDataRow - AF3_FILTER_DATA_ROW + 1) & "-" & _
        CStr(endDataRow - AF3_FILTER_DATA_ROW + 1)

    Exit Function

CopyFail:

    Application.StatusBar = False

    MsgBox _
        "Не удалось скопировать Transaction ID: " & Err.Description, _
        vbExclamation

End Function


Private Sub AF3_NormalizeRows( _
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

    ' Повторяем итог ручного переключения "Переносить текст":
    ' данные остаются в одну строку, а случайно растянутые строки сбрасываются.
    ws.UsedRange.WrapText = False

    If lastRow >= dataStartRow Then
        ws.Rows(dataStartRow & ":" & lastRow).RowHeight = ws.StandardHeight
    End If

End Sub


' ============================================================
' СЛУЖЕБНЫЕ ФУНКЦИИ
' ============================================================

Private Function AF3_BuildHeaderMap( _
    ByRef srcData As Variant, _
    ByVal lastCol As Long _
) As Object

    Dim d As Object
    Dim c As Long
    Dim h As String

    Set d = CreateObject("Scripting.Dictionary")

    For c = 1 To lastCol

        h = Trim$(LCase$(CStr(srcData(1, c))))

        If h <> "" Then

            ' Первое вхождение выигрывает: в выгрузке
            ' встречаются похожие заголовки.
            If Not d.Exists(h) Then d(h) = c

        End If

    Next c

    Set AF3_BuildHeaderMap = d

End Function


Private Function AF3_CellValue( _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByVal headerName As String _
) As Variant

    If srcCols.Exists(headerName) Then
        AF3_CellValue = srcData(r, srcCols(headerName))
    Else
        AF3_CellValue = ""
    End If

End Function


Private Function AF3_CellText( _
    ByRef srcData As Variant, _
    ByVal r As Long, _
    ByVal srcCols As Object, _
    ByVal headerName As String _
) As String

    Dim v As Variant

    v = AF3_CellValue(srcData, r, srcCols, headerName)

    If IsError(v) Then
        AF3_CellText = ""
    ElseIf IsNull(v) Then
        AF3_CellText = ""
    Else
        AF3_CellText = CStr(v)
    End If

End Function


' Приводит числовое поле к строке из одних цифр.
' Нужно, чтобы 1 810 353 629 и 1810353629 сравнивались одинаково.
Private Function AF3_CleanNumber(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim res As String

    s = Trim$(s)

    ' Экспоненциальная запись означает, что Excel
    ' уже потерял точность. Такое значение не используем.
    If InStr(1, s, "E", vbTextCompare) > 0 Then

        AF3_CleanNumber = ""
        Exit Function

    End If

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            res = res & ch
        ElseIf ch = " " Or ch = Chr$(160) Then
            ' разделитель разрядов пропускаем
        Else
            ' любой другой символ означает, что это не чистое число
            AF3_CleanNumber = ""
            Exit Function
        End If

    Next i

    AF3_CleanNumber = res

End Function


Private Function AF3_StripLeadingZeros(ByVal s As String) As String

    Do While Len(s) > 1

        If Left$(s, 1) = "0" Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If

    Loop

    AF3_StripLeadingZeros = s

End Function


Private Function AF3_NormalizeFrom(ByVal s As String) As String

    Dim v As String

    If Not AF3_NORMALIZE_FROM Then

        AF3_NormalizeFrom = s
        Exit Function

    End If

    v = Trim$(LCase$(s))

    Select Case v

        Case "through customer support", "by customer support"
            AF3_NormalizeFrom = "PS"

        Case "by yourself", "by customer"
            AF3_NormalizeFrom = "User"

        Case Else
            AF3_NormalizeFrom = s

    End Select

End Function


Private Function AF3_IsServiceSheet(ByVal sheetName As String) As Boolean

    Select Case LCase$(Trim$(sheetName))

        Case LCase$("Фильтр"), LCase$("Источник"), _
             LCase$("На проверку"), LCase$("Для копирования"), _
             LCase$("Закрыть статусом 225"), _
             LCase$("Из выгрузки"), LCase$("Из комментария"), _
             LCase$(AF3_SHEET_REPORT), LCase$(AF3_SHEET_SUMMARY_OLD)

            AF3_IsServiceSheet = True

        Case Else

            AF3_IsServiceSheet = False

    End Select

End Function


Private Function AF3_FindHeaderCol( _
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

            AF3_FindHeaderCol = c
            Exit Function

        End If

    Next c

    AF3_FindHeaderCol = 0

End Function


Private Sub AF3_DeleteSheetIfExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
)

    On Error Resume Next
    wb.Worksheets(sheetName).Delete
    On Error GoTo 0

End Sub


Private Sub AF3_OpenFilter(ByVal ws As Worksheet)

    ws.Activate
    ws.Range("A1").Select

    On Error Resume Next

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    On Error GoTo 0

End Sub
