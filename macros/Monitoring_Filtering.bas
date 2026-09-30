Sub Monitoring_Фильтрация()

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim headers As Variant
    Dim colIndex As Long
    Dim i As Long
    Dim targetCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerFound As Boolean
    Dim fromCol As Long
    Dim procDateCol As Long
    Dim procTimeCol As Long
    Dim cell As Range
    Dim currentTime As Date
    Dim procDate As Variant
    Dim diffHours As Double
    
    ' Исходный лист
    Set wsSource = ActiveSheet
    
    ' Удаляем старый лист Filtered, если есть
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Filtered").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Создаём новый лист
    Set wsTarget = Worksheets.Add
    wsTarget.Name = "Filtered"
    
    ' Находим последнюю строку и столбец
    lastRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    lastCol = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
    
    ' Список нужных колонок (Internal comment — в конце)
    headers = Array("Date created", "Processing date", "Processing time", "From", "External Status", _
                    "Country", "Ticket ID", "Transaction ID", "User ID", "Check amount", _
                    "Ticket currency", "Subagent ID", "Subagent", "Agent's wallet", "User wallet", _
                    "Unique transfer number", "Topic", "Ticket type", "Department", "Agent", "Internal comment")
    
    targetCol = 1
    
    ' Копируем нужные колонки по порядку
    For i = LBound(headers) To UBound(headers)
        headerFound = False
        
        For colIndex = 1 To lastCol
            If Trim(LCase(wsSource.Cells(1, colIndex).Value)) = Trim(LCase(headers(i))) Then
                wsSource.Range(wsSource.Cells(1, colIndex), wsSource.Cells(lastRow, colIndex)).Copy
                wsTarget.Cells(1, targetCol).PasteSpecial xlPasteValues
                wsTarget.Cells(1, targetCol).PasteSpecial xlPasteFormats
                headerFound = True
                
                ' Запоминаем позиции нужных колонок
                Select Case headers(i)
                    Case "From": fromCol = targetCol
                    Case "Processing date": procDateCol = targetCol
                    Case "Processing time": procTimeCol = targetCol
                End Select
                
                Exit For
            End If
        Next colIndex
        
        ' Если колонки нет — создаём пустую
        If Not headerFound Then
            wsTarget.Cells(1, targetCol).Value = headers(i)
            Select Case headers(i)
                Case "From": fromCol = targetCol
                Case "Processing date": procDateCol = targetCol
                Case "Processing time": procTimeCol = targetCol
            End Select
        End If
        
        targetCol = targetCol + 1
    Next i
    
    ' Форматируем заголовки
    wsTarget.Rows(1).Font.Bold = True
    wsTarget.Columns.AutoFit
    
    ' Устанавливаем ширину 20 для "Internal comment"
    For Each cell In wsTarget.Rows(1).Cells
        If Trim(LCase(cell.Value)) = "internal comment" Then
            wsTarget.Columns(cell.Column).ColumnWidth = 20
            Exit For
        End If
    Next cell
    
    ' --- ЗАМЕНА ЗНАЧЕНИЙ В КОЛОНКЕ "From" ---
    If fromCol > 0 Then
        For i = 2 To wsTarget.Cells(wsTarget.Rows.count, fromCol).End(xlUp).row
            Select Case Trim(LCase(wsTarget.Cells(i, fromCol).Value))
                Case "through customer support"
                    wsTarget.Cells(i, fromCol).Value = "PS"
                Case "by yourself"
                    wsTarget.Cells(i, fromCol).Value = "User"
            End Select
        Next i
    End If
    
    ' --- ВЫЧИСЛЕНИЕ Processing time (в часах) ---
    If procDateCol > 0 And procTimeCol > 0 Then
        currentTime = Now
        
        For i = 2 To wsTarget.Cells(wsTarget.Rows.count, procDateCol).End(xlUp).row
            procDate = wsTarget.Cells(i, procDateCol).Value
            
            If IsDate(procDate) Then
                diffHours = (currentTime - CDate(procDate)) * 24
                If diffHours < 0 Then diffHours = 0 ' на случай будущих дат
                wsTarget.Cells(i, procTimeCol).Value = Round(diffHours, 2)
            Else
                wsTarget.Cells(i, procTimeCol).Value = ""
            End If
        Next i
    End If
    
    MsgBox "? Готово! Всё выполнено: колонки переставлены, 'Internal comment' в конце (ширина=20), 'From' заменён, 'Processing time' рассчитан.", vbInformation

End Sub




