Option Explicit

Sub Monitoring_SFP_SMP_NR72ч()

    Dim wsSource As Worksheet
    Dim wsFiltered As Worksheet
    Dim sheetNames As Variant
    Dim i As Long
    
    sheetNames = Array("Filtered", _
                       "SFP (M) 72h+", _
                       "New request (M) 72h+", _
                       "SMP SFP 72h+", _
                       "SMP NR 72h+", _
                       "SMP подсчет 72h+", _
                       "Для копирования")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        If ActiveSheet.Name = CStr(sheetNames(i)) Then
            MsgBox "Запусти макрос с исходного листа выгрузки, а не с итогового листа.", vbExclamation
            Exit Sub
        End If
    Next i
    
    Set wsSource = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    For i = LBound(sheetNames) To UBound(sheetNames)
        Worksheets(CStr(sheetNames(i))).Delete
    Next i
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    Set wsFiltered = CreateFilteredSheet(wsSource, "Filtered")
    
    CreateBTStatusSheet wsFiltered, "SFP (M) 72h+", "Sent for processing (M)", 72
    CreateBTStatusSheet wsFiltered, "New request (M) 72h+", "New request (M)", 72
    
    CreateSMPStatusSheet wsFiltered, "SMP SFP 72h+", "Sent for processing", 72
    CreateSMPStatusSheet wsFiltered, "SMP NR 72h+", "New request", 72
    
    CreateSMPSummarySheet wsFiltered, "SMP подсчет 72h+", 72
    CreateCopySheet
    
    Application.ScreenUpdating = True
    
    MsgBox "Готово! Макрос отработал: данные отфильтрованы, 72h+ посчитаны, лист для копирования создан.", vbInformation

End Sub

Private Function CreateFilteredSheet(ByVal wsSource As Worksheet, ByVal resultSheetName As String) As Worksheet

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
    
    Dim currentTime As Date
    Dim procDate As Variant
    Dim parsedDate As Date
    Dim diffHours As Double
    
    headers = Array("Date created", "Processing date", "Processing time", "From", "External Status", _
                    "Country", "Ticket ID", "Transaction ID", "User ID", "Check amount", _
                    "Ticket currency", "Agent ID", "Referal", "Subagent ID", "Subagent", _
                    "Agent's wallet", "User wallet", "Unique transfer number", "Topic", _
                    "Ticket type", "Department", "Agent", "Internal comment")
    
    Set wsTarget = Worksheets.Add(After:=wsSource)
    wsTarget.Name = resultSheetName
    
    lastRow = GetLastRow(wsSource)
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    targetCol = 1
    
    For i = LBound(headers) To UBound(headers)
        headerFound = False
        
        For colIndex = 1 To lastCol
            If NormalizeText(wsSource.Cells(1, colIndex).Value) = NormalizeText(headers(i)) Then
                
                wsSource.Range(wsSource.Cells(1, colIndex), wsSource.Cells(lastRow, colIndex)).Copy
                wsTarget.Cells(1, targetCol).PasteSpecial xlPasteValues
                
                headerFound = True
                
                Select Case CStr(headers(i))
                    Case "From"
                        fromCol = targetCol
                    Case "Processing date"
                        procDateCol = targetCol
                    Case "Processing time"
                        procTimeCol = targetCol
                End Select
                
                Exit For
            End If
        Next colIndex
        
        If Not headerFound Then
            wsTarget.Cells(1, targetCol).Value = headers(i)
            
            Select Case CStr(headers(i))
                Case "From"
                    fromCol = targetCol
                Case "Processing date"
                    procDateCol = targetCol
                Case "Processing time"
                    procTimeCol = targetCol
            End Select
        End If
        
        targetCol = targetCol + 1
    Next i
    
    Application.CutCopyMode = False
    
    wsTarget.Rows(1).Font.Bold = True
    
    If fromCol > 0 Then
        For i = 2 To GetLastRow(wsTarget)
            Select Case NormalizeText(wsTarget.Cells(i, fromCol).Value)
                Case NormalizeText("through customer support")
                    wsTarget.Cells(i, fromCol).Value = "PS"
                Case NormalizeText("by yourself")
                    wsTarget.Cells(i, fromCol).Value = "User"
            End Select
        Next i
    End If
    
    If procDateCol > 0 And procTimeCol > 0 Then
        currentTime = Now
        
        For i = 2 To GetLastRow(wsTarget)
            procDate = wsTarget.Cells(i, procDateCol).Value
            
            If TryParseDate(procDate, parsedDate) Then
                diffHours = (currentTime - parsedDate) * 24
                If diffHours < 0 Then diffHours = 0
                
                wsTarget.Cells(i, procTimeCol).Value = Round(diffHours, 2)
            Else
                wsTarget.Cells(i, procTimeCol).Value = ""
            End If
        Next i
        
        wsTarget.Columns(procTimeCol).NumberFormat = "0.00"
    End If
    
    wsTarget.Cells.WrapText = False
    wsTarget.Columns.AutoFit
    
    If wsTarget.AutoFilterMode Then wsTarget.AutoFilterMode = False
    wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(1, UBound(headers) + 1)).AutoFilter
    
    Set CreateFilteredSheet = wsTarget

End Function

Private Sub CreateBTStatusSheet(ByVal wsFiltered As Worksheet, _
                                ByVal resultSheetName As String, _
                                ByVal statusFilter As String, _
                                ByVal minHours As Double)

    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long
    
    Dim countryCol As Long
    Dim ticketCol As Long
    Dim subagentCol As Long
    Dim statusCol As Long
    Dim procTimeCol As Long
    
    Dim currentStatus As String
    Dim currentTime As Variant
    
    countryCol = GetHeaderCol(wsFiltered, "Country")
    ticketCol = GetHeaderCol(wsFiltered, "Ticket ID")
    subagentCol = GetHeaderCol(wsFiltered, "Subagent")
    statusCol = GetHeaderCol(wsFiltered, "External Status")
    procTimeCol = GetHeaderCol(wsFiltered, "Processing time")
    
    If countryCol = 0 Or ticketCol = 0 Or subagentCol = 0 Or statusCol = 0 Or procTimeCol = 0 Then
        MsgBox "Не найдены обязательные колонки для листа: " & resultSheetName, vbCritical
        Exit Sub
    End If
    
    Set wsTarget = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsTarget.Name = resultSheetName
    
    WriteResultHeaders wsTarget
    
    lastRow = GetLastRow(wsFiltered)
    targetRow = 2
    
    For sourceRow = 2 To lastRow
        
        currentStatus = NormalizeText(wsFiltered.Cells(sourceRow, statusCol).Value)
        currentTime = wsFiltered.Cells(sourceRow, procTimeCol).Value
        
        If currentStatus = NormalizeText(statusFilter) Then
            If IsNumeric(currentTime) Then
                If CDbl(currentTime) >= minHours Then
                    wsTarget.Cells(targetRow, 1).Value = wsFiltered.Cells(sourceRow, countryCol).Value
                    wsTarget.Cells(targetRow, 2).Value = wsFiltered.Cells(sourceRow, ticketCol).Value
                    wsTarget.Cells(targetRow, 3).Value = wsFiltered.Cells(sourceRow, subagentCol).Value
                    targetRow = targetRow + 1
                End If
            End If
        End If
        
    Next sourceRow
    
    FormatResultSheet wsTarget

End Sub

Private Sub CreateSMPStatusSheet(ByVal wsFiltered As Worksheet, _
                                 ByVal resultSheetName As String, _
                                 ByVal statusFilter As String, _
                                 ByVal minHours As Double)

    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long
    
    Dim countryCol As Long
    Dim ticketCol As Long
    Dim subagentCol As Long
    Dim statusCol As Long
    Dim procTimeCol As Long
    Dim agentIdCol As Long
    Dim referalCol As Long
    
    Dim currentStatus As String
    Dim currentTime As Variant
    Dim currentCountry As String
    Dim currentAgentID As String
    Dim currentReferal As String
    
    countryCol = GetHeaderCol(wsFiltered, "Country")
    ticketCol = GetHeaderCol(wsFiltered, "Ticket ID")
    subagentCol = GetHeaderCol(wsFiltered, "Subagent")
    statusCol = GetHeaderCol(wsFiltered, "External Status")
    procTimeCol = GetHeaderCol(wsFiltered, "Processing time")
    agentIdCol = GetHeaderCol(wsFiltered, "Agent ID")
    referalCol = GetHeaderCol(wsFiltered, "Referal")
    
    If countryCol = 0 Or ticketCol = 0 Or subagentCol = 0 Or statusCol = 0 Or procTimeCol = 0 Or agentIdCol = 0 Or referalCol = 0 Then
        MsgBox "Не найдены обязательные колонки для SMP: Country, Ticket ID, Subagent, External Status, Processing time, Agent ID или Referal.", vbCritical
        Exit Sub
    End If
    
    Set wsTarget = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsTarget.Name = resultSheetName
    
    WriteResultHeaders wsTarget
    
    lastRow = GetLastRow(wsFiltered)
    targetRow = 2
    
    For sourceRow = 2 To lastRow
        
        currentCountry = CStr(wsFiltered.Cells(sourceRow, countryCol).Value)
        currentStatus = NormalizeText(wsFiltered.Cells(sourceRow, statusCol).Value)
        currentAgentID = NormalizeText(wsFiltered.Cells(sourceRow, agentIdCol).Value)
        currentReferal = NormalizeText(wsFiltered.Cells(sourceRow, referalCol).Value)
        currentTime = wsFiltered.Cells(sourceRow, procTimeCol).Value
        
        If IsSMPAgentID(currentAgentID) Then
            If IsSMPCountry(currentCountry) Then
                If IsAllowedReferal(currentReferal) Then
                    If currentStatus = NormalizeText(statusFilter) Then
                        If IsNumeric(currentTime) Then
                            If CDbl(currentTime) >= minHours Then
                                wsTarget.Cells(targetRow, 1).Value = wsFiltered.Cells(sourceRow, countryCol).Value
                                wsTarget.Cells(targetRow, 2).Value = wsFiltered.Cells(sourceRow, ticketCol).Value
                                wsTarget.Cells(targetRow, 3).Value = wsFiltered.Cells(sourceRow, subagentCol).Value
                                targetRow = targetRow + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    Next sourceRow
    
    FormatResultSheet wsTarget

End Sub

Private Sub CreateSMPSummarySheet(ByVal wsFiltered As Worksheet, _
                                  ByVal resultSheetName As String, _
                                  ByVal minHours As Double)

    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim sourceRow As Long
    
    Dim countryCol As Long
    Dim statusCol As Long
    Dim procTimeCol As Long
    Dim agentIdCol As Long
    Dim referalCol As Long
    
    Dim currentCountry As String
    Dim currentStatus As String
    Dim currentAgentID As String
    Dim currentReferal As String
    Dim currentTime As Variant
    
    Dim countSFP As Long
    Dim countNR As Long
    
    countryCol = GetHeaderCol(wsFiltered, "Country")
    statusCol = GetHeaderCol(wsFiltered, "External Status")
    procTimeCol = GetHeaderCol(wsFiltered, "Processing time")
    agentIdCol = GetHeaderCol(wsFiltered, "Agent ID")
    referalCol = GetHeaderCol(wsFiltered, "Referal")
    
    If countryCol = 0 Or statusCol = 0 Or procTimeCol = 0 Or agentIdCol = 0 Or referalCol = 0 Then
        MsgBox "Не найдены обязательные колонки для листа SMP подсчет 72h+.", vbCritical
        Exit Sub
    End If
    
    lastRow = GetLastRow(wsFiltered)
    
    For sourceRow = 2 To lastRow
        
        currentCountry = CStr(wsFiltered.Cells(sourceRow, countryCol).Value)
        currentStatus = NormalizeText(wsFiltered.Cells(sourceRow, statusCol).Value)
        currentAgentID = NormalizeText(wsFiltered.Cells(sourceRow, agentIdCol).Value)
        currentReferal = NormalizeText(wsFiltered.Cells(sourceRow, referalCol).Value)
        currentTime = wsFiltered.Cells(sourceRow, procTimeCol).Value
        
        If IsSMPAgentID(currentAgentID) Then
            If IsSMPCountry(currentCountry) Then
                If IsAllowedReferal(currentReferal) Then
                    If IsNumeric(currentTime) Then
                        If CDbl(currentTime) >= minHours Then
                            If currentStatus = NormalizeText("Sent for processing") Then
                                countSFP = countSFP + 1
                            ElseIf currentStatus = NormalizeText("New request") Then
                                countNR = countNR + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    Next sourceRow
    
    Set wsTarget = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsTarget.Name = resultSheetName
    
    wsTarget.Cells(1, 1).Value = "Показатель"
    wsTarget.Cells(1, 2).Value = "Количество"
    
    wsTarget.Cells(2, 1).Value = "SMP Sent for processing 72h+"
    wsTarget.Cells(2, 2).Value = countSFP
    
    wsTarget.Cells(3, 1).Value = "SMP New request 72h+"
    wsTarget.Cells(3, 2).Value = countNR
    
    wsTarget.Cells(4, 1).Value = "SMP Total 72h+"
    wsTarget.Cells(4, 2).Value = countSFP + countNR
    
    wsTarget.Rows(1).Font.Bold = True
    wsTarget.Cells.WrapText = False
    wsTarget.Columns.AutoFit

End Sub

Private Sub CreateCopySheet()

    Dim wsCopy As Worksheet
    Dim nextRow As Long
    
    Set wsCopy = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsCopy.Name = "Для копирования"
    
    nextRow = 1
    
    CopyResultBlockToSheet "SFP (M) 72h+", wsCopy, nextRow, "Sent for processing (M) 72h+"
    CopyResultBlockToSheet "New request (M) 72h+", wsCopy, nextRow, "New request (M) 72h+"
    CopyResultBlockToSheet "SMP SFP 72h+", wsCopy, nextRow, "SMP Sent for processing 72h+"
    CopyResultBlockToSheet "SMP NR 72h+", wsCopy, nextRow, "SMP New request 72h+"
    
    wsCopy.Cells.WrapText = False
    wsCopy.Columns.AutoFit

End Sub

Private Sub CopyResultBlockToSheet(ByVal sourceSheetName As String, _
                                   ByVal wsCopy As Worksheet, _
                                   ByRef nextRow As Long, _
                                   ByVal blockTitle As String)

    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Set wsSource = Worksheets(sourceSheetName)
    lastRow = GetLastRow(wsSource)
    
    wsCopy.Cells(nextRow, 1).Value = blockTitle
    wsCopy.Cells(nextRow, 1).Font.Bold = True
    nextRow = nextRow + 1
    
    wsCopy.Cells(nextRow, 1).Value = "Гео"
    wsCopy.Cells(nextRow, 2).Value = "Ticket ID"
    wsCopy.Cells(nextRow, 3).Value = "Subagent"
    wsCopy.Rows(nextRow).Font.Bold = True
    nextRow = nextRow + 1
    
    If lastRow >= 2 Then
        For r = 2 To lastRow
            wsCopy.Cells(nextRow, 1).Value = wsSource.Cells(r, 1).Value
            wsCopy.Cells(nextRow, 2).Value = wsSource.Cells(r, 2).Value
            wsCopy.Cells(nextRow, 3).Value = wsSource.Cells(r, 3).Value
            nextRow = nextRow + 1
        Next r
    Else
        wsCopy.Cells(nextRow, 1).Value = "Нет тикетов"
        nextRow = nextRow + 1
    End If
    
    nextRow = nextRow + 2

End Sub

Private Sub WriteResultHeaders(ByVal ws As Worksheet)

    ws.Cells(1, 1).Value = "Гео"
    ws.Cells(1, 2).Value = "Ticket ID"
    ws.Cells(1, 3).Value = "Subagent"
    ws.Rows(1).Font.Bold = True

End Sub

Private Sub FormatResultSheet(ByVal ws As Worksheet)

    ws.Cells.WrapText = False
    ws.Columns.AutoFit
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:C1").AutoFilter

End Sub

Private Function IsSMPAgentID(ByVal value As Variant) As Boolean

    IsSMPAgentID = (NormalizeText(value) = "1014")

End Function

Private Function IsSMPCountry(ByVal value As Variant) As Boolean

    Dim s As String
    s = NormalizeText(value)
    
    IsSMPCountry = False
    
    If s = NormalizeText("Egypt") Then IsSMPCountry = True
    If s = NormalizeText("Mauritania") Then IsSMPCountry = True
    If s = NormalizeText("Sudan") Then IsSMPCountry = True
    
    If s = NormalizeText("Египет") Then IsSMPCountry = True
    If s = NormalizeText("Мавритания") Then IsSMPCountry = True
    If s = NormalizeText("Судан") Then IsSMPCountry = True

End Function

Private Function IsAllowedReferal(ByVal value As Variant) As Boolean

    Dim s As String
    s = NormalizeText(value)
    
    IsAllowedReferal = False
    
    Select Case s
        Case NormalizeText("webdefault"), _
             NormalizeText("1xbet22.com"), _
             NormalizeText("melbet"), _
             NormalizeText("1xir.com"), _
             NormalizeText("1xgames"), _
             NormalizeText("bo.1xbet.com"), _
             NormalizeText("1xbet.tn"), _
             NormalizeText("1xbet.et"), _
             NormalizeText("bizbet"), _
             NormalizeText("1xcasino"), _
             NormalizeText("afropari"), _
             NormalizeText("onjabet"), _
             NormalizeText("bizbet africa (ar)"), _
             NormalizeText("1xbet.pa")
             
             IsAllowedReferal = True
    End Select

End Function

Private Function GetHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long

    Dim lastCol As Long
    Dim colIndex As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For colIndex = 1 To lastCol
        If NormalizeText(ws.Cells(1, colIndex).Value) = NormalizeText(headerName) Then
            GetHeaderCol = colIndex
            Exit Function
        End If
    Next colIndex
    
    GetHeaderCol = 0

End Function

Private Function NormalizeText(ByVal value As Variant) As String

    Dim s As String
    
    s = LCase(Trim(CStr(value)))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, "ё", "е")
    s = Replace(s, "м", "m")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeText = s

End Function

Private Function TryParseDate(ByVal value As Variant, ByRef resultDate As Date) As Boolean

    Dim s As String
    Dim datePart As String
    Dim timePart As String
    Dim partsDate As Variant
    Dim partsTime As Variant
    
    On Error GoTo ParseError
    
    If IsDate(value) Then
        resultDate = CDate(value)
        TryParseDate = True
        Exit Function
    End If
    
    s = Trim(CStr(value))
    s = Replace(s, "T", " ")
    s = Replace(s, "/", "-")
    
    If InStr(s, " ") > 0 Then
        datePart = Split(s, " ")(0)
        timePart = Split(s, " ")(1)
    Else
        datePart = s
        timePart = "00:00:00"
    End If
    
    If InStr(datePart, "-") > 0 Then
        partsDate = Split(datePart, "-")
        
        If Len(partsDate(0)) = 4 Then
            partsTime = Split(timePart, ":")
            
            resultDate = DateSerial(CInt(partsDate(0)), CInt(partsDate(1)), CInt(partsDate(2))) + _
                         TimeSerial(GetTimePart(partsTime, 0), GetTimePart(partsTime, 1), GetTimePart(partsTime, 2))
            
            TryParseDate = True
            Exit Function
        End If
    End If
    
    If InStr(datePart, ".") > 0 Then
        partsDate = Split(datePart, ".")
        partsTime = Split(timePart, ":")
        
        resultDate = DateSerial(CInt(partsDate(2)), CInt(partsDate(1)), CInt(partsDate(0))) + _
                     TimeSerial(GetTimePart(partsTime, 0), GetTimePart(partsTime, 1), GetTimePart(partsTime, 2))
        
        TryParseDate = True
        Exit Function
    End If
    
ParseError:
    TryParseDate = False

End Function

Private Function GetTimePart(ByVal partsTime As Variant, ByVal index As Long) As Integer

    On Error GoTo SafeExit
    
    If IsArray(partsTime) Then
        If UBound(partsTime) >= index Then
            GetTimePart = CInt(partsTime(index))
            Exit Function
        End If
    End If
    
SafeExit:
    GetTimePart = 0

End Function

Private Function GetLastRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range
    
    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
    
    If lastCell Is Nothing Then
        GetLastRow = 1
    Else
        GetLastRow = lastCell.Row
    End If

End Function