Option Explicit

Sub Monitoring_TTR_FTT()

    Const HEADER_COUNTRY_OLD As String = "Страна"
    Const HEADER_COUNTRY_NEW As String = "Гранулярность численности тикетов"
    Const HEADER_VALUE_FTT As String = "first_touch_hms_med"
    Const HEADER_VALUE_TTR As String = "time_to_resolve_hms_med"
    Const HEADER_VALUE_OLD As String = "med_dt_prev_diff_emp_sec_hms"

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim countries As Variant

    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim headerRow As Long
    Dim colCountry As Long
    Dim colValue As Long
    Dim lastRow As Long
    Dim lastCol As Long

    Dim currentHeader As String
    Dim valueHeaderFound As String
    Dim cleanSearch As String
    Dim cleanSource As String
    Dim foundValue As Variant
    Dim isFound As Boolean

    On Error GoTo ErrorHandler

    Set wb = ActiveWorkbook

    ' Источник всегда ВТОРОЙ лист выгрузки.
    ' На первом листе находится среднее значение (avg),
    ' на втором - медианное значение (med).
    ' Метрика может быть либо First Touch Time (first_touch_hms_med),
    ' либо Time to Resolve (time_to_resolve_hms_med) - макрос понимает обе.
    If wb.Worksheets.Count < 2 Then
        MsgBox _
            "В книге нет второго листа с медианным значением.", _
            vbCritical, _
            "Monitoring_TTR_FTT"
        Exit Sub
    End If

    Set wsSource = wb.Worksheets(2)

    ' Ищем нужные заголовки только на втором листе.
    headerRow = 0
    colCountry = 0
    colValue = 0
    valueHeaderFound = ""

    For r = 1 To 15

        lastCol = wsSource.Cells(r, wsSource.Columns.Count).End(xlToLeft).Column

        For c = 1 To lastCol

            currentHeader = Trim$(CStr(wsSource.Cells(r, c).Value))

            If StrComp(currentHeader, HEADER_COUNTRY_NEW, vbTextCompare) = 0 _
               Or StrComp(currentHeader, HEADER_COUNTRY_OLD, vbTextCompare) = 0 Then
                colCountry = c
            End If

            If StrComp(currentHeader, HEADER_VALUE_FTT, vbTextCompare) = 0 _
               Or StrComp(currentHeader, HEADER_VALUE_TTR, vbTextCompare) = 0 _
               Or StrComp(currentHeader, HEADER_VALUE_OLD, vbTextCompare) = 0 Then
                colValue = c
                valueHeaderFound = currentHeader
            End If

        Next c

        If colCountry > 0 And colValue > 0 Then
            headerRow = r
            Exit For
        End If

    Next r

    If headerRow = 0 Or colCountry = 0 Or colValue = 0 Then
        MsgBox _
            "На втором листе не найдены колонки страны и медианного значения." & vbCrLf & vbCrLf & _
            "Ожидаются: «" & HEADER_COUNTRY_NEW & "» и одна из: «" & HEADER_VALUE_FTT & "» или «" & HEADER_VALUE_TTR & "».", _
            vbCritical, _
            "Monitoring_TTR_FTT"
        Exit Sub
    End If

    lastRow = wsSource.Cells(wsSource.Rows.Count, colCountry).End(xlUp).Row

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Упорядоченный список").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsDest = wb.Worksheets.Add(After:=wsSource)
    wsDest.Name = "Упорядоченный список"

    wsDest.Cells(1, 1).Value = "Страна"
    wsDest.Cells(1, 2).Value = valueHeaderFound
    wsDest.Columns("B").NumberFormat = "[h]:mm:ss"

    countries = Array( _
        "Азербайджан", "Алжир", "Афганистан", "Бахрейн", "Боливия", _
        "Гватемала", "Гондурас", "Джибути", _
        "Египет", "Иордания", "Ирак", "Иран", "Йемен", _
        "Канада", "Катар", "Кувейт", "Кыргызстан", "Ливан", _
        "Ливия", "Мавритания", "Марокко", "Никарагуа", "ОАЭ", _
        "Оман", "Палестина", "Панама", "Парагвай", _
        "Саудовская Аравия", "Сирия", "Сомали", "Тайвань", "Тунис", _
        "Турция", "Южный Судан", "Судан", "Ямайка" _
    )

    For i = LBound(countries) To UBound(countries)

        wsDest.Cells(i + 2, 1).Value = countries(i)

        cleanSearch = NormalizeCountry_TTR_FTT(CStr(countries(i)))
        isFound = False
        foundValue = ""

        For r = headerRow + 1 To lastRow

            cleanSource = NormalizeCountry_TTR_FTT( _
                CStr(wsSource.Cells(r, colCountry).Value))

            If cleanSource = cleanSearch Then

                If Len(Trim$(CStr(wsSource.Cells(r, colValue).Value))) > 0 Then
                    foundValue = wsSource.Cells(r, colValue).Value
                    isFound = True
                    Exit For
                End If

            End If

        Next r

        If isFound Then
            wsDest.Cells(i + 2, 2).Value = foundValue
        Else
            wsDest.Cells(i + 2, 2).Value = ""
        End If

    Next i

    wsDest.Columns("A:B").AutoFit

    AddCopyButton_CMB _
        wsDest, _
        "btnMonitoringTtrFtt", _
        wsDest.Range("B1"), _
        "B2:B" & (UBound(countries) + 2), _
        RGB(31, 78, 121)

    MsgBox _
        "Готово! Взяты медианные значения со второго листа (" & valueHeaderFound & ").", _
        vbInformation, _
        "Успех"

    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True

    MsgBox _
        "Ошибка " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "Monitoring_TTR_FTT"

End Sub

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
        .OnAction = "'" & macroWorkbookName & "'!CMB_CopyBlock"

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
