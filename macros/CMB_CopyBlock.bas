Option Explicit

' Обработчик кнопок "Копировать" для Monitoring_TTR_FTT.
' Написан по образцу остальных обработчиков библиотеки
' (Monitoring_Перемещение_тикетов_CopyBlock, Monitoring_Выводы_CopyBlock, L2_CopyBlock):
' адрес диапазона лежит в .AlternativeText кнопки, имя кнопки приходит
' через Application.Caller.
'
' Function, а не Sub - не отображается в списке макросов (Alt+F8).
'
' Если CMB_CopyBlock уже есть в вашей книге / PERSONAL.XLSB, этот модуль
' импортировать не нужно: две Public-процедуры с одним именем в одном
' проекте дают ошибку "Ambiguous name detected".

Public Function CMB_CopyBlock() As Variant

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
    MsgBox "Не удалось скопировать данные: " & Err.Description, vbExclamation, "Monitoring_TTR_FTT"

End Function
