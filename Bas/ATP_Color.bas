Attribute VB_Name = "ATP_Color"
Sub ATP_FonColor()
' Application.ScreenUpdating = False
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim rowRange As Range, cellText As String, themeColor As Long
    Dim tint As Double, found As Boolean
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    ' Проверка, что выделен диапазон
    If rng Is Nothing Then
        MsgBox "Выделите диапазон ячеек!", vbExclamation
        Exit Sub
    End If
    
    ' Проходим по каждой ячейке выделенного диапазона
    For Each cell In rng
        'cellText = Trim(cell.Value)
        On Error Resume Next  '''''''''''''''''''''''''''
        cellText = Trim(cell.Value)
        If Err.Number <> 0 Then
            cellText = ""
            Err.Clear
        End If
        On Error GoTo 0 ''''''''''''''''''''''''''''''''
        
        found = False
        
        ' Пропускаем пустые ячейки
        If Len(cellText) > 0 Then
            
            ' Батайск
            If InStr(cellText, "Батайск") > 0 And InStr(cellText, "1") > 0 Then
                themeColor = 7: tint = 0.7: found = True
            ElseIf InStr(cellText, "Батайск") > 0 And InStr(cellText, "2") > 0 Then
                themeColor = 7: tint = 0.5: found = True
            ElseIf InStr(cellText, "Батайск") > 0 And InStr(cellText, "3") > 0 Then
                themeColor = 7: tint = 0.3: found = True
                
            'Волгоград
            ElseIf InStr(cellText, "Волгоград") > 0 Then
                themeColor = 11: tint = 0.95: found = True
            
            
            ' Минеральные Воды
            ElseIf InStr(cellText, "Мин Воды") > 0 Or InStr(cellText, "Минеральные Воды") > 0 Then
                themeColor = 11: tint = 0.95: found = True
            
            ' Омск
            ElseIf InStr(cellText, "Омск") > 0 Then
                themeColor = 6: tint = 0.95: found = True
            
            ' Самара
            ElseIf InStr(cellText, "Самара") > 0 Then
                themeColor = 9: tint = 0.95: found = True
            
            ' Краснодар 1
            ElseIf InStr(cellText, "Краснодар") > 0 And InStr(cellText, "1") > 0 Then
                themeColor = 10: tint = 0.8: found = True
            
            ' Краснодар 2
            ElseIf InStr(cellText, "Краснодар") > 0 And InStr(cellText, "2") > 0 Then
                themeColor = 10: tint = 0.9: found = True
            
            ' Тула
            ElseIf InStr(cellText, "Тула") > 0 Then
                themeColor = 8: tint = 0.9: found = True
            
            ' Усть-Лабинск
            ElseIf InStr(cellText, "Усть-Лабинск") > 0 Then
                themeColor = 10: tint = 0.9: found = True
            
            'Тех. служба, длительный ремонт
            ElseIf InStr(cellText, "Тех. служба, длительный ремонт") > 0 Then
                themeColor = 12: tint = 0.85: found = True
            
            End If
            
            ' Если найдено совпадение - красим только ячейки выделенного диапазона в этой строке
            If found Then
                ' Находим пересечение выделенного диапазона и строки с найденной ячейкой
                Set rowRange = Intersect(rng, ws.Rows(cell.row))
                
                If Not rowRange Is Nothing Then
                    With rowRange.Interior
                        .themeColor = themeColor
                        .TintAndShade = tint
                    End With
                End If
            End If
        End If
    Next cell
    'Application.ScreenUpdating = True
End Sub

Sub ThemeColorByColumn()
    Dim wsNew As Worksheet, i As Integer
    Set wsNew = ActiveWorkbook.Sheets.Add()
    wsNew.Name = "ThemeColors_" & Format(Now, "ss") 'Format(Now, "ddmmyy_hhmmss")
    
    For i = 1 To 12
        With wsNew.Cells(i, 1)
            .Value = i
            .Interior.themeColor = i
            .Interior.TintAndShade = 0.5
            If i >= 11 Then .Font.Bold = True
        End With
    Next i
    
    wsNew.Columns("A").AutoFit
    MsgBox "Готово! Лист: " & wsNew.Name, vbInformation
End Sub
