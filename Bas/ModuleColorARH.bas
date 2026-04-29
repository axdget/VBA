Attribute VB_Name = "ModuleColorARH"
Sub Кегль20()
Attribute Кегль20.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells(ActiveCell.row, 1).Select
    
    ActiveCell.EntireRow.Select
    
    With Selection.Font
        .Name = "Calibri"
        .Size = 20
    End With
    
    Cells(ActiveCell.row, 1).Select
End Sub

Sub ColorARH()
Attribute ColorARH.VB_ProcData.VB_Invoke_Func = "й\n14"
    Dim rng As Range
    Dim cell As Range
    Dim startRow As Long, endRow As Long
    Dim startCol As Long, endCol As Long
    Dim row As Long, col As Long
    Dim val1, val2 As Variant
    
    ' Проверяем выделение
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count < 2 Then
        MsgBox "Выделите диапазон как минимум из 2 столбцов", vbExclamation
        Exit Sub
    End If
    
    ' Получаем границы выделения
    Set rng = Selection
    With rng
        startRow = .row
        endRow = startRow + .Rows.Count - 1
        startCol = .Column
        endCol = startCol + .Columns.Count - 1
    End With
    
    Application.ScreenUpdating = False
    
    ' Обрабатываем каждую строку
    For row = startRow To endRow
        ' 1. Сначала все значения делаем черными
        For col = startCol To endCol
            With Cells(row, col)
                .Font.Color = vbBlack
                .Font.Bold = False
            End With
        Next col
        
        ' 2. Сравниваем пары слева направо
        For col = startCol To endCol - 1
            val1 = Cells(row, col).Value
            val2 = Cells(row, col + 1).Value
            
            ' Пропускаем пустые ячейки
            If IsEmpty(Cells(row, col)) Or IsEmpty(Cells(row, col + 1)) Then
                GoTo NextPair
            End If
            
            ' Сравниваем значения (учитывая числа как строки)
            If CStr(val1) <> CStr(val2) Then
                ' Красим обе ячейки в фиолетовый
                 Cells(row, col).Font.Color = RGB(148, 0, 211)
                 Cells(row, col + 1).Font.Color = RGB(148, 0, 211)
            End If
            
NextPair:
        Next col
    Next row
    Application.ScreenUpdating = True
    ' MsgBox "Обработано строк: " & (endRow - startRow + 1), vbInformation
End Sub
