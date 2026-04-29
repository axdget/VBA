Attribute VB_Name = "ZirnimGroup2"
Option Explicit

Sub ЖирнимГруппы()
Attribute ЖирнимГруппы.VB_ProcData.VB_Invoke_Func = "Й\n14"
    Dim cell As Range, prevValue, groupCounter As Integer
    Dim selT As Range

    prevValue = ""
    Set selT = Range(Cells(1, Selection.Column), Cells(Selection.Rows.Count, Selection.Column).End(xlUp))
    selT.Font.Bold = False
    
    For Each cell In selT
        If cell.Value <> prevValue Then
            groupCounter = groupCounter + 1
            prevValue = cell.Value
        End If
        
        If groupCounter Mod 2 = 1 Then cell.Font.Bold = True
    Next cell
End Sub

Function IsBold(cell As Range) As Boolean
    IsBold = cell.Font.Bold
End Function

Sub В_Числа_ОтрОтсечь()
    Dim cell As Range, prevValue, groupCounter As Integer
    Dim selT As Range

    prevValue = ""
    Set selT = Range(Cells(1, Selection.Column), Cells(Selection.Rows.Count, Selection.Column).End(xlUp))
    
    For Each cell In Selection  'selT
        If IsNumeric(cell.Value) Then
            If cell.Value < 0 Then
                cell.ClearContents
                cell.Font.Bold = True
            End If
        Else
            cell.ClearContents
            'cell.Font.Bold = True
        End If
       
    Next cell
End Sub


Sub ДатаКонверт()
    Dim cell As Range, prevValue, groupCounter As Integer
    Dim selT As Range

    prevValue = ""
    Set selT = Range(Cells(1, Selection.Column), Cells(Selection.Rows.Count, Selection.Column).End(xlUp))
    selT.Font.Bold = False
    
    For Each cell In selT
        If IsDate(cell.Value) Then
            cell.NumberFormat = "dd.mm.yyyy"
            cell.Value = CDate(cell.Value)
        End If
        
    Next cell
End Sub

Sub ДатаВчера_ДругиеСтрУбрать()
    Dim row As Long, lastRow As Long, sht As Worksheet, flag As Boolean
    Set sht = ActiveSheet
    lastRow = sht.Cells(sht.Rows.Count, 1).End(xlUp).row
        
    For row = lastRow To 2 Step -1
        If IsDate(sht.Cells(row, 1).Value) Then
            If (Not DateValue(sht.Cells(row, 1).Value) = Date - 1) Then
                sht.Rows(row).Delete
                GoTo NextIteration
            End If
'            If (Not DateValue(sht.Cells(row, 1).Value) = Date - 2) Then
'                sht.Rows(row).Delete
'            End If
        Else
            sht.Range("A" & row).Activate
            MsgBox "A" & row & " Не дата! ": End
        End If
NextIteration:
    Next row
End Sub
