Attribute VB_Name = "Module1"
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос2 Макрос
'
Dim dic As Dictionary
'
    Range("K26").Select
    ActiveCell.FormulaR1C1 = "ee"
    Range("K26").Select
    Selection.ClearContents
End Sub

Sub Прописные()
    Dim cell As Range
    For Each cell In Selection
        cell.Value = UCase(cell.Value)
    Next cell
End Sub

