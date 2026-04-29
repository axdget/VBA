Attribute VB_Name = "NomeraEnRu"
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxTimeoutW Lib "user32" _
        (ByVal hWnd As LongPtr, _
         ByVal lpText As LongPtr, _
         ByVal lpCaption As LongPtr, _
         ByVal uType As Long, _
         ByVal wLanguageId As Long, _
         ByVal dwMilliseconds As Long) As Long
#Else
    Private Declare Function MessageBoxTimeoutW Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal lpText As Long, _
         ByVal lpCaption As Long, _
         ByVal uType As Long, _
         ByVal wLanguageId As Long, _
         ByVal dwMilliseconds As Long) As Long
#End If

Public Sub ShowPopup(ByVal msg As String, Optional ByVal zaderzka As Integer = 300, Optional ByVal title As String = "Готово")
    Const MB_OK As Long = &H0&
    Const MB_ICONINFORMATION As Long = &H40&
    Call MessageBoxTimeoutW(Application.hWnd, StrPtr(msg), StrPtr(title), _
                            MB_OK Or MB_ОК, 0, 900&)       ' 500 миллисекунд
End Sub

Sub Номера_АнгРу()
    ' Массив соответствий английских и русских букв (используемых в автомобильных номерах)
    Dim letters() As Variant
    letters = Array( _
        Array("A", "А"), _
        Array("B", "В"), _
        Array("C", "С"), _
        Array("E", "Е"), _
        Array("H", "Н"), _
        Array("K", "К"), _
        Array("M", "М"), _
        Array("O", "О"), _
        Array("P", "Р"), _
        Array("T", "Т"), _
        Array("X", "Х"), _
        Array("Y", "У") _
    )
    
    Dim selectedRange As Range
    Dim cell As Range, i As Integer, newValue As String
    Application.ScreenUpdating = False:     Application.Calculation = xlCalculationManual
   
    If Selection Is Nothing Then
        MsgBox "Пожалуйста, выделите ячейки для обработки", vbExclamation
        Exit Sub
    End If
    
    Set selectedRange = Selection
    
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) Then
            newValue = CStr(cell.Value)
            
            For i = LBound(letters) To UBound(letters)
                ' Заменяем заглавные буквы
                newValue = Replace(newValue, letters(i)(0), letters(i)(1))
                ' Заменяем строчные буквы
                newValue = Replace(newValue, LCase(letters(i)(0)), LCase(letters(i)(1)))
            Next i
            
            cell.Value = newValue
        End If
    Next cell
    
    Application.Calculation = xlCalculationAutomatic:  Application.ScreenUpdating = True
    ShowPopup "Замена завершена!", 900
End Sub


Sub xx()
ShowPopup "Замена завершена!", 900
End Sub
