Sub countHeight()
    Debug.Print "================"
    Dim heightArray() As Variant
    Dim lastRow As Long, i As Long
    
    lastRow = ActiveSheet.UsedRange.Rows.count ' находим номер последней заполненной строки на листе
    
    ReDim heightArray(1 To lastRow) ' определяем размер массива
    
    For i = 1 To lastRow
        heightArray(i) = ActiveSheet.Rows(i).Height ' записываем высоту каждой заполненной строки в массив
    Next i

    Dim sum As Double
    Dim count As Integer
    Dim stop_ As Integer
    
    sum = 0
    count = 1
    stop_ = 0
    
    For i = 1 To UBound(heightArray) ' перебираем массив
        sum = sum + heightArray(i) ' добавляем значение высоты к сумме
        If sum > 200 And stop_ = 0 Then ' если сумма больше 300
            count = count + 1 ' увеличиваем счетчик на 1
            sum = heightArray(i) 'сбрасываем сумму до текущего значения высоты
            stop_ = 1
        End If
        
        If sum > 300 Then ' если сумма больше 500
            count = count + 1 ' увеличиваем счетчик на 1
            sum = heightArray(i) 'сбрасываем сумму до текущего значения высоты
        End If
    Next i
    
    Debug.Print count ' выводим результат в окно отладки
    
End Sub
