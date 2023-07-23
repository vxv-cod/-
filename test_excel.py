Sub countHeight()
    Debug.Print "================"
    Dim heightArray() As Variant
    Dim lastRow As Long, i As Long
    
    lastRow = ActiveSheet.UsedRange.Rows.count ' ������� ����� ��������� ����������� ������ �� �����
    
    ReDim heightArray(1 To lastRow) ' ���������� ������ �������
    
    For i = 1 To lastRow
        heightArray(i) = ActiveSheet.Rows(i).Height ' ���������� ������ ������ ����������� ������ � ������
    Next i

    Dim sum As Double
    Dim count As Integer
    Dim stop_ As Integer
    
    sum = 0
    count = 1
    stop_ = 0
    
    For i = 1 To UBound(heightArray) ' ���������� ������
        sum = sum + heightArray(i) ' ��������� �������� ������ � �����
        If sum > 200 And stop_ = 0 Then ' ���� ����� ������ 300
            count = count + 1 ' ����������� ������� �� 1
            sum = heightArray(i) '���������� ����� �� �������� �������� ������
            stop_ = 1
        End If
        
        If sum > 300 Then ' ���� ����� ������ 500
            count = count + 1 ' ����������� ������� �� 1
            sum = heightArray(i) '���������� ����� �� �������� �������� ������
        End If
    Next i
    
    Debug.Print count ' ������� ��������� � ���� �������
    
End Sub
