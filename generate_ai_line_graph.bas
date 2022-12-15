Attribute VB_Name = "Module3"
Sub AverageEachSecond()

    Dim i As Integer
    i = 0
    j = 1
    Do
        Floor = "B" + CStr(i + 1)
        Ceiling = "B" + CStr(i + 10)
        CopyRange = Floor + ":" + Ceiling
        i = i + 10
        resultCell = "E" + CStr(j)
        Range(resultCell) = Application.Average(Range(CopyRange))
        j = j + 1
    Loop Until i = 18010
    
End Sub

Sub PrintSeconds()

    Dim DateStr As Date
    Dim x As Integer
    x = 0
    
    'MAGIC #s
    where_date_ends_digital_time_begins = 12
    length_of_digital_time = 8
    num_of_measurements = 10
    
    'see Excel column A
    '00:01.4 maps to 10:00:01
    '00:01.5 maps to 10:00:02
    mid_value = 4
    
    Do
        x = x + 1
        'Copies 10:00:0x
        DateStr = Range("A" + CStr((x * num_of_measurements) - mid_value)).Value
        digital_time = Mid(CStr(DateStr), where_date_ends_digital_time_begins, length_of_digital_time)
        Range("D" + CStr(x)) = digital_time
    Loop Until x = 1801

End Sub
