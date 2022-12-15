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
    Do
        x = x + 1
        'Copies 10:00:0x
        DateStr = Range("A" + CStr((x * 10) - 4)).Value
        analog = Mid(CStr(DateStr), 12, 8)
        Range("D" + CStr(x)) = analog
    Loop Until x = 1801

End Sub
