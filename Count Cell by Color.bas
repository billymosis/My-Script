Function CountColor(Rng As Range, RngColor As Range) As Integer
Dim Cll As Range
Dim Clr As Long
Clr = RngColor.Range("A1").Interior.Color
For Each Cll In Rng
    If Cll.Interior.Color = Clr Then
    CountColor = CountColor + 1
    End If
Next Cll
    
End Function
