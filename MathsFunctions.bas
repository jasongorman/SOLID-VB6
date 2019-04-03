Attribute VB_Name = "MathsFunctions"
Public Function Ceiling(ByVal number As Double) As Double
    Dim rounded As Double
    rounded = Round(number)
    If rounded < number Then
        Ceiling = rounded + 1
    Else
        Ceiling = rounded
    End If
End Function
