Attribute VB_Name = "CarpetFactory"
Public Function CreateCarpet(ByVal pricePerSqMtr As Double, ByVal roundUp As Boolean) As carpet
    Dim carpet As New carpet
    carpet.pricePerSqMtr = pricePerSqMtr
    carpet.roundUp = roundUp
    Set CreateCarpet = carpet
End Function
