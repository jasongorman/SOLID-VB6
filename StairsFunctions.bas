Attribute VB_Name = "StairsFunctions"
Public Function StairsFromFloorLevel(ByVal Level As String) As Integer
    Select Case Level
    Case "G"
        StairsFromFloorLevel = 0
    Case "B"
        StairsFromFloorLevel = -1
    Case Else
        StairsFromFloorLevel = CInt(Level)
    End Select
End Function
