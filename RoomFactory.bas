Attribute VB_Name = "RoomFactory"
Public Function CreateRectangularRoom(ByVal Width As Double, ByVal Length As Double, ByVal Level As String) As Room
    Dim Room As New RectangularRoom
    Room.Width = Width
    Room.Length = Length
    Room.Level = Level
    Set CreateRectangularRoom = Room
End Function

Public Function CreateCircularRoom(ByVal Radius As Double, ByVal Level As String) As Room
    Dim Room As New CircularRoom
    Room.Radius = Radius
    Room.Level = Level
    Set CreateCircularRoom = Room
End Function
