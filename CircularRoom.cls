VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CircularRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Room
Implements FloorLevel


Private mRadius As Double
Private mLevel As String

Public Property Let Radius(ByVal value As Double)
    mRadius = value
End Property

Public Property Let Level(ByVal value As String)
    mLevel = value
End Property

Public Function Room_Area() As Double
    Room_Area = (mRadius * 2) * (mRadius * 2)
End Function

Public Function FloorLevel_FlightsOfStairs() As Integer
    FloorLevel_FlightsOfStairs = StairsFromFloorLevel(mLevel)
End Function
