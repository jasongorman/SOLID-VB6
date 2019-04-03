VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Carpet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPricePerSqMtr As Double
Private mRoundUp As Boolean

Public Property Let PricePerSqMtr(ByVal value As Double)
    mPricePerSqMtr = value
End Property

Public Property Let RoundUp(ByVal value As Boolean)
    mRoundUp = value
End Property

Public Function Price(ByVal area As Double)
    If mRoundUp Then
        area = Ceiling(area)
    End If
    
    Price = mPricePerSqMtr * area
End Function


