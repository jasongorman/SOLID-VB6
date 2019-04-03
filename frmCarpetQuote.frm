VERSION 5.00
Begin VB.Form frmCarpetQuote 
   Caption         =   "Carpet Quote"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFloor 
      Height          =   315
      ItemData        =   "frmCarpetQuote.frx":0000
      Left            =   2880
      List            =   "frmCarpetQuote.frx":0019
      TabIndex        =   16
      Text            =   "G"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton btnQuoteCircular 
      Caption         =   "Quote"
      Height          =   495
      Left            =   10560
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtRadius 
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Text            =   "2.5"
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton btnQuoteRectangular 
      Caption         =   "Quote"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox chkRoundUp 
      Caption         =   "Round up"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtPricePerSqMtr 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Text            =   "10.0"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtLength 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "2.5"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Text            =   "2.5"
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Floor Level"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Room Radius"
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblStairsPremium 
      Caption         =   "£0.00"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Stairs premium"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblPrice 
      Caption         =   "£0.00"
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "The total price for this carpet will be"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Price/sq m"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Length"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Room Width"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmCarpetQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnQuoteCircular_Click()
    Dim Room As Room
    Set Room = CreateCircularRoom(CDbl(txtRadius.Text), cboFloor.Text)
    Call DisplayQuote(Room)
End Sub

Private Sub btnQuoteRectangular_Click()
    Dim Room As Room
    Set Room = CreateRectangularRoom(CDbl(txtWidth.Text), CDbl(txtLength.Text), cboFloor.Text)
    Call DisplayQuote(Room)
End Sub

Private Sub DisplayQuote(ByRef Room As Room)
    Dim Carpet As Carpet
    Set Carpet = CreateCarpet(CDbl(txtPricePerSqMtr.Text), chkRoundUp.value)
    
    Dim quote As New CarpetQuote
    
    lblPrice.Caption = "£" & quote.quote(Room, Carpet)
    
    Dim fitting As New FittingCalculator
    
    lblStairsPremium.Caption = "£" & fitting.CalculateStairsPremium(Room)
End Sub
