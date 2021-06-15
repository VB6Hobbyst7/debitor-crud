VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SideMargins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ValidateUserInput.Model.ResizeView")
Option Explicit

Private Type TSideMargins
    BottomMargin As Double
    LeftMargin As Double
    RightMargin As Double
    TopMargin As Double
End Type

Private this As TSideMargins

Public Property Get BottomMargin() As Double
    BottomMargin = this.BottomMargin
End Property

Public Property Let BottomMargin(ByVal bottom As Double)
    this.BottomMargin = bottom
End Property

Public Property Get LeftMargin() As Double
    LeftMargin = this.LeftMargin
End Property

Public Property Let LeftMargin(ByVal left As Double)
    this.LeftMargin = left
End Property

Public Property Get RightMargin() As Double
    RightMargin = this.RightMargin
End Property

Public Property Let RightMargin(ByVal right As Double)
    this.RightMargin = right
End Property

Public Property Get TopMargin() As Double
    TopMargin = this.TopMargin
End Property

Public Property Let TopMargin(ByVal top As Double)
    this.TopMargin = top
End Property
