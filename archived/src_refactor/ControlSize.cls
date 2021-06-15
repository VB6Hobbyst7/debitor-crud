VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MaintainCustomers.Model.ResizeView")
Option Explicit

Private Type TControlSize
    height As Single
    width As Single
End Type

Private this As TControlSize

Public Property Get height() As Single
    height = this.height
End Property

Public Property Let height(ByVal height As Single)
    this.height = height
End Property

Public Property Get width() As Single
    width = this.width
End Property

Public Property Let width(ByVal width As Single)
    this.width = width
End Property

