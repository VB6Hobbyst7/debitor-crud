VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ValidateUserInput.Model"

Option Explicit

Private Type TModel
    field1 As String
    field2 As String
End Type

Private this As TModel

Public Property Get field1() As String
    field1 = this.field1
End Property

Public Property Let field1(ByVal value As String)
    this.field1 = value
End Property

Public Property Get field2() As String
    field2 = this.field2
End Property

Public Property Let field2(ByVal value As String)
    this.field2 = value
End Property
