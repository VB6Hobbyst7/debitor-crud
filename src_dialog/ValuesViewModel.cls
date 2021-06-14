VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("ValidateUserInput.ViewModel")
Option Explicit
'@PredeclaredId

Private Type TValuesModel
    Id As Long
    Filter As String
    ModelWidth As Double
    ModelHeight As Double
    DataSourceTable As String
End Type

Private this As TValuesModel

Public Function Create() As ValuesViewModel
    GuardClauses.GuardNonDefaultInstance Me, ValuesViewModel, TypeName(Me)
    
    Dim result As ValuesViewModel
    Set result = New ValuesViewModel
    
    Set Create = result
    
End Function

Public Property Get DataSourceTable() As String
    DataSourceTable = this.DataSourceTable
End Property

Public Property Let DataSourceTable(ByVal RHS As String)
    this.DataSourceTable = RHS
End Property

Public Property Get ModelWidth() As Double
    ModelWidth = this.ModelWidth
End Property

Public Property Let ModelWidth(ByVal RHS As Double)
    this.ModelWidth = RHS
End Property

Public Property Get ModelHeight() As Double
    ModelHeight = this.ModelHeight
End Property

Public Property Let ModelHeight(ByVal RHS As Double)
    this.ModelHeight = RHS
End Property

Public Property Get Id() As Long
    Id = this.Id
End Property

Public Property Let Id(ByVal RHS As Long)
    If this.Id <> RHS Then
        this.Id = setNewId
    End If
End Property

Public Property Get Filter() As String
    Filter = this.Filter
End Property

Public Property Let Filter(ByVal RHS As String)
    this.Filter = RHS
End Property

Private Function setNewId() As Long
    With Application.ThisWorkbook.Sheets(DataSourceTable)
        setNewId = Application.WorksheetFunction.Max(.Range("A1", .Cells(.Rows.Count, 1).End(xlUp))) + 1
    End With
End Function
