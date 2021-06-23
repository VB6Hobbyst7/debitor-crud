VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a binding validation error."
'@Folder AccountsManager.Infrastructure.Validation
'@ModuleDescription "Represents a binding validation error."
'@PredeclaredId
Option Explicit

Implements IValidationError

Private Type TState
    Binding As IPropertyBinding
    message As String
End Type

Private this As TState

'@Description "Creates a new instance of a validation error object."
Public Function Create(ByVal Binding As IPropertyBinding, ByVal message As String) As IValidationError
Attribute Create.VB_Description = "Creates a new instance of a validation error object."
    GuardClauses.GuardNonDefaultInstance Me, ValidationError, TypeName(Me)
    Dim result As ValidationError
    Set result = New ValidationError
    Set result.Binding = Binding
    result.message = message
    Set Create = result
End Function

'@Description "The validation error message."
Public Property Get message() As String
Attribute message.VB_Description = "The validation error message."
    message = this.message
End Property

Public Property Let message(ByVal RHS As String)
    this.message = RHS
End Property

'@Description "Gets the binding responsible for this validation error."
Public Property Get Binding() As IPropertyBinding
Attribute Binding.VB_Description = "Gets the binding responsible for this validation error."
    Set Binding = this.Binding
End Property

Public Property Set Binding(ByVal RHS As IPropertyBinding)
    Set this.Binding = RHS
End Property

Private Property Get IValidationError_Message() As String
    IValidationError_Message = this.message
End Property

Private Property Get IValidationError_Binding() As IPropertyBinding
    Set IValidationError_Binding = this.Binding
End Property
