VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetValidationAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target Excel.Range object."
'@Folder AccountsManager.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "An object that dynamically decorates a target Excel.Range object."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicAdorner

Private Type TState
    Target As Excel.Range
    TargetFormatter As IValidationErrorFormatter
End Type

Private this As TState

Public Function Create(ByVal Target As Excel.Range, ByVal Formatter As IValidationErrorFormatter) As IDynamicAdorner
    Dim result As WorksheetValidationAdorner
    Set result = New WorksheetValidationAdorner
    Set result.Target = Target
    Set result.Formatter = Formatter
    Set Create = result
End Function

Friend Property Get Formatter() As IValidationErrorFormatter
    Set Formatter = this.TargetFormatter
End Property

Friend Property Set Formatter(ByVal RHS As IValidationErrorFormatter)
    Set this.TargetFormatter = RHS
End Property

Friend Property Get Target() As Excel.Range
    Set Target = this.Target
End Property

Friend Property Set Target(ByVal RHS As Excel.Range)
    GuardClauses.GuardDefaultInstance Me, WorksheetValidationAdorner
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set this.Target = RHS
End Property

Private Sub IDynamicAdorner_Hide()
    this.TargetFormatter.Restore this.Target
End Sub

Private Sub IDynamicAdorner_Show(ByVal message As String)
    this.TargetFormatter.Apply this.Target, message
End Sub
