VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing command bindings."
'@Folder AccountsManager.Infrastructure.Bindings.CommandBindings
'@ModuleDescription "An object responsible for managing command bindings."
'@Exposed
Option Explicit
Implements ICommandManager
Implements IDisposable

Private Type TState
    CommandBindings As Collection
    DebugOutput As Boolean
End Type

Private this As TState

'@Description "Gets/sets a value indicating whether the object produces debugger output."
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value indicating whether the object produces debugger output."
    DebugOutput = this.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    this.DebugOutput = RHS
End Property

'@Description "Releases all held ICommandBinding references, prepares the object for proper destruction."
Public Sub Dispose()
Attribute Dispose.VB_Description = "Releases all held ICommandBinding references, prepares the object for proper destruction."
    '@Ignore VariableNotUsed
    Dim Index As Long
    For Index = 1 To this.CommandBindings.Count
        Disposable.TryDispose this.CommandBindings.Item(this.CommandBindings.Count)
        this.CommandBindings.Remove this.CommandBindings.Count
    Next
    Set this.CommandBindings = Nothing
End Sub

Private Sub EvaluateCanExecute(ByVal BindingContext As Object)
    Dim Binding As ICommandBinding
    For Each Binding In this.CommandBindings
        Binding.EvaluateCanExecute BindingContext
    Next
End Sub

Private Sub DebugMessage(ByVal Message As String)
'    If this.DebugOutput Then Debug.Print TypeName(Me) & ": " & Message
    If this.DebugOutput Then Logging.logINFO TypeName(Me) & ": " & Message
End Sub

Private Sub Class_Initialize()
    Set this.CommandBindings = New Collection
End Sub

Private Function ICommandManager_BindCommand(ByVal BindingContext As Object, ByVal Target As Object, ByVal command As ICommand) As ICommandBinding
    GuardClauses.GuardNullReference BindingContext, TypeName(Me)
    GuardClauses.GuardNullReference Target, TypeName(Me)
    GuardClauses.GuardNullReference command, TypeName(Me)
    'GuardClauses.GuardDefaultInstance Me, CommandManager
    
    Dim Binding As ICommandBinding
    
    Select Case True
    
        Case TypeOf Target Is MSForms.CommandButton
            Set Binding = CommandBinding.ForCommandButton(Target, command, BindingContext)
            
        Case TypeOf Target Is MSForms.Image
            Set Binding = CommandBinding.ForImage(Target, command, BindingContext)
            
        Case TypeOf Target Is MSForms.Label
            Set Binding = CommandBinding.ForLabel(Target, command, BindingContext)
            
        Case TypeOf Target Is MSForms.CheckBox
            Set Binding = CommandBinding.ForCheckBox(Target, command, BindingContext)
            
        Case Else
            GuardClauses.GuardExpression True, TypeName(BindingContext), "Target type '" & TypeName(Target) & "' does not currently support command bindings."
            
    End Select
    
    this.CommandBindings.add Binding
    Set ICommandManager_BindCommand = Binding
    
    DebugMessage "Created CommandBinding for " & TypeName(command) & "."
    
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal BindingContext As Object)
    EvaluateCanExecute BindingContext
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

