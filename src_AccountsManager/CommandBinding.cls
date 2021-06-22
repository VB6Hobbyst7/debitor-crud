VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
'@Folder AccountsManager.Infrastructure.Bindings.CommandBindings
'@ModuleDescription "An object responsible for binding a command to a UI element."
'@PredeclaredId
'@Exposed
Implements ICommandBinding
Implements IDisposable
Option Explicit

Private Type TCommandBinding
    ViewModel As Object
    Target As Object
    command As ICommand
End Type

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ImageEvents As MSForms.Image
Attribute ImageEvents.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private this As TCommandBinding

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
Public Function ForCommandButton(ByVal Target As MSForms.CommandButton, ByVal command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCommandButton.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCommandButton = Create(Target, command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
Public Function ForCheckBox(ByVal Target As MSForms.CheckBox, ByVal command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCheckBox.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCheckBox = Create(Target, command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
Public Function ForImage(ByVal Target As MSForms.Image, ByVal command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForImage.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForImage = Create(Target, command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
Public Function ForLabel(ByVal Target As MSForms.Label, ByVal command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForLabel.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForLabel = Create(Target, command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
Public Function Create(ByVal Target As Object, ByVal command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute Create.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    
    Dim result As CommandBinding
    Set result = New CommandBinding
    
    Set result.ViewModel = ViewModel
    Set result.Target = Target
    Set result.command = command
    
    Set Create = result
    
End Function

Public Property Get ViewModel() As Object
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set this.ViewModel = RHS
End Property

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    Set this.Target = RHS
    
    Select Case True
        Case TypeOf RHS Is MSForms.CommandButton
            Set CommandButtonEvents = RHS
            
        Case TypeOf RHS Is MSForms.CheckBox
            Set CheckBoxEvents = RHS
            
        Case TypeOf RHS Is MSForms.Image
            Set ImageEvents = RHS
            
        Case TypeOf RHS Is MSForms.Label
            Set LabelEvents = RHS
            
        Case Else
            GuardClauses.GuardExpression _
                Throw:=True, _
                Source:=TypeName(Me), _
                Message:="Type '" & TypeName(RHS) & "' does not support command bindings at the moment."
    End Select
    
End Property

Public Property Get command() As ICommand
    Set command = this.command
End Property

Public Property Set command(ByVal RHS As ICommand)
    GuardClauses.GuardDoubleInitialization this.command, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.command = RHS
End Property

Private Property Get AsInterface() As ICommandBinding
    Set AsInterface = Me
End Property

Private Sub OnExecute()
    If Not this.command Is Nothing Then
        this.command.Execute this.ViewModel
    Else
        Logging.logINFO "BUG in " & TypeName(Me) & ": Command is 'Nothing', cannot execute."
'        Debug.Print "BUG in " & TypeName(Me) & ": Command is 'Nothing', cannot execute."
        Debug.Assert False ' should not happen, break here if it does.
    End If
End Sub

Private Sub EvaluateCanExecute(ByVal Source As Object)
    If this.Target Is Nothing Then Exit Sub
    If this.command Is Nothing Then
        this.Target.Enabled = False
    Else
        On Error Resume Next
        this.Target.Enabled = this.command.CanExecute(Source)
        'a validation error formatter might have taken over the tooltip?
        this.Target.ControlTipText = IIf(this.Target.ControlTipText = vbNullString, this.command.Description, this.Target.ControlTipText)
        On Error GoTo 0
    End If
End Sub

Public Function ToString() As String
    ToString = TypeName(this.Target) & " -> " & TypeName(this.command)
End Function

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = this.Target
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = this.command
End Property

Private Sub ICommandBinding_EvaluateCanExecute(ByVal Context As Object)
    EvaluateCanExecute Context
End Sub

Private Sub CheckBoxEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.CheckBox, TypeName(Me)
    OnExecute
End Sub

Private Sub CommandButtonEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.CommandButton, TypeName(Me)
    OnExecute
End Sub

Private Sub IDisposable_Dispose()
    Set this.command = Nothing
    Set this.Target = Nothing
    Set this.ViewModel = Nothing
End Sub

Private Sub ImageEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.Image, TypeName(Me)
    OnExecute
End Sub

Private Sub LabelEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.Label, TypeName(Me)
    OnExecute
End Sub
