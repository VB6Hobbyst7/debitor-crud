VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing property bindings."
'@Folder AccountsManager.Infrastructure.Bindings
'@ModuleDescription "An object responsible for managing property bindings."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IBindingManager
Implements IHandlePropertyChanged
Implements IDisposable

Private Type TState
    Context As IAppContext
    
    PropertyBindings As Collection
    NotifierFactory As INotifierFactory
    StringFormatterFactory As IStringFormatterFactory
    DebugOutput As Boolean
    BindingTargetStrategies As Dictionary
End Type

Private this As TState

Public Function Create(ByVal Context As IAppContext, ByVal StringFormatFactory As IStringFormatterFactory, Optional ByVal Factory As INotifierFactory, Optional ByVal DebugOutput As Boolean = False) As IBindingManager
    GuardClauses.GuardNonDefaultInstance Me, BindingManager, TypeName(Me)
    
    Dim result As BindingManager
    Set result = New BindingManager
    Set result.Context = Context
    result.DebugOutput = DebugOutput
    
    If StringFormatFactory Is Nothing Then
        Set result.StringFormatterFactory = New StringFormatterNetFactory
    Else
        Set result.StringFormatterFactory = StringFormatFactory
    End If
    
    If Factory Is Nothing Then
        Set result.NotifierFactory = New NotifierBaseFactory
    Else
        Set result.NotifierFactory = Factory
    End If
    
    Set Create = result
End Function

Public Property Get NotifierFactory() As INotifierFactory
    Set NotifierFactory = this.NotifierFactory
End Property

Public Property Set NotifierFactory(ByVal RHS As INotifierFactory)
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.NotifierFactory, TypeName(Me)
    Set this.NotifierFactory = RHS
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is BindingManager
End Property

Private Sub Apply(ByVal Source As Object)
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me), "Member call is invalid against stateless default instance."
    Dim Binding As IPropertyBinding
    For Each Binding In this.PropertyBindings
        If Source Is Binding.Source.Context Then Binding.Apply
    Next
End Sub

Public Property Get Context() As IAppContext
    Set Context = this.Context
End Property

Friend Property Set Context(ByVal RHS As IAppContext)
    GuardClauses.GuardDefaultInstance Me, BindingManager
    GuardClauses.GuardNullReference RHS
    GuardClauses.GuardDoubleInitialization this.Context, TypeName(Me)
    Set this.Context = RHS
End Property

Public Property Get Handlers() As Collection
    Set Handlers = this.PropertyBindings
End Property

Public Property Get PropertyBindings() As Collection
    Set PropertyBindings = this.PropertyBindings
End Property

Public Property Get StringFormatterFactory() As IStringFormatterFactory
    Set StringFormatterFactory = this.StringFormatterFactory
End Property

Public Property Set StringFormatterFactory(ByVal RHS As IStringFormatterFactory)
    GuardClauses.GuardDefaultInstance Me, BindingManager
    GuardClauses.GuardNullReference RHS
    GuardClauses.GuardDoubleInitialization this.StringFormatterFactory, TypeName(Me)
    Set this.StringFormatterFactory = RHS
End Property

Public Property Get DebugOutput() As Boolean
    DebugOutput = this.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    this.DebugOutput = RHS
End Property

'@Description "Releases all held property bindings, prepares the object for proper destruction."
Public Sub Terminate()
Attribute Terminate.VB_Description = "Releases all held property bindings, prepares the object for proper destruction."
    '@Ignore VariableNotUsed
    Dim Index As Long
    For Index = 1 To this.PropertyBindings.Count
        this.PropertyBindings.Remove 1
    Next
    Set this.PropertyBindings = Nothing
End Sub

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set this.BindingTargetStrategies = New Dictionary
        Set this.PropertyBindings = New Collection
        
        With this.BindingTargetStrategies
            .add "MSForms.CheckBox", New CheckBoxBindingStrategy
            .add "MSForms.CommandButton", New CommandButtonBindingStrategy
            .add "MSForms.ComboBox", New ComboBoxBindingStrategy
            .add "MSForms.Frame", New CaptionBindingStrategy
            .add "MSForms.Label", New CaptionBindingStrategy
            .add "MSForms.ListBox", New ListBoxBindingStrategy
            .add "MSForms.MultiPage", New MultiPageBindingStrategy
            .add "MSForms.OptionButton", New OptionButtonBindingStrategy
            .add "MSForms.ScrollBar", New ScrollBarBindingStrategy
            .add "MSForms.SpinButton", New SpinButtonBindingStrategy
            .add "MSForms.TabStrip", New TabStripBindingStrategy
            .add "MSForms.TextBox", New TextBoxBindingStrategy
            .add "Excel.Range", New WorksheetCellBindingStrategy
        End With
    End If
End Sub

Private Sub Class_Terminate()
    If Not IsDefaultInstance Then
        Set this.BindingTargetStrategies = Nothing
        Set this.PropertyBindings = Nothing
    End If
End Sub

Private Function IBindingManager_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As String, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
    
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me), "Member call is invalid against stateless default instance."
    
    Dim Formatter As IStringFormatter
    If StringFormat <> vbNullString And Not this.StringFormatterFactory Is Nothing Then
        Set Formatter = this.StringFormatterFactory.Create(StringFormat)
    End If
    
    Dim Binding As IPropertyBinding
    Dim Strategy As IBindingTargetStrategy
    If TryGetBindingStrategyFor(Target, outStrategy:=Strategy) Then
        If TargetProperty = vbNullString Then
            Set Binding = Strategy.DefaultPropertyBindingFor(this.Context, BindingPath.Create(Source, PropertyPath), Target, Mode, UpdateTrigger, Validator, Converter, Formatter, ValidationAdorner)
        Else
            Set Binding = Strategy.PropertyBindingFor(this.Context, BindingPath.Create(Source, PropertyPath), BindingPath.Create(Target, TargetProperty), Mode, UpdateTrigger, Validator, Converter, Formatter, ValidationAdorner)
        End If
    Else
        Set Binding = OneWayPropertyBinding _
            .Create(this.Context, _
                BindingPath.Create(Source, PropertyPath), _
                BindingPath.Create(Target, TargetProperty), _
                Validator:=Validator, _
                Converter:=Converter, _
                StringFormat:=Formatter, _
                ValidationAdorner:=ValidationAdorner)
    End If
    
    If TypeOf Source Is INotifyPropertyChanged Then
        Dim Notifier As INotifyPropertyChanged
        Set Notifier = Source
        Notifier.RegisterHandler Binding
    End If
    
    this.PropertyBindings.add Binding
    Set IBindingManager_BindPropertyPath = Binding
    If Not this.DebugOutput Then Exit Function
    
    On Error Resume Next
    If this.DebugOutput And TargetProperty = vbNullString Then
        Debug.Print TypeName(Me) & ": Binding property path '" & PropertyPath & "' to the default-binding property of type '" & TypeName(Target) & "'."
    ElseIf this.DebugOutput Then
        Debug.Print TypeName(Me) & ": Binding property path '" & PropertyPath & "' to '" & TypeName(Target) & "." & TargetProperty & "'."
    End If
    On Error GoTo 0
    
End Function

Private Function TryGetBindingStrategyFor(ByVal Target As Object, ByRef outStrategy As IBindingTargetStrategy) As Boolean
    
    Dim LibraryName As String
    If TypeOf Target Is MSForms.Control Then
        LibraryName = "MSForms"
    ElseIf TypeOf Target Is Excel.Range Then
        LibraryName = "Excel"
    ElseIf TypeOf Target Is MSForms.UserForm Then
        LibraryName = "MSForms"
    Else
        'FIXME now what, prefix other target libraries here? need a better way.
    End If
    
    Dim Key As String
    Key = LibraryName & "." & TypeName(Target)
    If this.BindingTargetStrategies.Exists(Key) Then
        Set outStrategy = this.BindingTargetStrategies.item(Key)
        TryGetBindingStrategyFor = True
    End If
    
End Function

Private Sub IBindingManager_Apply(ByVal Source As Object)
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me)
    Apply Source
End Sub

Private Property Get IBindingManager_DebugOutput() As Boolean
    IBindingManager_DebugOutput = this.DebugOutput
End Property

Private Property Get IBindingManager_NotifierFactory() As INotifierFactory
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me)
    Set IBindingManager_NotifierFactory = this.NotifierFactory
End Property

Private Property Get IBindingManager_StringFormatterFactory() As IStringFormatterFactory
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me)
    Set IBindingManager_StringFormatterFactory = this.StringFormatterFactory
End Property

Private Property Get IBindingManager_Validation() As IValidationManager
    GuardClauses.GuardDefaultInstance Me, BindingManager, TypeName(Me)
    Set IBindingManager_Validation = this.Context.Validation
End Property

Private Sub IDisposable_Dispose()
    Set this.Context = Nothing
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Context.Commands.EvaluateCanExecute Source
End Sub
