VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBindingBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder AccountsManager.Infrastructure.Bindings
'@PredeclaredId
'@Exposed
Option Explicit
Implements IPropertyBinding
Implements IDisposable

Implements IControlEvents   '<~ relays MSForms.Control events to more specific IPropertyBinding implementations
Implements IHandleControlEvents '<~ ControlEventPunk callbacks
Implements IHandlePropertyChanged '<~ handles INotifyPropertyChange events from ViewModel (i.e. "applies" the binding)

Private Const DefaultTargetControlProperty As String = "Value"

Public Enum ApplyResult
    BindingSkipped
    BindingSuccess
    BindingValidationError
    BindingConversionError
    BindingFailed
End Enum

Private Type TState
    Applied As Boolean
    Applying As Boolean
    
    Punk As ControlEventsPunk
    Handlers As Collection
    
    Context As IAppContext
    Source As IBindingPath
    Target As IBindingPath
    
    Mode As BindingMode
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    
    Converter As IValueConverter
    StringFormat As IStringFormatter
    
    Validator As IValueValidator
    ValidationAdorner As IDynamicAdorner ' data validations (might be unset)
    CancelExitOnValidationError As Boolean
    
End Type

Private this As TState

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal TargetContext As Object, _
Optional ByVal TargetPropertyPath As String = DefaultTargetControlProperty, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnExit, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal ValidationAdorner As IDynamicAdorner, _
Optional ByVal CancelExitOnValidationError As Boolean = True) As IPropertyBinding

    GuardClauses.GuardNonDefaultInstance Me, PropertyBindingBase, TypeName(Me)
    
    Dim result As PropertyBindingBase
    Set result = New PropertyBindingBase
    
    result.CancelExitOnValidationError = CancelExitOnValidationError
    result.Mode = Mode
    result.UpdateSourceTrigger = UpdateSource
    
    Set result.Source = Source
    Set result.Target = BindingPath.Create(TargetContext, TargetPropertyPath)
    
    Set result.Context = Context
    
    Set result.Converter = Converter
    Set result.StringFormat = StringFormat
    
    Set result.Validator = Validator
    Set result.ValidationAdorner = ValidationAdorner
    
    Set Create = result
    
End Function

Public Property Get AsIControlEvents() As IControlEvents
    Set AsIControlEvents = Me
End Property

Public Property Get AsINotifyValidationError() As INotifyValidationError
    Set AsINotifyValidationError = Me
End Property

Public Property Get Source() As IBindingPath
    Set Source = this.Source
End Property

Friend Property Set Source(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization this.Source, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Source = RHS
End Property

Public Property Get Target() As IBindingPath
    Set Target = this.Target
End Property

Friend Property Set Target(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Target = RHS
    If TypeOf this.Target.Context Is MSForms.Control Then
        On Error Resume Next
        Set this.Punk.Target = this.Target.Context
        If this.Punk.Connect Then
            this.Punk.RegisterHandler Me
        Else
            DebugMessage "Unable to connect target IConnectionPoint."
        End If
        If Err.number <> 0 Then
            DebugMessage "Unable to handle MSForms.Control events. Binding is effectively one-way. " & Err.Description
        End If
        On Error GoTo 0
    Else
        Set this.Punk = Nothing
    End If
End Property

Public Property Get Mode() As BindingMode
    Mode = this.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    this.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    this.UpdateSourceTrigger = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = this.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    GuardClauses.GuardDoubleInitialization this.Converter, TypeName(Me)
    Set this.Converter = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = this.Validator
End Property

Friend Property Set Validator(ByVal RHS As IValueValidator)
    GuardClauses.GuardDoubleInitialization this.Validator, TypeName(Me)
    Set this.Validator = RHS
    If Not this.Validator Is Nothing Then this.UpdateSourceTrigger = this.Validator.Trigger
End Property

Public Property Get StringFormat() As IStringFormatter
    Set StringFormat = this.StringFormat
End Property

Friend Property Set StringFormat(ByVal RHS As IStringFormatter)
    GuardClauses.GuardDoubleInitialization this.StringFormat, TypeName(Me)
    Set this.StringFormat = RHS
End Property

Public Property Get ValidationAdorner() As IDynamicAdorner
    Set ValidationAdorner = this.ValidationAdorner
End Property

Friend Property Set ValidationAdorner(ByVal RHS As IDynamicAdorner)
    GuardClauses.GuardDoubleInitialization this.ValidationAdorner, TypeName(Me)
    Set this.ValidationAdorner = RHS
End Property

Public Property Get Context() As IAppContext
    Set Context = this.Context
End Property

Friend Property Set Context(ByVal RHS As IAppContext)
    GuardClauses.GuardDoubleInitialization this.Context, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Context = RHS
End Property

Public Property Get CancelExitOnValidationError() As Boolean
    CancelExitOnValidationError = this.CancelExitOnValidationError
End Property

Friend Property Let CancelExitOnValidationError(ByVal RHS As Boolean)
    this.CancelExitOnValidationError = RHS
End Property

Private Function TryConvert(ByRef Value As Variant) As Boolean
    
    If this.Converter Is Nothing Then
        TryConvert = True
        Exit Function
    End If
    
    On Error Resume Next
    Value = this.Converter.Convert(Value)
    If Err.number = 0 Then
        DebugMessage "Value was successfully converted."
        TryConvert = True
    Else
        DebugMessage "IValueConverter.Convert raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

Private Function TryConvertBack(ByRef Value As Variant) As Boolean
    
    If this.Converter Is Nothing Then
        TryConvertBack = True
        Exit Function
    End If
    
    On Error Resume Next
    Value = this.Converter.ConvertBack(Value)
    If Err.number = 0 Then
        DebugMessage "Value was successfully converted back."
        TryConvertBack = True
    Else
        DebugMessage "IValueConverter.ConvertBack raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

Public Sub Apply()
    If this.Applying Then Exit Sub '<~ could break accidental recursion... leave this here...
    
    If this.Mode = OneTimeBinding Then
        If Not this.Applied Then ApplyToTarget
        
    ElseIf this.Mode = OneWayBinding Or this.Mode = TwoWayBinding Then
        ApplyToTarget
        
    ElseIf this.Mode = OneWayToSource Then
        ApplyToSource
        
    End If

End Sub

Private Property Get CanApplyToTarget() As Boolean
    'one-way to source and already-applied one-time binding modes do not apply to target
    CanApplyToTarget = Not this.Applying And this.Mode <> OneWayToSource And (this.Mode <> OneTimeBinding Or Not this.Applied)
End Property

Public Sub ApplyToTarget()
'reads from the source and writes to the target.
    
    If Not CanApplyToTarget Then Exit Sub
    this.Applying = True
    
    Dim SourceValue As Variant
    Dim UseFallbackDefault As Boolean
    
    If this.Source.object Is Nothing Then
        UseFallbackDefault = TryGetDefaultBindingValue(outValue:=SourceValue)
    End If
    
    this.Source.Resolve
    
    If this.Source.object Is Nothing And Not UseFallbackDefault Then
        DebugMessage "ApplyToTarget aborted: source object in path '" & this.Source.path & "' is Nothing and there is no fallback default value for this binding."
        this.Applying = False
        Exit Sub
        
    ElseIf this.Source.object Is Nothing Then
        DebugMessage "ApplyToTarget: source object in path '" & this.Source.path & "' is Nothing; binding target property '" & this.Target.PropertyName & "' to default/fallback value."
        
    Else
        If Not this.Source.TryReadPropertyValue(outValue:=SourceValue) Then
            DebugMessage "ApplyToTarget failed to read current source property value."
            this.Applying = False
            Exit Sub
        End If
    End If
    
    Validate SourceValue, Propagate:=False
    SourceValue = Convert(SourceValue)
    
    this.Target.Resolve
    Dim CurrentValue As Variant
    If Not this.Target.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToTarget failed to read current target property value."
        this.Applying = False
        Exit Sub
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(SourceValue) Then
        Proceed = SourceValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        If Not this.StringFormat Is Nothing Then
            SourceValue = this.StringFormat.Format(SourceValue)
        End If
        
        If Not this.Target.TryWritePropertyValue(SourceValue) Then
            DebugMessage "ApplyToTarget failed to write to target property."
        Else
            DebugMessage "ApplyToTarget successfully applied to target."
            this.Applied = True
        End If
    End If
    this.Applying = False
End Sub

Private Property Get CanApplyToSource() As Boolean
    'one-time and one-way binding modes do not write values to the binding source (ViewModel).
    CanApplyToSource = Not this.Applying And this.Mode <> OneTimeBinding And this.Mode <> OneWayBinding
End Property

Public Function ApplyToSource(Optional ByVal TargetValidationValue As Variant) As ApplyResult
'reads from the target and writes to the source.
'when This.Mode=KeyPress, the TargetValidationValue contains the KeyAscii value.
    On Error GoTo CleanFail

    If Not CanApplyToSource Then Exit Function
    this.Applying = True
    
    this.Target.Resolve
    Dim TargetValue As Variant
    If Not this.Target.TryReadPropertyValue(outValue:=TargetValue) Then
        DebugMessage "ApplyToSource failed to read target property value."
        ApplyToSource = ApplyResult.BindingFailed
        this.Applying = False
        Exit Function
    End If
    
    If IsMissing(TargetValidationValue) Then
        TargetValidationValue = TargetValue
        
    ElseIf this.UpdateSourceTrigger = BindingUpdateSourceTrigger.OnKeyPress Then
        'keypress hasn't reached the control yet:
        TargetValue = InsertKeypress(TargetValue, TargetValidationValue)
        
    End If
    
    If Not TryConvertBack(TargetValue) Then
        ApplyToSource = ApplyResult.BindingConversionError
        DebugMessage TypeName(this.Converter) & " ConvertBack failed to convert the target value."
        OnEnter
        this.Applying = False
        Exit Function
    End If
    
    If Not Validate(TargetValidationValue) Then
        ApplyToSource = ApplyResult.BindingValidationError
        DebugMessage TypeName(this.Validator) & " failed to validate the target value."
        this.Applying = False
        Exit Function
    End If
    
    this.Source.Resolve
    Dim CurrentValue As Variant
    If Not this.Source.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToSource failed to read current source property value."
        ApplyToSource = ApplyResult.BindingFailed
        this.Applying = False
        Exit Function
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(CurrentValue) Then
        Proceed = TargetValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        ClearValidationErrors
        
        If Not this.Source.TryWritePropertyValue(TargetValue) Then
            DebugMessage "ApplyToSource failed to write to source property."
            ApplyToSource = ApplyResult.BindingFailed
            Exit Function
        End If
        
        DebugMessage "Binding was successfully applied to source."
        ApplyToSource = ApplyResult.BindingSuccess
        
    End If
    
CleanExit:
    this.Applying = False
    Exit Function
    
CleanFail:
    DebugMessage "ApplyToSource failed unexpectedly."
    ApplyToSource = ApplyResult.BindingFailed
    Resume CleanExit
    Resume
End Function

Private Function InsertKeypress(ByVal InitialValue As String, ByVal Char As String) As String
'FIXME this function is at the wrong abstraction level.
    
    If Char = vbNullString Then
        InsertKeypress = InitialValue
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim SelectionStart As Long
    SelectionStart = this.Target.object.SelStart
    
    Dim SelectionLength As Long
    SelectionLength = this.Target.object.SelLength
    
    On Error GoTo 0
    
    Select Case True
        
        Case SelectionStart > 0 And SelectionLength = 0
            
            'insert at SelStart
            With StringBuilderNet.Append(InitialValue) _
                              .Insert(SelectionStart, Char)
                InsertKeypress = .ToString
            End With
        
        Case SelectionStart > 0 And SelectionLength > 0
            
            'replace selection
            With StringBuilderNet.Append(InitialValue) _
                              .Remove(SelectionStart, SelectionLength) _
                              .Insert(SelectionStart, Char)
                     
                InsertKeypress = .ToString
            End With
            
        Case Else
            
            'replace initial value
            InsertKeypress = Char
    
    End Select
    
End Function

Private Function IsValid() As Boolean
    IsValid = this.Context.Validation.IsValid(this.Source.object, this.Source.PropertyName)
End Function

Private Sub ClearValidationErrors()
    With this.Context
        .Validation.ClearValidationError this.Source
        .Commands.EvaluateCanExecute this.Source
    End With
End Sub

Private Function ToString() As String
    ToString = TypeName(this.Source.Context) & "." & this.Source.path & " -> " & TypeName(this.Target.Context) & "." & this.Target.path
End Function

Private Function ResolvePropertyPath(ByVal Source As Object, ByVal PropertyPath As String) As Object
    
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    
    If UBound(Parts) = LBound(Parts) Then
        Set ResolvePropertyPath = Source
    Else
        Dim RecursiveProperty As Object
        Set RecursiveProperty = CallByName(Source, Parts(0), VbGet)
        If RecursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(RecursiveProperty, right$(PropertyPath, Len(PropertyPath) - Len(Parts(0)) - 1))
    End If
    
End Function

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    ResolvePropertyName = Parts(UBound(Parts))
End Function

Private Function TryGetDefaultBindingValue(ByRef outValue As Variant) As Boolean
'Gets a default value for certain specific target properties, used when source path cannot be fully resolved,
'e.g. when target binds to "SomeObjectProperty.SomeProperty" and "SomeObjectProperty" is Nothing.

    Select Case this.Target.PropertyName
    
        Case "Text", "Caption"
            outValue = vbNullString
            TryGetDefaultBindingValue = True
            
        Case "Enabled", "Visible"
            outValue = False
            TryGetDefaultBindingValue = True
            
        Case "Value"
            If TypeOf this.Target Is MSForms.CheckBox _
                Or TypeOf this.Target Is MSForms.OptionButton _
            Then
                outValue = False
                TryGetDefaultBindingValue = True
            End If
            
    End Select
End Function

Private Function Convert(ByVal Value As Variant) As Variant
    
    On Error GoTo CleanFail
    If Not this.Converter Is Nothing Then
        Convert = this.Converter.Convert(Value)
    Else
        Convert = Value
    End If
    
CleanExit:
    Exit Function
    
CleanFail:
    DebugMessage "IValueConverter failed with error " & Err.number & ": " & Err.Description
    OnEnter
    RethrowOnError
    Resume CleanExit 'NOTE: 'VF: is never reached
End Function

Public Function Validate(ByVal TargetValue As Variant, Optional ByVal Propagate As Boolean = True) As Boolean

    If this.Validator Is Nothing Then
        Validate = True
        Exit Function
    End If
    
    On Error GoTo CleanFail
    If this.Validator.IsValid(TargetValue, this.Source, this.Target) Then
        
        If Propagate Then
            this.Context.Validation.ClearValidationError this.Source
            If Not this.ValidationAdorner Is Nothing Then this.ValidationAdorner.Hide
        End If
        
        Validate = True
        
    Else
        
        If Propagate Then
            Dim Message As String
            Message = this.Validator.Message
            
            this.Context.Validation.OnValidationError _
                Context:=this.Source.Context, _
                ValidationError:=ValidationError.Create(Me, Message)
            
            If Not this.ValidationAdorner Is Nothing Then this.ValidationAdorner.Show Message
            'VF: commented out as it leads to select all, which is not desired in text and numeric textboxes even though there may be more than one eventhandler
            'OnEnter
        End If
    End If
        
CleanExit:
    If Propagate Then
        this.Context.Commands.EvaluateCanExecute this.Source.Context
    End If
    Exit Function
    
CleanFail:
    DebugMessage "IValueValidator failed with error " & Err.number & ": " & Err.Description
    OnEnter
    Resume CleanExit
    Resume
End Function

Private Sub Format()
    If this.StringFormat Is Nothing Then Exit Sub

    Dim BindingValue As Variant
    If Not this.Source.TryReadPropertyValue(outValue:=BindingValue) Then Exit Sub
        
    On Error GoTo CleanFail
    Dim FormattedValue As String
    FormattedValue = this.StringFormat.Format(BindingValue)
    On Error GoTo 0
    
    this.Applying = True
    If Not this.Target.TryWritePropertyValue(FormattedValue) Then
        DebugMessage "Could not format the binding value."
    End If
    this.Applying = False
    
CleanExit:
    Exit Sub
CleanFail:
    DebugMessage "IStringFormatter.Format() raised an error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub OnAfterUpdate()
    Format
    Dim Handler As IHandleControlEvents
    For Each Handler In this.Handlers
        Handler.HandleAfterUpdate
    Next
End Sub

Private Sub OnBeforeUpdate(ByRef Cancel As Boolean)
    OnBindingUpdateSourceOpportunity Cancel, BindingUpdateSourceTrigger.OnPropertyChanged
    Dim Handler As IHandleControlEvents
    For Each Handler In this.Handlers
        Handler.HandleBeforeUpdate Cancel
        If Cancel Then Exit For
    Next
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub OnChange()
    Format
    Dim Handler As IHandleControlEvents
    For Each Handler In this.Handlers
        Handler.HandleChange
    Next
End Sub

Private Sub OnEnter()
    Dim Handler As IHandleControlEvents
    For Each Handler In this.Handlers
        Handler.HandleEnter
    Next
End Sub

Private Sub OnExit(ByRef Cancel As Boolean)
    OnBindingUpdateSourceOpportunity Cancel, BindingUpdateSourceTrigger.OnExit
    Dim Handler As IHandleControlEvents
    For Each Handler In this.Handlers
        Handler.HandleExit Cancel
    Next
End Sub

Public Sub OnBindingUpdateSourceOpportunity(ByRef Cancel As Boolean, ByVal Trigger As BindingUpdateSourceTrigger, Optional ByVal TargetValidationValue As Variant)
    If Not CanApplyToSource Then Exit Sub
    
    Dim result As ApplyResult
    Select Case Trigger
        
        'VF: ? remove OnKeyPress here?, at least in TextBox no longer used
        'VF: changed: required to capture Cut/Paste and Backspace/Delete in TextBox
        'Case BindingUpdateSourceTrigger.OnKeyPress
        Case BindingUpdateSourceTrigger.OnKeyPress, BindingUpdateSourceTrigger.OnChange
        
            'If This.UpdateSourceTrigger = OnKeyPress And Not IsMissing(TargetValidationValue) Then
            If (this.UpdateSourceTrigger = OnKeyPress Or this.UpdateSourceTrigger = BindingUpdateSourceTrigger.OnChange) _
                And Not IsMissing(TargetValidationValue) Then
                'just validate the keypress 'VF: or string pasted
                result = IIf(Validate(TargetValidationValue), _
                    ApplyResult.BindingSuccess, _
                    ApplyResult.BindingValidationError)
                            
            End If
            
        Case BindingUpdateSourceTrigger.OnPropertyChanged
            
            If this.UpdateSourceTrigger = OnPropertyChanged Or this.UpdateSourceTrigger = OnKeyPress Then
                result = ApplyToSource
            End If
        
        Case BindingUpdateSourceTrigger.OnExit
            
            If this.UpdateSourceTrigger <> Never Then
                
                Dim outValue As Variant
                If this.Target.TryReadPropertyValue(outValue) Then
                
                    result = IIf(Validate(outValue), _
                        ApplyResult.BindingSuccess, _
                        ApplyResult.BindingValidationError)
                        
                End If
            End If
            
    End Select
    
    Cancel = this.CancelExitOnValidationError And (Cancel Or (result = BindingValidationError) Or (result = BindingConversionError))
    
End Sub

Private Sub OnValidationError()
    this.Context.Validation.OnValidationError this.Source.Context, ValidationError.Create(Me, this.Validator.Message)
    this.ValidationAdorner.Show this.Validator.Message
End Sub

Private Sub DebugMessage(ByVal Message As String)
'    If this.Context.Bindings.DebugOutput Then Debug.Print "PropertyBinding: [" & ToString & "] " & Message
    If this.Context.Bindings.DebugOutput Then Logging.logINFO "PropertyBinding: [" & ToString & "] " & Message
End Sub

Private Sub Class_Initialize()
    If Not Me Is PropertyBindingBase Then
        Set this.Handlers = New Collection
        Set this.Punk = New ControlEventsPunk
    End If
End Sub

Private Sub Class_Terminate()
    Set this.Handlers = Nothing
    If Not this.Punk Is Nothing Then
        this.Punk.Disconnect
        Set this.Punk = Nothing
    End If
End Sub

Private Sub IControlEvents_OnAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IControlEvents_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    OnBeforeUpdate LocalCancel
    Cancel.Value = LocalCancel
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IControlEvents_OnChange()
    OnChange
End Sub

Private Sub IControlEvents_OnEnter()
    OnEnter
End Sub

Private Sub IControlEvents_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    OnExit LocalCancel
    Cancel = LocalCancel
End Sub

Private Sub IControlEvents_RegisterHandler(ByVal Handler As IHandleControlEvents)
    this.Handlers.add Handler
End Sub

Private Sub IDisposable_Dispose()
    
    Set this.Context = Nothing
    
    this.Punk.Disconnect
    Set this.Punk = Nothing
    Disposable.TryDispose this.ValidationAdorner
    Set this.ValidationAdorner = Nothing
    
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef Cancel As Boolean)
    OnBeforeUpdate Cancel
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IHandleControlEvents_HandleChange()
    OnChange
End Sub

Private Sub IHandleControlEvents_HandleEnter()
    OnEnter
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef Cancel As Boolean)
    OnExit Cancel
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    'IHandlePropertyChanged -> INotifyPropertyChanged: source is the ViewModel, so we're propagating to Target here:
    If Not this.Applying And Source Is this.Source.object And PropertyName = this.Source.PropertyName Then
        ApplyToTarget
    End If
End Sub

Private Sub IPropertyBinding_Apply()
    Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = this.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = this.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = DefaultTargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = this.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = this.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = this.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = this.Validator
End Property


