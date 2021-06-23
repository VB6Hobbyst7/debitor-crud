VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder AccountsManager.Infrastructure.Bindings.PropertyBindings
'@PredeclaredId
'@Exposed
Option Explicit
Implements IPropertyBinding
Implements IDisposable
Implements IHandlePropertyChanged
Implements IHandleControlEvents

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents TargetEventSource As MSForms.ComboBox
Attribute TargetEventSource.VB_VarHelpID = -1

Private Const DefaultTargetControlProperty As String = "Value"
Private Const FormsControlProgId As String = "Forms.ComboBox.1"

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private this As TState

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Public Property Get ProgID() As String
    ProgID = FormsControlProgId
End Property

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As MSForms.ComboBox, _
Optional ByVal TargetProperty As Variant, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim LocalTargetProperty As String
    If IsMissing(TargetProperty) Then
        LocalTargetProperty = InferTargetPropertyFromSource(Source, outMode:=Mode)
    Else
        LocalTargetProperty = TargetProperty
    End If
    
    Dim BindingBase As PropertyBindingBase
    Set BindingBase = PropertyBindingBase _
        .Create(Context, Source, Target, LocalTargetProperty, _
            Mode:=Mode, _
            UpdateSource:=UpdateSource, _
            Converter:=Converter, _
            StringFormat:=StringFormat, _
            Validator:=Validator, _
            ValidationAdorner:=ValidationAdorner)
    
    Dim result As ComboBoxPropertyBinding
    Set result = New ComboBoxPropertyBinding
    
    If Not Validator Is Nothing Then
        BindingBase.AsINotifyValidationError.RegisterHandler ValidationManager
    End If

    result.InjectBindingInfo BindingBase
    Set Create = result
    
End Function

Private Function InferTargetPropertyFromSource(ByVal Source As IBindingPath, ByRef outMode As BindingMode) As String
    Dim result As String
    Dim SourceValue As Variant
    If Source.TryReadPropertyValue(outValue:=SourceValue) Then
        Select Case True
            Case VarType(SourceValue) = VbVarType.vbLong
                result = "ListIndex"
            Case VarType(SourceValue) = VbVarType.vbString
                result = "Text"
            Case IsArray(SourceValue) Or VarType(SourceValue) = VbVarType.vbArray
                Debug.Print TypeName(Me) & ": binding [" & Source.ToString & "] one-way."
                outMode = BindingMode.OneWayBinding
                result = "List"
            Case Else
                Debug.Print TypeName(Me) & ": binding [" & Source.ToString & "] one-way to source."
                outMode = BindingMode.OneWayToSource
                result = DefaultTargetControlProperty
        End Select
    Else
        Debug.Print TypeName(Me) & ": Could not read source property value; binding to default target property."
        result = DefaultTargetControlProperty
    End If
    InferTargetPropertyFromSource = result
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, ComboBoxPropertyBinding, TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Handler, TypeName(Me)
    Set this.Base = BindingInfo
    Set this.Handler = BindingInfo
    this.Base.AsIControlEvents.RegisterHandler Me
End Sub

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is TextBoxPropertyBinding
End Property

Private Sub IDisposable_Dispose()
    Set this.Handler = Nothing
    Disposable.TryDispose this.Base
    Set this.Base = Nothing
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef Cancel As Boolean)
End Sub

Private Sub IHandleControlEvents_HandleChange()
End Sub

Private Sub IHandleControlEvents_HandleEnter()
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef Cancel As Boolean)
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Handler.HandlePropertyChanged Source, PropertyName
End Sub

Private Sub IPropertyBinding_Apply()
    this.Base.Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = this.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = this.Base.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = DefaultTargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Base.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = this.Base.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = this.Base.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = this.Base.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = this.Base.Validator
End Property

Private Sub CheckBoxEventSource_Change()
    If this.Base.UpdateSourceTrigger = OnPropertyChanged Then this.Base.ApplyToSource
End Sub


