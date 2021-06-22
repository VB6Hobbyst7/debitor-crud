VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an IPropertyBinding dependency in a unit test."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub an IPropertyBinding dependency in a unit test."
'@PredeclaredId
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private this As TState

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
Optional ByVal TargetProperty As String = "Value", _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnExit, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As ITestStub
    Dim result As TestPropertyBinding
    Set result = New TestPropertyBinding
    
    Dim Base As PropertyBindingBase
    Set Base = PropertyBindingBase.Create(Context, Source, Target, TargetProperty, Mode, UpdateSource, Converter, StringFormat, Validator, ValidationAdorner)
    result.InjectBindingInfo Base
    Set Create = result
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, TestPropertyBinding, TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Handler, TypeName(Me)
    Set this.Base = BindingInfo
    Set this.Handler = BindingInfo
End Sub

Public Property Get AsIHandlePropertyChanged() As IHandlePropertyChanged
    Set AsIHandlePropertyChanged = Me
End Property

Public Property Get Source() As IBindingPath
    Set Source = this.Base.Source
End Property

Public Property Set Source(ByVal RHS As IBindingPath)
    Set this.Base.Source = RHS
End Property

Public Property Get Target() As IBindingPath
    Set Target = this.Base.Target
End Property

Public Property Set Target(ByVal RHS As IBindingPath)
    Set this.Base.Target = RHS
End Property

Public Property Get Mode() As BindingMode
    Mode = this.Base.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    this.Base.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = this.Base.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    this.Base.UpdateSourceTrigger = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = this.Base.Validator
End Property

Public Property Set Validator(ByVal RHS As IValueValidator)
    Set this.Base.Validator = RHS
End Property

Public Property Get ValidationAdorner() As IDynamicAdorner
    Set ValidationAdorner = this.Base.ValidationAdorner
End Property

Public Property Set ValidationAdorner(ByVal RHS As IDynamicAdorner)
    Set this.Base.ValidationAdorner = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = this.Base.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    Set this.Base.Converter = RHS
End Property

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Stub.OnInvoke "HandlePropertyChanged"
    this.Handler.HandlePropertyChanged Source, PropertyName
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = this.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    this.Stub.OnInvoke "DefaultTargetProperty.Get"
    IPropertyBinding_DefaultTargetProperty = vbNullString
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    this.Stub.OnInvoke "StringFormat.Get"
    Set IPropertyBinding_StringFormat = this.Base.StringFormat
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    this.Stub.OnInvoke "Validator.Get"
    Set IPropertyBinding_Validator = this.Base.Validator
End Property

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = this.Stub.MemberInvokes(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = this.Stub.ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    this.Stub.Verify Assert, MemberName, ExpectedInvokes
End Sub

Private Property Get IPropertyBinding_Mode() As BindingMode
    this.Stub.OnInvoke "Mode.Get"
    IPropertyBinding_Mode = this.Base.Mode
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    this.Stub.OnInvoke "UpdateSourceTrigger.Get"
    IPropertyBinding_UpdateSourceTrigger = this.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    this.Stub.OnInvoke "Source.Get"
    Set IPropertyBinding_Source = this.Base.Source
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    this.Stub.OnInvoke "Target.Get"
    Set IPropertyBinding_Target = this.Base.Target
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    this.Stub.OnInvoke "Converter.Get"
    Set IPropertyBinding_Converter = this.Base.Converter
End Property

Private Sub IPropertyBinding_Apply()
    this.Stub.OnInvoke "Apply"
End Sub

