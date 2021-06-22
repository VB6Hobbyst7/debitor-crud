VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBindingObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub a binding source or target for unit tests."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub a binding source or target for unit tests."
'@PredeclaredId
Option Explicit
Implements ITestStub
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged ' intercept events handled in PropertyBindingBase

Private Type TState
    Stub As TestStubBase
    Handlers As Collection
    TestStringProperty As String
    TestNumericProperty As Long
    TestBindingObjectProperty As TestBindingObject
End Type

Private this As TState

Public Function Create() As TestBindingObject
    GuardClauses.GuardNonDefaultInstance Me, TestBindingObject, TypeName(Me)
    Dim result As TestBindingObject
    Set result = New TestBindingObject
    Set Create = result
End Function

Public Property Get TestStringProperty() As String
    this.Stub.OnInvoke "TestStringProperty.Get"
    TestStringProperty = this.TestStringProperty
End Property

Public Property Let TestStringProperty(ByVal RHS As String)
    this.Stub.OnInvoke "TestStringProperty.Let"
    If this.TestStringProperty <> RHS Then
        this.TestStringProperty = RHS
        OnPropertyChanged Me, "TestStringProperty"
    End If
End Property

Public Property Get TestNumericProperty() As Long
    this.Stub.OnInvoke "TestNumericProperty.Get"
    TestNumericProperty = this.TestNumericProperty
End Property

Public Property Let TestNumericProperty(ByVal RHS As Long)
    this.Stub.OnInvoke "TestNumericProperty.Let"
    If this.TestNumericProperty <> RHS Then
        this.TestNumericProperty = RHS
        OnPropertyChanged Me, "TestNumericProperty"
    End If
End Property

Public Property Get TestBindingObjectProperty() As TestBindingObject
    this.Stub.OnInvoke "TestBindingObjectProperty.Get"
    Set TestBindingObjectProperty = this.TestBindingObjectProperty
End Property

Public Property Set TestBindingObjectProperty(ByVal RHS As TestBindingObject)
    this.Stub.OnInvoke "TestBindingObjectProperty.Set"
    If Not this.TestBindingObjectProperty Is RHS Then
        Set this.TestBindingObjectProperty = RHS
        OnPropertyChanged Me, "TestBindingObjectProperty"
    End If
End Property

Private Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase
    Set this.Handlers = New Collection
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Handlers.add Handler
End Sub

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = this.Stub.MemberInvokes(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = this.Stub.ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    this.Stub.Verify Assert, MemberName, ExpectedInvokes
End Sub
