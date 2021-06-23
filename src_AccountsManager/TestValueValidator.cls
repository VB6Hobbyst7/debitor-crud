VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit
Implements IValueValidator
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    IsValid As Boolean
    message As String
End Type

Private this As TState

'@Description "Gets/sets a value that determines whether this validator evaluates to True or False."
Public Property Get IsValid() As Boolean
Attribute IsValid.VB_Description = "Gets/sets a value that determines whether this validator evaluates to True or False."
    IsValid = this.IsValid
End Property

Public Property Let IsValid(ByVal RHS As Boolean)
    this.IsValid = RHS
End Property

'@Description "Gets/sets the validation error message for this validator."
Public Property Get message() As String
Attribute message.VB_Description = "Gets/sets the validation error message for this validator."
    message = this.message
End Property

Public Property Let message(ByVal RHS As String)
    this.message = RHS
End Property

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase
End Sub

Private Function IValueValidator_IsValid(ByVal value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    this.Stub.OnInvoke "IsValid"
    IValueValidator_IsValid = this.IsValid
End Function

Private Property Get IValueValidator_Message() As String
    this.Stub.OnInvoke "Message"
    IValueValidator_Message = this.message
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

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
