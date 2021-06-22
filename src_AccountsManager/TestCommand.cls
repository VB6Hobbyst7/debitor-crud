VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an ICommand dependency in a unit test."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub an ICommand dependency in a unit test."
Option Explicit
Implements ICommand
Implements ITestStub

Private Const CanExecuteDefaultValue As Boolean = True

Private Type TState
    Stub As TestStubBase
    CanExecute As Boolean
End Type

Private this As TState

'@Description "Controls the return value of ICommand.CanExecute. True by default."
Public Property Get CanExecute() As Boolean
Attribute CanExecute.VB_Description = "Controls the return value of ICommand.CanExecute. True by default."
    CanExecute = this.CanExecute
End Property

Public Property Let CanExecute(ByVal RHS As Boolean)
    this.CanExecute = RHS
End Property

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase
    this.CanExecute = CanExecuteDefaultValue
End Sub

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    this.Stub.OnInvoke "CanExecute"
    ICommand_CanExecute = this.CanExecute
End Function

Private Property Get ICommand_Description() As String
    this.Stub.OnInvoke "Description.Get"
    ICommand_Description = "A " & TypeName(Me) & " instance tracks the number of times its members are invoked."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    this.Stub.OnInvoke "Execute"
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
