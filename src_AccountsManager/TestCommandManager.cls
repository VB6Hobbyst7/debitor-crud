VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit
Implements ICommandManager
Implements ITestStub

Private Type TState
    Stub As TestStubBase
End Type

Private this As TState

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase
End Sub

Private Function ICommandManager_BindCommand(ByVal BindingContext As Object, ByVal Target As Object, ByVal command As ICommand) As ICommandBinding
    this.Stub.OnInvoke "BindCommand"
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal BindingContext As Object)
    this.Stub.OnInvoke "EvaluateCanExecute"
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


