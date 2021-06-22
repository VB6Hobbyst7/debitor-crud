VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyChangeNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Stubs")
Option Explicit
Implements INotifyPropertyChanged
Implements ITestStub

Private Type TState
    Handlers As Collection
    Stub As TestStubBase
End Type
Private this As TState

Public Property Get Handlers() As Collection
    Set Handlers = this.Handlers
End Property

Public Property Get Stub() As ITestStub
    Set Stub = this.Stub
End Property

Private Sub Class_Initialize()
    Set this.Handlers = New Collection
    Set this.Stub = New TestStubBase
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Stub.OnInvoke "OnPropertyChanged"
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Stub.OnInvoke "RegisterHandler"
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

