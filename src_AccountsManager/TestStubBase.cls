VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestStubBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit

Private Type TState
    MemberInvokes As Dictionary
End Type

Private this As TState

'@Description "Tracks a new invoke of the specified member."
Public Sub OnInvoke(ByVal MemberName As String)
Attribute OnInvoke.VB_Description = "Tracks a new invoke of the specified member."
    Dim NewValue As Long
    If this.MemberInvokes.Exists(MemberName) Then
        NewValue = this.MemberInvokes.item(MemberName) + 1
        this.MemberInvokes.Remove MemberName
    Else
        NewValue = 1
    End If
    this.MemberInvokes.add MemberName, NewValue
    Debug.Print TypeName(Me) & ": OnInvoke(" & MemberName & ") was invoked"
End Sub

'@Description "Gets the number of invokes made against the specified member in the lifetime of this object."
Public Property Get MemberInvokes(ByVal MemberName As String) As Long
Attribute MemberInvokes.VB_Description = "Gets the number of invokes made against the specified member in the lifetime of this object."
    If this.MemberInvokes.Exists(MemberName) Then
        MemberInvokes = this.MemberInvokes.item(MemberName)
    Else
        MemberInvokes = 0
    End If
End Property

'@Description "Gets a string listing the MemberInvokes cache content."
Public Function ToString() As String
Attribute ToString.VB_Description = "Gets a string listing the MemberInvokes cache content."
    Dim MemberNames As Variant
    MemberNames = this.MemberInvokes.Keys
    
    With New StringBuilderNet
        Dim i As Long
        For i = LBound(MemberNames) To UBound(MemberNames)
            Dim Name As String
            Name = MemberNames(i)
            .AppendFormat "{0} was invoked {1} time(s)", Name, this.MemberInvokes.item(Name)
        Next
        ToString = .ToString
    End With
    
End Function

Public Sub Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    Dim Invokes As Long
    Invokes = MemberInvokes(MemberName)
    Assert.AreEqual ExpectedInvokes, Invokes, "Member was invoked " & Invokes & " time(s), but test expected " & ExpectedInvokes & "."
End Sub

Private Sub Class_Initialize()
    Set this.MemberInvokes = New Dictionary
End Sub

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = this.MemberInvokes.item(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    Verify Assert, MemberName, ExpectedInvokes
End Sub

