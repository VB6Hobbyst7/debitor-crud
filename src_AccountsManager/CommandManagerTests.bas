Attribute VB_Name = "CommandManagerTests"
'@Folder Tests.Bindings
'@TestModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
Private Assert As Object
#Else
Private Assert As Rubberduck.AssertClass
#End If

Private Type TState
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As CommandManager
    AbstractSUT As ICommandManager
    
    BindingContext As TestBindingObject
    command As TestCommand
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
#If LateBind Then
    'requires HKCU registration of the Rubberduck COM library.
    Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
#Else
    'requires project reference to the Rubberduck COM library.
    Set Assert = New Rubberduck.PermissiveAssertClass
#End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Test.ConcreteSUT = New CommandManager
    Set Test.AbstractSUT = Test.ConcreteSUT
    Set Test.BindingContext = New TestBindingObject
    Set Test.command = New TestCommand
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
    Set Test.AbstractSUT = Nothing
    Set Test.BindingContext = Nothing
    Set Test.command = Nothing
End Sub

Private Sub ExpectError()
    Dim message As String
    If Err.number = Test.ExpectedErrNumber Then
        If (Test.ExpectedErrSource = vbNullString) Or (Err.Source = Test.ExpectedErrSource) Then
            Test.ExpectedErrorCaught = True
        Else
            message = "An error was raised, but not from the expected source. " & _
                      "Expected: '" & TypeName(Test.ConcreteSUT) & "'; Actual: '" & Err.Source & "'."
        End If
    ElseIf Err.number <> 0 Then
        message = "An error was raised, but not with the expected number. Expected: '" & Test.ExpectedErrNumber & "'; Actual: '" & Err.number & "'."
    Else
        message = "No error was raised."
    End If
    
    If Not Test.ExpectedErrorCaught Then Assert.Fail message
End Sub

Private Function DefaultTargetCommandBindingFor(ByVal ProgID As String, ByRef outTarget As Object) As ICommandBinding
    Set outTarget = CreateObject(ProgID)
    Set DefaultTargetCommandBindingFor = Test.AbstractSUT.BindCommand(Test.BindingContext, outTarget, Test.command)
End Function

'@TestMethod("DefaultCommandTargetBindings")
Private Sub BindCommand_BindsCommandButton()
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.CommandButtonProgId, outTarget:=Target)
        Assert.AreSame Test.command, .command
        Assert.AreSame Target, .Target
    End With
End Sub

'@TestMethod("DefaultCommandTargetBindings")
Private Sub BindCommand_BindsCheckBox()
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.CheckBoxProgId, outTarget:=Target)
        Assert.AreSame Test.command, .command
        Assert.AreSame Target, .Target
    End With
End Sub

'@TestMethod("DefaultCommandTargetBindings")
Private Sub BindCommand_BindsImage()
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.ImageProgId, outTarget:=Target)
        Assert.AreSame Test.command, .command
        Assert.AreSame Target, .Target
    End With
End Sub

'@TestMethod("DefaultCommandTargetBindings")
Private Sub BindCommand_BindsLabel()
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.LabelProgId, outTarget:=Target)
        Assert.AreSame Test.command, .command
        Assert.AreSame Target, .Target
    End With
End Sub


