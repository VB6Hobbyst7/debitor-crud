Attribute VB_Name = "ValidationManagerTests"
'@Folder Tests
'@TestModule
Option Explicit
Option Private Module

Private Type TState
    
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    Validator As IValueValidator
    
    ConcreteSUT As ValidationManager
    NotifyValidationErrorSUT As INotifyValidationError
    HandleValidationErrorSUT As IHandleValidationError
    
    BindingManager As IBindingManager
    BindingManagerStub As ITestStub
    
    CommandManager As ICommandManager
    CommandManagerStub As ITestStub
    
    BindingSource As TestBindingObject
    BindingSourceStub As ITestStub
    BindingTarget As TestBindingObject
    BindingTargetStub As ITestStub
    
    SourcePropertyPath As String
    TargetPropertyPath As String
    command As TestCommand
    
End Type

Private Test As TState

#Const LateBind = LateBindTests
#If LateBind Then
Private Assert As Object
#Else
Private Assert As Rubberduck.AssertClass
#End If

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
    Set Test.ConcreteSUT = ValidationManager.Create(New TestNotifierFactory)
    Set Test.NotifyValidationErrorSUT = Test.ConcreteSUT
    Set Test.HandleValidationErrorSUT = Test.ConcreteSUT
    Set Test.BindingSource = TestBindingObject.Create(Test.ConcreteSUT)
    Set Test.BindingSourceStub = Test.BindingSource
    Set Test.BindingTarget = TestBindingObject.Create(Test.ConcreteSUT)
    Set Test.BindingTargetStub = Test.BindingTarget
    Set Test.command = New TestCommand
    Set Test.CommandManager = New TestCommandManager
    Set Test.CommandManagerStub = Test.CommandManager
    Set Test.Validator = New TestValueValidator
    Dim Manager As TestBindingManager
    Set Manager = New TestBindingManager
    Set Test.BindingManager = Manager
    Set Test.BindingManagerStub = Test.BindingManager
    Test.SourcePropertyPath = "TestStringProperty"
    Test.TargetPropertyPath = "TestStringProperty"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
    Set Test.NotifyValidationErrorSUT = Nothing
    Set Test.HandleValidationErrorSUT = Nothing
    Set Test.BindingSource = Nothing
    Set Test.BindingTarget = Nothing
    Set Test.command = Nothing
    Set Test.Validator = Nothing
    Set Test.BindingManager = Nothing
    Set Test.BindingManagerStub = Nothing
    Test.SourcePropertyPath = vbNullString
    Test.TargetPropertyPath = vbNullString
    Test.ExpectedErrNumber = 0
    Test.ExpectedErrorCaught = False
    Test.ExpectedErrSource = vbNullString
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

