Attribute VB_Name = "BindingPathTests"
'@Folder Tests
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
    
    ConcreteSUT As BindingPath
    AbstractSUT As IBindingPath
    
    BindingContext As TestBindingObject
    BindingSource As TestBindingObject
    PropertyName As String
    path As String
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
    Dim Context As TestBindingObject
    Set Context = New TestBindingObject
    
    Set Context.TestBindingObjectProperty = New TestBindingObject
    
    Test.path = "TestBindingObjectProperty.TestStringProperty"
    Test.PropertyName = "TestStringProperty"
    Set Test.BindingSource = Context.TestBindingObjectProperty
    
    Set Test.BindingContext = Context
    Set Test.ConcreteSUT = BindingPath.Create(Test.BindingContext, Test.path)
    Set Test.AbstractSUT = Test.ConcreteSUT
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
    Set Test.AbstractSUT = Nothing
    Set Test.BindingSource = Nothing
    Set Test.BindingContext = Nothing
    Test.path = vbNullString
    Test.PropertyName = vbNullString
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

'@TestMethod("GuardClauses")
Private Sub Create_GuardsNullBindingContext()
    Test.ExpectedErrNumber = GuardClauseErrors.ObjectCannotBeNothing
    On Error Resume Next
        BindingPath.Create Nothing, Test.path
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Create_GuardsEmptyPath()
    Test.ExpectedErrNumber = GuardClauseErrors.StringCannotBeEmpty
    On Error Resume Next
        BindingPath.Create Test.BindingContext, vbNullString
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Create_GuardsNonDefaultInstance()
    Test.ExpectedErrNumber = GuardClauseErrors.InvalidFromNonDefaultInstance
    On Error Resume Next
        With New BindingPath
            .Create Test.BindingContext, Test.path
            ExpectError
        End With
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Context_GuardsDefaultInstance()
    Test.ExpectedErrNumber = GuardClauseErrors.InvalidFromDefaultInstance
    On Error Resume Next
        Set BindingPath.Context = Test.BindingContext
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Context_GuardsDoubleInitialization()
    Test.ExpectedErrNumber = GuardClauseErrors.ObjectAlreadyInitialized
    On Error Resume Next
        Set Test.ConcreteSUT.Context = New TestBindingObject
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Context_GuardsNullReference()
    Test.ExpectedErrNumber = GuardClauseErrors.ObjectCannotBeNothing
    On Error Resume Next
        Set Test.ConcreteSUT.Context = Nothing
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Path_GuardsDefaultInstance()
    Test.ExpectedErrNumber = GuardClauseErrors.InvalidFromDefaultInstance
    On Error Resume Next
        BindingPath.path = Test.path
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Path_GuardsDoubleInitialization()
    Test.ExpectedErrNumber = GuardClauseErrors.ObjectAlreadyInitialized
    On Error Resume Next
        Test.ConcreteSUT.path = Test.path
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("GuardClauses")
Private Sub Path_GuardsEmptyString()
    Test.ExpectedErrNumber = GuardClauseErrors.StringCannotBeEmpty
    On Error Resume Next
        Test.ConcreteSUT.path = vbNullString
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("Bindings")
Private Sub Resolve_SetsBindingSource()
    With New BindingPath
        .path = Test.path
        Set .Context = Test.BindingContext
        
        If Not .object Is Nothing Then Assert.Inconclusive "Object reference is unexpectedly set."
        .Resolve
        
        Assert.AreSame Test.BindingSource, .object
    End With
End Sub

'@TestMethod("Bindings")
Private Sub Resolve_SetsBindingPropertyName()
    With New BindingPath
        .path = Test.path
        Set .Context = Test.BindingContext
        
        If .PropertyName <> vbNullString Then Assert.Inconclusive "PropertyName is unexpectedly non-empty."
        .Resolve
        
        Assert.AreEqual Test.PropertyName, .PropertyName
    End With
End Sub

'@TestMethod("Bindings")
Private Sub Create_ResolvesPropertyName()
    Dim SUT As BindingPath
    Set SUT = BindingPath.Create(Test.BindingContext, Test.path)
    Assert.IsFalse SUT.PropertyName = vbNullString
End Sub

'@TestMethod("Bindings")
Private Sub Create_ResolvesBindingSource()
    Dim SUT As BindingPath
    Set SUT = BindingPath.Create(Test.BindingContext, Test.path)
    Assert.IsNotNothing SUT.object
End Sub
