VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Holds the validation errors state for a ViewModel."
'@Folder AccountsManager.Infrastructure.Validation
'@ModuleDescription "Holds the validation errors state for a ViewModel."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationManager

Private Type TState
    ErrorAdornerFactory As IDynamicAdornerFactory
    ValidationErrors As Dictionary
    IsValid As Boolean
End Type

Private this As TState

Public Function Create(Optional ByVal ErrorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    Dim result As ValidationManager
    Set result = New ValidationManager
    InitAdornerFactory result, ErrorAdornerFactory
    Set Create = result
End Function

Private Sub InitAdornerFactory(ByVal Manager As ValidationManager, ByVal AdornerFactory As IDynamicAdornerFactory)
    Dim Factory As IDynamicAdornerFactory
    
    If AdornerFactory Is Nothing Then
        Dim DefaultFactory As IDynamicAdornerFactory
        Set DefaultFactory = New DefaultErrorAdornerFactory
        Set Factory = DefaultFactory
    Else
        'use injected instance (could be a test stub):
        Set Factory = AdornerFactory
        
    End If
    
    Set Manager.ErrorAdornerFactory = Factory
End Sub

Public Property Get ErrorAdornerFactory() As IDynamicAdornerFactory
    Set ErrorAdornerFactory = this.ErrorAdornerFactory
End Property

Friend Property Set ErrorAdornerFactory(ByVal RHS As IDynamicAdornerFactory)
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    GuardClauses.GuardNullReference RHS, TypeName(ValidationManager)
    GuardClauses.GuardDoubleInitialization this.ErrorAdornerFactory, TypeName(Me)
    Set this.ErrorAdornerFactory = RHS
End Property

Public Property Get IsValid(ByVal BindingContext As Object, Optional ByVal PropertyName As String) As Boolean
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    
    Dim result As Boolean
    result = True
    
    Dim Key As Variant
    For Each Key In this.ValidationErrors.Keys
        
        Dim Errors As ValidationErrors
        Set Errors = this.ValidationErrors.item(Key)
        
        If BindingContext Is Errors.BindingContext Or Errors.BindingContext Is Nothing Then
            
            If PropertyName = vbNullString Then
                result = Errors.Count = 0
            Else
                Dim Error As IValidationError
                For Each Error In Errors
                    result = (Error.Binding.Source.PropertyName <> PropertyName)
                    If Not result Then Exit For
                Next
            End If
            
        End If
        
        If Not result Then Exit For
    Next
    
    IsValid = result
    
End Property

Public Property Get ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not this.ValidationErrors.Exists(ContextKey) Then Exit Property
    
    Dim Errors As ValidationErrors
    Set Errors = this.ValidationErrors(ContextKey)
    
    'ValidationErrors.Item getter returns Nothing if there's no error for the given property:
    Set ValidationError = Errors.item(BindingPath.PropertyName)
End Property

Private Sub ClearValidationErrors(ByVal BindingPath As IBindingPath)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not this.ValidationErrors.Exists(ContextKey) Then Exit Sub
    
    Dim Errors As ValidationErrors
    Set Errors = this.ValidationErrors(ContextKey)
    Errors.Remove BindingPath.PropertyName
End Sub

Private Sub OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(ValidationError.Binding.Source.Context))
    
    Dim Errors As ValidationErrors
    If this.ValidationErrors.Exists(ContextKey) Then
        Set Errors = this.ValidationErrors.item(ContextKey)
    Else
        Set Errors = ValidationErrors.Create(Context)
    End If
    
    Errors.add ValidationError
    Set this.ValidationErrors.item(ContextKey) = Errors
End Sub

Private Sub Class_Initialize()
    Set this.ValidationErrors = New Dictionary
    this.IsValid = True
End Sub

Private Sub Class_Terminate()
    Set this.ValidationErrors = Nothing
End Sub

Private Property Get IValidationManager_AdornerFactory() As IDynamicAdornerFactory
    Set IValidationManager_AdornerFactory = this.ErrorAdornerFactory
End Property

Private Sub IValidationManager_ClearValidationError(ByVal BindingPath As IBindingPath)
    ClearValidationErrors BindingPath
End Sub

Private Property Get IValidationManager_IsValid(ByVal Context As Object, Optional ByVal PropertyName As String) As Boolean
    IValidationManager_IsValid = IsValid(Context, PropertyName)
End Property

Private Sub IValidationManager_OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    OnValidationError Context, ValidationError
End Sub

Private Property Get IValidationManager_ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Set IValidationManager_ValidationError = ValidationError(BindingPath)
End Property
