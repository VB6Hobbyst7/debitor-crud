VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigureView 
   Caption         =   "[Titel]"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   OleObjectBlob   =   "ConfigureView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigureView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of an abstract View, which displays for example a configuration View to add new Customers."

'@Folder AccountsManager.View
'@ModuleDescription "Implementation of an abstract View, which displays for example a configuration View to add new Customers."

Implements IView
Implements ICancellable

Option Explicit

Private Type TView
    Context As IAppContext
    ViewModel As ConfigureViewModel
    LayoutBindings As list
    IsCancelled As Boolean
End Type

Private This As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ConfigureViewModel) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, ConfigureView, TypeName(Me)
    GuardClauses.GuardNullReference Context, TypeName(Me)
    GuardClauses.GuardNullReference ViewModel, TypeName(Me)

    Dim result As ConfigureView
    Set result = New ConfigureView

    Set result.Context = Context
    Set result.ViewModel = ViewModel
    Set Create = result
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ConfigureView
End Property

'@Description "Gets/sets the MaintainCustomers application context."
Public Property Get Context() As IAppContext
Attribute Context.VB_Description = "Gets/sets the MaintainCustomers application context."
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As IAppContext)
    GuardClauses.GuardDefaultInstance Me, ConfigureView, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Context, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Context = RHS
End Property

'@Description "Gets/sets the ViewModelManager to use as a context for property and command bindings."
Public Property Get ViewModel() As ConfigureViewModel
Attribute ViewModel.VB_Description = "Gets/sets the ViewModelManager to use as a context for property and command bindings."
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal model As ConfigureViewModel)
    GuardClauses.GuardDefaultInstance Me, ConfigureView, TypeName(Me)
    GuardClauses.GuardNullReference model
    Set This.ViewModel = model
End Property

Private Sub InitializeBindings()
    If ViewModel Is Nothing Then Exit Sub
    
    BindViewModelCommands
    BindViewModelProperties
    
    This.Context.Bindings.Apply ViewModel
End Sub

Private Sub BindViewModelProperties()
    With Context.Bindings
'        .BindPropertyPath ViewModel, "LanguageIDUI", Me.LanguageIDUI
'        .BindPropertyPath ViewModel, "Instructions", Me.Instructions, Mode:=OneTimeBinding, UpdateTrigger:=OnPropertyChanged
'
'        .BindPropertyPath ViewModel, "AccountGroup", Me.AccountGroup, "List", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged
'        .BindPropertyPath ViewModel, "AccountGroupValue", Me.AccountGroup, "Value", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged, _
'            Validator:=New RequiredValueValidator, _
'            ValidationAdorner:=ValidationErrorAdorner.Create( _
'            Target:=Me.AccountGroup, _
'            TargetFormatter:=ValidationErrorFormatter.WithErrorBorderColor.WithErrorBackgroundColor)
'
'        .BindPropertyPath ViewModel, "SalesOrganization", Me.SalesOrganization, "List", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged
'        .BindPropertyPath ViewModel, "SalesOrganizationValue", Me.SalesOrganization, "Value", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged, _
'            Validator:=New RequiredValueValidator, _
'            ValidationAdorner:=ValidationErrorAdorner.Create( _
'            Target:=Me.SalesOrganization, _
'            TargetFormatter:=ValidationErrorFormatter.WithErrorBorderColor.WithErrorBackgroundColor)
'
'        .BindPropertyPath ViewModel, "Channel", Me.Channel, "List", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged
'        .BindPropertyPath ViewModel, "ChannelValue", Me.Channel, "Value", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged, _
'            Validator:=New RequiredValueValidator, _
'            ValidationAdorner:=ValidationErrorAdorner.Create( _
'            Target:=Me.Channel, _
'            TargetFormatter:=ValidationErrorFormatter.WithErrorBorderColor.WithErrorBackgroundColor)
        
'        .BindPropertyPath ViewModel, "NewCustomer", Me.OptionButtonNewCustomer, Mode:=OneTimeBinding
'        .BindPropertyPath ViewModel, "Reactivate", Me.OptionButtonReactivate, Mode:=OneTimeBinding
        
'        .BindPropertyPath ViewModel, "AccountID", Me.AccountID
'        .BindPropertyPath ViewModel, "UserCreated", Me.UserCreated
'        .BindPropertyPath ViewModel, "TimeStampCreated", Me.TimeStampCreated
    End With
End Sub

Private Sub BindViewModelCommands()
    With Context.Commands
'        .BindCommand ViewModel, Me.ApplyButton, ApplyConfigCommand.Create(Me, This.Context.Validation)
'        .BindCommand ViewModel, Me.CancelButton, CancelCommand.Create(Me)
        
    End With
End Sub

Private Sub Localize(ByVal title As String)

    Dim Source As String: Source = "CaptionSource"

    Me.Caption = title
    FrameAccountGroup.Caption = GetResourceString("ConfigureFormFrameAccountGroup", 2, Source)
    FrameSalesOrganization.Caption = GetResourceString("ConfigureFormFrameSalesOrganization", 2, Source)
    FrameChannel.Caption = GetResourceString("ConfigureFormFrameChannel", 2, Source)
    FrameOptions.Caption = GetResourceString("ConfigureFormFrameOptions", 2, Source)
'    OptionButtonNewCustomer.Caption = GetResourceString("ConfigureFormOptionButtonNewCustomer", 2)
'    OptionButtonReactivate.Caption = GetResourceString("ConfigureFormOptionButtonReactivate", 2)
'    FrameID.Caption = GetResourceString("ConfigureFormFrameID", 2)
'    FrameUser.Caption = GetResourceString("ConfigureFormFrameUser", 2)
'    ApplyButton.Caption = GetResourceString("ConfigureFormApplyButton", 2, source)
'
'    CancelButton.Caption = GetResourceString("CancelButton", 2, source)
'
'    ApplyButton.ControlTipText = GetResourceString("ConfigureFormApplyButton", 3, source)
'    CancelButton.ControlTipText = GetResourceString("CancelButton", 3, source)
    
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    'Not implemented
End Sub

Private Function IView_ShowDialog() As Boolean
    
    Localize GetResourceString("ConfigureFormTitel", 2, "CaptionSource")
    InitializeBindings

    'EnableMouseScroll Me
'
'    This.ViewModel.AccountGroupValue = " "
'    This.ViewModel.SalesOrganizationValue = " "
'    This.ViewModel.ChannelValue = " "
        
    Me.Show vbModal
    IView_ShowDialog = Not This.IsCancelled

End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
