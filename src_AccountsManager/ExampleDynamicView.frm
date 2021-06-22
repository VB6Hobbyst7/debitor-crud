VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleDynamicView 
   Caption         =   "ExampleDynamicView"
   ClientHeight    =   5475
   ClientLeft      =   -450
   ClientTop       =   -1515
   ClientWidth     =   4815
   OleObjectBlob   =   "ExampleDynamicView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleDynamicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder AccountsManager.View
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As IAppContext
    ViewModel As ExampleViewModel
    IsCancelled As Boolean
End Type

Private this As TState

'@Description "Creates a new instance of this form."
Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ExampleViewModel, ViewDims As TViewDims) As IView
Attribute Create.VB_Description = "Creates a new instance of this form."
    Dim result As ExampleDynamicView
    Set result = New ExampleDynamicView
    Set result.Context = Context
    Set result.ViewModel = ViewModel
    With result
        .Height = ViewDims.Height
        .Width = ViewDims.Width
    End With
    Set Create = result
End Function

Public Property Get Context() As IAppContext
    Set Context = this.Context
End Property

Public Property Set Context(ByVal RHS As IAppContext)
    Set this.Context = RHS
End Property

Public Property Get ViewModel() As Object
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set this.ViewModel = RHS
End Property

Public Sub SizeView(Height As Long, Width As Long)
    With Me
        .Height = Height
        .Width = Width
    End With
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()
    
    Dim Layout As IContainerLayout
    Set Layout = ContainerLayout.Create(Me.Controls, TopToBottom)
    
    With DynamicControls.Create(this.Context, Layout)
        
        With .LabelFor("All controls on this form are created at run-time.")
            .Font.Bold = True
        End With
        
        .LabelFor BindingPath.Create(this.ViewModel, "Instructions")
        
        'VF: refactor free string to some enum PropertyName ("StringProperty", "CurrencyProperty") throughout (?) [when I frame a question mark in parentheses is not really a question but a rhetorical question, meaning I am pretty sure of the correct answer]
        .TextBoxFor BindingPath.Create(this.ViewModel, "StringProperty"), _
                    Validator:=New RequiredStringValidator, _
                    TitleSource:="Some String:"
                    
        .TextBoxFor BindingPath.Create(this.ViewModel, "CurrencyProperty"), _
                    FormatString:="{0:C2}", _
                    Validator:=New DecimalKeyValidator, _
                    TitleSource:="Some Amount:"
        
        'ToDo: 'VF: needs validation .CanExecute(This.Context) before .Show
        '(as textbox1 has focus and is empty and when moving to this close button, tb1 is validated and OnClick is disabled leaving the user out in the rain)
        .CommandButtonFor AcceptCommand.Create(Me, this.Context.Validation), this.ViewModel, "Close"
        
    End With
    
    this.Context.Bindings.Apply this.ViewModel
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = this.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeView
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    Me.Show vbModal
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
