VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("AccountsManager.Common.Commands")
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICommand

Private Type TState
    View As IView
    Context As IAppContext
End Type

Private this As TState

'@Description "Creates a new instance of this command."
Public Function Create(ByVal View As IView, ByVal Context As IAppContext) As ICommand
Attribute Create.VB_Description = "Creates a new instance of this command."
    Dim result As AddCommand
    Set result = New AddCommand
    Set result.View = View
    Set result.Context = Context
    Set Create = result
End Function

Public Property Get View() As IView
    Set View = this.View
End Property

Public Property Set View(ByVal object As IView)
    GuardClauses.GuardDoubleInitialization this.View, TypeName(Me)
    Set this.View = object
End Property

Public Property Get Context() As IAppContext
    Set Context = this.Context
End Property

Public Property Set Context(ByVal object As IAppContext)
    GuardClauses.GuardDoubleInitialization this.Context, TypeName(Me)
    Set this.Context = object
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Add New."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    this.View.Hide
    InitiateConfigure
End Sub

Private Sub InitiateConfigure()
    Dim ViewModel As ConfigureViewModel
    Set ViewModel = ConfigureViewModel.Create()

    With ViewModel
        .LanguageIDUI = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'        .Instructions = GetResourceString("ConfigureFormInstructions", 2)
'        .AccountGroup = GetRowSourceList("AccountGroup")
'        .SalesOrganization = GetRowSourceList("SalesOrganization")
        .Channel = Array()
'        .AccountGroupValue = "validate"
'        .SalesOrganizationValue = "validate"
'        .ChannelValue = "validate"
'        .UserCreated = Application.UserName
'        .TimeStampCreated = VBA.DateTime.Now
'        .AccountID = NewId
'        .NewCustomer = False
'        .Reactivate = False
    End With

    Dim View As IView
    Set View = ConfigureView.Create(this.Context, ViewModel)
    
    If View.ShowDialog() Then
        Debug.Print "Configure Loaded"
    Else
        Debug.Print "Configure was cancelled."
    End If
    
    
End Sub

Private Function NewId(ByVal book As Workbook, ByVal Source As String) As Long
    Dim sourceSheet As Worksheet
    Set sourceSheet = book.Worksheets(Source)
    With sourceSheet
        NewId = Application.WorksheetFunction.Max(.Range("A1", .Cells(.Rows.Count, 1).End(xlUp))) + 1
    End With
End Function
