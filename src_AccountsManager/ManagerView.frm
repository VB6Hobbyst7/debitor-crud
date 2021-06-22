VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManagerView 
   Caption         =   "[Titel]"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "ManagerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManagerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of an abstract View, which displays for example all entries in given WorkSheet"
'@Folder("AccountsManager.View")
'@ModuleDescription "Implementation of an abstract View, which displays for example all entries in given WorkSheet"

Implements IView
Implements ICancellable

Option Explicit

Private Type TView
    Context As IAppContext
    ViewModel As ManagerViewModel
    LayoutBindings As list
    IsCancelled As Boolean

End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ManagerViewModel) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, ManagerView, TypeName(Me)
    GuardClauses.GuardNullReference ViewModel, TypeName(Me)
    GuardClauses.GuardNullReference Context, TypeName(Me)
    
    Dim result As ManagerView
    Set result = New ManagerView
    
    Set result.Context = Context
    Set result.ViewModel = ViewModel
    Set Create = result

End Function

'@Description "Gets/sets the ViewModelManager to use as a context for property and command bindings."
Public Property Get ViewModel() As ManagerViewModel
Attribute ViewModel.VB_Description = "Gets/sets the ViewModelManager to use as a context for property and command bindings."
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal model As ManagerViewModel)
    GuardClauses.GuardDefaultInstance Me, ManagerView, TypeName(Me)
    GuardClauses.GuardNullReference model
    Set this.ViewModel = model
End Property

'@Description "Gets/sets the AccountsManager application context."
Public Property Get Context() As IAppContext
Attribute Context.VB_Description = "Gets/sets the AccountsManager application context."
    Set Context = this.Context
End Property

Public Property Set Context(ByVal app As IAppContext)
    GuardClauses.GuardDefaultInstance Me, ManagerView, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Context, TypeName(Me)
    GuardClauses.GuardNullReference app
    Set this.Context = app
End Property

Private Sub BindViewModelLayouts()

    Set this.LayoutBindings = New list

    Dim BackgroundFrameLayout As New ControlLayout
    BackgroundFrameLayout.Bind Me, ManagerFrame, AnchorAll

    Dim InstructionsLayout As New ControlLayout
    InstructionsLayout.Bind Me, Instructions, LeftAnchor + RightAnchor
    
    Dim Logo As New ControlLayout
    Logo.Bind Me, LogoImage, TopAnchor + RightAnchor
    
    Dim ListViewLayout As New ControlLayout
    ListViewLayout.Bind Me, TablesValuesList, AnchorAll
    
    Dim FrameFilterLayout As New ControlLayout
    FrameFilterLayout.Bind Me, FrameFilter, LeftAnchor + BottomAnchor
    
    Dim QuitButtonLayout As New ControlLayout
    QuitButtonLayout.Bind Me, QuitButton, BottomAnchor + RightAnchor
    
    Dim AddButtonLayout As New ControlLayout
    AddButtonLayout.Bind Me, AddButton, BottomAnchor + RightAnchor
        
    Dim EditButtonLayout As New ControlLayout
    EditButtonLayout.Bind Me, EditButton, BottomAnchor + RightAnchor
    
    this.LayoutBindings.add BackgroundFrameLayout, _
                            InstructionsLayout, _
                            Logo, _
                            ListViewLayout, _
                            FrameFilterLayout, _
                            AddButtonLayout, _
                            EditButtonLayout, _
                            QuitButtonLayout

End Sub

Private Sub BindViewModelProperties()
    With Context.Bindings
        .BindPropertyPath ViewModel, "Instructions", Me.Instructions, Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged
        .BindPropertyPath ViewModel, "FilterCaption", Me.FrameFilter, Mode:=OneWayBinding
        .BindPropertyPath ViewModel, "AddButtonCaption", Me.AddButton, Mode:=OneWayBinding
        .BindPropertyPath ViewModel, "EditButtonCaption", Me.EditButton, Mode:=OneWayBinding
        .BindPropertyPath ViewModel, "QuitButtonCaption", Me.QuitButton, Mode:=OneWayBinding
        
        .BindPropertyPath ViewModel, "Filter", Me.Filter, "List", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged
        .BindPropertyPath ViewModel, "FilterValue", Me.Filter, "Value", Mode:=TwoWayBinding, UpdateTrigger:=OnPropertyChanged

    End With
End Sub

Private Sub BindViewModelCommands()
    With Context.Commands
        .BindCommand ViewModel, Me.AddButton, AddCommand.Create(Me, this.Context)
'        .BindCommand ViewModel, Me.CancelButton, CancelCommand.Create(Me)

    End With
End Sub

Private Sub InitializeLayouts()
    BindViewModelLayouts
End Sub

Private Sub InitializeAccountsList(ByVal workSheetName As String)
    
    'Application.ThisWorkbook.Sheets("").PivotTables("").RefreshTable
    
    InitializeListView Me.TablesValuesList, _
                       ThisWorkbook.FullName, _
                       workSheetName, "*", vbNullString, vbNullString, "ID DESC"
End Sub

Private Sub InitializeBindings()
    BindViewModelProperties
    BindViewModelCommands
    this.Context.Bindings.Apply ViewModel
End Sub

Private Sub InitializeView()

    If ViewModel Is Nothing Then Exit Sub
    
    InitializeBindings
    InitializeLayouts
    InitializeAccountsList this.ViewModel.SourceTable

    Me.Width = GetSystemMetrics32(0) * PointsPerPixel * 0.6 'UF Width in Resolution * DPI * 60%
    Me.Height = GetSystemMetrics32(1) * PointsPerPixel * 0.4 'UF Height in Resolution * DPI * 40%

    MakeFormResizable Me, True
    ShowMinimizeButton Me, False
    ShowMaximizeButton Me, False
    
    Me.Caption = ViewModel.Titel
    Me.FrameFilter.ControlTipText = ViewModel.FilterControlTipText
    Me.AddButton.ControlTipText = ViewModel.AddControlTipText
    Me.EditButton.ControlTipText = ViewModel.EditControlTipText
    Me.QuitButton.ControlTipText = ViewModel.QuitControlTipText

    Me.AddButton.SetFocus
    
End Sub

Private Sub TablesValuesList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnSort Me.TablesValuesList, ColumnHeader
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Sub TablesValuesList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EditButton.Enabled = Item.selected
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

Private Sub UserForm_Resize()

    On Error Resume Next
    Application.ScreenUpdating = False
    
    If Me.Width < ViewModel.ModelWidth Then Me.Width = ViewModel.ModelWidth
    If Me.Height < ViewModel.ModelHeight Then Me.Height = ViewModel.ModelHeight
    
    Dim Layout As ControlLayout
    For Each Layout In this.LayoutBindings
        Layout.Resize Me
    Next

    Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub

Private Sub QuitButton_Click()
    '@Ignore FunctionReturnValueDiscarded
    OnCancel
    'To Do: Implement Workbook close and CheckIn procedure
End Sub

Private Sub AddButton_Click()
    Me.Hide
    
End Sub

Private Sub EditButton_Click()
    Me.Hide
    
End Sub
