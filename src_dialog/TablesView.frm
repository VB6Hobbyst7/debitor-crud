VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TablesView 
   Caption         =   "[Titel]"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "TablesView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TablesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of an abstract View, which displays for example all entries in given WorkSheet"
'@Folder("ValidateUserInput.View")
'@ModuleDescription "Implementation of an abstract View, which displays for example all entries in given WorkSheet"

Implements IView
Implements ICancellable

Option Explicit

Private Type TView
    ViewModel As ValuesViewModel
    layoutBindings As List
    isCancelled As Boolean

End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal ViewModel As ValuesViewModel) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, TablesView, TypeName(Me)
    GuardClauses.GuardNullReference ViewModel, TypeName(Me)
    
    Dim result As TablesView
    Set result = New TablesView
    
    Set result.ViewModel = ViewModel
    Set Create = result
    
End Function

'@Description "Gets/sets the ViewModelManager to use as a context for property and command bindings."
Public Property Get ViewModel() As ValuesViewModel
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal model As ValuesViewModel)
    GuardClauses.GuardDefaultInstance Me, TablesView, TypeName(Me)
    GuardClauses.GuardNullReference model
    Set this.ViewModel = model
End Property

Private Sub BindViewModelLayouts()

    Set this.layoutBindings = New List

    Dim BackgroundFrameLayout As New ControlLayout
    BackgroundFrameLayout.Bind Me, ManagerFrame, AnchorAll

    Dim InstructionsLayout As New ControlLayout
    InstructionsLayout.Bind Me, Instructions, LeftAnchor + RightAnchor
    
    Dim Logo As New ControlLayout
    Logo.Bind Me, LogoImage, TopAnchor + RightAnchor
    
    Dim ListViewLayout As New ControlLayout
    ListViewLayout.Bind Me, TablesValuesList, AnchorAll
    
    Dim QuitButtonLayout As New ControlLayout
    QuitButtonLayout.Bind Me, QuitButton, BottomAnchor + RightAnchor
    
    Dim AddButtonLayout As New ControlLayout
    AddButtonLayout.Bind Me, AddButton, BottomAnchor + RightAnchor
        
    Dim EditButtonLayout As New ControlLayout
    EditButtonLayout.Bind Me, EditButton, BottomAnchor + RightAnchor
    
    this.layoutBindings.Add BackgroundFrameLayout, _
                            InstructionsLayout, _
                            Logo, _
                            ListViewLayout, _
                            AddButtonLayout, _
                            EditButtonLayout, _
                            QuitButtonLayout

End Sub

Private Sub InitializeLayouts()
    BindViewModelLayouts
End Sub

Private Sub InitializeDisplayLanguage()

    Me.Caption = GetResourceString("TablesView.Caption", 2)
    Me.Instructions = GetResourceString("TablesView.Instructions", 2)
    
    AddButton.Caption = GetResourceString("TablesView.AddButton", 2)
    EditButton.Caption = GetResourceString("TablesView.EditButton", 2)
    QuitButton.Caption = GetResourceString("TablesView.QuitButton", 2)

    AddButton.ControlTipText = GetResourceString("TablesView.AddButton", 3)
    EditButton.ControlTipText = GetResourceString("TablesView.EditButton", 3)
    QuitButton.ControlTipText = GetResourceString("TablesView.QuitButton", 3)
    
End Sub

Private Sub InitializeAccountsList(ByVal workSheetName As String)
    
'    ThisWorkbook.Sheets("CustomersList").PivotTables("CustomersList").RefreshTable
    InitializeListView Me.TablesValuesList, _
                       ThisWorkbook.FullName, _
                       workSheetName, "*", vbNullString, vbNullString, "ID DESC"
End Sub

Private Sub TablesValuesList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnSort Me.TablesValuesList, ColumnHeader
End Sub

Private Sub OnCancel()
    this.isCancelled = True
    Me.Hide
    Application.Visible = True
End Sub

Private Sub TablesValuesList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EditButton.Enabled = Item.selected
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = this.isCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show(ByVal ViewModel As Object)
    'Not implemented
End Sub

Private Function IView_ShowDialog() As Boolean

    InitializeDisplayLanguage
    InitializeLayouts
    InitializeAccountsList this.ViewModel.DataSourceTable

    Me.width = GetSystemMetrics32(0) * PointsPerPixel * 0.6 'UF Width in Resolution * DPI * 60%
    Me.height = GetSystemMetrics32(1) * PointsPerPixel * 0.4 'UF Height in Resolution * DPI * 40%

    MakeFormResizable Me, True
    ShowMinimizeButton Me, False
    ShowMaximizeButton Me, False
    
    Me.AddButton.SetFocus
    
    Application.Visible = False
    
    Me.Show vbModal
    IView_ShowDialog = Not this.isCancelled
    
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub UserForm_Resize()

    On Error Resume Next
    Application.ScreenUpdating = False
    
    If Me.width < ViewModel.ModelWidth Then Me.width = ViewModel.ModelWidth
    If Me.height < ViewModel.ModelHeight Then Me.height = ViewModel.ModelHeight
    
    Dim Layout As ControlLayout
    For Each Layout In this.layoutBindings
        Layout.Resize Me
    Next

    Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub

Private Sub QuitButton_Click()
    '@Ignore FunctionReturnValueDiscarded
    OnCancel
End Sub

Private Sub AddButton_Click()
    OnCancel
    
End Sub

Private Sub EditButton_Click()
    OnCancel
    
End Sub
