VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomersListForm 
   Caption         =   "[Titel]"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "CustomersListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomersListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implementation of an abstract View, which displays for example all entries in DBCustomers"


'@Folder("MaintainCustomers.CustomersList")
'@ModuleDescription "Implementation of an abstract View, which displays for example all entries in DBCustomers"

Implements IView
Implements ICancellable

Option Explicit

Private Type TView
    width As Double
    height As Double
    layoutBindings As MaintainCustomers.List
    isCancelled As Boolean
End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create() As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, CustomersListForm, TypeName(Me)
    
    Dim result As CustomersListForm
    Set result = New CustomersListForm
    
    Set Create = result
    
End Function

Private Sub BindViewModelLayouts()

    Set this.layoutBindings = New MaintainCustomers.List

    Dim BackgroundFrameLayout As New ControlLayout
    BackgroundFrameLayout.Bind Me, ManagerFrame, AnchorAll

    Dim InstructionsLayout As New ControlLayout
    InstructionsLayout.Bind Me, Instructions, LeftAnchor + RightAnchor
    
    Dim ListViewLayout As New ControlLayout
    ListViewLayout.Bind Me, CustomersListView, AnchorAll
    
    Dim CancelButtonLayout As New ControlLayout
    CancelButtonLayout.Bind Me, CancelButton, BottomAnchor + RightAnchor
    
    Dim AddButtonLayout As New ControlLayout
    AddButtonLayout.Bind Me, AddButton, BottomAnchor + RightAnchor
        
    Dim ViewButtonLayout As New ControlLayout
    ViewButtonLayout.Bind Me, ViewButton, BottomAnchor + RightAnchor
    
    this.layoutBindings.Add BackgroundFrameLayout, _
                            InstructionsLayout, _
                            ListViewLayout, _
                            AddButtonLayout, _
                            ViewButtonLayout, _
                            CancelButtonLayout

End Sub

Private Sub InitializeLayouts()
    BindViewModelLayouts
End Sub

Private Sub Localize(ByVal title As String)

    Me.Caption = title
    Me.Instructions = GetResourceString("CustomersListInstructions", 2)
    
    AddButton.Caption = GetResourceString("CustomersListAddButton", 2)
    ViewButton.Caption = GetResourceString("CustomersListViewButton", 2)
    CancelButton.Caption = GetResourceString("CustomersListCancelButton", 2)

    AddButton.ControlTipText = GetResourceString("CustomersListAddButton", 3)
    ViewButton.ControlTipText = GetResourceString("CustomersListViewButton", 3)
    CancelButton.ControlTipText = GetResourceString("CustomersListCancelButton", 3)
    
End Sub

Private Sub InitializeAccountsList(ByVal workSheetName As String)
    
'    ThisWorkbook.Sheets("CustomersList").PivotTables("CustomersList").RefreshTable
    InitializeListView Me.CustomersListView, _
                       ThisWorkbook.FullName, _
                       workSheetName, "*", vbNullString, vbNullString, "ID DESC"
End Sub

Private Sub CustomersListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnSort Me.CustomersListView, ColumnHeader
End Sub

Private Sub OnCancel()
    this.isCancelled = True
    Me.Hide
End Sub

Private Sub CustomersListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ViewButton.Enabled = Item.selected
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

Private Function IView_ShowDialog(Optional ByVal workSheetName As String) As Boolean

    Localize GetResourceString("CustomersListFormTitel", 2)
    InitializeLayouts
    InitializeAccountsList workSheetName

    Me.width = GetSystemMetrics32(0) * PointsPerPixel * 0.6 'UF Width in Resolution * DPI * 60%
    Me.height = GetSystemMetrics32(1) * PointsPerPixel * 0.4 'UF Height in Resolution * DPI * 40%

    MakeFormResizable Me, True
    ShowMinimizeButton Me, False
    ShowMaximizeButton Me, False

    Me.Show vbModal
    IView_ShowDialog = Not this.isCancelled
    
End Function

Private Sub IView_MinimumWidth(ByVal width As Single)
    this.width = width
End Sub

Private Sub IView_MinimumHeight(ByVal height As Single)
    this.height = height
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub UserForm_Resize()

    On Error Resume Next
    Application.ScreenUpdating = False
    
    If Me.width < this.width Then Me.width = this.width
    If Me.height < this.height Then Me.height = this.height
    
    Dim Layout As MaintainCustomers.ControlLayout
    For Each Layout In this.layoutBindings
        Layout.Resize Me
    Next

    Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub

Private Sub CancelButton_Click()
    '@Ignore FunctionReturnValueDiscarded
    OnCancel
End Sub
