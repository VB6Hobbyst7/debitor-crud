VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Table2View 
   Caption         =   "[Titel]"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "Table2View.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Table2View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("AccountsManager.View")
Option Explicit

'Implements IView
Implements ICancellable

Private Type TView
    Width As Double
    Height As Double
    LayoutBindings As list
    IsCancelled As Boolean
End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create() As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, Table2View, TypeName(Me)
    
    Dim result As Table2View
    Set result = New Table2View
    
    Set Create = result
    
End Function

Private Sub BindViewModelLayouts()

'    Set this.layoutBindings = New List
'
'    Dim BackgroundFrameLayout As New ControlLayout
'    BackgroundFrameLayout.Bind Me, ManagerFrame, AnchorAll
'
'    Dim InstructionsLayout As New ControlLayout
'    InstructionsLayout.Bind Me, Instructions, LeftAnchor + RightAnchor
'
'    Dim Logo As New ControlLayout
'    Logo.Bind Me, LogoImage, TopAnchor + RightAnchor
'
'    Dim ListViewLayout As New ControlLayout
'    ListViewLayout.Bind Me, TablesValuesList, AnchorAll
'
'    Dim QuitButtonLayout As New ControlLayout
'    QuitButtonLayout.Bind Me, QuitButton, BottomAnchor + RightAnchor
'
'    Dim AddButtonLayout As New ControlLayout
'    AddButtonLayout.Bind Me, AddButton, BottomAnchor + RightAnchor
'
'    Dim EditButtonLayout As New ControlLayout
'    EditButtonLayout.Bind Me, EditButton, BottomAnchor + RightAnchor
'
'    this.layoutBindings.Add BackgroundFrameLayout, _
'                            InstructionsLayout, _
'                            Logo, _
'                            ListViewLayout, _
'                            AddButtonLayout, _
'                            EditButtonLayout, _
'                            QuitButtonLayout

End Sub

Private Sub InitializeLayouts()
    BindViewModelLayouts
End Sub

Private Sub Localize(ByVal title As String)

'    Me.Caption = title
'    Me.Instructions = GetResourceString("TablesValuesInstructions", 2)
'
'    AddButton.Caption = GetResourceString("TablesValuesAddButton", 2)
'    EditButton.Caption = GetResourceString("TablesValuesEditButton", 2)
'    QuitButton.Caption = GetResourceString("TablesValuesQuitButton", 2)
'
'    AddButton.ControlTipText = GetResourceString("TablesValuesAddButton", 3)
'    EditButton.ControlTipText = GetResourceString("TablesValuesEditButton", 3)
'    QuitButton.ControlTipText = GetResourceString("TablesValuesQuitButton", 3)
    
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
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

Private Sub IView_Show(ByVal ViewModel As Object)
    'Not implemented
End Sub

Private Function IView_ShowDialog() As Boolean

'    Localize GetResourceString("TablesValuesTitel", 2)
'    InitializeLayouts
'    InitializeAccountsList workSheetName

'    Me.width = GetSystemMetrics32(0) * PointsPerPixel * 0.6 'UF Width in Resolution * DPI * 60%
'    Me.height = GetSystemMetrics32(1) * PointsPerPixel * 0.4 'UF Height in Resolution * DPI * 40%

    MakeFormResizable Me, True
    ShowMinimizeButton Me, False
    ShowMaximizeButton Me, False
    
    Me.Show vbModal
    IView_ShowDialog = Not this.IsCancelled
    
End Function

Private Sub IView_MinimumWidth(ByVal Width As Single)
    this.Width = Width
End Sub

Private Sub IView_MinimumHeight(ByVal Height As Single)
    this.Height = Height
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
    
    If Me.Width < this.Width Then Me.Width = this.Width
    If Me.Height < this.Height Then Me.Height = this.Height
    
    Dim Layout As ControlLayout
    For Each Layout In this.LayoutBindings
        Layout.Resize Me
    Next

    Application.ScreenUpdating = True
    On Error GoTo 0
    
End Sub

Private Sub CancelButton_Click()
    '@Ignore FunctionReturnValueDiscarded
    OnCancel
End Sub

Private Sub BackButton_Click()
    OnCancel
'    AddValues.Add
End Sub

Private Sub NextButton_Click()
    OnCancel
    Dim workSheetName As String
    workSheetName = "Table3Values"

    Dim View As IView
    Set View = Table3View.Create()

'    View.MinimumHeight 346
'    View.MinimumWidth 318

    If View.ShowDialog() Then
        Debug.Print "Table3 Values Loaded."
    Else
        Debug.Print "Table3 Values cancelled."
    End If
    
End Sub
