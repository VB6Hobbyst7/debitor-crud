Attribute VB_Name = "Main"
Attribute VB_Description = "Main entry in Application"
'@Folder("AccountsManager")
Option Explicit

'@ModuleDescription "Main entry in Application"
Option Private Module

Public Sub AppStart()

    Dim Source As String: Source = "Manager"
    
    Dim ViewModel As ManagerViewModel
    Set ViewModel = ManagerViewModel.Create()
    
    Dim CaptionSource As String: CaptionSource = "CaptionSource"
    
    ViewModel.Titel = GetResourceString("Manager.Caption", 2, CaptionSource)
    ViewModel.Instructions = GetResourceString("Manager.Instructions", 2, CaptionSource)
    ViewModel.FilterCaption = GetResourceString("Manager.FrameFilter", 2, CaptionSource)
    ViewModel.AddButtonCaption = GetResourceString("Manager.AddButton", 2, CaptionSource)
    ViewModel.EditButtonCaption = GetResourceString("Manager.EditButton", 2, CaptionSource)
    ViewModel.QuitButtonCaption = GetResourceString("Manager.QuitButton", 2, CaptionSource)
    
    ViewModel.FilterControlTipText = GetResourceString("Manager.FrameFilter", 3, CaptionSource)
    ViewModel.AddControlTipText = GetResourceString("Manager.AddButton", 3, CaptionSource)
    ViewModel.EditControlTipText = GetResourceString("Manager.EditButton", 3, CaptionSource)
    ViewModel.QuitControlTipText = GetResourceString("Manager.QuitButton", 3, CaptionSource)
    
    ViewModel.Filter = GetRowSourceList("RPAStatus", "RPAStatusDescription", "RowSource")
    
    ViewModel.SourceTable = Source
    ViewModel.FilterValue = "OPEN"
    ViewModel.ModelHeight = 330
    ViewModel.ModelWidth = 540
    
    Dim app As AppContext
    Set app = AppContext.Create(DebugOutput:=True)

    Dim View As IView
    Set View = ManagerView.Create(app, ViewModel)

    If View.ShowDialog Then
        Debug.Print ViewModel.SourceTable, ViewModel.Filter
    Else
        Debug.Print "Manager cancelled."
    End If
    
    Disposable.TryDispose app

End Sub

Private Function NewId(ByVal book As Workbook, ByVal Source As String) As Long
    Dim sourceSheet As Worksheet
    Set sourceSheet = book.Worksheets(Source)
    With sourceSheet
        NewId = Application.WorksheetFunction.Max(.Range("A1", .Cells(.Rows.Count, 1).End(xlUp))) + 1
    End With
End Function
