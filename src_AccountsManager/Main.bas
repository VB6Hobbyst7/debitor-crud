Attribute VB_Name = "Main"
Attribute VB_Description = "Main entry in Application"
'@Folder("AccountsManager")
Option Explicit

'@ModuleDescription "Main entry in Application"
Option Private Module

Public Sub AppStart()

Dim LogToConsole As Boolean: LogToConsole = False
Dim LogToBuffer As Boolean: LogToBuffer = False
Dim LogToFile As Boolean: LogToFile = True

#If Debugging Then
    LogToConsole = True
    LogToBuffer = True
#End If
    
    Logging.setLoggigParams Logging.lgALL, LogToConsole, LogToBuffer, LogToFile
    Logging.logINFO ("***Starting Logger***")

    Dim Source As String: Source = "Manager"
    
    Dim XLProperties As ExcelProperties
    Set XLProperties = New ExcelProperties
    
    XLProperties.Save
    
    Dim ViewModel As ManagerViewModel
    Set ViewModel = ManagerViewModel.Create()
    
    Dim CaptionSource As String: CaptionSource = "CaptionSource"
    
    With ViewModel
        .Titel = GetResourceString("Manager.Caption", 2, CaptionSource)
        .Instructions = GetResourceString("Manager.Instructions", 2, CaptionSource)
        .FilterCaption = GetResourceString("Manager.FrameFilter", 2, CaptionSource)
        .AddButtonCaption = GetResourceString("Manager.AddButton", 2, CaptionSource)
        .EditButtonCaption = GetResourceString("Manager.EditButton", 2, CaptionSource)
        .QuitButtonCaption = GetResourceString("Manager.QuitButton", 2, CaptionSource)
    
        .FilterControlTipText = GetResourceString("Manager.FrameFilter", 3, CaptionSource)
        .AddControlTipText = GetResourceString("Manager.AddButton", 3, CaptionSource)
        .EditControlTipText = GetResourceString("Manager.EditButton", 3, CaptionSource)
        .QuitControlTipText = GetResourceString("Manager.QuitButton", 3, CaptionSource)
    
        .Filter = GetRowSourceList("RPAStatus", "RPAStatusDescription", "RowSource")
    
        .SourceTable = Source
        .FilterValue = "OPEN"
        .ModelHeight = 330
        .ModelWidth = 540
    End With
    
    Dim app As AppContext
    Set app = AppContext.Create(DebugOutput:=True)

    Dim View As IView
    Set View = ManagerView.Create(app, ViewModel)

    If View.ShowDialog Then
        Logging.logINFO ViewModel.SourceTable
    Else
        Logging.logINFO "Manager cancelled."
    End If
    
    Disposable.TryDispose app
    XLProperties.Restore
    
End Sub
