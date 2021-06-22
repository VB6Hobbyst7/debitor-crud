VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManagerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "ViewModel implementation for Manager dialog."
'@Folder("AccountsManager.ViewModel")
'@ModuleDescription "ViewModel implementation for Manager dialog."
'@PredeclaredId
'@Exposed
Implements INotifyPropertyChanged
Option Explicit

 Private Type TValuesModel
    Notifier As INotifyPropertyChanged
    
    Titel As String
    Instructions As String
    
    FilterCaption As String
    AddButtonCaption As String
    EditButtonCaption As String
    QuitButtonCaption As String
    
    FilterControlTipText As String
    AddControlTipText As String
    EditControlTipText As String
    QuitControlTipText As String
    
    Filter As Variant
    FilterValue As String
    ModelWidth As Double
    ModelHeight As Double
    SourceTable As String
    
    AddCommand As ICommand
    EditCommand As ICommand
    QuitCommand As ICommand

End Type

Private this As TValuesModel

Public Property Get Titel() As String
    Titel = this.Titel
End Property

Public Property Let Titel(ByVal text As String)
    this.Titel = text
End Property

Public Property Get Instructions() As String
    Instructions = this.Instructions
End Property

Public Property Let Instructions(ByVal text As String)
    If this.Instructions <> text Then
        this.Instructions = text
        OnPropertyChanged "Instructions"
    End If
End Property

Public Property Get FilterCaption() As String
    FilterCaption = this.FilterCaption
End Property

Public Property Let FilterCaption(ByVal text As String)
    this.FilterCaption = text
End Property

Public Property Get AddButtonCaption() As String
    AddButtonCaption = this.AddButtonCaption
End Property

Public Property Let AddButtonCaption(ByVal text As String)
    this.AddButtonCaption = text
End Property

Public Property Get EditButtonCaption() As String
    EditButtonCaption = this.EditButtonCaption
End Property

Public Property Let EditButtonCaption(ByVal text As String)
    this.EditButtonCaption = text
End Property

Public Property Get QuitButtonCaption() As String
    QuitButtonCaption = this.QuitButtonCaption
End Property

Public Property Let QuitButtonCaption(ByVal text As String)
    this.QuitButtonCaption = text
End Property

Public Property Get FilterControlTipText() As String
    FilterControlTipText = this.FilterControlTipText
End Property

Public Property Let FilterControlTipText(ByVal text As String)
    this.FilterControlTipText = text
End Property

Public Property Get AddControlTipText() As String
    AddControlTipText = this.AddControlTipText
End Property

Public Property Let AddControlTipText(ByVal text As String)
    this.AddControlTipText = text
End Property

Public Property Get EditControlTipText() As String
    EditControlTipText = this.EditControlTipText
End Property

Public Property Let EditControlTipText(ByVal text As String)
    this.EditControlTipText = text
End Property

Public Property Get QuitControlTipText() As String
    QuitControlTipText = this.QuitControlTipText
End Property

Public Property Let QuitControlTipText(ByVal text As String)
    this.QuitControlTipText = text
End Property

Public Property Get SourceTable() As String
    SourceTable = this.SourceTable
End Property

Public Property Let SourceTable(ByVal table As String)
    If this.SourceTable <> table Then
        this.SourceTable = table
        OnPropertyChanged "SourceTable"
    End If
End Property

Public Property Get ModelWidth() As Double
    ModelWidth = this.ModelWidth
End Property

Public Property Let ModelWidth(ByVal Value As Double)
    this.ModelWidth = Value
End Property

Public Property Get ModelHeight() As Double
    ModelHeight = this.ModelHeight
End Property

Public Property Let ModelHeight(ByVal Value As Double)
    this.ModelHeight = Value
End Property

Public Property Get Filter() As Variant
    Filter = this.Filter
End Property

Public Property Let Filter(ByVal list As Variant)
    this.Filter = list
    OnPropertyChanged "Filter"
End Property

Public Property Get FilterValue() As String
    FilterValue = this.FilterValue
End Property

Public Property Let FilterValue(ByVal Value As String)
    If this.FilterValue <> Value Then
        this.FilterValue = Value
        OnPropertyChanged "FilterValue"
    End If
End Property

Public Property Get AddCommand() As ICommand
    Set AddCommand = this.AddCommand
End Property

Public Property Set AddCommand(ByVal command As ICommand)
    Set this.AddCommand = command
End Property

Public Property Get EditCommand() As ICommand
    Set EditCommand = this.EditCommand
End Property

Public Property Set EditCommand(ByVal command As ICommand)
    Set this.EditCommand = command
End Property

Public Property Get QuitCommand() As ICommand
    Set QuitCommand = this.QuitCommand
End Property

Public Property Set QuitCommand(ByVal command As ICommand)
    Set this.QuitCommand = command
End Property

Public Function Create() As ManagerViewModel
    GuardClauses.GuardNonDefaultInstance Me, ManagerViewModel, TypeName(Me)
    
    Dim result As ManagerViewModel
    Set result = New ManagerViewModel
    
    Set Create = result
    
End Function

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set this.Notifier = New PropertyChangeNotifierBase
    
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Notifier.RegisterHandler Handler
End Sub
