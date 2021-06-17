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

End Type

Private This As TValuesModel

Public Property Get Titel() As String
    Titel = This.Titel
End Property

Public Property Let Titel(ByVal text As String)
    This.Titel = text
End Property

Public Property Get Instructions() As String
    Instructions = This.Instructions
End Property

Public Property Let Instructions(ByVal text As String)
    If This.Instructions <> text Then
        This.Instructions = text
        OnPropertyChanged "Instructions"
    End If
End Property

Public Property Get FilterCaption() As String
    FilterCaption = This.FilterCaption
End Property

Public Property Let FilterCaption(ByVal text As String)
    This.FilterCaption = text
End Property

Public Property Get AddButtonCaption() As String
    AddButtonCaption = This.AddButtonCaption
End Property

Public Property Let AddButtonCaption(ByVal text As String)
    This.AddButtonCaption = text
End Property

Public Property Get EditButtonCaption() As String
    EditButtonCaption = This.EditButtonCaption
End Property

Public Property Let EditButtonCaption(ByVal text As String)
    This.EditButtonCaption = text
End Property

Public Property Get QuitButtonCaption() As String
    QuitButtonCaption = This.QuitButtonCaption
End Property

Public Property Let QuitButtonCaption(ByVal text As String)
    This.QuitButtonCaption = text
End Property

Public Property Get FilterControlTipText() As String
    FilterControlTipText = This.FilterControlTipText
End Property

Public Property Let FilterControlTipText(ByVal text As String)
    This.FilterControlTipText = text
End Property

Public Property Get AddControlTipText() As String
    AddControlTipText = This.AddControlTipText
End Property

Public Property Let AddControlTipText(ByVal text As String)
    This.AddControlTipText = text
End Property

Public Property Get EditControlTipText() As String
    EditControlTipText = This.EditControlTipText
End Property

Public Property Let EditControlTipText(ByVal text As String)
    This.EditControlTipText = text
End Property

Public Property Get QuitControlTipText() As String
    QuitControlTipText = This.QuitControlTipText
End Property

Public Property Let QuitControlTipText(ByVal text As String)
    This.QuitControlTipText = text
End Property

Public Property Get SourceTable() As String
    SourceTable = This.SourceTable
End Property

Public Property Let SourceTable(ByVal table As String)
    If This.SourceTable <> table Then
        This.SourceTable = table
        OnPropertyChanged "SourceTable"
    End If
End Property

Public Property Get ModelWidth() As Double
    ModelWidth = This.ModelWidth
End Property

Public Property Let ModelWidth(ByVal value As Double)
    This.ModelWidth = value
End Property

Public Property Get ModelHeight() As Double
    ModelHeight = This.ModelHeight
End Property

Public Property Let ModelHeight(ByVal value As Double)
    This.ModelHeight = value
End Property

Public Property Get Filter() As Variant
    Filter = This.Filter
End Property

Public Property Let Filter(ByVal list As Variant)
    This.Filter = list
    OnPropertyChanged "Filter"
End Property

Public Property Get FilterValue() As String
    FilterValue = This.FilterValue
End Property

Public Property Let FilterValue(ByVal value As String)
    If This.FilterValue <> value Then
        This.FilterValue = value
        OnPropertyChanged "FilterValue"
    End If
End Property

Public Function Create() As ManagerViewModel
    GuardClauses.GuardNonDefaultInstance Me, ManagerViewModel, TypeName(Me)
    
    Dim result As ManagerViewModel
    Set result = New ManagerViewModel
    
    Set Create = result
    
End Function

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifierBase
    
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub
