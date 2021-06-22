VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An example ViewModel implementation for some dialog."
'@Folder AccountsManager.ViewModel
'@ModuleDescription "An example ViewModel implementation for some dialog."
'@PredeclaredId
'@Exposed
Implements INotifyPropertyChanged
Option Explicit

Private Type TViewModel
    Notifier As INotifyPropertyChanged
    
    SomeCommand As ICommand
    SomeFilePath As String
    SomeAmount As Currency
    SomeDate As Date
    SomeProperty As String
    SomeOption As Boolean
    SomeOtherOption As Boolean
    SomeItems As Variant
    SelectedItemText As String
    SelectedItemIndex As Long
    BooleanProperty As Boolean
    ByteProperty As Byte
    CurrencyProperty As Currency
    DateProperty As Date
    DoubleProperty As Double
    LongProperty As Long
    StringProperty As String
End Type

Private this As TViewModel

Public Function Create() As ExampleViewModel
    GuardClauses.GuardNonDefaultInstance Me, ExampleViewModel, TypeName(Me)
    
    Dim result As ExampleViewModel
    Set result = New ExampleViewModel
    
    Set Create = result
End Function

Public Property Get BooleanProperty() As Boolean
    BooleanProperty = this.BooleanProperty
End Property

Public Property Let BooleanProperty(ByVal RHS As Boolean)
    If this.BooleanProperty <> RHS Then
        this.BooleanProperty = RHS
        OnPropertyChanged "BooleanProperty"
    End If
End Property

Public Property Get ByteProperty() As Byte
    ByteProperty = this.ByteProperty
End Property

Public Property Let ByteProperty(ByVal RHS As Byte)
    If this.ByteProperty <> RHS Then
        this.ByteProperty = RHS
        OnPropertyChanged "ByteProperty"
    End If
End Property

Public Property Get CurrencyProperty() As Currency
    CurrencyProperty = this.CurrencyProperty
End Property

Public Property Let CurrencyProperty(ByVal RHS As Currency)
    If this.CurrencyProperty <> RHS Then
        this.CurrencyProperty = RHS
        OnPropertyChanged "CurrencyProperty"
    End If
End Property

Public Property Get DateProperty() As Date
    DateProperty = this.DateProperty
End Property

Public Property Let DateProperty(ByVal RHS As Date)
    If this.DateProperty <> RHS Then
        this.DateProperty = RHS
        OnPropertyChanged "DateProperty"
    End If
End Property

Public Property Get DoubleProperty() As Double
    DoubleProperty = this.DoubleProperty
End Property

Public Property Let DoubleProperty(ByVal RHS As Double)
    If this.DoubleProperty <> RHS Then
        this.DoubleProperty = RHS
        OnPropertyChanged "DoubleProperty"
    End If
End Property

Public Property Get LongProperty() As Long
    LongProperty = this.LongProperty
End Property

Public Property Let LongProperty(ByVal RHS As Long)
    If this.LongProperty <> RHS Then
        this.LongProperty = RHS
        OnPropertyChanged "LongProperty"
    End If
End Property

Public Property Get StringProperty() As String
    StringProperty = this.StringProperty
End Property

Public Property Let StringProperty(ByVal RHS As String)
    If this.StringProperty <> RHS Then
        this.StringProperty = RHS
        OnPropertyChanged "StringProperty"
    End If
End Property

Public Property Get SomeFilePath() As String
    SomeFilePath = this.SomeFilePath
End Property

Public Property Let SomeFilePath(ByVal RHS As String)
    If this.SomeFilePath <> RHS Then
        this.SomeFilePath = RHS
        OnPropertyChanged "SomeFilePath"
    End If
End Property

Public Property Get SomeItems() As Variant
    SomeItems = this.SomeItems
End Property

Public Property Get SelectedItemText() As String
    SelectedItemText = this.SelectedItemText
End Property

Public Property Let SelectedItemText(ByVal RHS As String)
    If this.SelectedItemText <> RHS Then
        this.SelectedItemText = RHS
        OnPropertyChanged "SelectedItemText"
    End If
End Property

Public Property Get SelectedItemIndex() As Long
    SelectedItemIndex = this.SelectedItemIndex
End Property

Public Property Let SelectedItemIndex(ByVal RHS As Long)
    If this.SelectedItemIndex <> RHS Then
        this.SelectedItemIndex = RHS
        OnPropertyChanged "SelectedItemIndex"
    End If
End Property

Public Property Get SomeAmount() As Currency
    SomeAmount = this.SomeAmount
End Property

Public Property Let SomeAmount(ByVal RHS As Currency)
    If this.SomeAmount <> RHS Then
        this.SomeAmount = RHS
        OnPropertyChanged "SomeAmount"
    End If
End Property

Public Property Get SomeDate() As Date
    SomeDate = this.SomeDate
End Property

Public Property Let SomeDate(ByVal RHS As Date)
    If this.SomeDate <> RHS Then
        this.SomeDate = RHS
        OnPropertyChanged "SomeDate"
    End If
End Property

Public Property Get SomeProperty() As String
    SomeProperty = this.SomeProperty
End Property

Public Property Let SomeProperty(ByVal RHS As String)
    If this.SomeProperty <> RHS Then
        this.SomeProperty = RHS
        OnPropertyChanged "SomeProperty"
    End If
End Property

Public Property Get SomeOption() As Boolean
    SomeOption = this.SomeOption
End Property

Public Property Let SomeOption(ByVal RHS As Boolean)
    If this.SomeOption <> RHS Then
        this.SomeOption = RHS
        OnPropertyChanged "SomeOption"
    End If
End Property

Public Property Get SomeOtherOption() As Boolean
    SomeOtherOption = this.SomeOtherOption
End Property

Public Property Let SomeOtherOption(ByVal RHS As Boolean)
    If this.SomeOtherOption <> RHS Then
        this.SomeOtherOption = RHS
        OnPropertyChanged "SomeOtherOption"
    End If
End Property

Public Property Get SomeCommand() As ICommand
    Set SomeCommand = this.SomeCommand
End Property

Public Property Set SomeCommand(ByVal RHS As ICommand)
    Set this.SomeCommand = RHS
End Property

Public Property Get SomeOptionName() As String
    SomeOptionName = "Auto"
End Property

Public Property Get SomeOtherOptionName() As String
    SomeOtherOptionName = "Manual/Browse"
End Property

Public Property Get Instructions() As String
    Instructions = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set this.Notifier = New PropertyChangeNotifierBase
    ReDim this.SomeItems(0 To 4)
    Dim i As Long
    For i = LBound(this.SomeItems) To UBound(this.SomeItems)
        this.SomeItems(i) = "Selection " & i
    Next
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Notifier.RegisterHandler Handler
End Sub


