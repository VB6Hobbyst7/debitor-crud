VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigureViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "ViewModel implementation to View Customer Account Initial Data."
'@Folder AccountsManager.ViewModel
'@ModuleDescription "ViewModel implementation to View Customer Account Initial Data."
Option Explicit
'@PredeclaredId

Implements INotifyPropertyChanged

Private Type TViewModel
    Notifier As INotifyPropertyChanged
    
    LanguageIDUI As Long
    Instructions As String
    
    AccountGroup As Variant
    AccountGroupValue As String
    
    SalesOrganization As Variant
    SalesOrganizationValue As String
    
    Channel As Variant
    ChannelValue As String
    
    NewCustomer As Boolean
    Reactivate As Boolean
    
    ApplyCommand As ICommand
    
    AccountID As Long
    UserCreated As String
    TimeStampCreated As Date
End Type

Private this As TViewModel

Public Function Create() As ConfigureViewModel
    GuardClauses.GuardNonDefaultInstance Me, ConfigureViewModel, TypeName(Me)
    
    Dim result As ConfigureViewModel
    Set result = New ConfigureViewModel
    Set Create = result
    
End Function

Public Property Get LanguageIDUI() As String
    LanguageIDUI = this.LanguageIDUI
End Property

Public Property Let LanguageIDUI(ByVal value As String)
    this.LanguageIDUI = value
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

Public Property Get AccountGroup() As Variant
    AccountGroup = this.AccountGroup
End Property

'Public Property Let AccountGroup(ByVal varArray As Variant)
'    This.AccountGroup = varArray
'    OnPropertyChanged "AccountGroup"
'End Property

Public Property Get AccountGroupValue() As String
    AccountGroupValue = this.AccountGroupValue
End Property

Public Property Let AccountGroupValue(ByVal value As String)
    If this.AccountGroupValue <> value Then
        this.AccountGroupValue = value
        OnPropertyChanged "AccountGroupValue"
    End If
End Property

Public Property Get SalesOrganization() As Variant
    SalesOrganization = this.SalesOrganization
End Property

'Public Property Let SalesOrganization(ByVal varArray As Variant)
'    This.SalesOrganization = varArray
'    OnPropertyChanged "SalesOrganization"
'End Property

Public Property Get SalesOrganizationValue() As String
    SalesOrganizationValue = this.SalesOrganizationValue
End Property

Public Property Let SalesOrganizationValue(ByVal value As String)
    If this.SalesOrganizationValue <> value Then
        this.SalesOrganizationValue = value
        OnPropertyChanged "SalesOrganizationValue"
        ChannelValue = vbNullString
        NewCustomer = False
        Reactivate = False
        FilterChannelsToOrganisations value
    End If
End Property

Public Property Get Channel() As Variant
    Channel = this.Channel
End Property

Public Property Let Channel(ByVal varArray As Variant)
    this.Channel = varArray
    OnPropertyChanged "Channel"
End Property

Public Property Get ChannelValue() As String
    ChannelValue = this.ChannelValue
End Property

Public Property Let ChannelValue(ByVal value As String)
    If this.ChannelValue <> value Then
        this.ChannelValue = value
        OnPropertyChanged "ChannelValue"
    End If
End Property

Public Property Get NewCustomer() As Boolean
    NewCustomer = this.NewCustomer
End Property

Public Property Let NewCustomer(ByVal opt As Boolean)
    If this.NewCustomer <> opt Then
        this.NewCustomer = opt
        OnPropertyChanged "NewCustomer"
    End If
End Property

Public Property Get Reactivate() As Boolean
    Reactivate = this.Reactivate
End Property

Public Property Let Reactivate(ByVal opt As Boolean)
    If this.Reactivate <> opt Then
        this.Reactivate = opt
        OnPropertyChanged "Reactivate"
    End If
End Property

Public Property Get AccountID() As Long
    AccountID = this.AccountID
End Property

Public Property Let AccountID(ByVal Id As Long)
    If this.AccountID <> Id Then
        this.AccountID = Id
        OnPropertyChanged "AccountID"
    End If
End Property

Public Property Get UserCreated() As String
    UserCreated = this.UserCreated
End Property

Public Property Let UserCreated(ByVal user As String)
    this.UserCreated = user
End Property

Public Property Get TimeStampCreated() As Date
    TimeStampCreated = this.TimeStampCreated
End Property

Public Property Let TimeStampCreated(ByVal timeStamp As Date)
    this.TimeStampCreated = timeStamp
End Property

Public Property Get ApplyCommand() As ICommand
    Set ApplyCommand = this.ApplyCommand
End Property

Public Property Set ApplyCommand(ByVal setObject As ICommand)
    Set this.ApplyCommand = setObject
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set this.Notifier = New PropertyChangeNotifierBase

    this.AccountGroup = GetRowSourceList("AccountGroup", "AccountGroupDescription", "RowSource")
    
    this.SalesOrganization = GetRowSourceList("SalesOrganization", "SalesOrganizationDescription", "RowSource")
    this.Channel = Array()
        
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    this.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Notifier.RegisterHandler Handler
End Sub

Private Sub FilterChannelsToOrganisations(ByVal salesOrg As String)
    GuardClauses.GuardEmptyString salesOrg
    
    If salesOrg Like "0361*" Then
        Channel = GetRowSourceList("Channel0361", "Channel0361Description", "RowSource")
        
    ElseIf salesOrg Like "2561*" Then
        Channel = GetRowSourceList("Channel2561", "Channel2561Description", "RowSource")
        
    ElseIf salesOrg Like "2961*" Then
        Channel = GetRowSourceList("Channel2961", "Channel2961Description", "RowSource")
        
    ElseIf salesOrg Like "3661*" Then
        Channel = GetRowSourceList("Channel3661", "Channel3661Description", "RowSource")
        
    End If
End Sub
