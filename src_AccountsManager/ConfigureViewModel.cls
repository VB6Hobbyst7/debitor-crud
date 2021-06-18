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

Private This As TViewModel

Public Function Create() As ConfigureViewModel
    GuardClauses.GuardNonDefaultInstance Me, ConfigureViewModel, TypeName(Me)
    
    Dim result As ConfigureViewModel
    Set result = New ConfigureViewModel
    Set Create = result
    
End Function

Public Property Get LanguageIDUI() As String
    LanguageIDUI = This.LanguageIDUI
End Property

Public Property Let LanguageIDUI(ByVal value As String)
    This.LanguageIDUI = value
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

Public Property Get AccountGroup() As Variant
    AccountGroup = This.AccountGroup
End Property

'Public Property Let AccountGroup(ByVal varArray As Variant)
'    This.AccountGroup = varArray
'    OnPropertyChanged "AccountGroup"
'End Property

Public Property Get AccountGroupValue() As String
    AccountGroupValue = This.AccountGroupValue
End Property

Public Property Let AccountGroupValue(ByVal value As String)
    If This.AccountGroupValue <> value Then
        This.AccountGroupValue = value
        OnPropertyChanged "AccountGroupValue"
    End If
End Property

Public Property Get SalesOrganization() As Variant
    SalesOrganization = This.SalesOrganization
End Property

'Public Property Let SalesOrganization(ByVal varArray As Variant)
'    This.SalesOrganization = varArray
'    OnPropertyChanged "SalesOrganization"
'End Property

Public Property Get SalesOrganizationValue() As String
    SalesOrganizationValue = This.SalesOrganizationValue
End Property

Public Property Let SalesOrganizationValue(ByVal value As String)
    If This.SalesOrganizationValue <> value Then
        This.SalesOrganizationValue = value
        OnPropertyChanged "SalesOrganizationValue"
        ChannelValue = vbNullString
        NewCustomer = False
        Reactivate = False
        FilterChannelsToOrganisations value
    End If
End Property

Public Property Get Channel() As Variant
    Channel = This.Channel
End Property

Public Property Let Channel(ByVal varArray As Variant)
    This.Channel = varArray
    OnPropertyChanged "Channel"
End Property

Public Property Get ChannelValue() As String
    ChannelValue = This.ChannelValue
End Property

Public Property Let ChannelValue(ByVal value As String)
    If This.ChannelValue <> value Then
        This.ChannelValue = value
        OnPropertyChanged "ChannelValue"
    End If
End Property

Public Property Get NewCustomer() As Boolean
    NewCustomer = This.NewCustomer
End Property

Public Property Let NewCustomer(ByVal opt As Boolean)
    If This.NewCustomer <> opt Then
        This.NewCustomer = opt
        OnPropertyChanged "NewCustomer"
    End If
End Property

Public Property Get Reactivate() As Boolean
    Reactivate = This.Reactivate
End Property

Public Property Let Reactivate(ByVal opt As Boolean)
    If This.Reactivate <> opt Then
        This.Reactivate = opt
        OnPropertyChanged "Reactivate"
    End If
End Property

Public Property Get AccountID() As Long
    AccountID = This.AccountID
End Property

Public Property Let AccountID(ByVal Id As Long)
    If This.AccountID <> Id Then
        This.AccountID = Id
        OnPropertyChanged "AccountID"
    End If
End Property

Public Property Get UserCreated() As String
    UserCreated = This.UserCreated
End Property

Public Property Let UserCreated(ByVal user As String)
    This.UserCreated = user
End Property

Public Property Get TimeStampCreated() As Date
    TimeStampCreated = This.TimeStampCreated
End Property

Public Property Let TimeStampCreated(ByVal timeStamp As Date)
    This.TimeStampCreated = timeStamp
End Property

Public Property Get ApplyCommand() As ICommand
    Set ApplyCommand = This.ApplyCommand
End Property

Public Property Set ApplyCommand(ByVal setObject As ICommand)
    Set This.ApplyCommand = setObject
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifierBase

    This.AccountGroup = GetRowSourceList("AccountGroup", "AccountGroupDescription", "RowSource")
    
    This.SalesOrganization = GetRowSourceList("SalesOrganization", "SalesOrganizationDescription", "RowSource")
    This.Channel = Array()
        
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
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
