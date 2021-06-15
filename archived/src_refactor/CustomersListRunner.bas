Attribute VB_Name = "CustomersListRunner"
Attribute VB_Description = "Display Customers as List in ListView Control"
'@Folder("MaintainCustomers.CustomersList")
'@ModuleDescription "Display Customers as List in ListView Control"
Option Explicit
'Option Private Module

'@Description "Main entry to Customers Manager UI Framework"
Public Sub RunCustomersList()
Attribute RunCustomersList.VB_Description = "Main entry to Customers Manager UI Framework"

    Dim workSheetName As String
    workSheetName = "CustomersList"

    Dim View As IView
    Set View = CustomersListForm.Create()
    
    View.MinimumHeight 330
    View.MinimumWidth 540

    If View.ShowDialog(workSheetName) Then
        Debug.Print "Manage Accounts Loaded."
    Else
        Debug.Print "Manage Accounts cancelled."
    End If

End Sub
