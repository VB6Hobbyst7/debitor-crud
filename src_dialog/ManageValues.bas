Attribute VB_Name = "ManageValues"
Attribute VB_Description = "Main entry in Application"
'@Folder("ValidateUserInput.Runner")
Option Explicit

'@ModuleDescription "Main entry in Application"
'Option Private Module

Public Sub Main()

    Dim workSheetName As String
    workSheetName = "TablesValues"

    Dim View As IView
    Set View = TablesView.Create()

    View.MinimumHeight 330
    View.MinimumWidth 540

    If View.ShowDialog(workSheetName) Then
        Debug.Print "Manage Accounts Loaded."
    Else
        Debug.Print "Manage Accounts cancelled."
    End If

End Sub
