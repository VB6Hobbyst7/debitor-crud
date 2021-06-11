VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ValidateUserInput.Presenter")

Option Explicit
Private WithEvents View As LastPageView
Attribute View.VB_VarHelpID = -1

'Private Property Get Dialog() As IDialogView
'    Set Dialog = View
'End Property
'
'Public Sub Show(Optional ByVal field1 As String, Optional ByVal field2 As String)
'
'    Set View = New SingleValuesView
'
'    Dim ViewModel As TablesModel
'    Set ViewModel = New TablesModel
'    ViewModel.field1 = field1
'    ViewModel.field2 = field2
'
'    If Dialog.ShowDialog(ViewModel) Then ApplyChanges ViewModel
'    Set View = Nothing
'
'End Sub
'
'Private Sub ApplyChanges(ByVal ViewModel As TablesModel)
''    Sheet1.Range("A1").value = ViewModel.field1
''    Sheet1.Range("A2").value = ViewModel.field2
'End Sub
'
'Private Sub view_ApplyChanges(ByVal ViewModel As TablesModel)
'    ApplyChanges ViewModel
'End Sub
