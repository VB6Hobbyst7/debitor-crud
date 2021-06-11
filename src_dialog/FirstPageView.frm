VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FirstPageView 
   Caption         =   "[Titel]"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "FirstPageView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FirstPageView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("ValidateUserInput.View")
Option Explicit

Public Event ApplyChanges(ByVal ViewModel As TablesModel)

'Private Type TView
'    isCancelled As Boolean
'    model As TablesModel
'End Type
'Private this As TView
'
'Implements IDialogView
'
'Private Sub AcceptButton_Click()
'    Me.Hide
'End Sub
'
'Private Sub ApplyButton_Click()
'    RaiseEvent ApplyChanges(this.model)
'End Sub
'
'Private Sub CancelButton_Click()
'    OnCancel
'End Sub
'
'Private Sub Field1Box_Change()
''    This.model.field1 = Field1Box.value
'End Sub
'
'Private Sub Field2Box_Change()
''    This.model.field2 = Field2Box.value
'End Sub
'
'Private Sub OnCancel()
'    this.isCancelled = True
'    Me.Hide
'End Sub
'
'Private Function IDialogView_ShowDialog(ByVal ViewModel As Object) As Boolean
'    Set this.model = ViewModel
'    Me.Show vbModal
'    IDialogView_ShowDialog = Not this.isCancelled
'End Function
'
'Private Sub UserForm_Activate()
''    Field1Box.value = this.Model.field1
''    Field2Box.value = this.Model.field2
'End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = VbQueryClose.vbFormControlMenu Then
'        Cancel = True
'        OnCancel
'    End If
'End Sub
