VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: clsCheckBox ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim UForm As Object, MPages As MSForms.Control, ctrl As MSForms.Control, X As Integer
Private WithEvents chbInForm As MSForms.CheckBox
Attribute chbInForm.VB_VarHelpID = -1
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Control ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Set Control(chbContainer As MSForms.CheckBox)

    Set chbInForm = chbContainer
End Property
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'chbInForm_Change ///////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub chbInForm_Change()

    Set UForm = chbInForm.Parent
    With chbInForm
        If .Tag Like "Must*" And .value <> False Then
            '.BackColor = RGB(242, 242, 242) 'Light Gray 'Light Gray
            .BackColor = RGB(255, 255, 255) 'White
        End If
        If .Name Like "*Reaktivieren*" Then
            If .value = True Then
                UForm.tbx_Debitor_SAPNr.Tag = "MustFill"
                UForm.tbx_Debitor_SAPNr.Visible = True
            Else
                UForm.tbx_Debitor_SAPNr.value = ""
                UForm.tbx_Debitor_SAPNr.Tag = ""
                'UForm.tbx_Debitor_SAPNr.BackColor = RGB(242, 242, 242) 'Button Face 'White on Win10
                UForm.tbx_Debitor_SAPNr.BackColor = RGB(255, 255, 255) 'White
            End If
        End If
        If .Name Like "*_E_*" And .value = True Then 'Electronic E-mail Setup MustFill
            Set MPages = .Parent
            For Each ctrl In MPages.Controls
                If ctrl.Name Like "*Email*" Then
                    If TypeOf ctrl Is MSForms.TextBox Then ctrl.Tag = "MustFill": _
                    If ctrl.value = "" Then ctrl.BackColor = RGB(255, 255, 204) 'Yellow
                End If
            Next
        ElseIf .Name Like "*_E_*" And .value = False Then  'Electronic E-mail Setup NotFill
            If Not chbInForm.Parent.Parent.Parent.cbx_Verkaufsorganisation.value Like "2961*" Then
                Set MPages = .Parent
                For Each ctrl In MPages.Controls
                    If ctrl.Name Like "*Email*" Then
                        If TypeOf ctrl Is MSForms.TextBox Then ctrl.Tag = "NotFill": ctrl.BackColor = RGB(255, 255, 255) 'White
                    End If
                Next
            End If
        End If
    End With
End Sub

