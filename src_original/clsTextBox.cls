VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBox"
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
Dim UForm As Object
Private WithEvents tbxInForm As MSForms.TextBox
Attribute tbxInForm.VB_VarHelpID = -1
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Control ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Set Control(tbxContainer As MSForms.TextBox)

    Set tbxInForm = tbxContainer
End Property
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'tbxInForm_Change ////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub tbxInForm_Change()
    
    Dim AbteilungIndex  As Integer, AbteilungName As String, AbteilungBox As Control
    Set UForm = tbxInForm.Parent.Parent.Parent
    With tbxInForm
        .ForeColor = RGB(0, 0, 0) 'Black
        If Application.Min(Int(Len(.Text))) > .Width Then
            .MultiLine = True
        End If
        If .Tag Like "Must*" And .value <> "" Then
            '.BackColor = RGB(242, 242, 242) 'Button Face 'White on Win10
            .BackColor = RGB(255, 255, 255) 'White
        End If
        If .Name Like "*Debitor_SAPNr*" Then
            If .value <> "" And Len(.value) < 6 Then
                UForm.chb_Reaktivieren.value = True
                UForm.tbx_Debitor_SAPNr.Tag = "MustFill"
            End If
        End If
        If UForm.Tag Like "New" Or UForm.Tag Like "Edit" Then
            Select Case UForm.cbx_Verkaufsorganisation.value
                Case Is = "2961 - Interlining DE"
                    If .Name Like "*Partner_*name*" Then
                        If Len(.value) >= 1 Then 'Set Department MustFill if there are CP's.
                            AbteilungName = "cbx_Partner_Abteilung"
                            AbteilungIndex = Right(.Name, 1)
                            Set AbteilungBox = UForm.Controls(AbteilungName & AbteilungIndex)
                            AbteilungBox.Tag = "MustFill"
                            Set AbteilungBox = Nothing 'Distroy Control from memory
                        Else    'Set Department NotFill if there are no CP's.
                            AbteilungName = "cbx_Partner_Abteilung"
                            AbteilungIndex = Right(.Name, 1)
                            Set AbteilungBox = UForm.Controls(AbteilungName & AbteilungIndex)
                            AbteilungBox.Tag = "NotFill"
                            Set AbteilungBox = Nothing 'Distroy Control from memory
                        End If
                    End If
                Case Else
                    'do nothing
            End Select
            Application.Run "Form_Rebuild.UpdateFormControls", UForm
        End If
    End With
End Sub
