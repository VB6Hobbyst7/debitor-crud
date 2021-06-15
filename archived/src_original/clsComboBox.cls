VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: clsComboBox ******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim UForm As Object
Private WithEvents cbxInForm As MSForms.ComboBox
Attribute cbxInForm.VB_VarHelpID = -1
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Control ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Property Set Control(cbxContainer As MSForms.ComboBox)

    Set cbxInForm = cbxContainer
End Property
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'chbInForm_Change ///////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub cbxInForm_Change()
    
    Dim comboIndex  As Integer, comboName As String, partnerBox As Control
    Set UForm = cbxInForm.Parent.Parent.Parent
    With cbxInForm
        If .Tag Like "Must*" And .value <> "" Then
            '.BackColor = RGB(242, 242, 242) 'Button Face 'White on Win10
            .BackColor = RGB(255, 255, 255) 'White
        End If
        If UForm.Tag Like "New" Or UForm.Tag Like "Edit" Then
            Select Case UForm.cbx_Verkaufsorganisation.value
                Case Is = "2961 - Interlining DE"
                    If .Name Like "*Partner_Abteilung*" Then 'Set Email MustFill if CP is Accounting
                        If .value Like "*Buchhaltung*" Or .value Like "*Accounting*" Then
                            comboName = "tbx_Partner_Email"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.Tag = "MustFill"
                            Set partnerBox = Nothing 'Distroy Control from memory
                        Else    'Set Email NotFill if CP is Other
                            comboName = "tbx_Partner_Email"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.Tag = "NotFill"
                            Set partnerBox = Nothing 'Distroy Control from memory
                        End If
                    End If
                    If .Name Like "*Partnerrolle*" Then
                        If .value Like "*AP*" Then 'AP - Ansprechpartner OR CP - Contact Person (AP)
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = "Ansprechpartner_2961"
                            partnerBox.Tag = "MustFill"
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                            Set partnerBox = Nothing 'Distroy Control from memory
                        ElseIf .value Like "*ZF*" Then 'ZF - Fax-/Mailempfänger OR ZF - Fax-/Email recipien
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = "Fax_Mailempfaenger_2961"
                            partnerBox.Tag = "MustFill"
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                            Set partnerBox = Nothing 'Distroy Control from memory
                        ElseIf .value Like "*ZP*" Then 'ZP - Provisionsvertreter OR ZP - Commision Repres.
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = "Provisionsvertreter_2961"
                            partnerBox.Tag = "MustFill"
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                            Set partnerBox = Nothing 'Distroy Control from memory
                        ElseIf .value Like "*ZH*" Then 'ZH - Hauptlaufspediteur OR ZH - Main shipper
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = "Spediteur_2961"
                            partnerBox.Tag = "MustFill"
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                            Set partnerBox = Nothing 'Distroy Control from memory
                        Else    'PartnerRolle is else then listed above, clear rowSource
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = ""
                            partnerBox.Tag = "NotFill"
                            partnerBox.value = ""
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenNever
                            Set partnerBox = Nothing 'Distroy Control from memory
                        End If
                    End If
                Case Is = "0361 - Interlining ES"
                    If .Name Like "*Partnerrolle*" Then
                        If .value Like "*ZH*" Then 'ZH - Hauptlaufspediteur OR ZH - Main shipper
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = "Spediteur_0361"
                            partnerBox.Tag = "MustFill"
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                            Set partnerBox = Nothing 'Distroy Control from memory
                        Else    'PartnerRolle is else then listed above, clear rowSource
                            comboName = "cbx_Partner_Nr"
                            comboIndex = Right(.Name, 1)
                            Set partnerBox = UForm.Controls(comboName & comboIndex)
                            partnerBox.RowSource = ""
                            partnerBox.Tag = "NotFill"
                            partnerBox.value = ""
                            partnerBox.ShowDropButtonWhen = fmShowDropButtonWhenNever
                            Set partnerBox = Nothing 'Distroy Control from memory
                        End If
                    End If
                Case Else
                    'do nothing
            End Select
            Application.Run "Form_Rebuild.UpdateFormControls", UForm
        End If
    End With
End Sub
