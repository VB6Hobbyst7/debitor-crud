Attribute VB_Name = "Form_TextSAP"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_TextSAP ****************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RemarksToDesigner //////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub RemarksToDesigner(UForm As Object)

    If UForm.Tag Like "Approve" Then
        With UForm
            'Custom SAPText 2961
            If .cbx_Kontengruppe.value Like "*KUNA*" And .cbx_Verkaufsorganisation.value Like "*2961*" Then
                If .tbx_UserLangVal.value = 49 Then 'German
                    If .cbx_Vertriebsweg.value Like "*IU*" And .cbx_Mindermengenzuschlag.value Like "*450 GBP*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Mindermengenzuschlag Text:" + vbCr + _
                                            "Minimum order 450 GBP"
                    End If
                    If .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*150*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Mindermengenzuschlag Text:" + vbCr + _
                                            "Kleinkunde mit einem Mindestbestellwert von 150 EUR." + vbCr + _
                                            "Ggf. ist ein Mindermengenzuschlag in Höhe von 30 EUR fällig."
                    ElseIf .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*300*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Mindermengenzuschlag Text:" + vbCr + _
                                            "Mindestbestellwert von EUR 300."
                    ElseIf Not .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*300*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Mindermengenzuschlag Text:" + vbCr + _
                                            "Kleinkunde mit einem Mindestbestellwert von 300 EUR." + vbCr + _
                                            "Ggf. ist ein Mindermengenzuschlag in Höhe von 30 EUR fällig."
                    End If
                    If .chb_Keine_Fracht.value = False Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Lieferbedingungen Anh.text 3:" + vbCr + _
                                            "INCOTERMS001" + vbCr + _
                                            "Sprache:  " & Left(.cbx_Sprache.value, 3)
                    End If
                    If .cbx_Zahlungsbedingung.value Like "K*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Zahlungsbedingung Text:" + vbCr + _
                                            "Vorkasse- bitte die Frachtkosten manuell in der Auftrag eingeben"
                    End If
                Else 'English
                    If .cbx_Vertriebsweg.value Like "*IU*" And .cbx_Mindermengenzuschlag.value Like "*450 GBP*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Min. orderquant. surcharge:" + vbCr + _
                                            "Minimum order 450 GBP"
                    End If
                    If .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*150*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Min. orderquant. surcharge:" + vbCr + _
                                            "Kleinkunde mit einem Mindestbestellwert von 150 EUR." + vbCr + _
                                            "Ggf. ist ein Mindermengenzuschlag in Höhe von 30 EUR fällig."
                    ElseIf .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*300*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Min. orderquant. surcharge:" + vbCr + _
                                            "Minimum order EUR 300."
                    ElseIf Not .cbx_Vertriebsweg.value Like "*HD*" And .cbx_Mindermengenzuschlag.value Like "*300*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Min. orderquant. surcharge:" + vbCr + _
                                            "Kleinkunde mit einem Mindestbestellwert von 300 EUR." + vbCr + _
                                            "Ggf. ist ein Mindermengenzuschlag in Höhe von 30 EUR fällig."
                    End If
                    If .chb_Keine_Fracht.value = False Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Terms of delivery Anh.text 3:" + vbCr + _
                                            "INCOTERMS001" + vbCr + _
                                            "Lang.:  " & Left(.cbx_Sprache.value, 3)
                    End If
                    If .cbx_Zahlungsbedingung.value Like "K*" Then
                        .tbx_TextSAP.Text = .tbx_TextSAP.Text + vbCr + _
                                            "Peyment term Text:" + vbCr + _
                                            "Payment in advance. Please introduce the Freight charges manually in the Order."
                    End If
                End If
            End If
        End With
    End If
End Sub

'"Vorkasse- bitte die Frachtkosten manuell in der Auftrag eingeben"
