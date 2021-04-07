Attribute VB_Name = "Form_Email"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Email *******************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SubmitError_Outlook ////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub SubmitError_Outlook()

    Dim OutApp As Object, OutMail As Object, strbody As String, MailErrorTo As String
    On Error Resume Next
        MailErrorTo = Join(Application.Transpose(xx_frmConst.Range("Admin_Email").value), ";")
    If Err.Number = 13 Then Err.Clear: MailErrorTo = xx_frmConst.Range("Admin_Email").value
    Application.Cursor = xlWait
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        strbody = "Hi there," & vbNewLine & vbNewLine & _
                  "an error ocured please verify." & vbNewLine & vbNewLine & _
                  "'...Breaf description and screenshot if possible...'" & vbNewLine & _
                  "" & vbNewLine & _
                  "Thank you."
        On Error Resume Next
            With OutMail
                .To = MailErrorTo
                .CC = ""
                .BCC = ""
                .Subject = "Error found in " & ThisWorkbook.Name
                .Body = strbody
                .Display 'or .Send
            End With
        On Error GoTo 0
    Application.Cursor = xlDefault
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'EmialToEDocumentsSetter ////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub EmialToEDocumentsSetter(UForm As Object)

    Dim OutApp As Object, OutMail As Object, MailTo As String, strbody1 As String, uFormValues As String
    Application.Cursor = xlWait
    With UForm
        uFormValues = "Reactivated: " & .chb_Reaktivieren.value & vbNewLine & _
                        "Sales Org.: " & .cbx_Verkaufsorganisation.value & vbNewLine & _
                        "Distr. Channel: " & .cbx_Vertriebsweg.value & vbNewLine
        If .chb_E_Rechnung.value = True Then 'AG E-Invoice data
            uFormValues = uFormValues & vbNewLine & _
                            "Name: " & .tbx_Name1.value & vbNewLine & _
                            "SAP Nr.: " & .tbx_NewKUNA_SAPNr.value & vbNewLine & _
                            "Country: " & .cbx_Land.value & vbNewLine & _
                            "Email: " & .tbx_Email.value & vbNewLine
        End If
        If .chb_WE_E_Lieferschein.value = True Then 'WE E-Deliverynote data
            uFormValues = uFormValues & vbNewLine & _
                            "WE Name: " & .tbx_WE_Name1.value & vbNewLine & _
                            "SAP Nr.: " & .tbx_NewKUNW_SAPNr.value & vbNewLine & _
                            "WE Country: " & .cbx_WE_Land.value & vbNewLine & _
                            "WE Email: " & .tbx_WE_Email.value & vbNewLine
        End If
        If .chb_RE_E_Rechnung.value = True Then 'RE E-Invoice data
            uFormValues = uFormValues & vbNewLine & _
                            "RE Name: " & .tbx_RE_Name1.value & vbNewLine & _
                            "SAP Nr.: " & .tbx_NewKUNR_SAPNr.value & vbNewLine & _
                            "RE Country: " & .cbx_RE_Land.value & vbNewLine & _
                            "RE Email: " & .tbx_RE_Email.value & vbNewLine
        End If
    End With
    On Error Resume Next
        MailTo = Join(Application.Transpose(xx_frmConst.Range("E_DocumentsSetter").value), ";")
    If Err.Number = 13 Then Err.Clear: MailTo = xx_frmConst.Range("E_DocumentsSetter").value
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    strbody1 = "Hallo," & vbNewLine & vbNewLine & _
              "please set E-Documents settings in SAP." & vbNewLine & vbNewLine & _
              "Thank you." & vbNewLine & vbNewLine
    On Error Resume Next
        With OutMail
            .To = MailTo
            .CC = ""
            .BCC = ""
            .Subject = "E-Documents SAP Settings"
            .Body = strbody1 & uFormValues
            .Display 'or .Send
        End With
    On Error GoTo 0
    Application.Cursor = xlDefault
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
