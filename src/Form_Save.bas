Attribute VB_Name = "Form_Save"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: Form_Save ********************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'VARIABLES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Const passWordLock As String = "fnextxx" 'Excel Application Password
Dim ctrl As MSForms.Control
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Save ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Save(UForm As Object, Optional overwrite As Long, Optional approved As Boolean)

Dim i As Long, n As Long, xObjects As Long, nxtEmptCell As Long, updateChart As Integer
Dim SettingColumns As Excel.ListColumns
Dim ControlsAndAnchors() As ControlAndAnchors
Dim rFound As Range, clickPos As Range, xRange As Range, xLastColumn As Range
    If overwrite = 0 Then ' Next Empty or Edit data refered to Cell Position
        nxtEmptCell = aa_valData.Cells(aa_valData.Rows.Count, 1).End(xlUp).Offset(1, 0).Row ' Write Values to new Row if new data
    Else
        nxtEmptCell = overwrite ' Overwrite Values if clicked on cell with data
    End If
    Set clickPos = aa_valData.Cells(nxtEmptCell, 1)
    Set xLastColumn = aa_valData.Cells(nxtEmptCell, aa_valData.UsedRange.Columns.Count)
    Set xRange = aa_valData.Range(clickPos, xLastColumn)
    Set SettingColumns = xx_frmConst.ListObjects(1).ListColumns
    xObjects = UForm.Controls.Count
    ReDim ControlsAndAnchors(1 To xObjects)
    n = 0
    If approved And Len(UForm.tbx_Name1.value) > 1 Then updateChart = updateChart + 1
    If approved And Len(UForm.tbx_WE_Name1.value) > 1 Then updateChart = updateChart + 1
    If approved And Len(UForm.tbx_RE_Name1.value) > 1 Then updateChart = updateChart + 1
    If UForm.chb_Approved.Tag = "Counted" Then
        If UForm.Tag = "Approve" And Not approved And Len(UForm.tbx_Name1.value) > 1 Then updateChart = updateChart - 1
        If UForm.Tag = "Approve" And Not approved And Len(UForm.tbx_WE_Name1.value) > 1 Then updateChart = updateChart - 1
        If UForm.Tag = "Approve" And Not approved And Len(UForm.tbx_RE_Name1.value) > 1 Then updateChart = updateChart - 1
    End If
'    aa_valData.Unprotect Password:=passWordLock ' unprotect "Form" Worksheet
        For i = 1 To xx_frmConst.ListObjects(1).DataBodyRange.Rows.Count
            With ControlsAndAnchors(i)
                Set .ctl = UForm.Controls(SettingColumns("ctrl").DataBodyRange.Rows(i).value)
                If TypeOf .ctl Is MSForms.TextBox Or TypeOf .ctl Is MSForms.ComboBox Or TypeOf .ctl Is MSForms.CheckBox Then
                    If TypeOf .ctl Is MSForms.TextBox Then
                        clickPos.Offset(0, n).value = .ctl.Name
                        If .ctl.Name Like "*_Name*" Then
                            clickPos.Offset(0, n + 1).value = .ctl.value
                        Else
                            clickPos.Offset(0, n + 1).value = StrConv(.ctl.value, vbProperCase)
                        End If
                        'Anforderer
                        If .ctl.Name Like "*_Anforder*" Then
                            clickPos.Offset(0, n + 1).value = .ctl.value
                        End If
                    Else
                        clickPos.Offset(0, n).value = .ctl.Name
                        clickPos.Offset(0, n + 1).value = .ctl.value
                    End If
                    If approved Then 'Approved checked then colorRowColor to green
                        clickPos.Offset(0, n + 1).Font.Color = RGB(55, 86, 35) 'Dark Green Font
                        clickPos.Offset(0, n + 1).Interior.Color = RGB(226, 239, 218) 'Light Green
                    ElseIf .ctl.Name Like "tbx_NewKUN*" Then ' else color first SAPNr Cell to Blue
                        clickPos.Offset(0, n + 1).Font.Color = RGB(0, 112, 192) 'Dark Blue Font
                        clickPos.Offset(0, n + 1).Interior.Color = RGB(221, 235, 247) 'Licht Blue
                    End If
                    If Not approved And UForm.Tag = "Approve" Then  'Approved checked then colorRowColor to green
                        clickPos.Offset(0, n + 1).Font.Color = RGB(0, 0, 0) 'Black Font
                        clickPos.Offset(0, n + 1).Interior.Color = RGB(255, 255, 255) 'White
                    End If
                n = n + 2
                End If
            End With
        Next i
        xRange.WrapText = False
        If overwrite = 0 And Not approved And Not UForm.Tag Like "Approve" Then
            aa_valData.Shapes("cht_Overview").IncrementTop 18 'Move one row down Overview Chart
        End If
        If UForm.Tag Like "Approve" Then
            
            Select Case UForm.cbx_Verkaufsorganisation.value 'Increment Overview Chart Values
                Case Is = "0361 - Interlining ES" '0361 - Interlining ES
                    If UForm.chb_Reaktivieren.value Then
                        xx_frmFeed_0361.Range("Reactivated_0361").value = xx_frmFeed_0361.Range("Reactivated_0361").value + updateChart
                    Else
                        xx_frmFeed_0361.Range("New_0361").value = xx_frmFeed_0361.Range("New_0361").value + updateChart
                    End If
                Case Is = "2561 - Interlining TR" '2561 - Interlining TR
                    If UForm.chb_Reaktivieren.value Then
                        xx_frmFeed_2561.Range("Reactivated_2561").value = xx_frmFeed_2561.Range("Reactivated_2561").value + updateChart
                    Else
                        xx_frmFeed_2561.Range("New_2561").value = xx_frmFeed_2561.Range("New_2561").value + updateChart
                    End If
                Case Is = "2961 - Interlining DE" '2961 - Interlining DE
                    If UForm.chb_Reaktivieren.value Then
                        xx_frmFeed_2961.Range("Reactivated_2961").value = xx_frmFeed_2961.Range("Reactivated_2961").value + updateChart
                    Else
                        xx_frmFeed_2961.Range("New_2961").value = xx_frmFeed_2961.Range("New_2961").value + updateChart
                    End If
                Case Is = "3661 - FPM Apparel IT" '3661 - FPM Apparel IT
                    If UForm.chb_Reaktivieren.value Then
                        xx_frmFeed_3661.Range("Reactivated_3661").value = xx_frmFeed_3661.Range("Reactivated_3661").value + updateChart
                    Else
                        xx_frmFeed_3661.Range("New_3661").value = xx_frmFeed_3661.Range("New_3661").value + updateChart
                    End If
                Case Else ' Do Nothing
            End Select
        End If
'    aa_valData.Protect Password:=passWordLock, AllowFiltering:=True ' protect "Form" Worksheet
End Sub
