VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
'@Folder AccountsManager.Infrastructure.View.Dynamic
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@Exposed
Option Explicit

Public Enum AnchorEdges
    LeftAnchor = 1
    TopAnchor = 2
    RightAnchor = 4
    BottomAnchor = 8
    AnchorAll = LeftAnchor + TopAnchor + RightAnchor + BottomAnchor
End Enum

Private Type TControlLayout
    Anchors As AnchorEdges
    Margins As SideMargins
    FormSize As ControlSize
    Size As ControlSize
    BoundControl As Control
End Type

Private This As TControlLayout

Public Sub Bind(ByVal Form As MSForms.UserForm, ByVal formControl As Control, ByVal Anchor As AnchorEdges)

    Set This.FormSize = New ControlSize
    Set This.Size = New ControlSize
    Set This.Margins = New SideMargins
    This.Anchors = Anchor

    This.FormSize.Height = Form.InsideHeight
    This.FormSize.Width = Form.InsideWidth

    This.Size.Height = formControl.Height
    This.Size.Width = formControl.Width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        This.Margins.BottomMargin = This.FormSize.Height - formControl.Top - formControl.Height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        This.Margins.LeftMargin = formControl.Left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        This.Margins.RightMargin = This.FormSize.Width - formControl.Left - formControl.Width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        This.Margins.TopMargin = formControl.Top
    End If

    Set This.BoundControl = formControl

End Sub

Public Sub Resize(ByVal Form As UserForm)

    This.FormSize.Height = Form.InsideHeight
    This.FormSize.Width = Form.InsideWidth

    Layout This.BoundControl

End Sub

Private Sub Layout(ByVal formControl As Control)

    If (Anchors And TopAnchor) = TopAnchor Then

    If (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.Height = This.FormSize.Height - formControl.Top - This.Margins.BottomMargin
    End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.Top = This.FormSize.Height - formControl.Height - This.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            formControl.Width = This.FormSize.Width - formControl.Left - This.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        formControl.Left = This.FormSize.Width - formControl.Width - This.Margins.RightMargin
    End If

End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = This.Anchors
End Property

Public Property Let Anchors(ByVal ancor As AnchorEdges)
    This.Anchors = ancor
End Property

Public Property Get BoundControl() As Control
    Set BoundControl = This.BoundControl
End Property

Public Property Set BoundControl(ByVal bound As Control)
    Set This.BoundControl = bound
End Property
