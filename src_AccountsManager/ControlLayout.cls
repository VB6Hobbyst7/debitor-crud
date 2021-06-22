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

Private this As TControlLayout

Public Sub Bind(ByVal Form As MSForms.UserForm, ByVal formControl As Control, ByVal Anchor As AnchorEdges)

    Set this.FormSize = New ControlSize
    Set this.Size = New ControlSize
    Set this.Margins = New SideMargins
    this.Anchors = Anchor

    this.FormSize.Height = Form.InsideHeight
    this.FormSize.Width = Form.InsideWidth

    this.Size.Height = formControl.Height
    this.Size.Width = formControl.Width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        this.Margins.BottomMargin = this.FormSize.Height - formControl.Top - formControl.Height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        this.Margins.LeftMargin = formControl.Left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        this.Margins.RightMargin = this.FormSize.Width - formControl.Left - formControl.Width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        this.Margins.TopMargin = formControl.Top
    End If

    Set this.BoundControl = formControl

End Sub

Public Sub Resize(ByVal Form As UserForm)

    this.FormSize.Height = Form.InsideHeight
    this.FormSize.Width = Form.InsideWidth

    Layout this.BoundControl

End Sub

Private Sub Layout(ByVal formControl As Control)

    If (Anchors And TopAnchor) = TopAnchor Then

    If (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.Height = this.FormSize.Height - formControl.Top - this.Margins.BottomMargin
    End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        formControl.Top = this.FormSize.Height - formControl.Height - this.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            formControl.Width = this.FormSize.Width - formControl.Left - this.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        formControl.Left = this.FormSize.Width - formControl.Width - this.Margins.RightMargin
    End If

End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = this.Anchors
End Property

Public Property Let Anchors(ByVal ancor As AnchorEdges)
    this.Anchors = ancor
End Property

Public Property Get BoundControl() As Control
    Set BoundControl = this.BoundControl
End Property

Public Property Set BoundControl(ByVal bound As Control)
    Set this.BoundControl = bound
End Property
