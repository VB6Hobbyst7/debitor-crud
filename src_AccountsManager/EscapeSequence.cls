VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder AccountsManager.Common.List
Private Type tEscapeSequence
    EscapeString As String
    ReplacementString As String
End Type

Private This As tEscapeSequence
Option Explicit

Public Property Get EscapeString() As String
    EscapeString = This.EscapeString
End Property

Friend Property Let EscapeString(value As String)
    This.EscapeString = value
End Property

Public Property Get ReplacementString() As String
    ReplacementString = This.ReplacementString
End Property

Friend Property Let ReplacementString(value As String)
    This.ReplacementString = value
End Property

Public Function Create(escape As String, replacement As String) As EscapeSequence
    Dim result As New EscapeSequence
    result.EscapeString = escape
    result.ReplacementString = replacement
    Set Create = result
End Function
