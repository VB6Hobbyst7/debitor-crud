VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToDateConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A value converter that can convert a String into a Date, or a Date into a String."
'@Folder AccountsManager.Common.Converters
'@ModuleDescription "A value converter that can convert a String into a Date, or a Date into a String."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValueConverter

Public Function Default() As IValueConverter
    GuardClauses.GuardNonDefaultInstance Me, StringToDateConverter
    Set Default = StringToDateConverter
End Function

Private Function IValueConverter_Convert(ByVal value As Variant) As Variant
    IValueConverter_Convert = Convert(value)
End Function

Private Function IValueConverter_ConvertBack(ByVal value As Variant) As Variant
    IValueConverter_ConvertBack = Convert(value)
End Function

Private Function Convert(ByVal value As Variant) As Variant
    GuardClauses.GuardExpression Not IsDate(value), TypeName(Me), "Value '" & value & "' was not recognized as a valid date."
    If VarType(value) = vbDate Then
        Convert = CStr(value)
    Else
        Convert = CDate(value)
    End If
End Function
