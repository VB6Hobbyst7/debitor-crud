Attribute VB_Name = "modResizeToScreen"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------
'@Module: modResizeToScreen ************************************************************************************************'
'@Autor: *******************************************************************************************************'
'@Contact:  **********************************************************************************'
'----------------------------------------------------------------------------------------------------------------------------
'Function to get screen resolution
#If VBA7 Then
    Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" ( _
                                    ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" ( _
                                    ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
                                    ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
                                    ByVal hwnd As Long, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" ( _
                                    ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" ( _
                                    ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
                                    ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" ( _
                                    ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If
'Functions to get DPI
Private Const LOGPIXELSX = 88 'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72 'A point is defined as 1/72 inches
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNCTION - PointsPerPixel //////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function PointsPerPixel() As Double

    Dim hDC As LongPtr
    Dim lDotsPerInch As LongPtr
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC
End Function
