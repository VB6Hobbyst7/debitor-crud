Attribute VB_Name = "modPearsonApiWindowCaption"
Option Explicit
Option Compare Text
Option Private Module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modWindowCaption
' By Chip Pearson, 15-March-2008, chip@cpearson.com, www.cpearson.com
' http://www.cpearson.com/Excel/FileExtensions.aspx
'
' This module contains code for working with Excel.Window captions. This code
' is necessary if you are going to use the FindWindowEx API call to get the
' HWnd of an Excel.Window.  Windows has a property called "Hide extensions of
' known file types". If this setting is TRUE, the file extension is not displayed
' (e.g., "Book1.xls" is displayed as just "Book1"). However, the Caption of an
' Excel.Window always includes the ".xls" file extension, regardless of the hide
' extensions setting. FindWindowEx requires that the ".xls" extension be removed
' if the "hide extensions" setting is True.
'
' This module contains a function named DoesWindowsHideFileExtensions, which returns
' TRUE if Windows is hiding file extensions or FALSE if Windows is not hiding file
' extensions. This is determined by a registry key. The module also contains a
' function named WindowCaption that returns the Caption of a specified Excel.Window
' with the ".xls" removed if necessary. The string returned by this function
' is suitable for use in FindWindowEx regardless of the value of the Windows
' "Hide Extensions" setting.
'
' This module also contains a function named WindowHWnd which returns the HWnd of
' a specified Excel.Window object. This function works regardless of the value of the
' Windows "Hide Extensions" setting.
'
' This module also includes the functions WindowText and WindowClassName which are
' just wrappers for the GetWindowText and GetClassName API functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_WINDOW_CLASSNAME = "EXCEL7"
Function DoesWindowsHideFileExtensions() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DoesWindowsHideFileExtensions
' This function looks in the registry key
'   HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced
' for the value named "HideFileExt" to determine whether the Windows Explorer
' setting "Hide Extensions Of Known File Types" is enabled. This function returns
' TRUE if this setting is in effect (meaning that Windows displays "Book1" rather
' than "Book1.xls"), or FALSE if this setting is not in effect (meaning that Windows
' displays "Book1.xls").
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim v As Long
#If VBA7 Then
Dim RegKey As LongPtr
Dim Res As LongPtr
#Else
Dim RegKey As Long
Dim Res As Long
#End If
Const KEY_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Const VALUE_NAME = "HideFileExt"
''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the registry key to get a handle (RegKey).
''''''''''''''''''''''''''''''''''''''''''''''''''
Res = RegOpenKeyEx(HKey:=HKCU, _
lpSubKey:=KEY_NAME, _
ulOptions:=0&, _
samDesired:=KEY_ALL_ACCESS, _
phkResult:=RegKey)
If Res <> ERROR_SUCCESS Then
Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the value of the "HideFileExt" named value.
''''''''''''''''''''''''''''''''''''''''''''''''''
Res = RegQueryValueEx(HKey:=RegKey, _
lpValueName:=VALUE_NAME, _
lpReserved:=0&, _
LPType:=REG_DWORD, _
LPData:=v, _
lpcbData:=Len(v))
If Res <> ERROR_SUCCESS Then
RegCloseKey RegKey
Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the key and return the result.
''''''''''''''''''''''''''''''''''''''''''''''''''
RegCloseKey RegKey
DoesWindowsHideFileExtensions = (v <> 0)
End Function
Function WindowCaption(w As Excel.Window) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowCaption
' This returns the Caption of the Excel.Window W with the ".xls" extension removed
' if required. The string returned by this function is suitable for use by
' the FindWindowEx API regardless of the value of the Windows "Hide Extensions"
' setting.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim HideExt As Boolean
Dim Cap As String
Dim Pos As Long
HideExt = DoesWindowsHideFileExtensions()
Cap = w.Caption
If HideExt = True Then
Pos = InStrRev(Cap, ".")
If Pos > 0 Then
Cap = Left(Cap, Pos - 1)
End If
End If
WindowCaption = Cap
End Function
Function WindowHWnd(w As Excel.Window) As LongPtr
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowHWnd
' This returns the HWnd of the Window referenced by W.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim AppHWnd As Long
Dim Cap As String
#If VBA7 Then
Dim DeskHWnd As LongPtr
Dim WHWnd As LongPtr
#Else
Dim DeskHWnd As Long
Dim WHWnd As Long
#End If
AppHWnd = Application.hwnd
DeskHWnd = FindWindowEx(AppHWnd, 0&, C_EXCEL_DESK_CLASSNAME, vbNullString)
If DeskHWnd > 0 Then
Cap = WindowCaption(w)
WHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, Cap)
End If
WindowHWnd = WHWnd
End Function
Function WindowText(hwnd As LongPtr) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowText
' This just wraps up GetWindowText.
'************Modified by Doug Glancy 2016-12-28 to split the Long variable N into n and N_temp, so would compile in 64-bit.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim S As String
Dim N_temp As Long
Dim n As LongPtr
N_temp = 255
S = String$(N_temp, vbNullChar)
n = GetWindowText(hwnd, S, N_temp)
If n > 0 Then
WindowText = Left(S, N_temp)
Else
WindowText = vbNullString
End If
End Function
Function WindowClassName(hwnd As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowClassName
' This just wraps up GetClassName.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim S As String
Dim n As Long
n = 255
S = String$(n, vbNullChar)
n = GetClassName(hwnd, S, n)
If n > 0 Then
WindowClassName = Left(S, n)
Else
WindowClassName = vbNullString
End If
End Function
