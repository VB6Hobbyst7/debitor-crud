Attribute VB_Name = "modApiDeclarations"
Option Explicit
Option Private Module
'----------------------------------------------------------------------------------------------------------------------------
'@Module: modApiDeclarations ***********************************************************************************************'
' By Chip Pearson, chip@cpearson.com, www.cpearson.com 21-March-2008 URL: http://www.cpearson.com/Excel/FormControl.aspx ***'
'----------------------------------------------------------------------------------------------------------------------------
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_USERS As Long = &H80000003
Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const ERROR_SUCCESS As Long = 0&
Public Const HKCU As Long = HKEY_CURRENT_USER
Public Const HKLM As Long = HKEY_LOCAL_MACHINE
Public Const C_USERFORM_CLASSNAME = "ThunderDFrame"
Public Const C_EXCEL_APP_CLASSNAME = "XLMain"
Public Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Public Const C_EXCEL_WINDOW_CLASSNAME = "Excel7"
Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const GWL_HWNDPARENT = (-8)
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2&
Public Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Public Const C_ALPHA_FULL_OPAQUE As Byte = 255
Public Const WS_DLGFRAME = &H400000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Enum REG_DATA_TYPE
    REG_DATA_TYPE_DEFAULT = 0   ' Default based on data type of value.
    REG_INVALID = -1            ' Invalid
    REG_SZ = 1                  ' String
    REG_DWORD = 4               ' Long
End Enum
Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum
#If VBA7 Then
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If
#If VBA7 Then
Public Declare PtrSafe Function SetParent Lib "user32" ( _
ByVal hWndChild As LongPtr, _
ByVal hWndNewParent As LongPtr) As LongPtr
#Else
Public Declare Function SetParent Lib "user32" ( _
ByVal hWndChild As Long, _
ByVal hWndNewParent As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As LongPtr
#Else
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
ByVal hwnd As LongPtr, _
ByVal nIndex As Long) As Long
#Else
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As LongPtr, _
ByVal nIndex As Long, _
ByVal dwNewLong As LongPtr) As Long
#Else
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hwnd As LongPtr, _
ByVal crey As Byte, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long
#Else
Public Declare Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hwnd As Long, _
ByVal crey As Byte, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
ByVal hWnd1 As LongPtr, _
ByVal hWnd2 As LongPtr, _
ByVal lpsz1 As String, _
ByVal lpsz2 As String) As LongPtr
#Else
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
ByVal hWnd1 As Long, _
ByVal hWnd2 As Long, _
ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
#Else
Public Declare Function GetActiveWindow Lib "user32" () As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
ByVal hwnd As LongPtr) As Long
#Else
Public Declare Function DrawMenuBar Lib "user32" ( _
ByVal hwnd As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetMenuItemCount Lib "user32" ( _
ByVal hMenu As LongPtr) As Long
#Else
Public Declare Function GetMenuItemCount Lib "user32" ( _
ByVal hMenu As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetSystemMenu Lib "user32" ( _
ByVal hwnd As LongPtr, _
ByVal bRevert As Long) As LongPtr
#Else
Public Declare Function GetSystemMenu Lib "user32" ( _
ByVal hwnd As Long, _
ByVal bRevert As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function RemoveMenu Lib "user32" ( _
ByVal hMenu As LongPtr, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
#Else
Public Declare Function RemoveMenu Lib "user32" ( _
ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
ByVal hwnd As LongPtr) As Long
#Else
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
ByVal hwnd As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
ByVal hwnd As LongPtr, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long
#Else
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
ByVal hwnd As Long, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function EnableMenuItem Lib "user32" ( _
ByVal hMenu As LongPtr, _
ByVal wIDEnableItem As Long, _
ByVal wEnable As Long) As Long
#Else
Public Declare Function EnableMenuItem Lib "user32" ( _
ByVal hMenu As Long, _
ByVal wIDEnableItem As Long, _
ByVal wEnable As Long) As Long
#End If
#If VBA7 Then
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
ByVal hwnd As LongPtr, _
ByVal lpString As String, _
ByVal cch As LongPtr) As LongPtr
#Else
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
ByVal hwnd As LongPtr, _
ByVal lpString As String, _
ByVal cch As LongPtr) As LongPtr
#End If
#If VBA7 Then
Public Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
ByVal HKey As LongPtr, _
ByVal lpSubKey As String, _
ByVal ulOptions As LongPtr, _
ByVal samDesired As LongPtr, _
phkResult As LongPtr) As LongPtr    'http://stackoverflow.com/questions/252297/why-is-regopenkeyex-returning-error-code-2-on-vista-64bit
#Else
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
ByVal HKey As LongPtr, _
ByVal lpSubKey As String, _
ByVal ulOptions As LongPtr, _
ByVal samDesired As LongPtr, _
phkResult As LongPtr) As LongPtr
#End If
#If VBA7 Then
Public Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
ByVal HKey As LongPtr, _
ByVal lpValueName As String, _
ByVal lpReserved As LongPtr, _
LPType As LongPtr, _
LPData As Any, _
lpcbData As LongPtr) As LongPtr
#Else
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
ByVal HKey As LongPtr, _
ByVal lpValueName As String, _
ByVal lpReserved As LongPtr, _
LPType As LongPtr, _
LPData As Any, _
lpcbData As LongPtr) As LongPtr
#End If
#If VBA7 Then
Public Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
ByVal HKey As LongPtr) As Long
#Else
Public Declare Function RegCloseKey Lib "advapi32.dll" ( _
ByVal HKey As LongPtr) As Long
#End If
