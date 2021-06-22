Attribute VB_Name = "Win32API"
Attribute VB_Description = "Win32 API's, Sets minimize and maximize Buttons to Form Toolbar, sets Form Opacity, makes Userform Resizable and mesures Points per Pixel to get screen resolution."
'@Folder AccountsManager.Infrastructure.Win32
'@ModuleDescription("Win32 API's, Sets minimize and maximize Buttons to Form Toolbar, sets Form Opacity, makes Userform Resizable and mesures Points per Pixel to get screen resolution.")
'@IgnoreModule UserMeaningfulName, HungarianNotation; Win32 parameter names are what they are
'All Rights reserverd to: By Chip Pearson, chip@cpearson.com, www.cpearson.com 21-March-2008
'URL: http://www.cpearson.com/Excel/FormControl.aspx
'URL: http://www.cpearson.com/Excel/FileExtensions.aspx

Option Explicit
Option Compare Text
Option Private Module

Public Const HKEY_CURRENT_user As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_userS As Long = &H80000003
Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const ERROR_SUCCESS As Long = 0&
Public Const HKCU As Long = HKEY_CURRENT_user
'@Ignore UseMeaningfulName
Public Const HKLM As Long = HKEY_LOCAL_MACHINE
Public Const C_userFORM_CLASSNAME = "ThunderDFrame"
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
Public Const LOGPIXELSX = 88                     'Pixels/inch in X
Public Const POINTS_PER_INCH As Long = 72        'A point is defined as 1/72 inches

'@Description "Get Points Per Pixel Screen resloution."
Public Function PointsPerPixel() As Double
Attribute PointsPerPixel.VB_Description = "Get Points Per Pixel Screen resloution."

    #If VBA7 Then
        '@Ignore UseMeaningfulName
        Dim hDC As LongPtr
        '@Ignore HungarianNotation
        Dim lDotsPerInch As LongPtr
    #Else
        Dim hDC As Long
        Dim lDotsPerInch As Long
    #End If

    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC

End Function

'@Description "Displays (if HideButton is False) or hides (if HideButton is True) a maximize window button"
    'NOTE: If EITHER a Minimize or Maximize button is displayed,
    'BOTH buttons are visible but may be disabled.
Public Sub ShowMaximizeButton(ByVal View As Object, HideButton As Boolean)
Attribute ShowMaximizeButton.VB_Description = "Displays (if HideButton is False) or hides (if HideButton is True) a maximize window button"

    Dim WinInfo As Long
    '@Ignore UseMeaningfulName
    Dim r As Long
    
    #If VBA7 Then
        Dim UFHWnd As LongPtr
    #Else
        Dim UFHWnd As Long
    #End If
    
    UFHWnd = HWndOfuserForm(View)
    
    If UFHWnd = 0 Then
        
        Exit Sub
    End If
    
    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MAXIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MAXIMIZEBOX)
    End If
    
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    
End Sub

'@Description "Displays (if HideButton is False) or hides (if HideButton is True) a minimize window button."
    ' NOTE: If EITHER a Minimize or Maximize button is displayed,
    ' BOTH buttons are visible but may be disabled.
Public Function ShowMinimizeButton(ByVal View As Object, HideButton As Boolean) As Boolean
Attribute ShowMinimizeButton.VB_Description = "Displays (if HideButton is False) or hides (if HideButton is True) a minimize window button."

    Dim WinInfo As Long
    '@Ignore UseMeaningfulName
    Dim r As Long
    
    #If VBA7 Then
        Dim UFHWnd As LongPtr
    #Else
        Dim UFHWnd As Long
    #End If
    
    UFHWnd = HWndOfuserForm(View)
    
    If UFHWnd = 0 Then
        ShowMinimizeButton = False
        Exit Function
    End If
    
    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MINIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MINIMIZEBOX)
    End If
    
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    ShowMinimizeButton = (r <> 0)
    
End Function

'@Description "This makes the userform UF resizable (if Sizable is TRUE) or not resizable (if Sizalbe is FALSE)."
Public Sub MakeFormResizable(ByRef View As Object, ByVal Sizable As Boolean)
Attribute MakeFormResizable.VB_Description = "This makes the userform UF resizable (if Sizable is TRUE) or not resizable (if Sizalbe is FALSE)."

    Dim WinInfo As Long
    '@Ignore UseMeaningfulName
    Dim r As Long
    
    #If VBA7 Then
        Dim UFHWnd As LongPtr
    #Else
        Dim UFHWnd As Long
    #End If
    
    UFHWnd = HWndOfuserForm(View)
    
    If UFHWnd = 0 Then
        
        Exit Sub
    End If
    
    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    
    If Sizable = True Then
        WinInfo = WinInfo Or WS_SIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_SIZEBOX)
    End If
    
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    
End Sub

'@Description "This function sets the opacity of the userForm referenced by the UF parameter. From 0 = fully transparent (invisible) to 255 = fully opaque."
Public Function SetFormOpacity(ByRef View As Object, ByVal Opacity As Byte) As Boolean
Attribute SetFormOpacity.VB_Description = "This function sets the opacity of the userForm referenced by the UF parameter. From 0 = fully transparent (invisible) to 255 = fully opaque."

    Dim WinL As Long
    Dim Res As Long
    
    #If VBA7 Then
        Dim UFHWnd As LongPtr
    #Else
        Dim UFHWnd As Long
    #End If
    
    SetFormOpacity = False
    UFHWnd = HWndOfuserForm(View)
    
    If UFHWnd = 0 Then
        Exit Function
    End If
    
    WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
    
    If WinL = 0 Then
        Exit Function
    End If
    
    Res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
    
    If Res = 0 Then
        Exit Function
    End If
    
    Res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
    
    If Res = 0 Then
        Exit Function
    End If
    
    SetFormOpacity = True
    
End Function

Private Function DoesWindowsHideFileExtensions() As Boolean
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
    '@Ignore UseMeaningfulName
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
    Res = RegOpenKeyEx(HKey:=HKCU, lpSubKey:=KEY_NAME, ulOptions:=0&, samDesired:=KEY_ALL_ACCESS, phkResult:=RegKey)
    
    If Res <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the value of the "HideFileExt" named value.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Res = RegQueryValueEx(HKey:=RegKey, lpValueName:=VALUE_NAME, lpReserved:=0&, LPType:=REG_DWORD, _
                          LPData:=v, lpcbData:=Len(v))
        
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

'@Ignore UseMeaningfulName
Private Function WindowCaption(w As Excel.Window) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' WindowCaption
    ' This returns the Caption of the Excel.Window W with the ".xls" extension removed
    ' if required. The string returned by this function is suitable for use by
    ' the FindWindowEx API regardless of the value of the Windows "Hide Extensions"
    ' setting.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim HideExt As Boolean
    Dim Cap As String
    Dim pos As Long
    
    HideExt = DoesWindowsHideFileExtensions()
    Cap = w.Caption
    
    If HideExt = True Then
        pos = InStrRev(Cap, ".")
        
        If pos > 0 Then
            Cap = Left$(Cap, pos - 1)
        End If
    End If
    WindowCaption = Cap
    
End Function

'@Ignore UseMeaningfulName
Private Function WindowHWnd(w As Excel.Window) As LongPtr
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' WindowHWnd
    ' This returns the HWnd of the Window referenced by W.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim AppHWnd As Long
    Dim Cap As String
    
    #If VBA7 Then
        Dim DeskHWnd As LongPtr
        '@Ignore UseMeaningfulName
        Dim WHWnd As LongPtr
    #Else
        Dim DeskHWnd As Long
        Dim WHWnd As Long
    #End If
    
    AppHWnd = Application.hWnd
    DeskHWnd = FindWindowEx(AppHWnd, 0&, C_EXCEL_DESK_CLASSNAME, vbNullString)
    
    If DeskHWnd > 0 Then
        Cap = WindowCaption(w)
        WHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, Cap)
    End If
    WindowHWnd = WHWnd
    
End Function

'@Ignore UseMeaningfulName
Private Function WindowText(hWnd As LongPtr) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' WindowText
    ' This just wraps up GetWindowText.
    '************Modified by Doug Glancy 2016-12-28 to split the Long variable N into n and N_temp,
    '************so would compile in 64-bit.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '@Ignore UseMeaningfulName
    Dim s As String
    Dim N_temp As Long
    '@Ignore UseMeaningfulName
    Dim N As LongPtr
    
    N_temp = 255
    s = String$(N_temp, vbNullChar)
    N = GetWindowText(hWnd, s, N_temp)
    
    If N > 0 Then
        WindowText = Left$(s, N_temp)
    Else
        WindowText = vbNullString
    End If
    
End Function

'@Ignore UseMeaningfulName
Private Function WindowClassName(hWnd As Long) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' WindowClassName
    ' This just wraps up GetClassName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '@Ignore UseMeaningfulName
    Dim s As String
    '@Ignore UseMeaningfulName
    Dim N As Long
    
    N = 255
    s = String$(N, vbNullChar)
    N = GetClassName(hWnd, s, N)
        
    If N > 0 Then
        WindowClassName = Left$(s, N)
    Else
        WindowClassName = vbNullString
    End If
    
End Function

Private Function HWndOfuserForm(ByRef View As Object) As LongPtr
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' HWndOfuserForm
    ' This returns the window handle (HWnd) of the userform referenced
    ' by UF. It first looks for a top-level window, then a child
    ' of the Application window, then a child of the ActiveWindow.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim AppHWnd As Long
    Dim Cap As String
    
    #If VBA7 Then
        Dim UFHWnd As LongPtr
        Dim WinHWnd As LongPtr
    #Else
        Dim UFHWnd As Long
        Dim WinHWnd As Long
    #End If
    
    Cap = View.Caption
    ' First, look in top level windows
    UFHWnd = FindWindow(C_userFORM_CLASSNAME, Cap)
    
    If UFHWnd <> 0 Then
        HWndOfuserForm = UFHWnd
        Exit Function
    End If
    
    ' Not a top level window. Search for child of application.
    AppHWnd = Application.hWnd
    UFHWnd = FindWindowEx(AppHWnd, 0&, C_userFORM_CLASSNAME, Cap)
    
    If UFHWnd <> 0 Then
        HWndOfuserForm = UFHWnd
        Exit Function
    End If
    
    ' Not a child of the application.
    ' Search for child of ActiveWindow (Excel's ActiveWindow, not
    ' Window's ActiveWindow).
    If Application.ActiveWindow Is Nothing Then
        HWndOfuserForm = 0
        Exit Function
    End If
    
    Dim WindowHWnd As Variant
    WinHWnd = WindowHWnd(Application.ActiveWindow)
    UFHWnd = FindWindowEx(WinHWnd, 0&, C_userFORM_CLASSNAME, Cap)
    HWndOfuserForm = UFHWnd
    
End Function

Private Function ClearBit(ByRef Value As Long, ByRef BitNumber As Long) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ClearBit
    ' Clears the specified bit in Value and returns the result. Bits are
    ' numbered, right (most significant) 31 to left (least significant) 0.
    ' BitNumber is made positive and then MOD 32 to get a valid bit number.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim SetMask As Long
    Dim ClearMask As Long
    
    BitNumber = Abs(BitNumber) Mod 32
    SetMask = Value
    
    If BitNumber < 30 Then
        ClearMask = Not (2 ^ (BitNumber - 1))
        ClearBit = SetMask And ClearMask
    Else
        ClearBit = Value And &H7FFFFFFF
    End If
End Function
