# Converted from WindowControl.py

# Attribute VB_Name = "WindowControl"
# Option Explicit

# ' Original Module Name: Ol_WindowOps
# ' (c) 2005 Wayne Phillips (http://www.everythingaccess.com)
# ' Written 02/06/2005
# ' http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground

# ' Custom structure for passing in the parameters in/out of the hook enumeration function

# Private Const SW_RESTORE = 9
# Private Const SW_Show = 5

# Public CountTries As Long

# '---------------------------------------------------------------------------------------
# ' Method : Function EnumWindowProc
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Callback-Function called by GetWindowText (Library Dll)
# '---------------------------------------------------------------------------------------
# Private Function EnumWindowProc(ByVal hwnd As Long, lParam As cFindWindowParms) As Long
# Dim strWindowTitle As String
# Const DebugMe As Boolean = False

# With lParam
# .hwnd = hwnd
# strWindowTitle = Space(260) & vbNullChar ' provide buffer
# ' Library GetWindowText calls back EnumWindowProc (addressed via lParam)
# ' to compare each open window, until the window title matches
# CountTries = CountTries + 1
# Call GetWindowText(hwnd, strWindowTitle, 260)
# strWindowTitle = TrimNul(strWindowTitle) ' Remove extra null terminator
if LenB(strWindowTitle) = 0 Then:
# EnumWindowProc = 1
elif LCase(strWindowTitle) Like LCase(.strTitle) Then:
if DebugMe And LenB(strWindowTitle) > 0 Then:
print(Debug.Print " Handle " & hwnd & ", Window name " & strWindowTitle & " found")
# .hwnd = hwnd                         'Store the result for later.
# .strTitle = strWindowTitle           ' this is literally what we found
if DebugMe Then:
print(Debug.Print " Handle " & hwnd & ", Window name matched " & .strTitle & " found")
# EnumWindowProc = 0                   'This will Debug.Assert False enumerating more windows
else:
if DebugMe Then:
if InStr(1, strWindowTitle, "outl", vbTextCompare) > 0 _:
# And DebugLogging Then
# DoVerify False
# EnumWindowProc = 1                   ' continue loop
# End With                                     ' lParam


# '---------------------------------------------------------------------------------------
# ' Method : Function FindWindowLike
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Custom structure for passing in the parameters in/out of the hook enumeration function
# '---------------------------------------------------------------------------------------
def findwindowlike():
    # Dim Parameters As cFindWindowParms

    # CountTries = 0
    if wHdl Is Nothing Then:
    # seekWin:
    # Set Parameters = New cFindWindowParms
    # With Parameters
    # .strTitle = strWindowTitle           ' Input parameter
    # ' pass callback routine ---V
    # Call EnumWindows(AddressOf EnumWindowProc, VarPtr(Parameters))
    if Not LCase(.strTitle) Like LCase(strWindowTitle) Then:
    # Parameters.hwnd = 0              ' no Match
    # .strTitle = vbNullString
    # FindWindowLike = .hwnd
    # End With                                 ' Parameters
    # Set wHdl = Parameters
    else:
    if Not LCase(wHdl.strTitle) Like LCase(strWindowTitle) Then:
    # Set wHdl = Nothing
    # GoTo seekWin

# '---------------------------------------------------------------------------------------
# ' Method : Function WindowSetForeground
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function WindowSetForeground(strWindowTitle As String, wHdl As cFindWindowParms) As Boolean
# Dim MyAppHWnd As Long
# Dim CurrentForegroundThreadID As Long
# Dim NewForegroundThreadID As Long
# Dim lngRetVal As Long
# Dim blnSuccessful As Boolean
# Dim canSkip As Boolean

# unskip:
if wHdl Is Nothing Then:
# canSkip = True
else:
# canSkip = False
# MyAppHWnd = FindWindowLike(strWindowTitle, wHdl)

if MyAppHWnd <> 0 Then:

# ' We have found the application window by the caption
# CurrentForegroundThreadID = _
# GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
# NewForegroundThreadID = GetWindowThreadProcessId(MyAppHWnd, ByVal 0&)

# ' AttachThreadInput is used to ensure SetForegroundWindow will work
# ' even if our application isn'W.xlTSheet currently the foreground window
# ' (e.g. an automated app running in the background)
# Call AttachThreadInput(CurrentForegroundThreadID, _
# NewForegroundThreadID, True)
# lngRetVal = SetForegroundWindow(MyAppHWnd)
# Call AttachThreadInput(CurrentForegroundThreadID, _
# NewForegroundThreadID, False)

if lngRetVal <> 0 Then:

# ' Now that the window is active, let's restore it from the taskbar
if IsIconic(MyAppHWnd) Then:
# Call ShowWindow(MyAppHWnd, SW_RESTORE)
else:
# Call ShowWindow(MyAppHWnd, SW_Show)

# blnSuccessful = True
# wHdl.hwnd = MyAppHWnd
if DebugLogging Then:
print('Found the window ')
# & "  and it should be in foreground"
else:
if DebugLogging Then:
print('Found the window ')
# & ", but failed to bring it to the foreground!"
else:
print(Debug.Print "Found the window " & Quote(strWindowTitle) _)
# & ", but failed to bring it to the foreground!"
if DebugLogging Then:
# Call IsModalWindow(MyAppHWnd)
# DebugLogging = True
else:
# 'Failed to find the window caption
# 'Therefore the app is probably closed.
# Set wHdl = Nothing
if Not canSkip Then                      ' window handle could have changed:
# GoTo unskip                          ' locate again
print('Application Window ')

# WindowSetForeground = blnSuccessful

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : IsModalWindow
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: True if Modal
# ' Source : https://devblogs.microsoft.com/oldnewthing/20111212-00/?p=8923
# ' Author : Raymond Chen
# '---------------------------------------------------------------------------------------
def ismodalwindow():
    # Const zKey As String = "WindowControl.IsModalWindow"

    # Dim hwndOwnerl As LongPtr
    # Dim wi As PWINDOWINFO
    # Dim wo As PWINDOWINFO
    # Dim ws As WINDOWS_STYLES

    # #If Win64 Then
    # Dim lngHwnd As LongPtr
    # Dim hwndOwner As LongLong
    # wi.cbSize = LenB(wi)
    # wo.cbSize = LenB(wo)
    # #Else
    # Dim lngHwnd As Long
    # Dim hwndOwner As Long
    # wi.cbSize = Len(wi)
    # wo.cbSize = Len(wo)
    # #End If

    # lngHwnd = hwnd
    if GetWindowInfo(lngHwnd, wi) Then           ' child windows cannot have owners:
    # Call ShowWindowStates(lngHwnd, wi)
    # hwndOwner = GetWindow(lngHwnd, GW_OWNER)
    if hwndOwner Then                        ' has owner window:
    if GetWindowInfo(hwndOwner, wo) Then ' owner is enabled:
    # Call ShowWindowStates(hwndOwner, wo)
    if Not (wo.dwWindowStatus And WS_DISABLED) Then:
    # IsModalWindow = True         ' owner is disabled: Modal ==> True

    # zExit:

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : ShowWindowStates
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: ...
# '---------------------------------------------------------------------------------------
def showwindowstates():

    # Const zKey As String = "WindowControl.ShowWindowStates"

    # Dim wiWinTitle As String

    # wiWinTitle = Space(260) & vbNullChar         ' provide buffer
    # Call GetWindowText(lngHwnd, wiWinTitle, 260)
    # wiWinTitle = TrimNul(wiWinTitle)             ' Remove extra null terminator
    print(Debug.Print vbCrLf, "Window data for " & wiWinTitle)
    print(Debug.Print String(80, "="))

    # Call PrintOkNok(WS_BORDER And wi.dwStyle, "Fenster hat einen Rahmen")
    # Call PrintOkNok(WS_CAPTION And wi.dwStyle, "Fenster hat Titelleiste")
    # Call PrintOkNok(WS_CHILD And wi.dwStyle, "Kindfenster. Kann kein Men haben und nicht Popup")
    # Call PrintOkNok(WS_CLIPCHILDREN And wi.dwStyle, "Verhindert Kindfenster auerhalb des Elternfensters")
    # Call PrintOkNok(WS_CLIPSIBLINGS And wi.dwStyle, "Verhindert sich gegenseitig berlappende Kindfenster")
    # Call PrintOkNok(WS_DISABLED And wi.dwStyle, "Fenster ist initial Disabled")
    # Call PrintOkNok(WS_DLGFRAME And wi.dwStyle, "Dialogfenster")
    # Call PrintOkNok(WS_HSCROLL And wi.dwStyle, "Fenster hat horizontale Scrollleiste")
    # Call PrintOkNok(WS_MINIMIZE And wi.dwStyle, "Fenster ist initial minimiert")
    # Call PrintOkNok(WS_MAXIMIZE And wi.dwStyle, "Fenster ist initial maximiert")
    # Call PrintOkNok(WS_MAXIMIZEBOX And wi.dwStyle, "Fenster hat Maximieren-Button in der Titelleiste")
    # Call PrintOkNok(WS_MINIMIZEBOX And wi.dwStyle, "Fenster hat Minimieren-Button in der Titelleiste")
    # Call PrintOkNok(WS_OVERLAPPED And wi.dwStyle, "berlappendes Fenster. Hat Titelleiste und Rahmen")
    # Call PrintOkNok(WS_popup And wi.dwStyle, "Popupfenster")
    # Call PrintOkNok(WS_SIZEBOX And wi.dwStyle, "Fenster hat grenvernderbaren Rahmen")
    # Call PrintOkNok(WS_SYSMENU And wi.dwStyle, "Fenster hat Windowsmen in der Titelleiste")
    # Call PrintOkNok(WS_VISIBLE And wi.dwStyle, "Fenster ist initial sichtbar")
    # Call PrintOkNok(WS_VSCROLL And wi.dwStyle, "Fenster hat vertikale Scrollleiste")
    # Call PrintOkNok(WS_eZ_clientedge And wi.dwExStyle, "Fenster hat abgesenkte Rahmen")
    # Call PrintOkNok(WS_EZ_CONTEXTHELP And wi.dwExStyle, "Fenster hat Hilfebutton in der Titelleiste")
    # Call PrintOkNok(WS_EZ_DLGMODALFRAME And wi.dwExStyle, "Fenster mit modalem Rahmen")
    # Call PrintOkNok(WS_eZ_noactivate And wi.dwExStyle, "Fenster kommt beim Klicken nicht in den Vordergrund")
    # Call PrintOkNok(WS_EZ_TOPMOST And wi.dwExStyle, "Fenster liegt selbst deaktiviert ber allen anderen")


# '---------------------------------------------------------------------------------------
# ' Method : PrintOkNok
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Prints text with condition and indicator " Not "
# '---------------------------------------------------------------------------------------
def printoknok():

    if OkNok Then:
    print(Debug.Print Text)
    else:
    print(Debug.Print "NOT: " & Text)


# '---------------------------------------------------------------------------------------
# ' Method : Sub demoWindowInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub demoWindowInfo()
# Dim wi As PWINDOWINFO
# Dim IE As Object

# Set IE = CreateObject("InternetExplorer.Application")
# IE.Visible = True

# #If Win64 Then
# wi.cbSize = LenB(wi)
# #Else
# wi.cbSize = Len(wi)
# #End If

if GetWindowInfo(IE.hwnd, wi) Then:
# With wi
print(Debug.Print "Fenstergre:",)
print(Debug.Print (.rcWindow.Right - .rcWindow.Left) _)
# & "*" & (.rcWindow.Bottom - .rcWindow.Top)

if .dwStyle And WINDOWS_STYLES.WS_VISIBLE Then:
print(Debug.Print "Fenster ist sichtbar")
else:
print(Debug.Print "Fenster ist nicht sichtbar")

if .dwStyle And WINDOWS_STYLES.WS_MINIMIZEBOX Then:
print(Debug.Print "Fenster hat einen Minimieren-Button")
# End With

# ProcRet:

