Attribute VB_Name = "WindowControl"
Option Explicit

' Original Module Name: Ol_WindowOps
' (c) 2005 Wayne Phillips (http://www.everythingaccess.com)
' Written 02/06/2005
' http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground

' Custom structure for passing in the parameters in/out of the hook enumeration function

Private Const SW_RESTORE = 9
Private Const SW_Show = 5

Public CountTries As Long

'---------------------------------------------------------------------------------------
' Method : Function EnumWindowProc
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Callback-Function called by GetWindowText (Library Dll)
'---------------------------------------------------------------------------------------
Private Function EnumWindowProc(ByVal hwnd As Long, lParam As cFindWindowParms) As Long
Dim strWindowTitle As String
Const DebugMe As Boolean = False

    With lParam
        .hwnd = hwnd
        strWindowTitle = Space(260) & vbNullChar ' provide buffer
        ' Library GetWindowText calls back EnumWindowProc (addressed via lParam)
        ' to compare each open window, until the window title matches
        CountTries = CountTries + 1
        Call GetWindowText(hwnd, strWindowTitle, 260)
        strWindowTitle = TrimNul(strWindowTitle) ' Remove extra null terminator
        If LenB(strWindowTitle) = 0 Then
            EnumWindowProc = 1
        ElseIf LCase(strWindowTitle) Like LCase(.strTitle) Then
            If DebugMe And LenB(strWindowTitle) > 0 Then
                Debug.Print " Handle " & hwnd & ", Window name " & strWindowTitle & " found"
            End If
            .hwnd = hwnd                         'Store the result for later.
            .strTitle = strWindowTitle           ' this is literally what we found
            If DebugMe Then
                Debug.Print " Handle " & hwnd & ", Window name matched " & .strTitle & " found"
            End If
            EnumWindowProc = 0                   'This will Debug.Assert False enumerating more windows
        Else
            If DebugMe Then
                If InStr(1, strWindowTitle, "outl", vbTextCompare) > 0 _
                And DebugLogging Then
                    DoVerify False
                End If
            End If
            EnumWindowProc = 1                   ' continue loop
        End If
    End With                                     ' lParam
    
End Function                                     ' WindowControl.EnumWindowProc

'---------------------------------------------------------------------------------------
' Method : Function FindWindowLike
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Custom structure for passing in the parameters in/out of the hook enumeration function
'---------------------------------------------------------------------------------------
Function FindWindowLike(strWindowTitle As String, wHdl As cFindWindowParms) As Long
Dim Parameters As cFindWindowParms
    
    CountTries = 0
    If wHdl Is Nothing Then
seekWin:
        Set Parameters = New cFindWindowParms
        With Parameters
            .strTitle = strWindowTitle           ' Input parameter
            ' pass callback routine ---V
            Call EnumWindows(AddressOf EnumWindowProc, VarPtr(Parameters))
            If Not LCase(.strTitle) Like LCase(strWindowTitle) Then
                Parameters.hwnd = 0              ' no Match
                .strTitle = vbNullString
            End If
            FindWindowLike = .hwnd
        End With                                 ' Parameters
        Set wHdl = Parameters
    Else
        If Not LCase(wHdl.strTitle) Like LCase(strWindowTitle) Then
            Set wHdl = Nothing
            GoTo seekWin
        End If
    End If
End Function                                     ' WindowControl.FindWindowLike

'---------------------------------------------------------------------------------------
' Method : Function WindowSetForeground
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Function WindowSetForeground(strWindowTitle As String, wHdl As cFindWindowParms) As Boolean
Dim MyAppHWnd As Long
Dim CurrentForegroundThreadID As Long
Dim NewForegroundThreadID As Long
Dim lngRetVal As Long
Dim blnSuccessful As Boolean
Dim canSkip As Boolean
    
unskip:
    If wHdl Is Nothing Then
        canSkip = True
    Else
        canSkip = False
    End If
    MyAppHWnd = FindWindowLike(strWindowTitle, wHdl)
    
    If MyAppHWnd <> 0 Then
        
        ' We have found the application window by the caption
        CurrentForegroundThreadID = _
                                  GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
        NewForegroundThreadID = GetWindowThreadProcessId(MyAppHWnd, ByVal 0&)
    
        ' AttachThreadInput is used to ensure SetForegroundWindow will work
        ' even if our application isn'W.xlTSheet currently the foreground window
        ' (e.g. an automated app running in the background)
        Call AttachThreadInput(CurrentForegroundThreadID, _
                               NewForegroundThreadID, True)
        lngRetVal = SetForegroundWindow(MyAppHWnd)
        Call AttachThreadInput(CurrentForegroundThreadID, _
                               NewForegroundThreadID, False)
            
        If lngRetVal <> 0 Then
        
            ' Now that the window is active, let's restore it from the taskbar
            If IsIconic(MyAppHWnd) Then
                Call ShowWindow(MyAppHWnd, SW_RESTORE)
            Else
                Call ShowWindow(MyAppHWnd, SW_Show)
            End If
            
            blnSuccessful = True
            wHdl.hwnd = MyAppHWnd
            If DebugLogging Then
                'msgbox "Found the window " & Quote(strWindowTitle) _
                & "  and it should be in foreground"
            End If
        Else
            If DebugLogging Then
                MsgBox "Found the window " & Quote(strWindowTitle) _
      & ", but failed to bring it to the foreground!"
            Else
                Debug.Print "Found the window " & Quote(strWindowTitle) _
      & ", but failed to bring it to the foreground!"
            End If
        End If
        If DebugLogging Then
            Call IsModalWindow(MyAppHWnd)
            DebugLogging = True
        End If
    Else
        'Failed to find the window caption
        'Therefore the app is probably closed.
        Set wHdl = Nothing
        If Not canSkip Then                      ' window handle could have changed
            GoTo unskip                          ' locate again
        End If
        MsgBox "Application Window " & Quote(strWindowTitle) & " not found!"
    End If
    
    WindowSetForeground = blnSuccessful
    
ProcRet:
End Function                                     ' WindowControl.WindowSetForeground

'---------------------------------------------------------------------------------------
' Method : IsModalWindow
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: True if Modal
' Source : https://devblogs.microsoft.com/oldnewthing/20111212-00/?p=8923
' Author : Raymond Chen
'---------------------------------------------------------------------------------------
Function IsModalWindow(hwnd As Long) As Boolean
Const zKey As String = "WindowControl.IsModalWindow"

Dim hwndOwnerl As LongPtr
Dim wi As PWINDOWINFO
Dim wo As PWINDOWINFO
Dim ws As WINDOWS_STYLES

    #If Win64 Then
Dim lngHwnd As LongPtr
Dim hwndOwner As LongLong
        wi.cbSize = LenB(wi)
        wo.cbSize = LenB(wo)
    #Else
Dim lngHwnd As Long
Dim hwndOwner As Long
        wi.cbSize = Len(wi)
        wo.cbSize = Len(wo)
    #End If

    lngHwnd = hwnd
    If GetWindowInfo(lngHwnd, wi) Then           ' child windows cannot have owners
        Call ShowWindowStates(lngHwnd, wi)
        hwndOwner = GetWindow(lngHwnd, GW_OWNER)
        If hwndOwner Then                        ' has owner window
            If GetWindowInfo(hwndOwner, wo) Then ' owner is enabled
                Call ShowWindowStates(hwndOwner, wo)
                If Not (wo.dwWindowStatus And WS_DISABLED) Then
                    IsModalWindow = True         ' owner is disabled: Modal ==> True
                End If
            End If
        End If
    End If

zExit:

ProcRet:
End Function                                     ' WindowControl.IsModalWindow

'---------------------------------------------------------------------------------------
' Method : ShowWindowStates
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: ...
'---------------------------------------------------------------------------------------
Sub ShowWindowStates(lngHwnd As LongPtr, wi As PWINDOWINFO)

Const zKey As String = "WindowControl.ShowWindowStates"

Dim wiWinTitle As String

    wiWinTitle = Space(260) & vbNullChar         ' provide buffer
    Call GetWindowText(lngHwnd, wiWinTitle, 260)
    wiWinTitle = TrimNul(wiWinTitle)             ' Remove extra null terminator
    Debug.Print vbCrLf, "Window data for " & wiWinTitle
    Debug.Print String(80, "=")
        
    Call PrintOkNok(WS_BORDER And wi.dwStyle, "Fenster hat einen Rahmen")
    Call PrintOkNok(WS_CAPTION And wi.dwStyle, "Fenster hat Titelleiste")
    Call PrintOkNok(WS_CHILD And wi.dwStyle, "Kindfenster. Kann kein Menü haben und nicht Popup")
    Call PrintOkNok(WS_CLIPCHILDREN And wi.dwStyle, "Verhindert Kindfenster außerhalb des Elternfensters")
    Call PrintOkNok(WS_CLIPSIBLINGS And wi.dwStyle, "Verhindert sich gegenseitig überlappende Kindfenster")
    Call PrintOkNok(WS_DISABLED And wi.dwStyle, "Fenster ist initial Disabled")
    Call PrintOkNok(WS_DLGFRAME And wi.dwStyle, "Dialogfenster")
    Call PrintOkNok(WS_HSCROLL And wi.dwStyle, "Fenster hat horizontale Scrollleiste")
    Call PrintOkNok(WS_MINIMIZE And wi.dwStyle, "Fenster ist initial minimiert")
    Call PrintOkNok(WS_MAXIMIZE And wi.dwStyle, "Fenster ist initial maximiert")
    Call PrintOkNok(WS_MAXIMIZEBOX And wi.dwStyle, "Fenster hat Maximieren-Button in der Titelleiste")
    Call PrintOkNok(WS_MINIMIZEBOX And wi.dwStyle, "Fenster hat Minimieren-Button in der Titelleiste")
    Call PrintOkNok(WS_OVERLAPPED And wi.dwStyle, "überlappendes Fenster. Hat Titelleiste und Rahmen")
    Call PrintOkNok(WS_popup And wi.dwStyle, "Popupfenster")
    Call PrintOkNok(WS_SIZEBOX And wi.dwStyle, "Fenster hat größenveränderbaren Rahmen")
    Call PrintOkNok(WS_SYSMENU And wi.dwStyle, "Fenster hat Windowsmenü in der Titelleiste")
    Call PrintOkNok(WS_VISIBLE And wi.dwStyle, "Fenster ist initial sichtbar")
    Call PrintOkNok(WS_VSCROLL And wi.dwStyle, "Fenster hat vertikale Scrollleiste")
    Call PrintOkNok(WS_eZ_clientedge And wi.dwExStyle, "Fenster hat abgesenkte Rahmen")
    Call PrintOkNok(WS_EZ_CONTEXTHELP And wi.dwExStyle, "Fenster hat Hilfebutton in der Titelleiste")
    Call PrintOkNok(WS_EZ_DLGMODALFRAME And wi.dwExStyle, "Fenster mit modalem Rahmen")
    Call PrintOkNok(WS_eZ_noactivate And wi.dwExStyle, "Fenster kommt beim Klicken nicht in den Vordergrund")
    Call PrintOkNok(WS_EZ_TOPMOST And wi.dwExStyle, "Fenster liegt selbst deaktiviert über allen anderen")

End Sub                                          ' WindowControl.ShowWindowStates

'---------------------------------------------------------------------------------------
' Method : PrintOkNok
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Prints text with condition and indicator " Not "
'---------------------------------------------------------------------------------------
Sub PrintOkNok(OkNok As Boolean, Text As String)

    If OkNok Then
        Debug.Print Text
    Else
        Debug.Print "NOT: " & Text
    End If

End Sub                                          ' WindowControl.PrintOkNok

'---------------------------------------------------------------------------------------
' Method : Sub demoWindowInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub demoWindowInfo()
Dim wi As PWINDOWINFO
Dim IE As Object

    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True

    #If Win64 Then
        wi.cbSize = LenB(wi)
    #Else
        wi.cbSize = Len(wi)
    #End If

    If GetWindowInfo(IE.hwnd, wi) Then
        With wi
            Debug.Print "Fenstergröße:",
            Debug.Print (.rcWindow.Right - .rcWindow.Left) _
      & "*" & (.rcWindow.Bottom - .rcWindow.Top)
    
            If .dwStyle And WINDOWS_STYLES.WS_VISIBLE Then
                Debug.Print "Fenster ist sichtbar"
            Else
                Debug.Print "Fenster ist nicht sichtbar"
            End If
    
            If .dwStyle And WINDOWS_STYLES.WS_MINIMIZEBOX Then
                Debug.Print "Fenster hat einen Minimieren-Button"
            End If
        End With
    End If

ProcRet:
End Sub                                          ' WindowControl.demoWindowInfo


