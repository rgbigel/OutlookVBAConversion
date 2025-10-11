Attribute VB_Name = "SystemIFs"

#If VBA7 Then
Public Const isVba7 As Boolean = True
#Else
Public Const isVba7 As Boolean = False
#End If
#If Win64 Then
Public Const isWin64 As Boolean = True
#Else
Public Const isWin64 As Boolean = False
#End If

' Declare System DLLs
' see also https://www.cadsharp.com/docs/Win32API_PtrSafe.txt
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const Flags As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HW_TOPMOST = -1
Public Const GWL_STYLE As Long = -16
Public Const GA_PARENT As Integer = 1

Public Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Public Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare PtrSafe Function AddFontMemResourceEx Lib "Gdi32.dll" (ByVal pbFont As LongPtr, ByVal cbFont As Integer, ByVal pdv As Integer, ByRef pcFonts As Integer) As LongPtr
Public Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As LongPtr
Public Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Public Declare PtrSafe Function GetWindowInfo Lib "user32" (ByVal hwnd As LongPtr, ByRef pwi As PWINDOWINFO) As Boolean
Public Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr) As Long
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Public Declare PtrSafe Function MoveWindow Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal x As LongPtr, ByVal y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal bRepaint As LongPtr) As LongPtr
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ReleaseDC Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
Public Declare PtrSafe Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFilename As String) As LongPtr
Public Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As LongPtr
Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongLong, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
Public Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLongPtr As LongPtr) As LongPtr
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PWINDOWINFO
    cbSize As Long                               'Strukturgröße
    rcWindow As RECT                             'Fensterkoordinaten
    rcClient As RECT                             'Clientkoordinaten
    dwStyle As Long                              'Gibt WINDOWS_STYLES zurück
    dwExStyle As Long                            'Gibt EXTENDED_WINDOWS_STYLES zurück
    dwWindowStatus As Long                       '1=Fenster ist aktives Fenster, 0=Fenster nicht aktiv
    cxWindowBorders As Long                      'Rahmenbreite in px
    cyWindowBorders As Long                      'Rahmenhöhe in px
    atomWindowType As Integer
    wCreatorVersion As Integer
End Type

Enum WINDOWS_STYLES
    WS_BORDER = &H800000                         'Fenster hat einen Rahmen
    WS_CAPTION = &HC00000                        'Fenster hat Titelleiste
    WS_CHILD = &H40000000                        'Kindfenster. Kann kein Menü haben und nicht Popup
    WS_CLIPCHILDREN = &H2000000                  'Verhindert Kindfenster außerhalb des Elternfensters
    WS_CLIPSIBLINGS = &H4000000                  'Verhindert sich gegenseitig überlappende Kindfenster
    WS_DISABLED = &H8000000                      'Fenster ist initial Disabled
    WS_DLGFRAME = &H400000                       'Dialogfenster
    WS_HSCROLL = &H100000                        'Fenster hat horizontale Scrollleiste
    WS_MINIMIZE = &H20000000                     'Fenster ist initial minimiert
    WS_MAXIMIZE = &H1000000                      'Fenster ist initial maximiert
    WS_MAXIMIZEBOX = &H10000                     'Fenster hat Maximieren-Button in der Titelleiste
    WS_MINIMIZEBOX = &H20000                     'Fenster hat Minimieren-Button in der Titelleiste
    WS_OVERLAPPED = &H0                          'überlappendes Fenster. Hat Titelleiste und Rahmen
    WS_popup = &H80000000                        'Popupfenster
    WS_SIZEBOX = &H40000                         'Fenster hat größenveränderbaren Rahmen
    WS_SYSMENU = &H80000                         'Fenster hat Windowsmenü in der Titelleiste
    WS_VISIBLE = &H10000000                      'Fenster ist initial sichtbar
    WS_VSCROLL = &H200000                        'Fenster hat vertikale Scrollleiste
End Enum

'EXTENDED_WINDOWS_STYLES ist stark unvollständig, da in VBA kaum Bedarf dafür besteht
Enum EXTENDED_WINDOWS_STYLES
    WS_eZ_clientedge = &H200                     'Fenster hat abgesenkte Rahmen
    WS_EZ_CONTEXTHELP = &H400                    'Fenster hat Hilfebutton in der Titelleiste
    WS_EZ_DLGMODALFRAME = &H1                    'Fenster mit modalem Rahmen
    WS_eZ_noactivate = &H8000000                 'Fenster kommt beim Klicken nicht in den Vordergrund
    WS_EZ_TOPMOST = &H8                          'Fenster liegt selbst deaktiviert über allen anderen
End Enum

Public Enum eCmd
    GW_FINDFIRST = 0                             'Erstes Fenster gleicher Ebene
    GW_FINDLAST = 1                              'Letztes Fenster gleicher Ebene
    GW_HWNDNEXT = 2                              'Nächstes Fenster gleicher Ebene
    GW_HWNDPREV = 3                              'Letztes Fenster gleicher Ebene
    GW_OWNER = 4                                 'Elternfenster
    GW_CHILD = 5                                 'Erstes Kindfenster
End Enum




