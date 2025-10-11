VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFindWindowParms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' all procs are Z_-Type, used for callback

Private stringAppTitle As String                   ' INPUT
Private handlerAppWindow As Long                   ' OUTPUT

Property Let hwnd(hw As Long)
    handlerAppWindow = hw
End Property                                       ' cFindWindowParms.hWnd Let

Property Get hwnd() As Long
    hwnd = handlerAppWindow
End Property                                       ' cFindWindowParms.hWnd Get

Property Let strTitle(sT As String)
    stringAppTitle = sT
End Property                                       ' cFindWindowParms.strTitle Let

Property Get strTitle() As String
    strTitle = stringAppTitle
End Property                                       ' cFindWindowParms.strTitle Get

