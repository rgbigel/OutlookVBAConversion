VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Sub Application_Startup
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_Startup()                  ' *** Entry Point ***
Const thisIsOutlook As String = "Outlook"          ' unfortunately, can't make this Public Const

Dim CallingApp As Object
    
    Call BugTimer.BugState_SetPause
        
    If Application.Name <> thisIsOutlook Then
        Debug.Assert False
        On Error Resume Next
        Set CallingApp = GetObject(, "Excel.Application") ' returns error 429 if excel not running
        If Err.Number = 0 Then
            If Excel.Application.ActiveSheet.Name = thisIsOutlook Then
                Debug.Print "Do not run Outlook Macros here, all should be done in Excel"
                TheMasterIsExcel = True
                Exit Sub
            End If
        End If
    End If
    
    Call N_CheckStartSession

End Sub                                            ' ThisOutlookSession.Application_Startup

'---------------------------------------------------------------------------------------
' Method : Sub Application_Quit
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_Quit()                     ' *** Entry Point ***
    Debug.Print "Application_Quit has been called"
    Call AllPublic.Quit
End Sub                                            ' ThisOutlookSession.Application_Quit

'---------------------------------------------------------------------------------------
' Method : Sub Application_Reminder
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_Reminder(ByVal Item As Object) ' *** Entry Point ***
    If OlSession Is Nothing Then
        Call N_CheckStartSession
    End If
    Call BugTimer.BugState_SetPause
    Call OlSession.Appl_Reminder(Item)
    Call BugTimer.BugState_UnPause
End Sub                                            ' ThisOutlookSession.Application_Reminder

'---------------------------------------------------------------------------------------
' Method : Sub Application_AdvancedSearchComplete
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) ' *** Entry Point ***
    If OlSession Is Nothing Then
        Call N_CheckStartSession
    End If
    Call BugTimer.BugState_SetPause
    Call OlSession.AdvSearchDone(SearchObject)
    Call BugTimer.BugState_UnPause
End Sub                                            ' ThisOutlookSession.Application_AdvancedSearchComplete

'---------------------------------------------------------------------------------------
' Method : Sub Application_ItemSend
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean) ' *** Entry Point ***
    If OlSession Is Nothing Then
        Call N_CheckStartSession
    End If
    Call BugTimer.BugState_SetPause
    Call OlSession.StartMainApp
    Call BugTimer.BugState_UnPause
End Sub                                            ' ThisOutlookSession.Application_ItemSend

'---------------------------------------------------------------------------------------
' Method : Sub Application_NewMail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Application_NewMail()                  ' *** Entry Point ***
    If OlSession Is Nothing Then
        Call N_CheckStartSession
    End If
    EventHappened = True
    Call BugTimer.BugState_SetPause
    Call OlSession.Appl_NewMail
    Call BugTimer.BugState_UnPause
End Sub                                            ' ThisOutlookSession.Application_NewMail

