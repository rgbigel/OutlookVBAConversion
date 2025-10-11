VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents EventOl_WEBInItems As Outlook.Items
Attribute EventOl_WEBInItems.VB_VarHelpID = -1
Public WithEvents EventOl_WEBSeItems As Outlook.Items
Attribute EventOl_WEBSeItems.VB_VarHelpID = -1
Public WithEvents EventOl_HotInItems As Outlook.Items
Attribute EventOl_HotInItems.VB_VarHelpID = -1
Public WithEvents EventOl_HotSeItems1 As Outlook.Items
Attribute EventOl_HotSeItems1.VB_VarHelpID = -1
Public WithEvents EventOl_HotSeItems2 As Outlook.Items
Attribute EventOl_HotSeItems2.VB_VarHelpID = -1
Public WithEvents EventOl_GooInItems As Outlook.Items
Attribute EventOl_GooInItems.VB_VarHelpID = -1
Public WithEvents EventOl_GooSeItems As Outlook.Items
Attribute EventOl_GooSeItems.VB_VarHelpID = -1
Public WithEvents EventOl_BackupHomeInItems As Outlook.Items
Attribute EventOl_BackupHomeInItems.VB_VarHelpID = -1
Public WithEvents EventOl_BackupHomeSeItems As Outlook.Items
Attribute EventOl_BackupHomeSeItems.VB_VarHelpID = -1
Public WithEvents EventOl_NewMail As Outlook.Items
Attribute EventOl_NewMail.VB_VarHelpID = -1

Public WithEvents EventOl_NewTask As Outlook.Items
Attribute EventOl_NewTask.VB_VarHelpID = -1
Public WithEvents EventOl_Contacts As Outlook.Items
Attribute EventOl_Contacts.VB_VarHelpID = -1
Public WithEvents EventOl_Calendar As Outlook.Items
Attribute EventOl_Calendar.VB_VarHelpID = -1

Public WithEvents myOlExplorers As Outlook.Explorers
Attribute myOlExplorers.VB_VarHelpID = -1
Public WithEvents objReminders As Outlook.Reminders
Attribute objReminders.VB_VarHelpID = -1
Public WithEvents olInsp As Inspectors
Attribute olInsp.VB_VarHelpID = -1
' no such Event Public WithEvents AdvancedSearchComplete As Outlook.Search

'---------------------------------------------------------------------------------------
' Method : StartMainApp
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Any externally called Macro (Entry Point) must use this as the first call
'---------------------------------------------------------------------------------------
Public Sub StartMainApp()
Dim zErr As cErr
Const zKey As String = "cOutlookSession.StartMainApp"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, eQAsMode, tSub)
    
ReStartProc:
    DebugSleep = DefaultDebugSleep
    Call Init_Item_Handlers
    
    IsEntryPoint = False

ProcReturn:
    Call ProcExit(zErr)
    UseStartUp = 2
pExit:
End Sub                                            ' cOutlookSession.StartMainApp

' from https://superuser.com/questions/251963/how-to-make-outlook-calendar-reminders-stay-on-top-in-windows-7
'---------------------------------------------------------------------------------------
' Method : Appl_Reminder
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Make sure reminder window is on top. Event Routine (not really Public)
'---------------------------------------------------------------------------------------
Public Sub Appl_Reminder(ByVal Item As Object)     ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.Appl_Reminder"
Dim ReminderWindowHWnd As Variant

'------------------- gated Entry -------------------------------------------------------
    If E_AppErr.EventBlock Then
        If DebugLogging Then Call LogEvent("*** Event blocked " & zKey, eLmin)
        GoTo pExit
    End If
    
    If DebugLogging Then Call LogEvent("*** Event " & zKey & b & Item.Subject, eLmin)
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    ReminderWindowHWnd = FindWindow(vbNullString, " Reminder")
    SetWindowPos ReminderWindowHWnd, HW_TOPMOST, 0, 0, 0, 0, Flags

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.Appl_Reminder

'---------------------------------------------------------------------------------------
' Method : AdvSearchDone
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: called when advanced search finished
'---------------------------------------------------------------------------------------
Sub AdvSearchDone(ByVal SearchObject As Search)    ' *** Entry Point ***

Const zKey As String = "cOutlookSession.AdvSearchDone"

' *** this is NOT GATED, Event never Blocked ***
    Call DoCall(zKey, tSub, eQzMode)
     
    Const MyId As String = "AdvSearchDone"
        
    If DebugLogging Then
        Call LogEvent("*** Event " & zKey & SearchObject.Tag & " in " & SearchObject.Scope, eLmin)
    End If
    
    If SearchObject.Tag = SpecialSearchFolderName Then
        SpecialSearchComplete = True
    End If

zExit:
    Call DoExit(zKey)

End Sub                                            ' cOutlookSession.AdvSearchDone

'---------------------------------------------------------------------------------------
' Method : Appl_ItemSend
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event (nop)
'---------------------------------------------------------------------------------------
Sub Appl_ItemSend(ByVal Item As Object, Cancel As Boolean) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.Appl_ItemSend"
'------------------- gated Entry -------------------------------------------------------
    If E_AppErr.EventBlock Then
        If DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLmin)
        GoTo pExit
    End If
    
    If DebugLogging Then Call LogEvent("*** Event ignored " & zKey & b & Item.Subject, eLmin)
    GoTo pExit
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.Appl_ItemSend

'---------------------------------------------------------------------------------------
' Method : Appl_NewMail
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event
'---------------------------------------------------------------------------------------
Sub Appl_NewMail()                                 ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.Appl_NewMail"
'------------------- gated Entry -------------------------------------------------------
    If E_AppErr.EventBlock Then
        If DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLdebug)
        GoTo pExit
    End If
    
    Call LogEvent(">> Event " & zKey & " will be processed", eLmin)
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    Call CollectItemsToLog(-1)                  ' do all eligible folders
    Call FldActions2Do                          ' we (must) have (at least 1) open items, do em now
    
ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                         ' cOutlookSession.Appl_NewMail

'---------------------------------------------------------------------------------------
' Method : Init_Item_Handlers
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Init all Item handlers (not all are needed)
'---------------------------------------------------------------------------------------
Public Sub Init_Item_Handlers()
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "cOutlookSession.Init_Item_Handlers"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    ' Set EventOl_Calendar = aNamespace.GetDefaultFolder(olFolderCalendar).Items
    ' Set EventOl_Contacts = aNameSpace.GetDefaultFolder(olFolderContacts).Items
    Set EventOl_NewTask = aNameSpace.GetDefaultFolder(olFolderTasks).Items

ProcReturn:
    Call ProcExit(zErr)

End Sub                                            ' cOutlookSession.Init_Item_Handlers

'---------------------------------------------------------------------------------------
' Method : App_ReInit
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Treated like Event Routine (but there is no such Event in Outlook)
'---------------------------------------------------------------------------------------
Public Sub App_ReInit()
Dim zErr As cErr
Const zKey As String = "cOutlookSession.App_ReInit"
    Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSub, ExplainS:="ReInit-Call")

    Debug.Print "ReInit-Call of Outlook-Application (Application)!"
    If DebugSleep = 0 And Not (DebugLogging Or DebugMode) Then
        DebugSleep = DefaultDebugSleep
    End If
    If DebugLogging Then
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    IsEntryPoint = False

ProcReturn:
    Call ProcExit(zErr)
End Sub                                            ' cOutlookSession.App_ReInit

'---------------------------------------------------------------------------------------
' Method : EventOl_Calendar_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine (nop)
'---------------------------------------------------------------------------------------
Sub EventOl_Calendar_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_Calendar_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_AppErr.EventBlock Then
        If DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLmin)
        GoTo pExit
    End If
    
    If DebugLogging Then Call LogEvent("*** Event ignored " & zKey, eLmin)
    GoTo pExit
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_Calendar_ItemAdd

'---------------------------------------------------------------------------------------
' Method : Sub EventOl_Calendar_ItemChange
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub EventOl_Calendar_ItemChange(ByVal Item As Object)
    Stop
End Sub

'---------------------------------------------------------------------------------------
' Method : Sub EventOl_Calendar_ItemRemove
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub EventOl_Calendar_ItemRemove()
    Stop
End Sub

'---------------------------------------------------------------------------------------
' Method : EventOl_Contacts_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_Contacts_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_Contacts_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_AppErr.EventBlock Then
        If DebugLogging Then Call LogEvent("*** Event blocked " & zKey, eLmin)
        GoTo pExit
    End If
    
    If DebugLogging Then Call LogEvent("*** Event ignored " & zKey & b & Item.Subject, eLmin)
    GoTo pExit
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_Contacts_ItemAdd

'---------------------------------------------------------------------------------------
' Method : Sub EventOl_Contacts_ItemChange
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub EventOl_Contacts_ItemChange(ByVal Item As Object)
    Stop
End Sub

'---------------------------------------------------------------------------------------
' Method : EventOl_NewTask_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_NewTask_ItemAdd(ByVal Item As Object)  ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_NewTask_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    AddType = "Task (plain)"
    Call commonTask(Item)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_NewTask_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_NewTaskRequestAcceptItem_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_NewTaskRequestAcceptItem_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestAcceptItem_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    AddType = "TaskRequestAcceptItem"
    Call commonTask(Item)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_NewTaskRequestAcceptItem_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_NewTaskRequestDeclineItem_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_NewTaskRequestDeclineItem_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestDeclineItem_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    AddType = "TaskRequestDeclineItem"
    Call commonTask(Item)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_NewTaskRequestDeclineItem_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_NewTaskRequest_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_NewTaskRequest_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_NewTaskRequest_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    AddType = "TaskRequest"
    Call commonTask(Item)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_NewTaskRequest_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_NewTaskRequestUpdateItem_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_NewTaskRequestUpdateItem_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestUpdateItem_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    AddType = "TaskRequestUpdateItem"
    Call commonTask(Item)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_NewTaskRequestUpdateItem_ItemAdd

'---------------------------------------------------------------------------------------
' Method : commonTask
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Common Event Task
'---------------------------------------------------------------------------------------
Sub commonTask(ByVal Item As Object)
Dim zErr As cErr
Const zKey As String = "cOutlookSession.commonTask"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    Debug.Print "new " & AddType & " arrived: " & Item.Subject

ProcReturn:
    Call ProcExit(zErr)
End Sub                                            ' cOutlookSession.commonTask

'---------------------------------------------------------------------------------------
' Method : Sub EventOl_NewTask_ItemRemove
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub EventOl_NewTask_ItemRemove()
    Stop
End Sub

'---------------------------------------------------------------------------------------
' Method : EventOl_WEBInItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_WEBInItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_WEBInItems_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
    
    Debug.Print zKey & "-Event of Outlook occurred!"
    
    Debug.Print "'NewMail' Event arrived for WEB: " & Now() _
        & b & Quote(Item.Subject) _
        & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    If DebugLogging Then
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    
    Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_WEBInItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : Sub EventOl_WEBInItems_ItemChange
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub EventOl_WEBInItems_ItemChange(ByVal Item As Object)
'Stop
End Sub

'---------------------------------------------------------------------------------------
' Method : EventOl_WEBSeItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_WEBSeItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_WEBSeItems_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call StartEP(zErr, zKey, tSubEP, eQEPMode)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    
    Debug.Print "'ItemAdd' Event sent for WEB: " & Now() _
        & b & Quote(Item.Subject) _
        & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    If DebugLogging Then
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    
    Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_WEBSeItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_HotInItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_HotInItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_HotInItems_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)
   
    Debug.Print zKey & "-Event of Outlook occurred!"
    
    Debug.Print "'ItemAdd' Event arrived for Outlook: " & Now() _
        & b & Quote(Item.Subject) _
        & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    If DebugLogging Then
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    
    Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_HotInItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_HotSeItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_HotSeItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
'''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cOutlookSession.EventOl_HotSeItems_ItemAdd"
Static zErr As New cErr

    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    '------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Debug.Print "'ItemAdd' Event sent for outlook: " & Now() _
        & b & Quote(Item.Subject) _
        & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    If DebugLogging Then
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    
    Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_HotSeItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_BackupHomeInItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
' Note   : currently, this event is not activated, so it will not be called or working
'---------------------------------------------------------------------------------------
Sub EventOl_BackupHomeInItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_BackupHomeInItems_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)
    
    If DebugLogging Then
        Debug.Print "'ItemAdd' Event arrived and IGNORED actions, item in Folder " _
                    & Item.Parent.FullFolderPath & ": " & Now() _
                    & vbCrLf & vbTab & Quote(Item.Subject) _
                    & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If

    If EventHappened Then
        If Not NoEventOnAddItem Then
            If DebugLogging Then
                Debug.Print "NoEventOnAddItem = True"
                Call Sleep(DebugSleep)             ' wait here with modal box open
            End If
        End If
    End If

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_BackupHomeInItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : EventOl_BackupHomeSeItems_ItemAdd
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Event Routine
'---------------------------------------------------------------------------------------
Sub EventOl_BackupHomeSeItems_ItemAdd(ByVal Item As Object) ' *** Entry Point ***
Dim zErr As cErr
Const zKey As String = "cOutlookSession.EventOl_BackupHomeSeItems_ItemAdd"
'------------------- gated Entry -------------------------------------------------------
    If E_Active.EventBlock Then
        GoTo pExit
    End If
    
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)
    
    If DebugLogging Then
        Debug.Print "'ItemAdd' Event sent and IGNORED actions for " _
                    & Item.Parent.FullFolderPath & ": " & Now() _
                    & vbCrLf & vbTab & Quote(Item.Subject) _
                    & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
        Call Sleep(DebugSleep)                     ' wait here with modal box open
    End If
    
    ' DEACTIVATED ??? Call DeferredActionAdd(Item, curAction:=3)

ProcReturn:
    Call ProcExit(zErr)
pExit:
End Sub                                            ' cOutlookSession.EventOl_BackupHomeSeItems_ItemAdd

'---------------------------------------------------------------------------------------
' Method : Sub Init_Evnt_Handlers
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub Init_Evnt_Handlers()
Dim zErr As cErr
Const zKey As String = "cOutlookSession.Init_Evnt_Handlers"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    Set myOlExplorers = olApp.Explorers
    Set objReminders = olApp.Reminders
    Set olInsp = olApp.Inspectors

ProcReturn:
    Call ProcExit(zErr)
End Sub                                            ' cOutlookSession.Init_Evnt_Handlers

'---------------------------------------------------------------------------------------
' Method : App_Test
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: .
'---------------------------------------------------------------------------------------
Public Sub App_Test()
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "cOutlookSession.App_Test"
Dim zErr As cErr

Dim i As Long
Dim aExplorer As Explorer
    
    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    If ActiveExplorer Is Nothing Then
        Debug.Print "no active explorers ... is this the end?"
    Else
        Debug.Print "there are " & Explorers.Count & " explorers open now"
    End If

    For i = 1 To Explorers.Count
        Set aExplorer = Explorers.Item(i)
        Debug.Print " look at me"
        aExplorer.Close
    
        Debug.Assert False
    Next i
    If Not xlApp Is Nothing Then
        Call xlEndApp
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                            ' cOutlookSession.App_Test


