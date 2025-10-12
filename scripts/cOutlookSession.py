# Converted from cOutlookSession.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cOutlookSession"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public WithEvents EventOl_WEBInItems As Outlook.Items
# Attribute EventOl_WEBInItems.VB_VarHelpID = -1
# Public WithEvents EventOl_WEBSeItems As Outlook.Items
# Attribute EventOl_WEBSeItems.VB_VarHelpID = -1
# Public WithEvents EventOl_HotInItems As Outlook.Items
# Attribute EventOl_HotInItems.VB_VarHelpID = -1
# Public WithEvents EventOl_HotSeItems1 As Outlook.Items
# Attribute EventOl_HotSeItems1.VB_VarHelpID = -1
# Public WithEvents EventOl_HotSeItems2 As Outlook.Items
# Attribute EventOl_HotSeItems2.VB_VarHelpID = -1
# Public WithEvents EventOl_GooInItems As Outlook.Items
# Attribute EventOl_GooInItems.VB_VarHelpID = -1
# Public WithEvents EventOl_GooSeItems As Outlook.Items
# Attribute EventOl_GooSeItems.VB_VarHelpID = -1
# Public WithEvents EventOl_BackupHomeInItems As Outlook.Items
# Attribute EventOl_BackupHomeInItems.VB_VarHelpID = -1
# Public WithEvents EventOl_BackupHomeSeItems As Outlook.Items
# Attribute EventOl_BackupHomeSeItems.VB_VarHelpID = -1
# Public WithEvents EventOl_NewMail As Outlook.Items
# Attribute EventOl_NewMail.VB_VarHelpID = -1

# Public WithEvents EventOl_NewTask As Outlook.Items
# Attribute EventOl_NewTask.VB_VarHelpID = -1
# Public WithEvents EventOl_Contacts As Outlook.Items
# Attribute EventOl_Contacts.VB_VarHelpID = -1
# Public WithEvents EventOl_Calendar As Outlook.Items
# Attribute EventOl_Calendar.VB_VarHelpID = -1

# Public WithEvents myOlExplorers As Outlook.Explorers
# Attribute myOlExplorers.VB_VarHelpID = -1
# Public WithEvents objReminders As Outlook.Reminders
# Attribute objReminders.VB_VarHelpID = -1
# Public WithEvents olInsp As Inspectors
# Attribute olInsp.VB_VarHelpID = -1
# ' no such Event Public WithEvents AdvancedSearchComplete As Outlook.Search

# '---------------------------------------------------------------------------------------
# ' Method : StartMainApp
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Any externally called Macro (Entry Point) must use this as the first call
# '---------------------------------------------------------------------------------------
# Public Sub StartMainApp()
# Dim zErr As cErr
# Const zKey As String = "cOutlookSession.StartMainApp"
# '------------------- gated Entry -------------------------------------------------------
if E_Active.EventBlock Then:
# GoTo pExit

# Call ProcCall(zErr, zKey, eQAsMode, tSub)

# ReStartProc:
# DebugSleep = DefaultDebugSleep
# Call Init_Item_Handlers

# IsEntryPoint = False

# ProcReturn:
# Call ProcExit(zErr)
# UseStartUp = 2
# pExit:

# ' from https://superuser.com/questions/251963/how-to-make-outlook-calendar-reminders-stay-on-top-in-windows-7
# '---------------------------------------------------------------------------------------
# ' Method : Appl_Reminder
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Make sure reminder window is on top. Event Routine (not really Public)
# '---------------------------------------------------------------------------------------
# Public Sub Appl_Reminder(ByVal Item As Object)     ' *** Entry Point ***
# Dim zErr As cErr
# Const zKey As String = "cOutlookSession.Appl_Reminder"
# Dim ReminderWindowHWnd As Variant

# '------------------- gated Entry -------------------------------------------------------
if E_AppErr.EventBlock Then:
if DebugLogging Then Call LogEvent("*** Event blocked " & zKey, eLmin):
# GoTo pExit

if DebugLogging Then Call LogEvent("*** Event " & zKey & b & Item.Subject, eLmin):
# Call StartEP(zErr, zKey, tSubEP, eQEPMode)

# ReminderWindowHWnd = FindWindow(vbNullString, " Reminder")
# SetWindowPos ReminderWindowHWnd, HW_TOPMOST, 0, 0, 0, 0, Flags

# ProcReturn:
# Call ProcExit(zErr)
# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : AdvSearchDone
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: called when advanced search finished
# '---------------------------------------------------------------------------------------
def advsearchdone():

    # Const zKey As String = "cOutlookSession.AdvSearchDone"

    # ' *** this is NOT GATED, Event never Blocked ***
    # Call DoCall(zKey, tSub, eQzMode)

    # Const MyId As String = "AdvSearchDone"

    if DebugLogging Then:
    # Call LogEvent("*** Event " & zKey & SearchObject.Tag & " in " & SearchObject.Scope, eLmin)

    if SearchObject.Tag = SpecialSearchFolderName Then:
    # SpecialSearchComplete = True

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Appl_ItemSend
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event (nop)
# '---------------------------------------------------------------------------------------
def appl_itemsend():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.Appl_ItemSend"
    # '------------------- gated Entry -------------------------------------------------------
    if E_AppErr.EventBlock Then:
    if DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLmin):
    # GoTo pExit

    if DebugLogging Then Call LogEvent("*** Event ignored " & zKey & b & Item.Subject, eLmin):
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Appl_NewMail
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event
# '---------------------------------------------------------------------------------------
def appl_newmail():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.Appl_NewMail"
    # '------------------- gated Entry -------------------------------------------------------
    if E_AppErr.EventBlock Then:
    if DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLdebug):
    # GoTo pExit

    # Call LogEvent(">> Event " & zKey & " will be processed", eLmin)
    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    # Call CollectItemsToLog(-1)                  ' do all eligible folders
    # Call FldActions2Do                          ' we (must) have (at least 1) open items, do em now

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Init_Item_Handlers
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Init all Item handlers (not all are needed)
# '---------------------------------------------------------------------------------------
# Public Sub Init_Item_Handlers()
# '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
# Const zKey As String = "cOutlookSession.Init_Item_Handlers"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

# ' Set EventOl_Calendar = aNamespace.GetDefaultFolder(olFolderCalendar).Items
# ' Set EventOl_Contacts = aNameSpace.GetDefaultFolder(olFolderContacts).Items
# Set EventOl_NewTask = aNameSpace.GetDefaultFolder(olFolderTasks).Items

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : App_ReInit
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Treated like Event Routine (but there is no such Event in Outlook)
# '---------------------------------------------------------------------------------------
# Public Sub App_ReInit()
# Dim zErr As cErr
# Const zKey As String = "cOutlookSession.App_ReInit"
# Call ProcCall(zErr, zKey, Qmode:=eQEPMode, CallType:=tSub, ExplainS:="ReInit-Call")

print(Debug.Print "ReInit-Call of Outlook-Application (Application)!")
if DebugSleep = 0 And Not (DebugLogging Or DebugMode) Then:
# DebugSleep = DefaultDebugSleep
if DebugLogging Then:
# Call Sleep(DebugSleep)                     ' wait here with modal box open
# IsEntryPoint = False

# ProcReturn:
# Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_Calendar_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine (nop)
# '---------------------------------------------------------------------------------------
def eventol_calendar_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_Calendar_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_AppErr.EventBlock Then:
    if DebugLogging Then Call LogEvent("*** Event blocked in Appl " & zKey, eLmin):
    # GoTo pExit

    if DebugLogging Then Call LogEvent("*** Event ignored " & zKey, eLmin):
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EventOl_Calendar_ItemChange
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EventOl_Calendar_ItemChange(ByVal Item As Object)
# Stop

# '---------------------------------------------------------------------------------------
# ' Method : Sub EventOl_Calendar_ItemRemove
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EventOl_Calendar_ItemRemove()
# Stop

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_Contacts_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_contacts_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_Contacts_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_AppErr.EventBlock Then:
    if DebugLogging Then Call LogEvent("*** Event blocked " & zKey, eLmin):
    # GoTo pExit

    if DebugLogging Then Call LogEvent("*** Event ignored " & zKey & b & Item.Subject, eLmin):
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EventOl_Contacts_ItemChange
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EventOl_Contacts_ItemChange(ByVal Item As Object)
# Stop

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_NewTask_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_newtask_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_NewTask_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")
    # AddType = "Task (plain)"
    # Call commonTask(Item)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_NewTaskRequestAcceptItem_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_newtaskrequestacceptitem_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestAcceptItem_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")
    # AddType = "TaskRequestAcceptItem"
    # Call commonTask(Item)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_NewTaskRequestDeclineItem_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_newtaskrequestdeclineitem_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestDeclineItem_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")
    # AddType = "TaskRequestDeclineItem"
    # Call commonTask(Item)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_NewTaskRequest_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_newtaskrequest_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_NewTaskRequest_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")
    # AddType = "TaskRequest"
    # Call commonTask(Item)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_NewTaskRequestUpdateItem_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_newtaskrequestupdateitem_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_NewTaskRequestUpdateItem_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")
    # AddType = "TaskRequestUpdateItem"
    # Call commonTask(Item)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : commonTask
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Common Event Task
# '---------------------------------------------------------------------------------------
def commontask():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.commonTask"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    print(Debug.Print "new " & AddType & " arrived: " & Item.Subject)

    # ProcReturn:
    # Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : Sub EventOl_NewTask_ItemRemove
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EventOl_NewTask_ItemRemove()
# Stop

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_WEBInItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_webinitems_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_WEBInItems_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")

    print(Debug.Print "'NewMail' Event arrived for WEB: " & Now() _)
    # & b & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    if DebugLogging Then:
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    # Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub EventOl_WEBInItems_ItemChange
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EventOl_WEBInItems_ItemChange(ByVal Item As Object)
# 'Stop

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_WEBSeItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_webseitems_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_WEBSeItems_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call StartEP(zErr, zKey, tSubEP, eQEPMode)

    print(Debug.Print zKey & "-Event of Outlook occurred!")

    print(Debug.Print "'ItemAdd' Event sent for WEB: " & Now() _)
    # & b & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    if DebugLogging Then:
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    # Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_HotInItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_hotinitems_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_HotInItems_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    print(Debug.Print zKey & "-Event of Outlook occurred!")

    print(Debug.Print "'ItemAdd' Event arrived for Outlook: " & Now() _)
    # & b & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    if DebugLogging Then:
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    # Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_HotSeItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_hotseitems_itemadd():
    # '''' Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cOutlookSession.EventOl_HotSeItems_ItemAdd"
    # Static zErr As New cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    print(Debug.Print "'ItemAdd' Event sent for outlook: " & Now() _)
    # & b & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    if DebugLogging Then:
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    # Call DeferredActionAdd(Item, curAction:=3)     ' defer this item (there may be more waiting)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_BackupHomeInItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# ' Note   : currently, this event is not activated, so it will not be called or working
# '---------------------------------------------------------------------------------------
def eventol_backuphomeinitems_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_BackupHomeInItems_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    if DebugLogging Then:
    print(Debug.Print "'ItemAdd' Event arrived and IGNORED actions, item in Folder " _)
    # & Item.Parent.FullFolderPath & ": " & Now() _
    # & vbCrLf & vbTab & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    if EventHappened Then:
    if Not NoEventOnAddItem Then:
    if DebugLogging Then:
    print(Debug.Print "NoEventOnAddItem = True")
    # Call Sleep(DebugSleep)             ' wait here with modal box open

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : EventOl_BackupHomeSeItems_ItemAdd
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Event Routine
# '---------------------------------------------------------------------------------------
def eventol_backuphomeseitems_itemadd():
    # Dim zErr As cErr
    # Const zKey As String = "cOutlookSession.EventOl_BackupHomeSeItems_ItemAdd"
    # '------------------- gated Entry -------------------------------------------------------
    if E_Active.EventBlock Then:
    # GoTo pExit

    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

    if DebugLogging Then:
    print(Debug.Print "'ItemAdd' Event sent and IGNORED actions for " _)
    # & Item.Parent.FullFolderPath & ": " & Now() _
    # & vbCrLf & vbTab & Quote(Item.Subject) _
    # & vbCrLf & " from " & Quote(Item.SenderEmailAddress)
    # Call Sleep(DebugSleep)                     ' wait here with modal box open

    # ' DEACTIVATED ??? Call DeferredActionAdd(Item, curAction:=3)

    # ProcReturn:
    # Call ProcExit(zErr)
    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Init_Evnt_Handlers
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub Init_Evnt_Handlers()
# Dim zErr As cErr
# Const zKey As String = "cOutlookSession.Init_Evnt_Handlers"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub)

# Set myOlExplorers = olApp.Explorers
# Set objReminders = olApp.Reminders
# Set olInsp = olApp.Inspectors

# ProcReturn:
# Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : App_Test
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: .
# '---------------------------------------------------------------------------------------
# Public Sub App_Test()
# '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
# Const zKey As String = "cOutlookSession.App_Test"
# Dim zErr As cErr

# Dim i As Long
# Dim aExplorer As Explorer

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

if ActiveExplorer Is Nothing Then:
print(Debug.Print "no active explorers ... is this the end?")
else:
print(Debug.Print "there are " & Explorers.Count & " explorers open now")

# Set aExplorer = Explorers.Item(i)
print(Debug.Print " look at me")
# aExplorer.Close

# Debug.Assert False
if Not xlApp Is Nothing Then:
# Call xlEndApp

# ProcReturn:
# Call ProcExit(zErr)


