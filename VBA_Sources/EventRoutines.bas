Attribute VB_Name = "EventRoutines"
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Sub GetMainProfileRoot
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub GetMainProfileRoot()
    MainProfileAccount = aNameSpace.DefaultStore.DisplayName
    Call LogEvent("* Default Profile: " & aNameSpace.CurrentProfileName _
                & ", Main Profile Account: " & MainProfileAccount, eLall)
    ContactFolderName = "\" & aNameSpace.AddressLists.Item(1).Name

ProcRet:
End Sub ' EventRoutines.GetMainProfileRoot

'---------------------------------------------------------------------------------------
' Method : Sub PrepareMainFolders
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Prepare all Email Folder Events and required search folders
'---------------------------------------------------------------------------------------
Sub PrepareMainFolders(Account As String, Fmail As Folder, finx As Long)
Dim zErr As cErr
Const zKey As String = "EventRoutines.PrepareMainFolders"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim SubFolder As Folder
Dim mgPart As String
Dim ScopeFolders As String
Dim seCtr As Long
Dim atpos As Long

    Call BugTimer.BugState_SetPause

    ' EventOl_... Events are defined in OlSession (default-) class
       
    aBugTxt = "get any folder in Account " & Account
    Call Try(testAll)
    Set Fmail = TopFolders.Item(Account)
    Call Trunc(1, Fmail.Name, "@", NLoggedName, atpos)
    If LenB(NLoggedName) = 0 Then
        NLoggedName = Fmail.Name
    Else
        NLoggedName = TrimTail(NLoggedName, ".")
    End If
    Catch DoClear:=False
    If Fmail Is Nothing Then
        Call LogEvent("Account '" & Account & "' is missing. " _
                    & "No Events for account possible", eLall)
        GoTo FuncExit
    End If
    
    aBugTxt = "get specific folders for Registering Event Routines"
    Call Try(testAll)
    Select Case (finx)  ' note Folder  name strings differ!!!
    Case 0
        Set SubFolder = Fmail.Folders(BackupInboxFolder)
        Catch
        If SubFolder Is Nothing Then
            Call LogEvent(LString("Ordner " _
                & Quote(BackupInboxFolder), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath), eLall)
            GoTo FuncExit
        Else
            aBugTxt = "Add Account to loggable folders " _
                            & Quote(SubFolder.FolderPath)
            Call Try(&H8004010F)
            LoggableFolders.Add SubFolder.FolderPath, SubFolder
            Catch
        End If
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        aBugTxt = "get specific folders for Registering Event Routines"
        Call Try(testAll)
        Set OlSession.EventOl_BackupHomeInItems = SubFolder.Items
        Call CatchNC
        Call LogEvent(LString("IN-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = Quote1(SubFolder.FolderPath)
        
        Set SubFolder = Fmail.Folders(BackupSentFolder) ' ? &H8004010F
        If SubFolder Is Nothing Then
            Call Catch(AddMsg:=LString("Ordner " _
                & Quote(BackupSentFolder), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath))
            GoTo FuncExit
        Else
            Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_BackupHomeSeItems = SubFolder.Items
        Call LogEvent(LString("SE-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)
        
        Call MakeNotLogged(Account, ScopeFolders)
        
        Set FolderAggregatedInbox = Fmail.Folders(BackupAggregatedInbox)
        Catch
        If FolderAggregatedInbox Is Nothing Then
            Call LogEvent(LString("Ordner " _
                & Quote(BackupAggregatedInbox), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath), eLall)
            GoTo FuncExit
        Else
            Debug.Assert FolderAggregatedInbox.FolderPath = Fmail.FolderPath & "\" & BackupAggregatedInbox
        End If
    Case 1
        Set SubFolder = Fmail.Folders(WebInboxFolder)
Alternative1:
        Catch
        If SubFolder Is Nothing Then
            If SubFolder = WebInboxFolder Then
                Call Try(testAll)
                Set SubFolder = Fmail.Folders(StdInboxFolder)
                GoTo Alternative1
            End If
            Call LogEvent(LString("Ordner " _
                & Quote(SubFolder), OffCal) _
                & " missing in " & Quote(Fmail.FolderPath), eLall)
            GoTo FuncExit
        Else
            Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath)) _
                            
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_WEBInItems = SubFolder.Items
        Call LogEvent(LString("IN-Event auf Ordner " _
            & Quote(SubFolder.FolderPath), OffObj) & mgPart, eLall)
        ScopeFolders = Quote1(SubFolder.FolderPath)
        
        Set SubFolder = Fmail.Folders(StdSentFolder)
        If SubFolder Is Nothing Then
            Catch
            Call LogEvent(LString("Ordner " & Quote(StdSentFolder), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath), eLall)
            GoTo FuncExit
        Else
            Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_WEBSeItems = SubFolder.Items
        Call LogEvent(LString("SE-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)
        
        Call MakeNotLogged(Account, ScopeFolders)
    Case 2
        Set SubFolder = Fmail.Folders(StdInboxFolder)
        If SubFolder Is Nothing Then
            Catch
            Call LogEvent(LString("Ordner " & Quote(StdInboxFolder), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath), eLall)
            GoTo FuncExit
        Else
            aBugTxt = "Events Activation"
            Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_HotInItems = SubFolder.Items
        Call LogEvent(LString("IN-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = Quote1(SubFolder.FolderPath)
        
        ' #1 of 2
        seCtr = 0
        Set SubFolder = Fmail.Folders("Gesendete Elemente")
        If SubFolder Is Nothing Then
            Call Catch(AddMsg:=LString("Ordner " _
                & Quote("Gesendete Elemente"), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath))
            GoTo FuncExit
        Else
            Call Catch(AddMsg:=Account & " Folder " _
                & Quote(SubFolder.FolderPath))
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            seCtr = seCtr + 1
            mgPart = IsActive
            ScopeFolders = ScopeFolders & ", " _
                & Quote1(SubFolder.FolderPath)
            
        Else
        ' #2 of 2
            Set OlSession.EventOl_HotSeItems1 = SubFolder.Items
            Call LogEvent(LString("SE-Event1 auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
            Set SubFolder = Fmail.Folders("Gesendet")
            If SubFolder Is Nothing Then
                Call Catch(AddMsg:=LString("Ordner " _
                    & Quote("Gesendet"), OffObj) _
                    & " missing in " & Quote(Fmail.FolderPath))
                GoTo FuncExit
            Else
                Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
            End If
            LoggableFolders.Add SubFolder.FolderPath, SubFolder
            Set OlSession.EventOl_HotSeItems2 = SubFolder.Items
            If E_AppErr.errNumber = 0 Then
                seCtr = seCtr + 1
                mgPart = IsActive
            Else
                mgPart = InActive
            End If
        End If
        
        If seCtr > 0 Then
            Call LogEvent(LString("SE-Event2 auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) & mgPart, eLall)
            ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)
        Else
            Call LogEvent("Kein SE-Event im Konto " & Quote(Account), eLall)
        End If
        
        Call MakeNotLogged(Account, ScopeFolders)
Case 3
        Set SubFolder = Fmail.Folders(StdInboxFolder)
        If SubFolder Is Nothing Then
            Call Catch(AddMsg:=LString("Ordner " & Quote(StdInboxFolder), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath))
            GoTo FuncExit
        Else
            Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
        End If
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_GooInItems = SubFolder.Items
        Call LogEvent(LString("IN-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = Quote1(SubFolder.FolderPath)
        
        Set SubFolder = Fmail.Folders("[Google Mail]")
        If SubFolder Is Nothing Then
            Call Catch(AddMsg:=LString("Ordner " _
                & Quote("[Google Mail]"), OffObj) _
                & " missing in " & Quote(Fmail.FolderPath))
            GoTo FuncExit
        End If
        'LoggableFolders.Add SubFolder.FolderPath, SubFolder --- Not used, never has Items
        Set SubFolder = SubFolder.Folders("Gesendet")
        Call Catch(AddMsg:=Account & " Folder " _
                            & Quote(SubFolder.FolderPath))
        LoggableFolders.Add SubFolder.FolderPath, SubFolder
        If E_AppErr.errNumber = 0 Then
            mgPart = IsActive
        Else
            mgPart = InActive
        End If
        Set OlSession.EventOl_GooSeItems = SubFolder.Items
        Call LogEvent(LString("SE-Event auf Ordner " _
                            & Quote(SubFolder.FolderPath), OffObj) _
                            & mgPart, eLall)
        ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)
        
        Call MakeNotLogged(Account, ScopeFolders)
    Case Else
        DoVerify False, " undefined file index"
    End Select
    FldCnt = FldCnt + 1
    If UInxDeferred > 0 Then
        Call LogEvent("Es wurde im Account " & UInxDeferred & b _
            & "(" & Account & ") ein Suchordner vom Typ " _
            & Quote(SpecialSearchFolderName) _
            & " gefunden, aktuell unbearbeitet " _
            & DeferredFolder(UInxDeferred).Items.Count & " Items")
    Else
        Call LogEvent("No " _
            & Quote(SpecialSearchFolderName & b & NLoggedName) _
            & " in Account '" & Account & "' gefunden", eLall)
    End If
                 
FuncExit:
    Call ErrReset(0)
    Call ProcExit(zErr)
    Call BugTimer.BugState_UnPause
    
pExit:
End Sub ' EventRoutines.PrepareMainFolders

'---------------------------------------------------------------------------------------
' Method : Sub initEvents4Mail
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub initEvents4Mail()
Dim zErr As cErr
Const zKey As String = "EventRoutines.initEvents4Mail"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    Call BugTimer.BugState_SetPause
    
    UInxDeferred = 0                   ' init for known special search folders
    ItemsToDoCount = 0
    If LoggableFolders.Count > 0 Then
        Set LoggableFolders = New Dictionary
    End If
    
    Call GetMainProfileRoot
    Call PrepareMainFolders(BackupPSTname, OlBackupHome, 0)
    Call PrepareMainFolders(MailAccount1, OlWEBmailHome, 1)
    Call PrepareMainFolders(MainProfileAccount, OlHotMailHome, 2)
    Call PrepareMainFolders(MailAccount2, OlGooMailHome, 3)

    Call ProcExit(zErr)
    Call BugTimer.BugState_UnPause

pExit:
End Sub ' EventRoutines.initEvents4Mail

'---------------------------------------------------------------------------------------
' Method : GetToDoItems
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Obtain items in Deferred Folders that need processing
'---------------------------------------------------------------------------------------
Sub GetToDoItems()
Const zKey As String = "EventRoutines.GetToDoItems"
Static zErr As New cErr
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="EventRoutines")

Dim i As Long
Dim sFolder As Folder
    
    ItemsToDoCount = 0
    For i = 0 To LoggableFolders.Count - 1
        Set sFolder = LoggableFolders.Items(i)
        Call LogEvent(LString("+ " & sFolder.FolderPath, OffObj) _
                & "contains " & RString(sFolder.Items.Count, 8) & " Items", eLall)
        ItemsToDoCount = ItemsToDoCount + sFolder.Items.Count
    Next i
    
    Set sFolder = Nothing

ProcReturn:
    Call ProcExit(zErr)
    
End Sub ' EventRoutines.GetToDoItems

'---------------------------------------------------------------------------------------
' Method : Sub InitEventTraps
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub InitEventTraps()
Dim zErr As cErr
Const zKey As String = "EventRoutines.InitEventTraps"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    Call BugTimer.BugState_SetPause
    
    If Not topFolder Is Nothing Then
        GoTo ProcReturn   ' event Traps were already set up
        Resume Next
    End If
    
    Call LogEvent("* Preparing Outlook Operations", eLall)
    
    ' Hilfsgrößen für InBox und Gesendete Elemente
    EventHappened = False
    FldCnt = 0
    Set Deferred = New Collection
    
    Set TopFolders = aNameSpace.Folders
    Set FolderTasks = aNameSpace.GetDefaultFolder(olFolderTasks)
    If TrashFolder Is Nothing Then
        Set TrashFolder = aNameSpace.GetDefaultFolder(olFolderDeletedItems)
        TrashFolderPath = TrashFolder.FolderPath
    End If
    Call FindTopFolder(TrashFolder)
    
    Set CalendarFolder = aNameSpace.GetDefaultFolder(olFolderCalendar)
    
'   Verweis auf Items im Default Posteingang setzen:
    Set MainFolderInbox = olApp.Session.GetDefaultFolder(olFolderInbox)
    Set FolderInbox = MainFolderInbox
    Set ContactFolder = olApp.Session.GetDefaultFolder(olFolderContacts)
        
    If ContactFolder Is Nothing Then
        ContactFolderPath = "\\" & MainProfileAccount & "Kontakte"
        Set ContactFolder = GetFolderByName(ContactFolderPath)
        DoVerify Not ContactFolder Is Nothing, "Folder " & ContactFolderPath _
                & " not found"
    End If
    
    'Set OlSession.EventOl_HotInItems = MainFolderInbox.Items
    
    If GetOrMakeOlFolder(BackupPSTname, FolderBackup, TopFolders) Then
        Call N_ClearAppErr
        Set FolderInbox = MainFolderInbox
        GoTo OtherFolders
        If GetOrMakeOlFolder(BackupSMSFolder, FolderSMS, FolderInbox.Parent.Folders) Then
            Call N_ClearAppErr                 ' ignore if no SMS Folder was created
        End If
    Else
        Set FolderInbox = CreateFolderIfNotExists(BackupInboxFolder, FolderBackup)
        If FolderInbox Is Nothing Then
            Set FolderInbox = MainFolderInbox
        End If
        Call N_ClearAppErr                     ' ignore if no BackupInboxFolder was created
OtherFolders:
        If GetOrMakeOlFolder(BackupSMSFolder, FolderSMS, FolderBackup.Folders) Then
            Call N_ClearAppErr                 ' ignore if no SMS Folder was created
        End If
        If GetOrMakeOlFolder(BackupUnknownFolder, FolderUnknown, FolderBackup.Folders, allowMissing:=0) Then
            Call N_ClearAppErr                 ' ignore if no Unk Folder was created
        End If
    End If
    
'   Verweis auf Items in "Gesendet" setzen:
    Set MainFolderSent = olApp.Session.GetDefaultFolder(olFolderSentMail)
'   Verweis auf Zielordners des Transports aus "Gesendet"
    If FolderBackup Is Nothing Then
        Set FolderSent = MainFolderSent
    Else
        Set FolderSent = CreateFolderIfNotExists(BackupSentFolder, FolderBackup)
    End If
    Call GetDateId(-1)  ' just set global vars, value not needed
    Call LogEvent("=== TopFolder defined as " _
        & MainFolderInbox.Parent.FolderPath, eLall)
    
    Call initEvents4Mail
    
    Call OlSession.Init_Evnt_Handlers
    
    Call DefineLocalEnvironment

ProcReturn:
    Call ProcExit(zErr)
    Call BugTimer.BugState_UnPause
    
pExit:
End Sub ' EventRoutines.InitEventTraps

'---------------------------------------------------------------------------------------
' Method : Sub Inspector_Close
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Inspector_Close()
Dim zErr As cErr
Const zKey As String = "EventRoutines.Inspector_Close"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    
    DoVerify False, "who called Inspector_Close ???"
    If Explorers.Count = 0 And Inspectors.Count <= 1 Then
         DoVerify False
    End If
    
FuncExit:
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub ' EventRoutines.Inspector_Close

'---------------------------------------------------------------------------------------
' Method : Sub Start_BugTimer
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Start_BugTimer() ' Private of ThisOutlookSession / Friends only
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.Start_BugTimer"
    Call DoCall(zKey, "Sub", eQzMode)
    
    Call BugTimerDeActivate
    
    If Not NoTimerEvent Then
        BugTimer.BugStateReCheck = True
        If Not NoTimerEvent Then
            Debug.Print ("Activating the Timer.")
            Call BugTimerActivate(0)
        End If
    End If

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub ' EventRoutines.Start_BugTimer

'---------------------------------------------------------------------------------------
' Method : Sub Stop_BugTimer
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub Stop_BugTimer() ' Private of ThisOutlookSession / Friends only
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.Stop_BugTimer"
    Call DoCall(zKey, "Sub", eQzMode)
    
    Call BugTimerDeActivate

zExit:
    Call DoExit(zKey)
ProcRet:
End Sub ' EventRoutines.Stop_BugTimer

'---------------------------------------------------------------------------------------
' Method : Sub BugTimerDeActivate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub BugTimerDeActivate()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.BugTimerDeActivate"
'   deactivated: Call DoCall/DoExit     (zKey, "Sub", eQzMode)
Dim lSuccess As Long

    With BugTimer
        If .BugTimerId <> 0 Then
            lSuccess = KillTimer(0, .BugTimerId)
            If lSuccess = 0 Then
                MsgBox "The BugTimer failed to deactivate.", vbQuestion
            Else
                .BugStateLast = Timer
            
                If DebugLogging Then
                    Debug.Print "Timer " & .BugTimerId & " De-activated at " & .BugStateLast _
                        & " elapsed time " & .BugStateElapsed & " TriggerCount " & .BugStateTrigCount
                End If
                .BugTimerId = 0
            End If
            .BugStateElapsed = 0
        End If
    End With ' BugTimer
    BugTimer.BugStateReCheck = False

ProcRet:
End Sub ' EventRoutines.BugTimerDeActivate

'---------------------------------------------------------------------------------------
' Method : Sub BugTimerActivate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub BugTimerActivate(ByVal TimerSeconds As Long)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.BugTimerActivate"
    Call DoCall(zKey, "Sub", eQzMode)

    With BugTimer
        If Not .BugStateReCheck Then
            If DebugLogging Then
                Debug.Print "Not re-checking BugState timer"
            End If
            Call BugTimerDeActivate
            GoTo zExit
        End If
        
        If TimerSeconds = 0 Or .BugStateTicks = 0 Then
            TimerSeconds = DftBugStateAge
        End If
        
        .BugStateTicks = TimerSeconds           ' The SetTimer call accepts milliseconds
        
        If .BugTimerId <> 0 Then
            Call BugTimerDeActivate        ' Check to see if timer is running before call to SetTimer
        End If
        If .BugStateTicks <= 0 Then
            Debug.Print "BugStateTicks must be > 0"
            Debug.Assert False
            .BugStateTicks = 1000
        End If
        .BugTimerId = SetTimer(0, 0, .BugStateTicks * 1000, AddressOf BugStateEvent)
        If .BugTimerId = 0 Then
            MsgBox "The BugTimerId failed to activate.", vbQuestion
        Else
            .BugStateLast = Timer
            .BugStateElapsed = 0
            .BugStateReCheck = True
        
            If DebugLogging Then
                Debug.Print "Timer " & .BugTimerId & " activated at " & .BugStateLast & " elapsed " & Timer - .BugStateLast _
                        & " tick Sec " & TimerSeconds & " TriggerCount " & .BugStateTrigCount
            End If
        End If
    End With ' BugTimer

zExit:
    Call DoExit(zKey)
    
ProcRet:
End Sub ' EventRoutines.BugTimerActivate

'---------------------------------------------------------------------------------------
' Method : Sub BugStateEvent
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Time Interrupt Handling
'---------------------------------------------------------------------------------------
Sub BugStateEvent(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.BugStateEvent"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If ErrTimerEventNotReady Then
        Debug.Print "??? BugStatevent not ready"
        Debug.Assert False
        
        If BugTimer Is Nothing Then
            Set BugTimer = New cBugTimer
        Else
            Call BugTimerDeActivate
        End If
        GoTo ProcRet
    End If
    
    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                        & P_Active.DbgId & " => " & zKey _
                        & " at " & Timer & " after " & (Timer - BugTimer.BugStateLast) & " Sec" _
                        & " TriggerCount " & BugTimer.BugStateTrigCount
        End If
        GoTo ProcRet
    End If
    Recursive = True                        ' restored by    Recursive = False ProcRet:
    
    Call DoCall(zKey, "Sub", eQzMode)
    Call BugTimer.BugState_SetPause
        
    Call BugTimerEvent(vbNullString)

    Recursive = False

zExit:
    Call DoExit(zKey)
    Call BugTimer.BugState_UnPause
    
ProcRet:
End Sub ' EventRoutines.BugStateEvent

'---------------------------------------------------------------------------------------
' Method : Sub BugTimerEvent
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub BugTimerEvent(iCallMode As String)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "EventRoutines.BugTimerEvent"
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If ErrTimerEventNotReady Then
        GoTo ProcRet
    End If
    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug >= 8 And DebugMode Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                        & P_Active.DbgId & " => " & zKey _
                        & " at " & Timer & " after " & (Timer - BugTimer.BugStateLast) & " Sec" _
                        & " TriggerCount " & BugTimer.BugStateTrigCount
        End If
        GoTo ProcRet
    End If
    
    Recursive = True                        ' restored by    Recursive = False ProcRet:
    
    With BugTimer
        .BugStateElapsed = Timer - .BugStateLast
        .BugStateLast = Timer
        If .BugStateElapsed > .BugStateTicks Then
            .BugStateTrigCount = .BugStateTrigCount + 1
            If DebugLogging Then
                If LenB(iCallMode) = 0 Then
                    Debug.Print "Timer " & .BugTimerId _
                        & " has triggered after " & .BugStateElapsed _
                        & " sec" & " TriggerCount " & .BugStateTrigCount
                Else
                    Debug.Print "Timer " & .BugTimerId & b & iCallMode _
                        & " at " & .BugStateElapsed _
                        & " sec" & " TriggerCount " & .BugStateTrigCount
                End If
            End If
            .BugStateReCheck = False
            frmErrStatus.fLastErrExplanations = Time _
                        & "started waiting for BugStateReCheck for " _
                        & CSng(.BugStateTicks) & " Sec"
            Call Wait(CSng(.BugStateTicks) * 1000, _
                        Title:="waiting for BugStateReCheck") ' using mSec*1000
            Call BugEval
            If Not NoTimerEvent Then
                .BugStateReCheck = True
            End If
        End If
    End With ' BugTimer

zExit:
    Recursive = False
ProcRet:
End Sub ' EventRoutines.BugTimerEvent

'---------------------------------------------------------------------------------------
' Method : ErrTimerEventNotReady
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Check availabilit of Timer Events
'---------------------------------------------------------------------------------------
Function ErrTimerEventNotReady() As Boolean

    If BugTimer Is Nothing Then
        GoTo NAV
    End If
    With BugTimer
        If Not .BugStateReCheck Then
            GoTo NAV
        End If
        If .BugTimerId = 0 Then
            GoTo NAV
        End If
    End With ' BugTimer
    GoTo ProcRet

NAV:
    ErrTimerEventNotReady = True
    
ProcRet:
End Function ' EventRoutines.ErrTimerEventNotReady

'---------------------------------------------------------------------------------------
' Method : Function GetTerminationState
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function GetTerminationState() As Boolean
    GetTerminationState = T_DC.TermRQ
End Function ' EventRoutines.GetTerminationState

