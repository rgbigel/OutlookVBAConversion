# Converted from EventRoutines.py

# Attribute VB_Name = "EventRoutines"
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub GetMainProfileRoot
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getmainprofileroot():
    # MainProfileAccount = aNameSpace.DefaultStore.DisplayName
    # Call LogEvent("* Default Profile: " & aNameSpace.CurrentProfileName _
    # & ", Main Profile Account: " & MainProfileAccount, eLall)
    # ContactFolderName = "\" & aNameSpace.AddressLists.Item(1).Name

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub PrepareMainFolders
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Prepare all Email Folder Events and required search folders
# '---------------------------------------------------------------------------------------
def preparemainfolders():
    # Dim zErr As cErr
    # Const zKey As String = "EventRoutines.PrepareMainFolders"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # Dim SubFolder As Folder
    # Dim mgPart As String
    # Dim ScopeFolders As String
    # Dim seCtr As Long
    # Dim atpos As Long

    # Call BugTimer.BugState_SetPause

    # ' EventOl_... Events are defined in OlSession (default-) class

    # aBugTxt = "get any folder in Account " & Account
    # Call Try(testAll)
    # Set Fmail = TopFolders.Item(Account)
    # Call Trunc(1, Fmail.Name, "@", NLoggedName, atpos)
    if LenB(NLoggedName) = 0 Then:
    # NLoggedName = Fmail.Name
    else:
    # NLoggedName = TrimTail(NLoggedName, ".")
    # Catch DoClear:=False
    if Fmail Is Nothing Then:
    # Call LogEvent("Account '" & Account & "' is missing. " _
    # & "No Events for account possible", eLall)
    # GoTo FuncExit

    # aBugTxt = "get specific folders for Registering Event Routines"
    # Call Try(testAll)
    match differ!!!:
        case 0:
    # Set SubFolder = Fmail.Folders(BackupInboxFolder)
    # Catch
    if SubFolder Is Nothing Then:
    # Call LogEvent(LString("Ordner " _
    # & Quote(BackupInboxFolder), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath), eLall)
    # GoTo FuncExit
    else:
    # aBugTxt = "Add Account to loggable folders " _
    # & Quote(SubFolder.FolderPath)
    # Call Try(&H8004010F)
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    # Catch
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # aBugTxt = "get specific folders for Registering Event Routines"
    # Call Try(testAll)
    # Set OlSession.EventOl_BackupHomeInItems = SubFolder.Items
    # Call CatchNC
    # Call LogEvent(LString("IN-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = Quote1(SubFolder.FolderPath)

    # Set SubFolder = Fmail.Folders(BackupSentFolder) ' ? &H8004010F
    if SubFolder Is Nothing Then:
    # Call Catch(AddMsg:=LString("Ordner " _
    # & Quote(BackupSentFolder), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath))
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_BackupHomeSeItems = SubFolder.Items
    # Call LogEvent(LString("SE-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)

    # Call MakeNotLogged(Account, ScopeFolders)

    # Set FolderAggregatedInbox = Fmail.Folders(BackupAggregatedInbox)
    # Catch
    if FolderAggregatedInbox Is Nothing Then:
    # Call LogEvent(LString("Ordner " _
    # & Quote(BackupAggregatedInbox), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath), eLall)
    # GoTo FuncExit
    else:
    # Debug.Assert FolderAggregatedInbox.FolderPath = Fmail.FolderPath & "\" & BackupAggregatedInbox
        case 1:
    # Set SubFolder = Fmail.Folders(WebInboxFolder)
    # Alternative1:
    # Catch
    if SubFolder Is Nothing Then:
    if SubFolder = WebInboxFolder Then:
    # Call Try(testAll)
    # Set SubFolder = Fmail.Folders(StdInboxFolder)
    # GoTo Alternative1
    # Call LogEvent(LString("Ordner " _
    # & Quote(SubFolder), OffCal) _
    # & " missing in " & Quote(Fmail.FolderPath), eLall)
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath)) _

    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_WEBInItems = SubFolder.Items
    # Call LogEvent(LString("IN-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) & mgPart, eLall)
    # ScopeFolders = Quote1(SubFolder.FolderPath)

    # Set SubFolder = Fmail.Folders(StdSentFolder)
    if SubFolder Is Nothing Then:
    # Catch
    # Call LogEvent(LString("Ordner " & Quote(StdSentFolder), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath), eLall)
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_WEBSeItems = SubFolder.Items
    # Call LogEvent(LString("SE-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)

    # Call MakeNotLogged(Account, ScopeFolders)
        case 2:
    # Set SubFolder = Fmail.Folders(StdInboxFolder)
    if SubFolder Is Nothing Then:
    # Catch
    # Call LogEvent(LString("Ordner " & Quote(StdInboxFolder), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath), eLall)
    # GoTo FuncExit
    else:
    # aBugTxt = "Events Activation"
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_HotInItems = SubFolder.Items
    # Call LogEvent(LString("IN-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = Quote1(SubFolder.FolderPath)

    # ' #1 of 2
    # seCtr = 0
    # Set SubFolder = Fmail.Folders("Gesendete Elemente")
    if SubFolder Is Nothing Then:
    # Call Catch(AddMsg:=LString("Ordner " _
    # & Quote("Gesendete Elemente"), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath))
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # seCtr = seCtr + 1
    # mgPart = IsActive
    # ScopeFolders = ScopeFolders & ", " _
    # & Quote1(SubFolder.FolderPath)

    else:
    # ' #2 of 2
    # Set OlSession.EventOl_HotSeItems1 = SubFolder.Items
    # Call LogEvent(LString("SE-Event1 auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # Set SubFolder = Fmail.Folders("Gesendet")
    if SubFolder Is Nothing Then:
    # Call Catch(AddMsg:=LString("Ordner " _
    # & Quote("Gesendet"), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath))
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    # Set OlSession.EventOl_HotSeItems2 = SubFolder.Items
    if E_AppErr.errNumber = 0 Then:
    # seCtr = seCtr + 1
    # mgPart = IsActive
    else:
    # mgPart = InActive

    if seCtr > 0 Then:
    # Call LogEvent(LString("SE-Event2 auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) & mgPart, eLall)
    # ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)
    else:
    # Call LogEvent("Kein SE-Event im Konto " & Quote(Account), eLall)

    # Call MakeNotLogged(Account, ScopeFolders)
        case 3:
    # Set SubFolder = Fmail.Folders(StdInboxFolder)
    if SubFolder Is Nothing Then:
    # Call Catch(AddMsg:=LString("Ordner " & Quote(StdInboxFolder), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath))
    # GoTo FuncExit
    else:
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_GooInItems = SubFolder.Items
    # Call LogEvent(LString("IN-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = Quote1(SubFolder.FolderPath)

    # Set SubFolder = Fmail.Folders("[Google Mail]")
    if SubFolder Is Nothing Then:
    # Call Catch(AddMsg:=LString("Ordner " _
    # & Quote("[Google Mail]"), OffObj) _
    # & " missing in " & Quote(Fmail.FolderPath))
    # GoTo FuncExit
    # 'LoggableFolders.Add SubFolder.FolderPath, SubFolder --- Not used, never has Items
    # Set SubFolder = SubFolder.Folders("Gesendet")
    # Call Catch(AddMsg:=Account & " Folder " _
    # & Quote(SubFolder.FolderPath))
    # LoggableFolders.Add SubFolder.FolderPath, SubFolder
    if E_AppErr.errNumber = 0 Then:
    # mgPart = IsActive
    else:
    # mgPart = InActive
    # Set OlSession.EventOl_GooSeItems = SubFolder.Items
    # Call LogEvent(LString("SE-Event auf Ordner " _
    # & Quote(SubFolder.FolderPath), OffObj) _
    # & mgPart, eLall)
    # ScopeFolders = ScopeFolders & "," & Quote1(SubFolder.FolderPath)

    # Call MakeNotLogged(Account, ScopeFolders)
        case _:
    # DoVerify False, " undefined file index"
    # FldCnt = FldCnt + 1
    if UInxDeferred > 0 Then:
    # Call LogEvent("Es wurde im Account " & UInxDeferred & b _
    # & "(" & Account & ") ein Suchordner vom Typ " _
    # & Quote(SpecialSearchFolderName) _
    # & " gefunden, aktuell unbearbeitet " _
    # & DeferredFolder(UInxDeferred).Items.Count & " Items")
    else:
    # Call LogEvent("No " _
    # & Quote(SpecialSearchFolderName & b & NLoggedName) _
    # & " in Account '" & Account & "' gefunden", eLall)

    # FuncExit:
    # Call ErrReset(0)
    # Call ProcExit(zErr)
    # Call BugTimer.BugState_UnPause

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub initEvents4Mail
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initevents4mail():
    # Dim zErr As cErr
    # Const zKey As String = "EventRoutines.initEvents4Mail"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    # Call BugTimer.BugState_SetPause

    # UInxDeferred = 0                   ' init for known special search folders
    # ItemsToDoCount = 0
    if LoggableFolders.Count > 0 Then:
    # Set LoggableFolders = New Dictionary

    # Call GetMainProfileRoot
    # Call PrepareMainFolders(BackupPSTname, OlBackupHome, 0)
    # Call PrepareMainFolders(MailAccount1, OlWEBmailHome, 1)
    # Call PrepareMainFolders(MainProfileAccount, OlHotMailHome, 2)
    # Call PrepareMainFolders(MailAccount2, OlGooMailHome, 3)

    # Call ProcExit(zErr)
    # Call BugTimer.BugState_UnPause

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : GetToDoItems
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Obtain items in Deferred Folders that need processing
# '---------------------------------------------------------------------------------------
def gettodoitems():
    # Const zKey As String = "EventRoutines.GetToDoItems"
    # Static zErr As New cErr
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="EventRoutines")

    # Dim i As Long
    # Dim sFolder As Folder

    # ItemsToDoCount = 0
    # Set sFolder = LoggableFolders.Items(i)
    # Call LogEvent(LString("+ " & sFolder.FolderPath, OffObj) _
    # & "contains " & RString(sFolder.Items.Count, 8) & " Items", eLall)
    # ItemsToDoCount = ItemsToDoCount + sFolder.Items.Count

    # Set sFolder = Nothing

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub InitEventTraps
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def initeventtraps():
    # Dim zErr As cErr
    # Const zKey As String = "EventRoutines.InitEventTraps"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")
    # Call BugTimer.BugState_SetPause

    if Not topFolder Is Nothing Then:
    # GoTo ProcReturn   ' event Traps were already set up
    # Resume Next

    # Call LogEvent("* Preparing Outlook Operations", eLall)

    # ' Hilfsgren fr InBox und Gesendete Elemente
    # EventHappened = False
    # FldCnt = 0
    # Set Deferred = New Collection

    # Set TopFolders = aNameSpace.Folders
    # Set FolderTasks = aNameSpace.GetDefaultFolder(olFolderTasks)
    if TrashFolder Is Nothing Then:
    # Set TrashFolder = aNameSpace.GetDefaultFolder(olFolderDeletedItems)
    # TrashFolderPath = TrashFolder.FolderPath
    # Call FindTopFolder(TrashFolder)

    # Set CalendarFolder = aNameSpace.GetDefaultFolder(olFolderCalendar)

    # '   Verweis auf Items im Default Posteingang setzen:
    # Set MainFolderInbox = olApp.Session.GetDefaultFolder(olFolderInbox)
    # Set FolderInbox = MainFolderInbox
    # Set ContactFolder = olApp.Session.GetDefaultFolder(olFolderContacts)

    if ContactFolder Is Nothing Then:
    # ContactFolderPath = "\\" & MainProfileAccount & "Kontakte"
    # Set ContactFolder = GetFolderByName(ContactFolderPath)
    # DoVerify Not ContactFolder Is Nothing, "Folder " & ContactFolderPath _
    # & " not found"

    # 'Set OlSession.EventOl_HotInItems = MainFolderInbox.Items

    if GetOrMakeOlFolder(BackupPSTname, FolderBackup, TopFolders) Then:
    # Call N_ClearAppErr
    # Set FolderInbox = MainFolderInbox
    # GoTo OtherFolders
    if GetOrMakeOlFolder(BackupSMSFolder, FolderSMS, FolderInbox.Parent.Folders) Then:
    # Call N_ClearAppErr                 ' ignore if no SMS Folder was created
    else:
    # Set FolderInbox = CreateFolderIfNotExists(BackupInboxFolder, FolderBackup)
    if FolderInbox Is Nothing Then:
    # Set FolderInbox = MainFolderInbox
    # Call N_ClearAppErr                     ' ignore if no BackupInboxFolder was created
    # OtherFolders:
    if GetOrMakeOlFolder(BackupSMSFolder, FolderSMS, FolderBackup.Folders) Then:
    # Call N_ClearAppErr                 ' ignore if no SMS Folder was created
    if GetOrMakeOlFolder(BackupUnknownFolder, FolderUnknown, FolderBackup.Folders, allowMissing:=0) Then:
    # Call N_ClearAppErr                 ' ignore if no Unk Folder was created

    # '   Verweis auf Items in "Gesendet" setzen:
    # Set MainFolderSent = olApp.Session.GetDefaultFolder(olFolderSentMail)
    # '   Verweis auf Zielordners des Transports aus "Gesendet"
    if FolderBackup Is Nothing Then:
    # Set FolderSent = MainFolderSent
    else:
    # Set FolderSent = CreateFolderIfNotExists(BackupSentFolder, FolderBackup)
    # Call GetDateId(-1)  ' just set global vars, value not needed
    # Call LogEvent("=== TopFolder defined as " _
    # & MainFolderInbox.Parent.FolderPath, eLall)

    # Call initEvents4Mail

    # Call OlSession.Init_Evnt_Handlers

    # Call DefineLocalEnvironment

    # ProcReturn:
    # Call ProcExit(zErr)
    # Call BugTimer.BugState_UnPause

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Inspector_Close
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def inspector_close():
    # Dim zErr As cErr
    # Const zKey As String = "EventRoutines.Inspector_Close"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    # DoVerify False, "who called Inspector_Close ???"
    if Explorers.Count = 0 And Inspectors.Count <= 1 Then:
    # DoVerify False

    # FuncExit:

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Start_BugTimer
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def start_bugtimer():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.Start_BugTimer"
    # Call DoCall(zKey, "Sub", eQzMode)

    # Call BugTimerDeActivate

    if Not NoTimerEvent Then:
    # BugTimer.BugStateReCheck = True
    if Not NoTimerEvent Then:
    print(Debug.Print ("Activating the Timer."))
    # Call BugTimerActivate(0)

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Stop_BugTimer
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def stop_bugtimer():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.Stop_BugTimer"
    # Call DoCall(zKey, "Sub", eQzMode)

    # Call BugTimerDeActivate

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub BugTimerDeActivate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bugtimerdeactivate():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.BugTimerDeActivate"
    # '   deactivated: Call DoCall/DoExit     (zKey, "Sub", eQzMode)
    # Dim lSuccess As Long

    # With BugTimer
    if .BugTimerId <> 0 Then:
    # lSuccess = KillTimer(0, .BugTimerId)
    if lSuccess = 0 Then:
    print('The BugTimer failed to deactivate.')
    else:
    # .BugStateLast = Timer

    if DebugLogging Then:
    print(Debug.Print "Timer " & .BugTimerId & " De-activated at " & .BugStateLast _)
    # & " elapsed time " & .BugStateElapsed & " TriggerCount " & .BugStateTrigCount
    # .BugTimerId = 0
    # .BugStateElapsed = 0
    # End With ' BugTimer
    # BugTimer.BugStateReCheck = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub BugTimerActivate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bugtimeractivate():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.BugTimerActivate"
    # Call DoCall(zKey, "Sub", eQzMode)

    # With BugTimer
    if Not .BugStateReCheck Then:
    if DebugLogging Then:
    print(Debug.Print "Not re-checking BugState timer")
    # Call BugTimerDeActivate
    # GoTo zExit

    if TimerSeconds = 0 Or .BugStateTicks = 0 Then:
    # TimerSeconds = DftBugStateAge

    # .BugStateTicks = TimerSeconds           ' The SetTimer call accepts milliseconds

    if .BugTimerId <> 0 Then:
    # Call BugTimerDeActivate        ' Check to see if timer is running before call to SetTimer
    if .BugStateTicks <= 0 Then:
    print(Debug.Print "BugStateTicks must be > 0")
    # Debug.Assert False
    # .BugStateTicks = 1000
    # .BugTimerId = SetTimer(0, 0, .BugStateTicks * 1000, AddressOf BugStateEvent)
    if .BugTimerId = 0 Then:
    print('The BugTimerId failed to activate.')
    else:
    # .BugStateLast = Timer
    # .BugStateElapsed = 0
    # .BugStateReCheck = True

    if DebugLogging Then:
    print(Debug.Print "Timer " & .BugTimerId & " activated at " & .BugStateLast & " elapsed " & Timer - .BugStateLast _)
    # & " tick Sec " & TimerSeconds & " TriggerCount " & .BugStateTrigCount
    # End With ' BugTimer

    # zExit:
    # Call DoExit(zKey)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub BugStateEvent
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Time Interrupt Handling
# '---------------------------------------------------------------------------------------
def bugstateevent():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.BugStateEvent"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if ErrTimerEventNotReady Then:
    print(Debug.Print "??? BugStatevent not ready")
    # Debug.Assert False

    if BugTimer Is Nothing Then:
    # Set BugTimer = New cBugTimer
    else:
    # Call BugTimerDeActivate
    # GoTo ProcRet

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey _
    # & " at " & Timer & " after " & (Timer - BugTimer.BugStateLast) & " Sec" _
    # & " TriggerCount " & BugTimer.BugStateTrigCount
    # GoTo ProcRet
    # Recursive = True                        ' restored by    Recursive = False ProcRet:

    # Call DoCall(zKey, "Sub", eQzMode)
    # Call BugTimer.BugState_SetPause

    # Call BugTimerEvent(vbNullString)

    # Recursive = False

    # zExit:
    # Call DoExit(zKey)
    # Call BugTimer.BugState_UnPause

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub BugTimerEvent
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bugtimerevent():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "EventRoutines.BugTimerEvent"
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if ErrTimerEventNotReady Then:
    # GoTo ProcRet
    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 And DebugMode Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey _
    # & " at " & Timer & " after " & (Timer - BugTimer.BugStateLast) & " Sec" _
    # & " TriggerCount " & BugTimer.BugStateTrigCount
    # GoTo ProcRet

    # Recursive = True                        ' restored by    Recursive = False ProcRet:

    # With BugTimer
    # .BugStateElapsed = Timer - .BugStateLast
    # .BugStateLast = Timer
    if .BugStateElapsed > .BugStateTicks Then:
    # .BugStateTrigCount = .BugStateTrigCount + 1
    if DebugLogging Then:
    if LenB(iCallMode) = 0 Then:
    print(Debug.Print "Timer " & .BugTimerId _)
    # & " has triggered after " & .BugStateElapsed _
    # & " sec" & " TriggerCount " & .BugStateTrigCount
    else:
    print(Debug.Print "Timer " & .BugTimerId & b & iCallMode _)
    # & " at " & .BugStateElapsed _
    # & " sec" & " TriggerCount " & .BugStateTrigCount
    # .BugStateReCheck = False
    # frmErrStatus.fLastErrExplanations = Time _
    # & "started waiting for BugStateReCheck for " _
    # & CSng(.BugStateTicks) & " Sec"
    # Call Wait(CSng(.BugStateTicks) * 1000, _
    # Title:="waiting for BugStateReCheck") ' using mSec*1000
    # Call BugEval
    if Not NoTimerEvent Then:
    # .BugStateReCheck = True
    # End With ' BugTimer

    # zExit:
    # Recursive = False
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : ErrTimerEventNotReady
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Check availabilit of Timer Events
# '---------------------------------------------------------------------------------------
def errtimereventnotready():

    if BugTimer Is Nothing Then:
    # GoTo NAV
    # With BugTimer
    if Not .BugStateReCheck Then:
    # GoTo NAV
    if .BugTimerId = 0 Then:
    # GoTo NAV
    # End With ' BugTimer
    # GoTo ProcRet

    # NAV:
    # ErrTimerEventNotReady = True

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function GetTerminationState
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getterminationstate():
    # GetTerminationState = T_DC.TermRQ
