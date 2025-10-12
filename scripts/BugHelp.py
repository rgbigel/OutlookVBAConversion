# Converted from BugHelp.py

# Attribute VB_Name = "BugHelp"
# Option Explicit

# ' Name Conventions: Procs in this module should be simple to recognize on Stacks
# '        names N_ are not visible on any stacks, do not use Z_Entry/DoExit etc.
# '                 these should (must ???) never call non-N_Type Procs
# '        names Z_ are defined but not usually visible
# '                   Mode must be <=eQyMode<=eQxDMode
# '              Proc/Live/Show/Query equivalent "Z_", e.g. ProcCall, ProcReturn etc.
# '        Z_    Procs which Show/manipulate stacks
# '              Z_ = x/y/z_Type procs visible on stacks, iff they call DoCall/DoExit or found on LiveStack
# '              Z_   not normally using D_ErrInterface
# '              ALL  Procs can use ProcCall and ProcExit, because they use DoCall / DoExit
# '        Z_Types are used for all Procs when FastMode = True (changing Qmode to y)
# '        All other Names can be of any Qmode-Type, most of these use:
# ' DoCall      Mark entry to  EVERY Proc (Sub, Function, ... )
# ' ProcCall      "  as above, but with special rules for error handling
# ' DoExit      Mark exit from EVERY Proc (Sub, Function, ... )
# ' ProcReturn    "  as above, but with special rules for error handling
# ' StartEP     like ProcCall, but used for external entry points/event handlers (EP-Procs)
# ' ReturnEP    like ProcReturn, but used when Ending EP-Procs

# ' Note      ' DscMode Determines the eQmode which sets the rules for ProcCall
# '               all eQmodes>eQzMode are in D_ErrInterface Collection and do the following:
# '               Watches for Err.Number<>0 and handles Errors according to eQmode
# '               Eliminate ClientDsc's marked invalid(???)
# '               Completes all values in ClientDsc and adds new ones to D_ErrInterfaceM
# '               Calls N_ConDscErr to define consistent cErr/cProcDsc
# '               for non-predefining calls, use N_StackDef to put on any Stacks,
# '                  with N_CallEnv to define Call Environment (like back ref to Caller etc),
# '                  Set Date/Time of ClientDsc's call
# '               Adds to D_AppStack Stack (if QMode >= eQASMode)
# '               Optionally Logs using N_LogEntry-->LogGen

# ' If ClientDsc not defined when using LiveStack:
# '                  create cProcDsc and add to D_ErrInterface (Dictionary)
# '                  create cErr    (as ClientDsc.ErrActive, using N_ConDscErr)
# ' ProcCall:  Notes the Call of a procedure of any QMode using DoCall
# '                  add 1 to atDsc's CallCounter
# '                  inserts atCalledBy from D_AppStack
# '                 Checks if recursion is allowed.
# '                    If so and recursion level>1, create cErr Instance and chain it
# '                 Pauses Callers
# '                 If QMode >= eQASMode adds to D_AppStack by Call Z_ToAppStack

# ' Hierarchy of Calls of Procedures in this Module

# ' General:      use the following general parameter list:
# '               (ClientDsc As cProcItem, DscMode As eQmode, tSub/tFunction As String, Mode As String, _
# '                Optional ClassInstance As Long, Optional RecursionRequested As Boolean)

# '---------------------------------------------------------------------------------------
# ' Method : BugEval
# ' Author : Rolf G. Berchtocall
# ' Date   : 20211108@11_47
# ' Purpose: Re-Evaluate settings in aNonModalForm
# '---------------------------------------------------------------------------------------
def bugeval():

    # Const zKey As String = "BugHelp.BugEval"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug > 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # Call DoCall(zKey, tSub, eQzMode)

    if ErrStatusFormUsable Then:
    # frmErrStatus.fModifications.Enabled = True
    # frmErrStatus.frmIgnoreErrStatusChange = True
    # frmErrStatus.fModifications = True
    # Call frmErrStatus.ReEvaluate
    # StackDebugOverride = StackDebug            ' this value is always >= 0
    # Call ShowOrHideForm(frmErrStatus, Not frmErrStatus.fHideMe)

    # FuncExit:
    # Recursive = False

    # zExit:
    # Call DoExit(zKey)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : BugSet
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Update settings in aNonModalForm from Program vars, espc. from E_Active
# '---------------------------------------------------------------------------------------
def bugset():

    # Const zKey As String = "BugHelp.BugEval"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if StackDebug >= 8 Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    if ErrStatusFormUsable Then:
    # ErrDisplayModify = True
    # Call frmErrStatus.UpdInfo

    # FuncExit:
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Function Catch
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: After Try, if any error was found, check the details and Clear if acceptable
# ' Note   : the Global Error handler has analized if the error was acceptable (by Try-Rules)
# '          if an error did occur, but was acceptable, the Error is resetted
# '                           unless DoClear=False is specified
# '          The Returnvalue of Catch indicates if there was an Error at all.
# '---------------------------------------------------------------------------------------
def catch():

    if T_DC.TermRQ Then:
    # Call TerminateRun                          ' aborting Entry if not handled

    # Dim msg As String

    # With E_Active
    if LenB(.Explanations) > 0 Then:
    # msg = .Explanations
    elif LenB(aBugTxt) > 0 Then:
    # msg = aBugTxt
    elif LenB(T_DC.DCerrMsg) > 0 Then:
    # msg = T_DC.DCerrMsg
    if LenB(msg) = 0 Then:
    # msg = "Catch"
    if .errNumber <> 0 Then:
    if IsMissing(HandleErr) Then:
    # Catch = True
    elif T_DC.DCerrNum <> HandleErr Then:
    # Catch = True
    else:
    # GoTo ErrOk
    if .FoundBadErrorNr = 0 And T_DC.DCerrNum = 0 Then:
    # ErrOk:
    if DoMessage And (DebugMode Or DebugLogging Or InStr(msg, "***") > 0) Then:
    # Call LogEvent("OK:" & msg)
    if DoClear Then:
    # Call ErrReset(4)
    else:
    if .FoundBadErrorNr <> 0 Then:
    if LenB(AddMsg) > 0 Then:
    # msg = msg & vbCrLf & AddMsg
    # Call LogEvent("!!! Failed: " & msg)
    # Debug.Assert False
    if DoClear Then:
    # Call ErrReset(4)
    # End With                                       ' E_Active
    # aBugTxt = vbNullString

    # ProcRet:

# ' ------ Variation of Catch with different parms ---------------------------------------
def catchnc():
    # Call Catch(DoClear, DoMessage, AddMsg, HandleErr)

# ' Global Interface variables of BugHelp see Module Z_ErrIf, generated by ZZIfGen (but long outdated)

# '---------------------------------------------------------------------------------------

# '---------------------------------------------------------------------------------------
# ' Method : DoCall
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: If #MoreDiagnostics, maintain entries in D_Errinterface
# '---------------------------------------------------------------------------------------
def docall():
    # Const zKey As String = "BugHelp.DoCall"

    # Dim aKey As String
    # Dim aCallType As String
    # Dim aMode As eQMode
    # Dim aDsc As cProcItem                              ' that is <Self>, not Client (CalledDsc, aDsc.ErrActive)
    # Dim aErr As cErr
    # Dim TimeIn As Double
    # Dim isNewDsc As Boolean

    if QuitStarted Then:
    # Exit Sub

    # TimeIn = Timer

    # Call BugTimer.BugState_SetPause             ' do not allow any timer and other events

    if D_ErrInterface.Count = 0 Then            ' first definition, come back for ActiveProc below:
    # aCallType = tSub
    # aMode = eQnoDef
    # Call N_ConDscErr(aDsc, "Extern.Caller", tSub, aMode, aErr)
    # Set aErr.atCalledBy = aErr              ' self reference, the only allowed case
    # Call N_StackDef(aDsc, aErr, False)      ' External Caller is the Stack root, but has no at_Live
    # Set ExternCaller = aDsc
    # Set P_Active = aDsc
    # Set E_Active = aErr

    # aKey = zKey                             ' use this to define atDsc / Err for DoCall
    # Set aDsc = Nothing
    # Set aErr = Nothing                      ' continue for really called Prog later!
    else:
    # aKey = ClientKey
    # aCallType = CallType
    # aMode = Mode
    # ActiveProc:
    if LenB(aKey) = 0 Then:
    # Stop                                ' is this ever possible ???
    # GoTo isNew
    if CalledDsc Is Nothing Then:
    # GoTo unknown
    elif CalledDsc.Key = vbNullString Then:
    # GoTo unknown
    elif CalledDsc.ErrActive Is Nothing Then:
    # unknown:
    # Set aDsc = Nothing
    # Set aErr = Nothing

    if aDsc Is Nothing Then:
    if Not aErr Is Nothing Then             ' should not ever happen ???:
    # DoVerify aErr.atKey = aKey, _
    # "** Key mismatch at definition of cErr ???"
    # Set aDsc = aErr.atDsc
    # DoVerify Not aDsc Is Nothing, _
    # "inconsistent cErr ???"
    # aDsc.Key = aKey
    # aCallType = CallType
    # aMode = Mode
    # GoTo isNew
    else:
    if Not dontIncrementCallDepth Then:
    # DoVerify aDsc.Key = aKey, _
    # "** Key mismatch at definition of aDsc=" _
    # & aDsc.Key & "<>" & aKey

    if D_ErrInterface.Exists(aKey) Then:
    if isEmpty(D_ErrInterface.Item(aKey)) Then:
    # DoVerify False, _
    # "*** VBA Bug: how can D_ErrInterface.Item be Empty ???"
    # Call N_ConDscErr(aDsc, aKey, tSub, aMode, aErr)
    # Set D_ErrInterface.Item(aKey) = aDsc
    if aErr Is Nothing Then:
    # Set aErr = aDsc.ErrActive
    # Set aDsc.ErrActive = aErr
    else:
    # Set aDsc = D_ErrInterface.Item(aKey)
    if aErr Is Nothing Then:
    # Set aErr = aDsc.ErrActive
    else:
    # Stop ' ??? need if ???

    # Set aDsc.ErrActive = aErr
    if aDsc Is Nothing Or aErr Is Nothing Then:
    # DoVerify False, _
    # "D_ErrInterface.Item atDsc/Err is Nothing ???"
    # GoTo isNew
    else:
    # aBugTxt = "D_ErrInterface.Item Key in atDsc <> Err ???"
    if DoVerify(aErr.atKey = aKey) Then:
    # GoTo isNew
    # GoTo isKnown
    else:
    # isNew:
    # Call N_ConDscErr(aDsc, aKey, CallType, aMode, aErr)
    if aErr.atCallDepth < 0 Then:
    # aErr.atCallDepth = 0                            ' show it is new in DoCall
    # aErr.atMessage = ExplainS
    # isNewDsc = True
    # isKnown:
    # Set CalledDsc = aDsc                                ' CalledDsc: delivering values
    if N_StackDef(aDsc, aErr, False) Then               ' False:: always GenCallData:
    if aKey = zKey Then:
    # Set NCall = aDsc                           ' self-Defining has been done
    # NCall.CallCounter = 0                      ' Inc happens below
    if ClientKey <> D_ErrInterface.Keys(0) Then ' not again for Extern.Caller:
    # aKey = ClientKey
    # Set aDsc = Nothing
    # Set aErr = Nothing
    # GoTo ActiveProc                         ' now define true Caller
    else:
    # Set P_Active = D_ErrInterface.Item(ClientKey)
    # Set E_Active = P_Active.ErrActive
    else:
    # Set aDsc.ErrActive = aErr

    # FuncExit:
    # NCall.CallCounter = NCall.CallCounter + 1
    # NCall.TotalProcTime = NCall.TotalProcTime + Timer - TimeIn

    if ExLiveDscGen And Not dontIncrementCallDepth Then    ' not predefining; if Live Stack use requested:
    # aErr.atLiveLevel = ExLiveCheck(ForcePrint:=ExLiveCheckLog, LocalStack:=D_LiveStack)

    if Not aErr.atCalledBy Is Nothing Then:
    # aErr.atCalledBy.atCallState = eCpaused
    # Set aDsc = Nothing
    # Set aErr = Nothing
    # Set CalledDsc = P_Active
    if LenB(ExplainS) > 0 And CallLogging Then:
    # Call LogEvent("Calling " & P_Active.DbgId & b & ExplainS)
    # Call BugTimer.BugState_UnPause                 ' do restore timer and other events

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : ExLiveCheck
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose: Get Call Level of running VBA Procs and the current top running proc
# ' Note   : ignores itself, Main result is in the Public LCA (cCallEnv)
# '---------------------------------------------------------------------------------------
def exlivecheck():

    # Dim LTimer As Double
    # Dim msg As String
    # Dim ModProc As String
    # Dim FoundModProc As String
    # Dim i As Long
    # Dim MatchFullQualified As Boolean
    # Dim ActiveProc As String
    # Dim HaveMatch As Boolean
    # Dim IsRelevant As Boolean

    # Dim aProcDsc As cProcItem
    # Dim nErr As cErr
    # Dim aErr As cErr

    if ExLiveCheck Then:
    # SkipRemainingStack = False
    # MustMatch = vbNullString
    # MatchFullQualified = True
    # LookUpDsc = True
    elif MustMatch <> vbNullString Then:
    # MatchFullQualified = InStr(MustMatch, ".")
    if LookUpDsc And Not ExLiveCheck Then:
    if DoVerify(Not SkipRemainingStack, "if we have MustMatch, must SkipRemainingStack") Then:
    # SkipRemainingStack = True
    else:
    # MatchFullQualified = True
    if SkipRemainingStack Then:
    if DoVerify(Not LookUpDsc, _:
    # "if we have LookUpDsc, can't SkipRemainingStack") Then
    # SkipRemainingStack = False
    # LTimer = Timer
    # Set LocalStack = N_GetLiveStack
    # msg = "Live Stack Root"

    # Set LCA = LocalStack.Items(i)
    # ExLiveCheck = LocalStack.Count - i
    # IsRelevant = N_RelevantProc(LCA, ForcePrint, ModProc)
    # IsEntryPoint = InStr(LCA.ProcedureName, "Appl") > 0
    if IsRelevant Then:
    # FoundPreCall:
    # ModProc = Left(LCA.ModProc & String(30, b), 30)
    if i = 1 Then:
    if ForcePrint Then:
    # msg = "This is the active Proc, called by " & ActiveProc
    else:
    if ForcePrint And i < LocalStack.Count - 1 Then:
    # msg = "Root at - " & i
    if i = 2 Then:
    # ActiveProc = ModProc

    if MustMatch = IIf(MatchFullQualified, LCA.ModProc, LCA.ProcedureName) Then:
    # HaveMatch = True
    if ForcePrint Then:
    # msg = Trim(msg & " found the MustMatch " & MustMatch)
    # FoundModProc = LCA.ModProc
    if ForcePrint Then:
    print(Debug.Print ExLiveCheck & " / " & LSD, _)
    # FoundModProc, "Used=" & Timer - LTimer, _
    # "Line " & LCA.LineNumber & " Code: " & LCA.LineCode _
    # & vbCrLf & String(80, "-")
    # GoTo FoundIt

    if ForcePrint And Not ShutUpMode Then:
    print(Debug.Print i, ModProc, _)
    # LString(msg, 20) & "--->", LCA.LineCode

    if LenB(FoundModProc) = 0 Then:
    if LSD > 0 Then:
    if LSD > ExLiveCheck Then:
    # GoTo PrChange
    elif MustMatch = vbNullString Then:
    # ChangeIt:
    if ForcePrint And Not ShutUpMode Then:
    # PrChange:
    print(Debug.Print i, ModProc, _)
    # LString("LSD has been set", 20), _
    # "to Call Depth=" & ExLiveCheck
    # LSD = ExLiveCheck
    # FoundIt:
    if ExLiveCheck Or IsEntryPoint _:
    # Or (LookUpDsc And LCA.CallerErr Is Nothing) Then
    if D_ErrInterface.Exists(LCA.ModProc) Then              ' N_ Procs never exist:
    if isEmpty(D_ErrInterface.Item(LCA.ModProc)) Then:
    # Set D_ErrInterface.Item(LCA.ModProc) = aProcDsc ' should be Nothing
    # GoTo mustCorrErr
    # Set aProcDsc = D_ErrInterface.Item(LCA.ModProc)
    # Set nErr = aProcDsc.ErrActive
    else:
    if ExLiveCheck Then               ' generate a description for all procs:
    # mustCorrErr:
    # Set nErr = New cErr
    # Call N_ConDscErr(aProcDsc, LCA.ModProc, LCA.DscKind, eQnoDef, nErr)
    # nErr.Explanations = "! ExLiveCheck=True !"
    # Set LCA.CallerErr = nErr
    if dontIncrementCallDepth Then:
    # Set aProcDsc = nErr.atDsc
    else:
    # nErr.atCallDepth = ExLiveCheck
    # dontIncrementCallDepth = True   ' no push because False at start
    # Call N_StackDef(aProcDsc, nErr, IsRelevant)
    # dontIncrementCallDepth = False  ' no pop necessary
    if LCA.CallerErr Is Nothing Then:
    # Set LCA.CallerErr = aProcDsc.ErrActive
    else:
    if nErr Is Nothing Then:
    # Stop ' ???
    else:
    # Set LCA.CallerErr = nErr
    # Set aErr = nErr
    # Set aErr.at_Live = LocalStack       ' maybe we do not need all of that ???
    # nErr.atCatDict = Trim(nErr.atCatDict & b & nErr.atKey)
    if Not aErr.atCalledBy Is Nothing Then:
    if DoVerify(nErr.atCalledBy Is aErr.atCalledBy, _:
    # "??? CalledBy correct? is " _
    # & nErr.atCalledBy.atKey & " to " _
    # & aErr.atCalledBy.atKey) Then
    # Set nErr.atCalledBy = aErr.atCalledBy

    if i > 1 Then:
    # nErr.atCallState = eCpaused             ' all procs below active are paused

    if SkipRemainingStack And FoundModProc <> vbNullString Then:
    # GoTo FuncExit

    # FuncExit:
    # Set aProcDsc = Nothing
    # Set aErr = Nothing
    # Set nErr = Nothing
    # Set LocalStack = Nothing

    if ClearLCS Then:
    # Set LCS = Nothing

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_RelevantProc
# ' Author : rgbig
# ' Date   : 13.03.2020
# ' Purpose: Determine if Proc is relevant based on Name
# '---------------------------------------------------------------------------------------
def n_relevantproc():

    # N_RelevantProc = True                          ' all the rest is relevant

    if Left(LCA.ProcedureName, 1) = "<" Then       ' immediate and external code (eg. events):
    # GoTo IsIrrelevant
    elif InStr("ExLi Live DoCa DoEx Proc Show Query ", _:
    # Left(LCA.ProcedureName, 4)) > 0 Then
    # IsIrrelevant:
    # N_RelevantProc = False                     ' DoCall etc. are irrelevant
    if ForcePrint Then:
    # ModProc = Left(LCA.ModProc & String(30, b), 30)
    print(Debug.Print LCA.StackDepth, ModProc, _)
    # LString("Irrelevant Proc", 20) _
    # & "--->", LCA.LineCode


# '---------------------------------------------------------------------------------------
# ' Method : DoExit
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Exit required after for DoCall. Used in ProcExit
# '---------------------------------------------------------------------------------------
def doexit():
    # Dim fromDsc As cProcItem
    # Dim fromErr As cErr
    # Dim toErr As cErr
    # Dim Consumed As String
    # Dim FuncVal As String

    if QuitStarted Then:
    # Exit Sub

    # Call BugTimer.BugState_SetPause                ' do not allow any timer and other events

    # aBugTxt = "Return without a call definition ???"
    if DoVerify(D_ErrInterface.Exists(zKey)) Then:
    # GoTo ERET
    # aBugVer = isEmpty(D_ErrInterface.Item(zKey))
    if aBugVer Then:
    if StackDebug >= 8 Then:
    print(Debug.Print zKey & "has empty cErr ???")
    # Debug.Assert False
    # GoTo ERET

    if E_Active.errNumber <> 0 Then                ' any pending errors?:
    # Call N_CaptureNewErr                       ' handle error as defined (likely will pause there)
    if T_DC.TermRQ Then                            ' aborting Entry if not handled by user:
    # Call TerminateRun                          ' never returns unless user says so
    if E_Active.errNumber = 0 Then:
    # T_DC.N_ClearTermination

    # Set fromDsc = D_ErrInterface.Item(zKey)
    # Set fromErr = fromDsc.ErrActive
    # Set toErr = E_Active.atCalledBy
    if toErr Is Nothing Then                       ' external Caller:
    # Set toErr = D_ErrInterface.Items(0)
    # DoVerify toErr.atCallDepth <= 1 Or Not fromErr Is toErr, _
    # "??? returning to itself=" & toErr.atKey
    # DoVerify fromErr Is fromDsc.ErrActive, "atDsc and Err not linked ???"
    if DoVerify(fromDsc.Key = fromErr.atKey, "Dsc.Key <> Err.atKey ??? in " & fromDsc.Key) Then:
    # fromErr.atKey = fromDsc.Key
    if DoVerify(Not E_Active.atDsc Is Nothing, "E_Active.atDsc is Nothing") Then:
    # Set E_Active = fromErr

    if LogPerformance Then:
    if fromErr.atThisEntrySec > 0 Then:
    # Call Z_UsedThisCall(fromErr, Timer)
    # Consumed = " Consumed " & ElapsedTime & " sec "
    else:
    # Consumed = vbNullString

    # fromErr.atCallState = eCExited

    if LogZProcs Or Not AppStartComplete Then:
    if LenB(DisplayValue) > 0 Then:
    # FuncVal = " >" & DisplayValue
    else:
    # FuncVal = vbNullString
    # Call N_ShowProgress(CallNr, fromDsc, "-Z", _
    # "zD=" & toErr.atCallDepth & "" & fromErr.atCallDepth _
    # & IIf(E_Active.EventBlock, " NoEvents", vbNullString), _
    # Consumed & FuncVal, ErrClient:=fromErr)

    # Set fromErr.atErrPrev = Nothing                ' no previous error instance
    # Set fromErr.at_Live = Nothing                  ' no call stack after exit
    # fromErr.atCatDict = vbNullString

    if toErr.atCallDepth <> 0 Then                 ' only extern caller has =0:
    if DoVerify(toErr.atRecursionLvl > 0, _:
    # "error in recursion level, too many returns ???") Then
    # fromErr.atRecursionLvl = 0             ' this does not really fix the problem!
    else:
    # fromErr.atRecursionLvl = fromErr.atRecursionLvl - 1
    if toErr.atDsc.CallMode > eQnoDef Then:
    # DoVerify toErr.atCallState = eCpaused, _
    # "exiting to caller: " & toErr.atKey _
    # & "=" & CStateNames(toErr.atCallState) _
    # ' unlikely call state, it should be Paused

    if ExLiveDscGen Then:
    if ExLiveDscGen = toErr.atLiveLevel Then:
    if DebugMode And LogZProcs Then:
    print(Debug.Print "DoExit setting CallState to Active for " & toErr.atKey)
    # toErr.atCallState = eCActive

    # Set P_Active = toErr.atDsc
    # Set E_Active = toErr
    # Call BugTimer.BugState_UnPause                 ' restore any timer and other events

    # ERET:
    # Set fromDsc = Nothing
    # Set fromErr = Nothing
    # Set toErr = Nothing
    # aBugTxt = vbNullString                         ' no active DoVerify
    # aBugVer = True

# '---------------------------------------------------------------------------------------
# ' Method : DoVerify
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Stop if Debug Condition false (with optional message)
# '---------------------------------------------------------------------------------------
def doverify():

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Long

    # Dim DebugHalt As Boolean

    if Recursive > 0 Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if Recursive > 1 Then:
    # GoTo blockit
    if NoStop Then                             ' not an error within error:
    if DebugLogging Then:
    # blockit:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from DoVerify")
    if Not NoStop Then:
    print(Debug.Print String(OffCal, b) & "Message: " & Message)
    if DebugLogging Then:
    # DebugHalt = True
    # GoTo FuncExit
    # Recursive = Recursive + 1                      ' restored by    Recursive = False ProcRet:

    if IsMissing(NoStop) Then:
    # NoStop = aBugVer
    if LenB(Message) = 0 Then:
    # Message = aBugTxt
    if NoStop Then:
    # DoVerify = False
    # Message = "Verified OK:" & vbCrLf & Message
    if LogAllErrors Then:
    if InStr(Message, testAll) > 0 Then:
    # Call LogEvent(Message)
    if InStr(Message, "***") > 0 And DebugMode Then:
    # DebugHalt = True
    else:
    # DoVerify = True
    # E_AppErr.Reasoning = Message
    # Call BugTimerDeActivate                    ' Function must not Recurse! (Gated)
    if LenB(Message) = 0 Then:
    if E_Active Is Nothing Then:
    # Message = "*** Debug Stop requested in uninited state "
    else:
    # Message = "*** Debug Stop requested in " & E_Active.atKey
    if Recursive = 1 Then:
    if ShowMsgBox Then:
    else:
    print(Debug.Print Message)
    if InStr(Message, "***") > 0 Or DebugMode Then:
    # DebugHalt = True

    # FuncExit:
    # E_AppErr.Reasoning = Message
    if DebugHalt Then:
    # Debug.Assert False

    # zExit:
    # Recursive = Recursive - 1
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_PrintNameInfo
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_printnameinfo():

    # Dim msg As String

    # With CallEnv
    # msg = LString(Right(String(5, b) & i, 5) _
    # & b & .ProcedureName, OffObj - lCallInfo) _
    # & b & LString(.CallerInfo, lCallInfo) _
    # & RString(.StackDepth, 3) _
    # & RString("L" & .LineNumber, 5) _
    # & b & Trim(.LineCode)
    # Call LogEvent(msg, eLall)
    # End With                                       ' CallEnv


# '---------------------------------------------------------------------------------------
# ' Method : ShowLiveStack
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show all info contained in D_LiveStack Dictionary
# '---------------------------------------------------------------------------------------
def n_showlivestack():
    # Dim i As Long
    # Dim CallEnv As cCallEnv

    # Set CallEnv = D_LiveStack.Items(i)
    # Call N_PrintNameInfo(i + 1, CallEnv)

    # Set CallEnv = Nothing


# '---------------------------------------------------------------------------------------
# ' Method : Sub getDebugMode
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def getdebugmode():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "BugHelp.getDebugMode"
    # Dim zErr As cErr

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if StackDebug > 8 Then:
    # GoTo ProcRet

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet

    # Dim originalTestVar As String

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, Recursive:=False)

    if forceGet Then:
    # Testvar = GetEnvironmentVar("Test")        ' Variable "Test" im "1. sichtbaren" Environment

    # originalTestVar = Testvar

    if LenB(Testvar) > 0 Then:
    # Call N_InterpretTestVar
    else:
    # Call Z_StateToTestVar
    if StackDebugOverride > 0 Then:
    if StackDebugOverride <> StackDebug Then:
    # Call Z_StateToTestVar
    # StackDebugOverride = -StackDebug           ' do not override again

    if originalTestVar <> Testvar Then:
    print(Debug.Print "Changing Variable 'Test' to: " & Quote(Testvar))
    print(Debug.Print "Previously was " & String(14, b) & Quote(originalTestVar))

    # Call SetEnvironmentVariable("Test", Testvar)
    # Call SetGlobal("Test", Testvar)

    # ProcReturn:
    # Call ProcExit(zErr)

    # ProcRet:

# ' Method : N_CaptureNewErr
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Capture change of Err.Number and pass to proc that analyzes it
# ' Note   : when called via ErrEx, use ErrEx.Number instead
# '---------------------------------------------------------------------------------------
def n_capturenewerr():
    # Const zKey As String = "BugHelp.N_CaptureNewErr"

    if E_Active Is Nothing Then:
    # Set E_Active = New cErr
    # E_Active.atKey = "** undefined proc entering " & zKey & " **"

    # With E_Active
    if ErrExEvent Then:
    # .errNumber = T_DC.DCerrNum
    # .Description = T_DC.DCerrMsg
    # .Source = T_DC.DCerrSource
    else:
    # .errNumber = Err.Number
    # .Description = Err.Description
    # .Source = Err.Source

    if .errNumber = 0 Then                     ' shunt if no error. So the Entry is not logged in that case:
    # DoVerify False, "don't call N_CaptureNewErr if no error has occurred"
    # Call ErrReset(4)
    # GoTo ProcRet

    if ZErrSnoCatch Then                      ' N_PublishBugState Recursion not allowed:
    # aBugTxt = "**** previous error handling not complete when a new error occurred"
    # DoVerify .FoundBadErrorNr = 0
    # Call ErrReset(4)                       ' escape from this invalid state
    # StackDebug = 9                         ' Trace this!
    # GoTo ProcRet
    if MayChangeErr Then:
    # Err.Clear                              ' this error now cleared to get subsequent

    # ZErrSnoCatch = True                       ' No ErrHandler Recursions for:
    # .ErrSnoCatch = ZErrSnoCatch               ' NO N_PublishBugState
    if isEmpty(T_DC.DCAllowedMatch) Then       ' simple version testing acceptable errors:
    # GoTo trapIt
    elif Left(T_DC.DCAllowedMatch, 1) = "*" Then:
    # GoTo ProcRet
    elif T_DC.DCAllowedMatch = T_DC.DCerrNum Then:
    # GoTo ProcRet
    elif InStr(T_DC.DCerrMsg, T_DC.DCAllowedMatch) > 0 Then:
    # GoTo ProcRet

    # trapIt:
    # .FoundBadErrorNr = .errNumber
    # StackDebug = 9                             ' Trace this!
    # End With                                       ' E_Active

    if AppStartComplete Then:
    # Call ShowErr                               ' for testing only: forces frmErrStatus to show
    # frmErrStatus.fErrNumber = T_DC.DCerrNum
    # frmErrStatus.Top = 245
    # frmErrStatus.Left = 1041
    # Set aNonModalForm = frmErrStatus
    # ErrStatusFormUsable = True

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_ChkRC
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Test on Err. Exits false if err.number is 0, sets ErrorCaught always
# ' Note   : if Try allows several errors as defined by E_Active
# '        : Multiple N_ChkRC calls may follow, with ErrorCaught match.
# '        : All but the last one must then code DoClear:=False or ErrorCaught missing
# '---------------------------------------------------------------------------------------
def n_chkrc():

    # Const zKey As String = "BugHelp.N_ChkRC"

    # Dim Putlead As String

    if T_DC.TermRQ Then:
    # Call TerminateRun                          ' aborting Entry if not handled

    if E_Active Is Nothing Then:
    # N_ChkRC = True
    # GoTo ProcRet

    # With E_Active
    if .FoundBadErrorNr = 0 And Err.Number <> 0 Then:
    # .errNumber = Err.Number
    # .Description = Err.Description
    # .FoundBadErrorNr = Err.Number
    if .errNumber = 0 Then:
    # GoTo FuncExit                          ' no problem!
    if LenB(.atKey) = 0 Then:
    # .atKey = "** undefined proc entering " & zKey

    if Not IsMissing(TryingThisCode) Then:
    # .Explanations = TryingThisCode
    if Not IsMissing(WhatThisMeans) Then:
    # .Reasoning = WhatThisMeans
    if IsMissing(MatchAllow) Then:
    # MatchAllow = 0                         ' .Permitted is not changed!
    if Left(T_DC.DCAllowedMatch, 1) = "*" Then ' ignore any error:
    # N_ChkRC = True
    # .Permitted = Empty                 ' but now it is!
    # MatchAllow = T_DC.DCerrNum
    # .FoundBadErrorNr = 0
    # GoTo FuncExit                      ' so, this will just finish Error Try
    else:
    # N_ChkRC = Z_IsUnacceptable(.Permitted) ' is it bad or not?
    else:
    # .Permitted = MatchAllow
    # N_ChkRC = Z_IsUnacceptable(.Permitted) ' is it bad or not?

    # Putlead = String(OffCal, b)
    if Catch Then:
    if DebugMode Or DebugLogging Then:
    print(Debug.Print Putlead _)
    # & " *#'* " & .atKey _
    # & " has caused Error " & .FoundBadErrorNr _
    # & " '" & .atMessage & "'"
    if LenB(TryingThisCode) > 0 Then:
    print(Debug.Print Putlead & "Purpose:     " & TryingThisCode)
    if LenB(WhatThisMeans) > 0 Then:
    print(Debug.Print Putlead & "Explanation: " & WhatThisMeans)
    print(Debug.Print Putlead & "currently permittted: " & .Permitted)
    # Debug.Assert False

    if .Permitted = testOne Then        ' all Errors allowed, only once, err is returned:
    # .FoundBadErrorNr = 0
    # .Permitted = Empty              ' error capure is now off again
    elif .Permitted = testAll Then    ' all allowed, Permitted stays, err is returned:
    # .FoundBadErrorNr = 0
    elif .Permitted = allowAll Then   ' all allowed, Permitted stays, err is cleared:
    # .FoundBadErrorNr = 0
    elif .Permitted = allowNew Then   ' all Errors allowed, only once, err is cleared:
    # .FoundBadErrorNr = 0
    # .Permitted = Empty              ' error capure is now off again
    else:
    # End With                                ' E_Active

    # FuncExit:
    if IsMissing(MatchAllow) Then:
    # Call ErrReset(0)                           ' error Try block has ended, N_ChkRC NOT changing
    elif MatchAllow = 0 And DoClear Then:
    if T_DC.DCAllowedMatch = testAll Then         ' Keep Permitted ANYTHING:
    # Call ErrReset(4)                       ' error Try done, T_DC NOT changing
    else:
    # Call ErrReset(0)                       ' error Try scope end, DCAllowedMatch := Empty
    elif MatchAllow = 0 Then                     ' And NOT DoClear:
    if T_DC.DCAllowedMatch = testAll Then         ' Keep Permitted ANYTHING:
    # Call ErrReset(4)                       ' error Try done, T_DC NOT changing
    # ' ErrorCaught not changing
    else:
    # Call ErrReset(0)                       ' error Try scope end, DCAllowedMatch := Empty
    elif Not N_ChkRC And DoClear Then:
    # Call ErrReset(3)                           ' error Try block has ended and no app bugs found
    # Call T_DC.N_ClearTermination               ' this ends the scope of Try

    # ZErrSnoCatch = E_Active.ErrSnoCatch

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ClearAppErr
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub N_ClearAppErr()
# Const zKey As String = "BugHelp.N_ClearAppErr"

if MayChangeErr Then:
# Err.Clear
if E_AppErr Is Nothing Then:
# DoVerify False, "Clear an App without it being started ???"
# Set E_AppErr = New cErr                    ' no need to init, just clear
else:
# Call N_ErrClean(0)                         ' clear Err data in E_AppErr only

# '---------------------------------------------------------------------------------------
# ' Method : N_ConDscErr
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Set up a consistent atDsc / Err Pair. Uses no Stacks
# '---------------------------------------------------------------------------------------
def n_condscerr():

    # Const zKey As String = "BugHelp.N_ConDscErr"
    # Static zDsc As cProcItem

    # Dim msg As String
    # Dim ModuleName As String
    # Dim ModeNameN As String
    # Dim ModeLetterN As String

    # ModeNameN = QModeNames(Qmode)
    # ModeLetterN = UCase(Left(ModeNameN, 1))

    if LogZProcs And DebugLogging Then:
    # msg = zKey & " defines atDsc/Err for " & ClientKey _
    # & " Qmode=" & ModeLetterN
    # Call LogEvent(msg)

    # DoVerify LenB(ClientKey) > 0, "ClientKey must not be empty string"
    if ClientDsc Is Nothing Then:
    # Set ClientDsc = New cProcItem
    # With ClientDsc
    if Not .ErrActive Is Nothing Then          ' if previously chained, N_ConDscErr done before:
    if .ErrActive.atKey <> .Key Then:
    # Stop                                ' how that????????
    # .ErrActive.atKey = .Key
    if .CallMode < Qmode And .CallCounter <= 0 Then:
    # ClientDsc.CallMode = Qmode         ' allow mode upgrade z->y->x
    # GoTo reEntry
    # GoTo zExit
    # .ProcIndex = 0                             ' more details added below iff CallType<>""
    # .CallCounter = -1
    # .Key = ClientKey
    # .DbgId = RTail(ClientKey, ".", ModuleName)
    if LenB(ModuleName) > 0 Then:
    # .Module = ModuleName
    else:
    # .Module = "Extern"

    # ' Purpose: Construct a pair of linked ClientDsc/ClientErr
    # ' Note   : in case of conflict, ClientDsc wins

    if ClientErr Is Nothing Then:
    # Set ClientErr = New cErr
    # implicit:
    # Set ClientErr.atDsc = ClientDsc
    # ClientErr.atKey = ClientDsc.Key
    # Set ClientErr.atDsc = ClientDsc
    # ClientErr.NrMT = "--- u" & ModeLetterN
    # Set ClientDsc.ErrActive = ClientErr
    else:
    if ClientErr.atKey = vbNullString Then:
    # GoTo implicit
    # DoVerify ClientErr.atDsc Is ClientDsc
    # DoVerify ClientErr.atKey = ClientDsc.Key

    if LenB(CallType) > 0 Then:
    if LenB(ClientDsc.CallType) > 0 Then:
    # DoVerify ClientDsc.CallType = CallType, "change in CallType ???"
    else:
    # .CallType = CallType
    if .CallMode = Qmode Then:
    # .ModeLetter = ModeLetterN
    else:
    # DoVerify .CallMode = eQnoDef, " ** analyze mode change"
    # ClientDsc.CallMode = Qmode
    # reEntry:
    # With .ErrActive
    if Qmode <= eQxMode Then           ' y or z Mode: do not autocheck:
    # .atRecursionOK = True          ' recursion is checked by the proc using individual rules
    # .atProcIndex = inv
    # .atLastInDate = Date
    # .atLastInSec = Timer
    # .atThisEntrySec = Timer
    # Set .at_Live = New Dictionary
    # .EventBlock = E_Active.EventBlock
    # End With                               ' .ErrActive
    # Set ClientErr = .ErrActive
    # End With                                       ' ClientDsc
    # Call N_SetErrLvl(ClientDsc)

    # zExit:

# '---------------------------------------------------------------------------------------
# ' Method : N_DebugStart
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Initialize debug start values
# '---------------------------------------------------------------------------------------
def n_debugstart():

    if StackDebug = 0 Then:
    # StackDebug = 8                             ' >=4 all, 3 inline, 2 verbose, 1 normal StackDebug
    # StackDebugOverride = StackDebug
    if UseTestStart Then:
    # ExLiveCheckLog = True
    # ExLiveDscGen = True
    # TraceMode = True
    # CallLogging = True
    # LogZProcs = True                ' Due UseTestStart
    # LogAppStack = True
    # ShowFunctionValues = True
    # StackDebugOverride = -8
    elif StackDebug >= 8 Then:
    # TraceMode = True
    # CallLogging = True
    # LogZProcs = True                ' Due StackDebug >=8
    # LogAppStack = True
    # ShowFunctionValues = True
    elif StackDebug = 7 Then:
    # CallLogging = True
    # LogZProcs = False               ' Due StackDebug = 7
    # ShowFunctionValues = ShowFunctionValues Or UseTestStart
    # LogPerformance = False                     ' no performance data for 7 and up and 4 down
    elif StackDebug = 6 Then:
    # CallLogging = False                        ' no CallLogging for 6 and up and 4 down
    # LogZProcs = False               ' Due StackDebug = 6
    # LogPerformance = False
    elif StackDebug = 5 Then:
    # StackDebug = Abs(StackDebugOverride)
    # ExLiveCheckLog = False
    # ExLiveDscGen = False
    # TraceMode = False
    # CallLogging = False
    # LogZProcs = False               ' Due StackDebug = 5
    # LogAppStack = False
    # ShowFunctionValues = False
    # LogPerformance = False
    elif StackDebug = 4 Then:
    # StackDebug = Abs(StackDebugOverride)
    # ExLiveDscGen = True
    # LogZProcs = True                ' Due StackDebug = 4
    elif StackDebug = 3 Then:
    # StackDebug = Abs(StackDebugOverride)
    # LogZProcs = True                ' Due StackDebug = 3
    elif StackDebug = 2 Then:
    # StackDebug = Abs(StackDebugOverride)
    # LogPerformance = True
    elif StackDebug = 1 Then:
    # StackDebug = Abs(StackDebugOverride)
    # CallLogging = True
    elif StackDebug = -1 Then:
    # StackDebug = Abs(StackDebugOverride)
    # FastMode = True                            ' true on -1 only
    # ExLiveCheckLog = False                     ' true on UseTeststart and 5
    # ExLiveDscGen = False                          ' true on UseTestStart and 4, 5
    # TraceMode = False                          ' true on UseTestStart and 5, 8
    # CallLogging = False                        ' true on UseTestStart and 1, 7, 8
    # LogZProcs = False                          ' true on UseTestStart and 1, 3, 7, 8
    # LogAppStack = False                        ' true on UseTestStart and 8
    # ShowFunctionValues = False                 ' true on UseTestStart and 1, 8, kept on 7
    # LogPerformance = False                     ' true on UseTestStart and 2
    # UseTestStart = False                       ' false on -1
    if doStop And Not DidStop Then:
    # DidStop = True
    # doStop = False
    # Call BugSet


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ErrClean
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_errclean():
    # Const zKey As String = "BugHelp.N_ErrClean"

    # Dim ExplainS As String

    if Not E_AppErr.errNumber = 0 And (DebugMode Or DebugLogging) Then:
    # ExplainS = "Err Nr.=" & Err.Number
    if E_AppErr.FoundBadErrorNr <> 0 Then:
    # ExplainS = ExplainS & " (bad)"
    # ExplainS = "Reset Err, Lvl=" & ForceLevel & ", " & ExplainS

    if MayChangeErr Then:
    if LenB(ExplainS) > 0 Then:
    print(Debug.Print ExplainS)
    # Err.Clear
    else:
    if LenB(ExplainS) > 0 Then:
    print(Debug.Print ExplainS & ", Err not cleared=" & Err.Number)

    # With E_AppErr
    if .errNumber = 0 Then:
    if ForceLevel < 1 Then:
    # GoTo simple
    if ForceLevel = inv Then                   ' do not log at all and like ForceLevel = 0:
    # simple:
    # .errNumber = 0
    # .Description = vbNullString
    # .atFuncResult = vbNullString
    # .FoundBadErrorNr = 0
    # .Permitted = Empty
    # .Reasoning = vbNullString
    # GoTo ProcRet

    # ' all cases 0-2
    # .errNumber = 0
    # .Description = vbNullString
    # .atFuncResult = vbNullString
    # .FoundBadErrorNr = 0
    # .Permitted = Empty
    # .Reasoning = vbNullString

    match ForceLevel:
        case 0:
    # ' no further N_ErrCleaning in E_AppErr. (Keeps the ErrTry-State tErr way)
        case 1:
    # .Explanations = vbNullString
    # .Reasoning = vbNullString
        case 2:
    # DoVerify False, " whatfor and whatelse-just use: New"
    # .atProcIndex = inv
    # .atCallState = eCUndef
    # .atMessage = vbNullString
    # .atKey = vbNullString
    # .atShowStack = vbNullString
    # .DebugState = False
    # .Explanations = vbNullString
    # .Reasoning = vbNullString
    # .atRecursionOK = False
    # .atPrevEntrySec = 0
    # .atThisEntrySec = 0
    # .atTraceStackPos = 0
    # Set .atCalledBy = Nothing
    # Set .atDsc = Nothing
    # Set .atErrPrev = Nothing
    # .atCallDepth = 0
    # Set .atErrPrev = Nothing
        case _:
    # DoVerify False, " Invalid ForceLevel"
    # End With                                       ' E_AppErr
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ErrClear
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_errclear():
    # Const zKey As String = "BugHelp.N_ErrClear"

    # Call N_Suppress(Push, zKey)

    if T_DC Is Nothing Then:
    if rsp <> vbCancel Then:
    # Call ShowDbgStatus
    if IgnoreUnhandledError Then:
    # GoTo FuncExit
    # T_DC.N_ClearTermination
    # Call N_ErrClean(ForceLevel)

    # FuncExit:
    # Call N_Suppress(Pop, zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ErrInterfacePrint
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_errinterfaceprint():

    # Const zKey As String = "BugHelp.N_ErrInterfacePrint"

    # Dim smi1 As Long
    # Dim ClientDsc As cProcItem

    # Call N_Suppress(Push, zKey)
    if IsMissing(ErrClient) Then:
    # Set ErrClient = E_AppErr
    if ErrClient Is Nothing Then:
    print(Debug.Print "can't N_ErrInterfacePrint, no Data!")
    # GoTo FuncExit
    if ClientDsc.ProcIndex > inv Then:
    # Set ClientDsc = D_ErrInterface.Items(ErrClient.atProcIndex)
    else:
    # Set ClientDsc = D_ErrInterface.Items(2 - ErrClient.atProcIndex)

    # With ErrClient
    if .atCallState = eCActive Then:
    if ClientDsc.CallMode = eQArMode Then:
    # smi1 = .atCallDepth - 1
    # .atMessage = Right("    " & CStr(smi1), 5) & String(smi1, ">") _
    # & b & sExplain _
    # & vbCrLf & Right("    " & CStr(smi1), 5) & String(smi1, b) _
    # & " called by " _
    # & Replace(.atCalledBy.atKey, dModuleWithP, vbNullString) _
    # & " (using Stack, Call Depth=" & smi1 & ")"
    else:
    # sExplain = "ProcCall/Exit, No ErrHandler: " & .atKey
    # .atMessage = String(.atCallDepth, ">") _
    # & b & sExplain & b & vbCrLf & String(.atCallDepth, b) _
    # & " Caller: " & Replace(.atCalledBy.atKey, dModuleWithP, vbNullString)
    elif .atCallState = eCExited Then:
    # sExplain = Replace("Exit to: " & .atCalledBy.atKey _
    # & " from: " & .atKey, dModuleWithP, vbNullString)
    # .atMessage = String(.atCallDepth, "<") _
    # & b & sExplain & vbTab
    elif .atCallState = eCpaused Then:
    # sExplain = Replace("Paused: " & .atKey _
    # & " from: " & .atCalledBy.atKey, dModuleWithP, vbNullString)
    # .atMessage = String(.atCallDepth, "<") _
    # & b & sExplain & vbTab
    else:
    # sExplain = "never called: " & .atKey
    # .atMessage = sExplain

    print(Debug.Print .atMessage)

    # End With                                   ' ErrClient

    # FuncExit:
    # Call N_Suppress(Pop, zKey)


# '---------------------------------------------------------------------------------------
# ' Method : N_ErrStackLines
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Create a string to output for File or Immediate window
# '---------------------------------------------------------------------------------------
def n_errstacklines():
    # Const zKey As String = "BugHelp.N_ErrStackLines"

    # Dim MN As String
    # Dim PN As String
    # Dim extra As String

    # LiveErrors = Now() _
    # & " - " & CStr(ErrEx.Number) & " - " & CStr(ErrEx.Description)
    # LiveErrors = LiveErrors & ", Saved=" & ErrEx.SourceProjectIsSaved _
    # & ", VBEVersion=" & ErrEx.VBEVersion
    # ' separate the call stack to single lines in the log

    # With ErrEx.LiveCallstack
    # Do
    if .ModuleName = Mid(dModuleWithP, 2) Then ' Omitting default Module for readability:
    # MN = vbNullString
    else:
    # MN = .ModuleName & "."
    if 1 = 0 Then                          ' Omitting ProjectName in Outlook:
    # PN = .ProjectName & "."
    else:
    # PN = vbNullString
    if .ModuleName = dModule Then:
    # extra = String(4, b)
    # MN = vbNullString
    else:
    # extra = vbNullString
    # LiveErrors = LiveErrors & vbCrLf _
    # & extra & "       --> " & PN _
    # & MN _
    # & .ProcedureName & ", " _
    # & "#" & .LineNumber & ", " _
    # & .LineCode
    # Loop While .NextLevel
    # End With                                       ' ErrEx.LiveCallstack

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : N_GenCallData
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Set up the values for this Call
# '---------------------------------------------------------------------------------------
def n_gencalldata():
    # Dim Consumed As String
    # Dim TEntry As cTraceEntry

    # aErr.atFastMode = FastMode                     ' set this for correct Exit Mode

    if LogZProcs Then:
    if LogPerformance Then:
    # Call Z_UsedThisCall(aErr, aErr.atThisEntrySec)
    # Consumed = " Consumed " & ElapsedTime & " sec "
    else:
    # Consumed = vbNullString

    if Not dontIncrementCallDepth Then              ' (not predefining):
    # aErr.atCallDepth = E_Active.atCallDepth + 1

    if TraceMode Or dontIncrementCallDepth Then:
    # Set TEntry = New cTraceEntry
    # Set TEntry.TErr = aErr
    # Call N_TraceEntry(TEntry)                  ' generate trace Entry for ExternCaller
    # Set TEntry = Nothing

    if Not BugTimer Is Nothing And Not ShutUpMode Then:
    if BugTimer.BugStateReCheck Then:
    if Timer - BugTimer.BugStateLast _:
    # > 2 * BugTimer.BugStateTicks Then
    # Call BugTimerEvent("Call")

    if Not dontIncrementCallDepth Then              ' Predefining does not inc::
    # With aErr
    # ' for each call we increment .atRecursionLvl until we DoExit
    if .atRecursionLvl > 0 Then:
    # Set aErr = aErr.Clone           ' new instance due to recursion
    # aErr.atRecursionLvl = .atRecursionLvl + 1
    else:
    # .atRecursionLvl = 1                 ' no new instance (no recursion)
    if aDsc.MaxRecursions < E_Active.atRecursionLvl Then:
    # aDsc.MaxRecursions = E_Active.atRecursionLvl
    # End With                                    ' aErr

    # With aErr
    # Set .atCalledBy = E_Active
    # .atLastInDate = Date
    # .atThisEntrySec = Timer
    # .atPrevEntrySec = .atThisEntrySec
    # .atLastInSec = .atThisEntrySec
    if CallLogging Then:
    # Call N_ShowProgress(CallNr, aDsc, "+" & aDsc.ModeLetter, _
    # "zD=" & .atCallDepth - 1 & "" & .atCallDepth, _
    # Consumed, ErrClient:=aErr)
    # Set P_Active = aDsc                     ' replacing Caller with called Program
    # Set P_Active.ErrActive = aErr
    # Set E_Active = aErr
    # E_Active.atCallState = eCActive
    # P_Active.CallCounter = P_Active.CallCounter + 1
    # End With                                    ' aErr
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_GetLive
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Use Live Stack info to get relevant LiveStack elements Needed
# '---------------------------------------------------------------------------------------
def n_getlive():

    # Const zKey As String = "BugHelp.N_GetLive"

    # Dim limit As Long
    # Dim PN As String
    # Dim StackIndex As Long
    # Dim msg As String
    # Dim ClientDsc As cProcItem
    # Dim gotThem As Boolean
    # Dim ModuleName As String

    # Set LiveStack = New Collection
    if Needing = 0 Then:
    # limit = ErrLifeKept                        ' we never deliver more results than that
    else:
    # limit = Needing

    # msg = vbNullString
    # StackIndex = 1                                 ' start counting real position in LiveStack
    # With ErrEx.LiveCallstack
    # Do
    # Set ClientDsc = Nothing

    # PN = .ProcedureName
    # ModuleName = .ModuleName

    # Call Z_GetProcDsc(ModuleName & "." & PN, ClientDsc, msg)
    if ClientDsc Is Nothing Then           ' msg given from Z_GetProcDsc:
    # GoTo noDsc

    if ClientDsc.CallMode = eQnoDef Then   ' it's a dummy:
    if Filtered Then:
    # GoTo DidStep                   ' do not Show in LiveStack !

    # LiveStack.Add ClientDsc
    if LiveStack.Count >= limit Then:
    # gotThem = True
    # GoTo FuncExit                      ' we only that many (max)
    # DidStep:
    # StackIndex = StackIndex + 1
    # noDsc:
    # Loop While .NextLevel

    # End With                                       ' ErrEx.LiveCallstack

    if Not gotThem Then:
    if Logging Then:
    if CallLogging Then:
    print(Debug.Print msg)
    print(Debug.Print String(15, b) & "limit reached= " & limit _)
    # & ". found only " & LiveStack.Count _
    # & " relevant entries. Full Stack: ";
    # Call ShowLiveStack(doPrint:=True, _
    # tSubFilter:=False, getNewStack:=True, Full:=Filtered)
    # Debug.Assert False

    # FuncExit:
    # Set ClientDsc = Nothing

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_InterpretTestVar
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_interprettestvar():
    # Const zKey As String = "BugHelp.N_InterpretTestVar"

    # Dim i As Long

    if StackDebugOverride > 0 Then:
    # StackDebug = StackDebugOverride            ' until this is saved
    # ' Override: TraceMode = InStr(1, Testvar, "TraceMode", vbTextCompare) > 0
    # ' Override: LogPerformance = InStr(1, Testvar, "LogPerformance", vbTextCompare) > 0
    # GoTo FromTestVar
    else:
    if InStr(1, Testvar, "StackDebug", vbTextCompare) > 0 Then:
    # i = InStr(1, Testvar, "StackDebug=", vbTextCompare)
    if i > 0 Then:
    # StackDebug = Mid(Testvar, i + Len("StackDebug="), 2)
    # TraceMode = InStr(1, Testvar, "TraceMode", vbTextCompare) > 0
    # LogPerformance = InStr(1, Testvar, "LogPerformance", vbTextCompare) > 0
    # FromTestVar:
    # DebugMode = InStr(1, Testvar, "DebugMode", vbTextCompare) > 0
    # DebugLogging = InStr(Testvar, "LOG") > 0
    # LogAllErrors = InStr(Testvar, "ERR") > 0
    # ShowFunctionValues = InStr(Testvar, "ShowFunctionValues") > 0


# '---------------------------------------------------------------------------------------
# ' Method : N_LogEntry
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Print Application-Relevant Call
# '---------------------------------------------------------------------------------------
def n_logentry():
    # Const zKey As String = "BugHelp.N_LogEntry"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # Dim noPrint As Boolean
    # Dim Lvl As Long
    # Dim Caller As String
    # Dim ObjInfo As String
    # Dim addInfo As String

    # Call N_ShowHeader("BugHelp Log " & TimerNow)

    if StackDebug <= 9 Then:
    if StackDebug > 8 Then:
    if ClientDsc.CallMode <= eQxMode Then:
    # GoTo testHidden
    if ClientDsc.CallMode = eQzMode Then       ' covers  .., Z_.., and O_Goodies, Classes:
    # testHidden:
    if StackDebug > 4 Then:
    # GoTo FuncExit

    if Left(moreEE, 1) = "!" Then                  ' do not print:
    # noPrint = True
    # addInfo = Mid(moreEE, 2)
    else:
    # addInfo = moreEE

    # Caller = Replace(ClientDsc.Key, dModuleWithP, vbNullString)
    # ObjInfo = ObjStr

    # Lvl = S_AppIndex + 1

    # GenOut:

    # Call Z_Protocol(ClientDsc, CallNr, Caller, "==>", ObjInfo, addInfo)

    # FuncExit:
    # Recursive = False

    # zExit:

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_LogErrEx
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Log ErrEx Data to File
# '---------------------------------------------------------------------------------------
def n_logerrex():
    # Const zKey As String = "BugHelp.N_LogErrEx"
    # Dim zErr As cErr

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Not ErrExActive Then:
    # GoTo ProcRet
    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # Dim LogLines As Variant
    # Dim sL As Variant
    # Dim logLine As String
    # Dim InitialSkip As Boolean

    # Call N_ErrStackLines(logLine)
    # LogLines = split(logLine, vbCrLf)

    for sl in loglines:
    if Not InitialSkip Then                    ' skip lines at start of dump output:
    if InStr(sL, "BugHelp.Z") > 0 Then:
    # GoTo SkipLine
    else:
    # InitialSkip = True                 ' stop when useful info reached

    # logLine = sL
    print(Debug.Print logLine)
    # SkipLine:

    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_OnError
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Fehler Abfangen mit ErrEx > Global Error Handler <
# '---------------------------------------------------------------------------------------
def n_onerror():
    # Const zKey As String = "BugHelp.N_OnError"

    # Dim msg As String
    # Dim ForceDialog As Boolean
    # Dim Prefix As String

    # '------------------- gated Entry -------------------------------------------------------
    # BugState = ErrEx.State
    # BugStateAsStr = ErrEx.StateAsStr
    # BugDlgRsp = BugState                           ' no Dialog yet, default as same
    # BugDlgRspAsStr = BugStateAsStr
    # ' all error data now in E_Active
    # msg = ErrEx.Description
    if LenB(msg) > 0 Then:
    # T_DC.DCAppl = S_AppKey
    # T_DC.DCerrMsg = msg
    # T_DC.DCerrNum = ErrEx.Number
    # ErrorCaught = T_DC.DCerrNum
    # ErrEx.Number = 0

    # msg = ErrEx.SourceModule & "." & ErrEx.SourceProcedure
    # msg = msg & vbCrLf & "Line " & ErrEx.SourceLineNumber & ": " & ErrEx.SourceLineCode
    if isEmpty(T_DC.DCAllowedMatch) Then:
    # msg = msg & ", UnExpected"
    else:
    # msg = msg & ", Accepting=" & T_DC.DCAllowedMatch _
    # & " Error Number=" & T_DC.DCerrNum
    # msg = msg & "=&H" & Hex8(T_DC.DCerrNum) _
    # & vbCrLf & "   ErrorMessage: " & T_DC.DCerrMsg
    # msg = msg & vbCrLf & "   ErrStatus=" & BugState _
    # & "(" & BugStateAsStr & ")"
    if Left(T_DC.DCAllowedMatch, 1) = "*" _:
    # Or IsNumeric(T_DC.DCAllowedMatch) Then
    # Prefix = String(3, b)
    else:
    # Prefix = vbCrLf & String(3, "!! ")
    # T_DC.DCerrSource = msg

    if DebugMode Or StackDebug > 7 Then:
    # msg = Prefix & "N_OnError called, " & msg

    if StackDebug > 8 _:
    # Or BugState > 3 Then                    ' not handled at caller's level
    print(Debug.Print String(20, "-") _)
    # & vbCrLf _
    # & msg & "   must handle by Catch"
    elif E_Active.atCallDepth > 0 Then:
    if T_DC.DCAllowedMatch <> T_DC.DCerrNum _:
    # And Left(T_DC.DCAllowedMatch, 1) <> "*" Then
    # msg = String(80, b) & vbCrLf & msg
    # Call LogEvent(msg, eLSome)         ' not accepting
    # msg = vbNullString                         ' do not repeat report later

    # Call N_CaptureNewErr(True)                     ' handle error as defined
    if AppStartComplete Then:
    # Call N_PublishBugState                     ' also "locks" Err until exit

    # ForceDialog = DebugMode Or DebugLogging

    # With E_Active
    match BugState:

        case 0, OnErrorGoto0, OnErrorCatch, OnErrorCatchAll ' 0, 1, 10, 11:
    if Left(.Permitted, 1) = "*" Then:
    if .Permitted = "*" Then       ' all are allowed, err is returned:
    # T_DC.DCAllowedMatch = Empty ' error capure is now off again
    elif .Permitted = testAll Then  ' all allowed, Permitted stays, err is returned:
    elif .Permitted = allowAll Then ' all allowed, Permitted stays, err is cleared:
    # Call ErrReset(4)           ' cleans up error trace!!
    elif .Permitted = allowNew Then:
    # Call ErrReset(0)
    else:
    # ForceDialog = ForceDialog Or (StackDebug > 8)
    # .FoundBadErrorNr = 0
    # GoTo itsOK
    else:
    if isEmpty(.Permitted) Then:
    # ForceDialog = True
    # itsOK:
    # .Permitted = T_DC.DCAllowedMatch
    if Not ForceDialog Then:
    # ErrEx.State = OnErrorResumeNext
    # GoTo zExit

    # ' ---------------------------------------------------------------
    # ' Unhandled errors
    # ' ---------------------------------------------------------------
    # ForceDialog = True

        case OnErrorResumeNext                 ' 2:
    # ' ---------------------------------------------------------------
    # ' Ignore errors when On Error Resume Next is set
    # ' ---------------------------------------------------------------

        case OnErrorGotoLabel                  ' 3:
    # ' ---------------------------------------------------------------
    # ' Ignore locally handled errors, so go where instructed
    # ' ---------------------------------------------------------------

        case CalledByLocalHandler              ' 6:
    # ' ---------------------------------------------------------------
    # ' ErrEx.CallGlobalBugHelp was called

    # ' This is a special case for when local error handling was in use
    # ' but the local error handler has not dealt with the error and
    # ' so has passed it on to the global error handler
    # ' ---------------------------------------------------------------

        case OnErrorPropagate                  ' 8:
    # ' ---------------------------------------------------------------
    # ' Propagation would not cause ProcExit to handle stack!
    # ' So Propagation must not be used: Only locally handled errors
    # ' (otherwise handled in a previous routine in the call stack)
    # ' ---------------------------------------------------------------
    if LenB(BugWillPropagateTo) > 0 Then:
    # Stop                           ' error during Propagation
    # BugWillPropagateTo = vbNullString

    # With ErrEx.Callstack
    # .FirstLevel
    # Do
    if .HasActiveErrorHandler = True Then:
    # BugWillPropagateTo = .ProjectName _
    # & "." & .ModuleName & "." & .ProcedureName
    # Exit Do
    # Loop While .NextLevel
    # End With                           ' ErrEx.Callstack

        case OnErrorInsideFinally              ' 14 = &HE:

    # ' An error occurred inside the ErrEx.Finally block (typically for cleanup code).
    # ' We will use OnErrorResumeNext to skip over these
    # ErrEx.State = OnErrorResumeNext

        case Else                              ' 4, 5, 7, 9, 12 = &HC, 13 = &HD, 15 = &HF, 17 = &H11, 18 = &H12:
    print(Debug.Print "ErrEx.State " & BugState _)
    # & "=" & BugStateAsStr _
    # & " is not handled in N_OnError"
    # ForceDialog = True
    # Debug.Assert False
    # End With                                       ' E_Active

    if ForceDialog Then:
    if LenB(msg) > 0 Then:
    # Call LogEvent(msg, eLSome)
    # Call N_LogErrEx

    # BugDlgRsp = ErrEx.ShowErrorDialog

    if BugDlgRsp = OnErrorResumeNext Then:
    print(Debug.Print "Request to clear Error: " & BugDlgRsp _)
    # & " Debugging(" & ErrEx.StateAsStr & ")"
    # Err.Clear
    elif BugDlgRsp = OnErrorDebug Then:
    # Call ErrReset(4)                       ' clear this error logically (but recognizable)
    print(Debug.Print "T_DC Status not cleared: " & BugDlgRsp _)
    # & " Debugging(" & ErrEx.StateAsStr & ")" & b;
    print(Debug.Print "retry the erroneous statement: " & ErrEx.StateAsStr)
    # GoTo zExit
    elif BugDlgRsp = OnErrorEnd Then:
    print(Debug.Print "ShowErrorDialog Choice is: " & BugDlgRsp & b & ErrEx.StateAsStr)
    if Left(E_Active.Permitted, 1) <> "*" Then:
    # Stop                               ' continue possible: Press Ctrl-Shift-F8
    # ' If the close button is pressed on the error dialog, we don't
    # ' want to end abruptly (OnErrorEnd), but instead want the program
    # ' flow to continue in our local error handler:
    if BugState = OnErrorEnd Then:
    # ErrEx.State = CalledByLocalHandler
    print(Debug.Print "Another End Request from ShowErrorDialog: " _)
    # & BugDlgRsp & b & ErrEx.StateAsStr
    else:
    # ErrEx.State = OnErrorResumeNext
    print(Debug.Print "Request to resume from ShowErrorDialog:" _)
    # ; " & ErrEx.StateAsStr"
    # Call ErrReset

    # GoTo zExit
    else:
    # SetState:
    # ErrEx.State = BugDlgRsp
    # BugDlgRspAsStr = ErrEx.StateAsStr
    print(Debug.Print "ShowErrorDialog Choice is: " _)
    # & BugDlgRsp & " (" & BugDlgRspAsStr & ")"

    # FuncExit:
    # IgnoreUnhandledError = False                   ' now we care about Unhandled again
    # MayChangeErr = True

    # zExit:
    # BugStateAsStr = ErrEx.StateAsStr


# ' ----------------------------------------------------------------
# ' Procedure Name: N_PopLog
# ' Purpose: Pop stack into LogAppStack
# ' Procedure Kind: Sub
# ' Procedure Access: Public
# ' Author: Rolf G. Bercht
# ' Date: 29.06.2017
# ' ----------------------------------------------------------------
def n_poplog():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                            ' Simple proc
    # Const zKey As String = "BugHelp.N_PopLog"

    # Dim oldCount As Long

    # oldCount = C_PushLogStack.Count
    # DoVerify oldCount > 0, "** nothing on stack to pop"
    # LogAppStack = C_PushLogStack.Item(oldCount)
    # C_PushLogStack.Remove oldCount


# ' ----------------------------------------------------------------
# ' Procedure Name: N_PopSimple
# ' Purpose       : Pop stack state variant from top of the stack C_AllPurposeStack
# ' Author        : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' ----------------------------------------------------------------
def n_popsimple():
    # Const zKey As String = "BugHelp.N_PopSimple"

    # Dim oldCount As Long

    if useStack Is Nothing Then:
    # Set useStack = C_AllPurposeStack
    # oldCount = useStack.Count

    if VariantObject Then:
    # Set aVarO = useStack.Item(oldCount)
    else:
    # aVarO = useStack.Item(oldCount)
    # useStack.Remove oldCount

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : N_Prepare
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Prepare indispensable variables used during initializations
# '---------------------------------------------------------------------------------------
def n_prepare():

    # Dim nextField As String
    # Dim i As Long

    # lHeadLine = 156

    # testNonValueProperties = "Nothing Empty"
    # ScalarTypes = split(ScalarTypeNames)
    # ScalarTypeV = Array(vbInteger, vbLong, vbSingle, vbDouble, _
    # vbDate, vbString, vbBoolean, 20&)
    # Set dSType = New Dictionary
    # dSType.Add InStr(ScalarTypeNames & b, _
    # ScalarTypes(i) & b), ScalarTypeV(i)

    # dModuleWithP = dModule & "."
    # QModeNames = split(QModeString)
    # QModeNames(i) = LString(QModeNames(i), 7)  ' 6+1 b
    # OkValueNames = split(OkValueString)

    # CStateNames = split(CStateString)
    # CStateNames(i) = LString(CStateNames(i), 8)

    # LogLevelNames = split(LogLevelString)
    # AccountTypeNames = split(AccountTypeString)
    # PushTypes = split(PushTypeString)
    # ExStackProcNames = split(ExStackProcString)
    # ExModeNames = split(ExModeNamesString)

    # Call N_SetErrExHdl

    # ModelLine(1) = "Call# Pr# MTS Lvl "
    # OffPrN = InStr(ModelLine(1), "Pr#")
    # OffMTS = InStr(ModelLine(1), "MTS")
    # OffLvl = InStr(ModelLine(1), "Lvl")
    # OffCal = Len(ModelLine(1))                     ' offset of Head Line Call-Column

    # ModelLine(3) = ModelLine(1) & LString("Caller", lKeyM - 15)
    # ModelLine(1) = ModelLine(1) & LString("Caller", lKeyM)
    # OffObj = Len(ModelLine(1))

    # ModelLine(2) = Left(LString("Call#", 5) _
    # & b & LString("Caller", OffObj - 6 - lCallInfo) _
    # & b & LString("CallerInfo", lCallInfo) _
    # & "sD=  -- Code " _
    # & String(lDbgM + lKeyM, "-"), lHeadLine)


    # ModelLine(1) = ModelLine(1) & Left("----- Object " & String(lDbgM, "-"), lDbgM + 7) & b
    # OffTim = Len(ModelLine(1))

    # ModelLine(1) = ModelLine(1) & nextField
    # OffAdI = Len(ModelLine(1))

    # ModelLine(3) = ModelLine(3) & nextField
    # ModelLine(3) = ModelLine(3) & "Lne -- Code " & String(lDbgM, "-")

    # ModelLine(1) = LString(ModelLine(1) & nextField, lHeadLine)

    # ModelLine(3) = LString(ModelLine(3) & String(9, "-") & nextField, lHeadLine)
    # Call N_ShowHeader("BugHelp Log " & TimerNow)

    # Call N_PreDefine

    if WannaKnow Then:
    print(Debug.Print "OffPrN=" & OffPrN, "OffMTS=" & OffMTS, "OffLvl=" & OffLvl, _)
    # "OffCal=" & OffCal, "OffObj=" & OffObj, _
    # "OffTim=" & OffTim, "OffAdI=" & OffAdI, _
    # "lHeadLine=" & lHeadLine
    print(Debug.Print ModelLine(1))
    print(Debug.Print ModelLine(2))
    print(Debug.Print ModelLine(3))


# '---------------------------------------------------------------------------------------
# ' Method : N_PublishBugState
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Interprets Error information from E_Active and then
# '          Decide what to do (using Z_IsUnacceptable)
# '             sets this into T_DC and E_AppErr
# '             if acceptable error, proceed,
# '             else, or if debugmode, Show Error Status form
# ' Note   : Data from N_OnError->N_CaptureNewErr, which puts Err->E_Active
# '---------------------------------------------------------------------------------------
def n_publishbugstate():
    # Const zKey As String = "BugHelp.N_PublishBugState"
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet

    if E_Active Is Nothing Then:
    print(Debug.Print "* Error before error-setup complete")
    print(Debug.Print RetLead & zKey & " -> " & S_DbgId)
    # Call TerminateRun

    # With E_Active
    if IgnoreUnhandledError Then:
    if DebugMode Then:
    print(Debug.Print "N_PublishBugState called when IgnoreUnhandledError = True")
    # DoVerify False
    # GoTo FuncExit
    if Err.Number <> 0 Then:
    print(Debug.Print "* unhandled error between error handling calls. Manual intervention?")
    # DoVerify False
    # GoTo FuncExit
    # Resume Next                            ' on manual intervention only: analyze problem area
    # Recursive = True


    # E_AppErr.errNumber = .errNumber            ' to inform outside world (App lvl)
    # E_AppErr.Description = .Description
    # E_AppErr.Source = .Source

    # IgnoreUnhandledError = True                ' anything called until Exit will not check Unhandled

    # End With                                       ' E_Active

    # FuncExit:
    # Recursive = False

    # ProcRet:

# ' ----------------------------------------------------------------
# ' Procedure Name: N_PushLog
# ' Purpose: Push state of LogAppStack
# ' Author: Rolf G. Bercht
# ' Date: 29.06.2017
# ' ----------------------------------------------------------------
def n_pushlog():
    # Const zKey As String = "BugHelp.N_PushLog"

    # C_PushLogStack.Add LogAppStack
    # LogAppStack = NewState


# ' ----------------------------------------------------------------
# ' Procedure Name: N_PushSimple
# ' Purpose: Push state variant on top of the stack C_AllPurposeStack
# ' Author: Rolf G. Bercht
# ' Date: 29.06.2017
# ' ----------------------------------------------------------------
def n_pushsimple():

    # Const zKey As String = "BugHelp.N_PushSimple"

    if useStack Is Nothing Then:
    # Set useStack = C_AllPurposeStack
    if (DebugMode Or aDebugState) And useStack.Count > 0 Then:
    if VariantObject Then:
    # DoVerify useStack(useStack.Count) Is OldState
    else:
    # DoVerify useStack(useStack.Count) = OldState

    # useStack.Add OldState
    if VariantObject Then:
    # Set OldState = NewState
    else:
    # OldState = NewState

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_SetErrLvl
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_seterrlvl():
    # Const zKey As String = "BugHelp.N_SetErrLvl"

    # With ClientDsc
    match .CallMode:
        case eQnoDef                           '  = 0:
    # .ErrLevel = eLdebug                ' = 0
        case eQzMode                           '  = 1:
    # .ErrLevel = eLdebug                ' = 0
        case eQyMode                           '  = 2:
    if .ErrLevel = 0 Then:
    # .ErrLevel = eLmin              ' = 3 + StackDebug > 9 only
    else:
    # .ErrLevel = eLmin
        case eQxMode                           '  = 3:
    # .ErrLevel = eLmin                  ' = 3 + StackDebug > 9 only
        case eQrMode                           '  = 4:
    # .ErrLevel = eLmin                  ' = 3
        case eQuMode                           '  = 5:
    # .ErrLevel = eLmin                  ' = 3
        case eQEPMode                          '  = 6:
    # .ErrLevel = eLall                  ' = 1
        case eQAsMode                          '  = 7:
    # .ErrLevel = eLmin                  ' = 3
        case eQArMode                          '  = 8:
    # .ErrLevel = eLSome                 ' = 2
        case _:
    # DoVerify False, " this is an incorrect Qmode"
    # End With                                       ' ClientDsc


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ShowErrInstance
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_showerrinstance():
    # Dim cPos As Long
    # Dim CommentPart As String
    # Dim ActionPart As String
    # Dim ExplainS As String

    # With ErrClient

    # cPos = InStr(.atShowStack, "#")
    if cPos > 0 Then:
    # CommentPart = Trim(Mid(.atShowStack, cPos + 1))
    # ActionPart = Trim(Left(.atShowStack, cPos - 1))
    else:
    if LenB(.atShowStack) > 0 Then:
    # CommentPart = Trim(.atShowStack)
    else:
    # CommentPart = Trim(.Reasoning)

    # ExplainS = CommentPart

    # Call N_ShowProcDsc(ErrClient.atDsc, IndexNr, WithInstances:=False, _
    # ExplainS:=ExplainS, ErrClient:=ErrClient)
    # End With                                       ' ErrClient


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ShowHeader
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_showheader():
    # Static lListName As Long
    # Static midOff As Long
    # Static modelNr As Long
    # Dim EndLine As String

    if force Or lListName = 0 Then:
    # lHeadLine = 0
    # MinusLine = vbNullString
    # HeadlineName = vbNullString
    # lListName = Len(ListName)
    # midOff = Round((lListName + 1) / 2)
    if modelNr = 0 Then:
    # modelNr = ModelType
    # lHeadLine = Len(ModelLine(modelNr))
    # lMinus = (lHeadLine - midOff) / 2 - 1

    if StartOrEnd Then                             ' StartOrEnd=True  ==> (=End) Force:
    # EndLine = Left(MinusLine, lMinus + lListName * 2 + 1) _
    # & " -- End " & String(lMinus, "-")
    # Call LogEvent(Left(EndLine, lHeadLine), eLall)
    # HeadlineName = vbNullString                ' force a new value for Headline next call
    elif HeadlineName <> ListName Then:
    if lHeadLine > midOff Then:
    # lListName = Len(ListName)
    # midOff = Round((lListName + 1) / 2)
    # lMinus = (lHeadLine - midOff) / 2 - 1
    # MinusLine = Left(String(lMinus, "-") & b & ListName & b _
    # & String(lMinus, "-"), lHeadLine)

    # Call LogEvent(MinusLine, eLall)        ' printing with LogEvent
    # Call LogEvent(ModelLine(ModelType), eLall)
    # HeadlineName = ListName                ' remember for change detection
    # modelNr = ModelType
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_ShowInstances
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show the instances of default or specified cProcItem
# '---------------------------------------------------------------------------------------
def n_showinstances():
    # Dim ErrClient As cErr

    # Dim instCount As Long
    # Dim msg As String
    # Dim ThisHeadLine As String

    if ClientDsc Is Nothing Then:
    print(Debug.Print "no ProcItem")
    # GoTo FuncExit

    if withItemDsc Then:
    # Call N_ShowProcDsc(ClientDsc, Ordinal)
    # Set ErrClient = ClientDsc.ErrActive

    if ErrClient Is Nothing Then:
    # DoVerify False
    # GoTo FuncExit

    if ErrClient.atTraceStackPos > 0 Then:
    # msg = ClientDsc.DbgId & " beginning at TracePos=" _
    # & ErrClient.atTraceStackPos _

    else:
    # msg = LString(ClientDsc.DbgId, lDbgM + 5)

    # ThisHeadLine = " Call Chain Instances belonging to " & msg

    # Call N_ShowHeader(ThisHeadLine, force:=True)

    if ErrClient Is Nothing Then:
    # Exit For
    # Call N_ShowErrInstance(ErrClient, ErrClient.atCallDepth)
    if ErrClient Is ErrClient.atErrPrev Then:
    # Exit For
    # Set ErrClient = ErrClient.atErrPrev

    # Call N_ShowHeader(ThisHeadLine, StartOrEnd:=True)

    # FuncExit:
    # Set ErrClient = Nothing

# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ShowProcDsc
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_showprocdsc():
    # Const zKey As String = "BugHelp.N_ShowProcDsc"

    # Dim sProc As String
    # Dim ObjInfo As String
    # Dim Consumed As String
    # Dim ObjStatus As String
    # Dim Tmsg As String
    # Dim rL As Long
    # Dim StackState As String

    if ClientDsc Is Nothing Then:
    if ErrClient Is Nothing Then:
    # Tmsg = "! ! ! There is an empty element on the stack in position " & i
    # GoTo FuncExit
    else:
    # Set ClientDsc = ErrClient.atDsc
    if ClientDsc Is Nothing Then:
    # Tmsg = "! ! ! unspecified Client in position " & i
    # GoTo FuncExit
    else:
    if ErrClient Is Nothing Then:
    # Set ErrClient = ClientDsc.ErrActive
    if ErrClient Is Nothing Then:
    # ObjInfo = " E"
    else:
    if Not ClientDsc Is ErrClient.atDsc Then:
    # ObjInfo = ObjInfo & "!"
    # DoVerify False
    if Not ErrClient Is ClientDsc.ErrActive Then:
    if ErrClient.atRecursionOK Then:
    # ObjInfo = ObjInfo & " "
    else:
    # ObjInfo = ObjInfo & " ??"     ' recursion instance call
    # ' DoVerify False
    # followChain:
    if ErrClient.atErrPrev Is Nothing Then:
    if ErrClient.atRecursionLvl > 1 Then:
    # ObjInfo = ObjInfo & "Chain?"
    # DoVerify False, "ErrPrev-Chain leaves recursion"
    else:
    # ObjInfo = ObjInfo & vbCrLf & vbTab & vbTab & "recursion from " _
    # & ErrClient.atErrPrev.atCalledBy.atKey
    else:
    # Set ErrClient = ErrClient.atErrPrev ' follow chain
    if LenB(ErrClient.Reasoning) > 0 Then:
    if InStr(" LsT", Left(ErrClient.Reasoning, 1)) = 0 Then:
    # ObjInfo = ObjInfo & "X"

    if ClientDsc.ErrActive Is Nothing Then:
    # Tmsg = LString(i, 5) & String(OffCal - 5, b) _
    # & ClientDsc.DbgId & " is an invalid Entry: .ErrActive Is Nothing."
    # GoTo FuncExit

    if LenB(ExplainS) = 0 Then:
    if E_AppErr Is ErrClient Then              ' ErrClient is the last in the loop:
    # ObjStatus = " current App"
    else:
    # ObjStatus = vbNullString
    else:
    if InStr(ExplainS, "IdsOK") > 0 Then       ' mode is Dump D_ErrInterface:
    # Tmsg = ExplainS & b
    else:
    # ObjStatus = ExplainS
    # ExplainS = vbNullString

    if ClientDsc.TotalRunTime > 0 Then:
    # Consumed = "rT=" & RString(ClientDsc.TotalRunTime, 6)
    if Not ErrClient.atCalledBy Is Nothing Then:
    if ErrClient.atCalledBy.atDsc Is Nothing Then:
    # Consumed = Trim(Consumed & " C: NoCaller")
    else:
    # Consumed = Trim(Consumed & " C: " & ErrClient.atCalledBy.atKey)

    if LenB(Caller) = 0 Then:
    # Caller = ClientDsc.DbgId

    # Call N_Suppress(Push, zKey, Value:=False)
    if Consumption Then:
    # ObjStatus = Consumed
    # Call N_ShowProgress(i, ClientDsc, Caller, _
    # ObjectMsg:=StackState & ObjInfo & ObjStatus, _
    # ExplainS:=sProc & b & ExplainS, _
    # ErrClient:=ErrClient, _
    # doPrint:=True)
    else:
    # Call N_ShowProgress(i, ClientDsc, Caller, _
    # ObjectMsg:=StackState & ObjInfo & ObjStatus, _
    # ExplainS:=sProc & Trim(b & ExplainS) & ": " & Consumed, _
    # ErrClient:=ErrClient, _
    # doPrint:=True)
    # Call N_Suppress(Pop, zKey)
    if WithInstances Then:
    # Call N_ShowInstances(ClientDsc)

    # FuncExit:
    if LenB(Tmsg) > 0 Then:
    print(Debug.Print Tmsg)


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ShowProgress
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_showprogress():
    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    # GoTo ProcRet                               ' simply ignore recursion
    if ShutUpMode Then:
    # GoTo ProcRet
    # Recursive = True

    # Dim P As Long
    # Dim px As String
    # Dim Target As String

    # Output = vbNullString

    if ErrClient Is Nothing Then:
    # Set ErrClient = ClientDsc.ErrActive

    if InStr(ObjectMsg, ".") > 0 Then:
    # GoTo PutDsc

    # P = InStr("+-", Left(ObjKey, 1))
    if P > 0 Then                                   ' only 0, 1, 2 possible:
    # px = Mid(ObjKey, 2, 1)
    # ObjKey = Mid(ObjKey, 3)

    if px = "D" Then                            ' +D_DefProc ...:
    # ObjectMsg = "#" & px & "=" & LString(ClientDsc.ProcIndex, 5) & ObjectMsg
    # P = 1                                   ' never removed
    elif px = "T" Then                        ' +C_CallTrace:
    # ObjectMsg = "#" & px & "=" & LString(TraceTop + 1, 5) & ObjectMsg
    # P = 1                                   ' never removed
    elif InStr("NZYXRQESR", px) > 0 Then      ' calls/exits:
    # Proc:
    if P = 2 Then                           ' "-": do not double info for exit:
    # Target = ErrClient.atCalledBy.atDsc.DbgId     ' ... Show only on exit
    # DoVerify LenB(Target) <> 0, _
    # "on exit, Target must be defined ???"
    # ObjKey = LRString(Target, "<= " & ErrClient.atDsc.DbgId, _
    # ErrClient.atCallDepth, lKeyM, rCutL:=Len(Target) + 3)
    # P = P + 1                           ' indent out for exits
    # GoTo PutDsc

    # Output = ClientDsc.DbgId & b & Quote(ClientDsc.Module, Bracket)
    # ObjKey = LRString(ObjKey & b & Output, _
    # Target, E_Active.atCallDepth, lKeyM)
    else:
    # ObjKey = String(ErrClient.atCallDepth, b) & ObjKey

    # PutDsc:
    # Output = RString(Line, 5) & b _
    # & ErrClient.NrMT _
    # & Left(CStateNames(ErrClient.atCallState), 1) & b _
    # & RString(ErrClient.atCallDepth _
    # + ErrClient.atLiveLevel - LSD - P + 1, 3) & b _
    # & LString(ObjKey, lKeyM) _
    # & ObjectMsg ' Indent-value includes Live Stack depth with Offset LSD

    # Output = LString(Output, OffTim) & b & LString(ErrClient.atLastInSec, 14) _
    # & " rL=" & ErrClient.atRecursionLvl _
    # & " cc=" & ClientDsc.CallCounter

    if InStr(Result, "IdsOK") = 0 Then:
    # Result = Output & b & ExplainS
    else:
    # Result = Result & b & Output

    if LenB(Result) > 0 Then:
    # Line = Line + 1
    if doPrint Then:
    # Call LogEvent(Result, eLSome)          ' this saves/restores .Permitted so it is not changed
    else:
    print(Debug.Print Result)

    # Recursive = False

    # ProcRet:

# ' a=all, b=both AllErr, AppStack and CallTrace,
# ' c=CallTrace, e=Defproc + AllErr
# ' IMPLEMENT: p=Perf
def n_showstacks():
    # Const zKey As String = "BugHelp.N_ShowStacks"

    # Call N_Suppress(Push, zKey)

    if what = "a" Then:
    # Call ShowDefProcs(WithInstances:=True, Full:=Full)
    # what = "b"

    if what = "b" Then:
    # what = "p"

    if what = "e" Then:
    if Full Then                               ' not full version shows the same as ShowErrStack:
    # Call ShowDefProcs(False, Full:=True)   ' so omit
    # Call ShowErrStack

    if what = "b" Then:
    # Call ShowErrStack
    # Call ShowCallTrace
    else:
    if what = "c" Then:
    # Call ShowCallTrace

    # Call N_Suppress(Pop, zKey)

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : Function N_Stackdef
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Set up the atDsc and Err for new items and stack these. Returns True for FirstItem called.
# '---------------------------------------------------------------------------------------
def n_stackdef():
    # Const zKey As String = "BugHelp.N_StackDef"
    # Const MyId As String = "N_StackDef"
    # Const MyClass As String = "BugHelp"

    if sDsc.ProcIndex = 0 Or sErr.atProcIndex = 0 Then:
    if D_ErrInterface.Exists(sDsc.Key) Then:
    if isEmpty(D_ErrInterface.Item(sDsc.Key)) Then:
    # DoVerify sDsc.ProcIndex = 0, sDsc.Key & " already defined at Position " & sDsc.ProcIndex
    # Set D_ErrInterface.Item(sDsc.Key) = sDsc
    # GoTo corrItem
    if sErr.atKey = vbNullString Then:
    # Set sErr = sDsc.ErrActive
    if DoVerify(D_ErrInterface.Item(sDsc.Key).ErrActive Is sErr, _:
    # "D_ErrInterface item(" & sDsc.Key & ") is incorrect") Then
    # Set D_ErrInterface.Items(sDsc.Key) = sErr.atKey ' conflict, sErr wins
    else:
    # D_ErrInterface.Add sDsc.Key, sDsc
    # corrItem:
    # sDsc.ProcIndex = D_ErrInterface.Count
    if sErr.atProcIndex <> sDsc.ProcIndex Then  ' this includes atProcIndex = Inval:
    # sErr.atProcIndex = sDsc.ProcIndex       ' conflict, atDsc wins
    if CallLogging And (LogZProcs Or dontIncrementCallDepth) Then:
    if dontIncrementCallDepth Then:
    # sErr.Explanations = "Predefining"
    # Call N_ShowProgress(CallNr, sDsc, "+D", _
    # "creator=" & P_Active.DbgId, _
    # sErr.Explanations, ErrClient:=sErr)

    if IgnoreCallData Then:
    # N_StackDef = True
    # sDsc.CallCounter = sDsc.CallCounter + 1
    else:
    if D_ErrInterface.Count <= 2 Then          ' real calls follow Extern.Caller and DoCall:
    # sDsc.CallCounter = sDsc.CallCounter + 1
    # sErr.atCallState = eCExited
    # N_StackDef = True                      ' Omitting N_GenCallData
    else:
    # Call N_GenCallData(sDsc, sErr)

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_Suppress
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Push or Pop Shutupmode and SuppressStatusFormUpdate (if allowed)
# '---------------------------------------------------------------------------------------
def n_suppress():
    # Dim IdentifiedVal As cPair
    # Dim LastEntry As Long
    # Dim PrevState As Boolean
    # Static LogMe As Boolean

    if ProtectStackState = inv Then:
    # Set C_ProtectedStack = New Collection
    if C_ProtectedStack.Count = 0 Then:
    # ProtectStackState = StackDebug

    if ProtectStackState < 1 Then                  ' do not ShutUp:
    # ShutUpMode = False                         ' Parameter "Value" is ignored
    # GoTo msgout                                ' Not on Entry Nor Exit
    if ProtectStackState > 8 Then                  ' do not change Shutupmode at all:
    # msgout:
    if LogMe Then:
    print(Debug.Print "Ignored! Caller=" & Caller, "ShutUpValue==" & ShutUpMode)
    # GoTo FuncExit                              ' Not on Entry Nor Exit

    # DoStacking:

    if Push Then:
    if LogMe Then:
    print(Debug.Print "Pushing Caller=" & Caller, "Count=" _)
    # & C_ProtectedStack.Count, "ShutUpValue=" _
    # & ShutUpMode & "->" & Value
    # Set IdentifiedVal = New cPair
    # IdentifiedVal.pValue = ShutUpMode
    # SuppressStatusFormUpdate = ShutUpMode
    # IdentifiedVal.pId = Caller
    # C_ProtectedStack.Add IdentifiedVal
    # ShutUpMode = Value
    else:
    # LastEntry = C_ProtectedStack.Count
    if LastEntry = 0 Then:
    print(Debug.Print "* Pop by Caller " _)
    # & Caller & " impossible, no items to pop: ", _
    # "Count=" & C_ProtectedStack.Count, _
    # "ShutUpValue(unch.)=" & ShutUpMode
    # DoVerify False
    else:
    # Set IdentifiedVal = C_ProtectedStack.Item(LastEntry)
    # DoVerify IdentifiedVal.pId = Caller, " popper must be pusher: " & Caller
    # PrevState = ShutUpMode
    # ShutUpMode = IdentifiedVal.pValue
    # SuppressStatusFormUpdate = ShutUpMode
    if LogMe Then:
    print(Debug.Print "Pop by  Caller=" & Caller, "Count=" _)
    # & C_ProtectedStack.Count, "ShutUpValue=" _
    # & ShutUpMode & "" & PrevState
    if LastEntry > 0 Then:
    # C_ProtectedStack.Remove LastEntry

    # FuncExit:
    # Set IdentifiedVal = Nothing
    # SuppressStatusFormUpdate = Not ShutUpMode


# '---------------------------------------------------------------------------------------
# ' Method : N_TraceCallReduce
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: At end of lifetime, drop elements from CallTrac
# '---------------------------------------------------------------------------------------
def n_tracecallreduce():

    # Const zKey As String = "BugHelp.N_TraceCallReduce"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim i As Long
    # Dim j As Long
    # Dim StackErr As cErr
    # Dim Removed As Long
    # Dim stopRemove As Boolean
    # Dim TEntry As cTraceEntry
    # Dim TSucc As Long

    if C_CallTrace.Count > ErrLifeTime Then        ' clear if too much history:
    if StackDebug > 8 Then:
    print(Debug.Print "* Cleaning C_CallTrace because LifeTime reached: " & ErrLifeTime)
    # Removed = 0                                ' for debug restart
    # j = i - Removed
    # Set TEntry = C_CallTrace.Item(j)
    # Set StackErr = TEntry.TErr
    if ErrLifeKept >= Removed And Not stopRemove Then:
    if TEntry.TRL = 0 Then:
    # C_CallTrace.Remove j           ' clean up outdated, leaving #1 unchanged
    # j = inv                        ' mark the Instance as outdated, removing from C_CallTrace
    # Removed = Removed + 1
    else:
    if TEntry.TLne > 0 Then:
    # 'stopRemove = True                  ' leave intact otherwise *** ???
    # StackErr.atTraceStackPos = j           ' correct this position
    # TSucc = TEntry.TSuc - Removed
    if TSucc > 0 Then:
    # C_CallTrace.Item(TSucc).TPre = Removed - i ' used to be there before...

    if TraceMode Then:
    if C_CallTrace.Count - ErrLifeKept - 1 > 0 Then:
    print(Debug.Print "* N_TraceCallReduce was not able to remove " _)
    # & C_CallTrace.Count - ErrLifeKept - 1 _
    # & " Items, CallTrace.Count=" _
    # & C_CallTrace.Count
    # DoVerify C_CallTrace.Count < ErrLifeTime, " REALLY bad!! ErrLifeTime may be too small"
    # Call ShowStatusUpdate

    # Set StackErr = Nothing
    # Set TEntry = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_TraceDef
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def n_tracedef():

    # Dim ClientDsc As cProcItem
    # Dim ClientErr As cErr
    # Dim PreEntry As cTraceEntry
    # Dim Caller As String
    # Dim Source As String

    # With TEntry
    # .TLog = vbNullString
    # Set ClientErr = .TErr
    # Set ClientDsc = ClientErr.atDsc
    # .TLD = ClientErr.atLastInDate              ' trace<>non-instance vals in cErr
    # .TLS = ClientErr.atLastInSec
    # .TPS = ClientErr.atPrevEntrySec
    # .TES = ClientErr.atThisEntrySec

    # .TSrc = ClientDsc.DbgId
    # Source = .TSrc & b & Quote(ClientDsc.Module, Bracket)
    # .TLog = LRString(Caller, LString(.TES, 14), _
    # ClientErr.atCallDepth, lKeyM - 1) & b _
    # & LString(.TLne, 3) & b _
    # & LString(Source, lHeadLine - OffObj) & b _
    # & ClientErr.Explanations
    # End With                                       ' TEntry

    # Set ClientDsc = Nothing
    # Set ClientErr = Nothing
    # Set PreEntry = Nothing

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : N_TraceEntry
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Trace the current stack top. Obviously can't trace itself or recurse
# '---------------------------------------------------------------------------------------
def n_traceentry():

    # Const zKey As String = "BugHelp.N_TraceEntry"
    # Const MyId As String = "N_TraceEntry"

    # Dim ClientDsc As cProcItem
    # Dim ClientErr As cErr
    # Dim PN As String
    # Dim MN As String
    # Dim LineNo As Long
    # Dim Line As String

    # Dim sLiveName As String                            ' if data is from LiveStack, use source Line
    # Dim sNamPos As Long                                ' position from Proc Name in sLiveName
    # Dim sType As String                                ' sub or function from sLiveName
    # Dim TdbgId As String
    # Dim Key As String
    # Dim sKey As String
    # Dim putS As Boolean
    # Dim iPat As Long
    # Dim HasMoreLevels As Boolean
    # Dim HaveCalledProc As Boolean
    # Dim IsInLiveStack As Long
    # Dim Lvl As Long

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then                       ' do not trace the Trace:
    print(' Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True

    # Lvl = 0
    # HaveCalledProc = False
    # IsInLiveStack = False
    # Line = vbNullString
    # sLiveName = vbNullString
    # sKey = vbNullString
    # TdbgId = vbNullString
    # putS = False
    if TEntry Is Nothing Then:
    # Set TEntry = New cTraceEntry
    else:
    if LenB(TEntry.TSrc) > 0 Then              ' this is derived from LiveStack at some earlier time:
    # putS = True                            ' Put on stack!
    # sLiveName = Trunc(1, TEntry.TSrc, "(") ' cut off Call Parameters

    # Set ClientErr = TEntry.TErr
    if ClientErr Is Nothing Then:
    # sNamPos = InStr(sLiveName, "Call ")
    if sNamPos > 0 Then:
    # sType = tSub
    # sLiveName = Mid(sLiveName, sNamPos + 5)
    else:
    # sNamPos = InStr(sLiveName, " = ")
    if sNamPos > 0 Then:
    # sType = tFunction
    # sLiveName = Mid(sLiveName, sNamPos + 3)
    else:
    # sNamPos = InStr(sLiveName, b)
    # sType = tSub               ' wild assumption
    # sLiveName = Mid(sLiveName, sNamPos + 1)
    # sLiveName = "Live." & sLiveName
    # GoTo dummyGen
    else:
    # sLiveName = ClientErr.atKey
    if DoVerify(D_ErrInterface.Exists(sLiveName), "no entry in D_ErrInterface ???") Then:
    # sLiveName = "Missing." & sLiveName
    # dummyGen:
    # Call N_ConDscErr(ClientDsc, sLiveName, sType, eQnoDef, ClientErr)
    # Set TEntry.TErr = ClientErr
    # GoTo LessDetails

    # Set ClientDsc = ClientErr.atDsc
    # TdbgId = ClientDsc.DbgId
    if InStr(sLiveName, TdbgId) > 0 Then   ' got a reference:
    # sKey = ClientErr.atKey
    else:
    # Set ClientErr = TEntry.TErr
    # Set ClientDsc = ClientErr.atDsc
    # sKey = ClientErr.atKey
    # putS = True
    if ClientErr.atCallDepth = 0 Then      ' must be Extern.Caller:
    # HaveCalledProc = True
    # GoTo LessDetails

    if LenB(ExplainS) = 0 Then:
    # ExplainS = ClientErr.atMessage
    # ClientErr.atMessage = vbNullString
    if Not ErrExActive Then:
    if DebugMode Then:
    print(Debug.Print "* ErrEx is not active, can't use Live Stack")
    # TEntry.TLne = 0
    # TEntry.TSrc = vbNullString
    # HaveCalledProc = Not ClientErr Is Nothing
    # GoTo LessDetails

    if Not (DebugMode Or DebugLogging) Then        ' determine if we want to use source line info:
    # HaveCalledProc = True
    # GoTo LessDetails                           ' no source line in trace printout

    # ' look for Caller match (not Proc reference here)
    # With ErrEx.LiveCallstack
    # Do
    # Lvl = Lvl + 1
    # Set ClientDsc = Nothing
    # ' get all data from ErrEx.LiveCallStack line
    # PN = .ProcedureName
    # MN = .ModuleName
    # LineNo = .LineNumber
    # Line = .LineCode
    # HasMoreLevels = .NextLevel

    print('    Debug.Print Lvl, "More: " & HasMoreLevels, LString(TdbgId, lDbgM), LString(Key, lDbgM), LString(PN, lDbgM), Line)

    # ' position to next ErrEx.LiveCallStack line if any, no more old line info avail.
    if PN = MyId Then:
    # GoTo noDsc
    if HasMoreLevels Then:
    if Z_SourceAnalyse(sLiveName, TdbgId) Then:
    # HaveCalledProc = True
    # Key = sKey
    # GoTo UseLastKey

    # Key = MN & "." & PN
    if LenB(sKey) > 0 Then                 ' want only this Proc:
    if iPat > 1 Then:
    if Not IsSimilar(sKey, Key) Then:
    # GoTo noDsc
    elif Key <> sKey Then            ' refuse others:
    # GoTo noDsc
    if IsInLiveStack = 0 Then:
    # IsInLiveStack = Lvl                ' youngest proc reference only
    # GoTo noDsc                         ' skip for potential later reference
    # UseLastKey:
    if D_ErrInterface.Exists(Key) Then     ' do search for it:
    # Set ClientDsc = D_ErrInterface.Item(Key)
    # Set ClientErr = ClientDsc.ErrActive
    else:
    # GoTo noDsc

    # LessDetails:
    if Not HaveCalledProc Then:
    # GoTo noDsc
    # DoVerify TEntry.TErr Is ClientErr, "design check ???"

    if putS And Not dontIncrementCallDepth Then:
    # With TEntry
    # .TDet = ExplainS
    # .TLne = LineNo
    if LenB(.TSrc) > 0 Then:
    # .TSrc = Trim(Trim(Trunc(1, Line, "'", sLiveName)) _
    # & " '" & sLiveName)
    if .TSrc = "'" Then:
    # .TSrc = vbNullString
    # .TRL = ClientErr.atRecursionLvl
    # End With                           ' TEntry
    # Call TEntry.TraceAdd(ExplainS)
    # Call N_TraceDef(TEntry)
    if HaveCalledProc Then:
    # GoTo FuncExit
    # noDsc:
    # Loop While HasMoreLevels
    # End With                                       ' ErrEx.LiveCallStack

    if IsInLiveStack > 0 Then                      ' did not find caller line:
    # Lvl = 0
    # With ErrEx.LiveCallstack
    # Do
    # Lvl = Lvl + 1
    if Lvl = IsInLiveStack + 1 Then    ' skip to item after ProcCall:
    # addLast:
    # Line = .LineCode
    # LineNo = .LineNumber
    # With TEntry
    # .TDet = ExplainS
    # .TLne = LineNo
    # .TSrc = Trim(Trim(Trunc(1, Line, "'", sLiveName)) _
    # & " '" & sLiveName)
    # .TRL = ClientErr.atRecursionLvl
    # End With                       ' TEntry
    # Call TEntry.TraceAdd(ExplainS)
    # Call N_TraceDef(TEntry)

    # GoTo FuncExit
    # HasMoreLevels = .NextLevel
    # Loop While HasMoreLevels
    # GoTo addLast
    # End With                                   ' ErrEx.LiveCallStack

    if putS And (DebugMode Or DebugLogging) Then:
    print(Debug.Print sKey & " is not on live stack, can not trace")
    # DoVerify Not DebugMode, " instructed to trace, but not on LiveStack"
    # Set TEntry.TErr = Nothing                      ' assume no success

    # FuncExit:
    # Recursive = False
    # Set ClientDsc = Nothing
    # Set ClientErr = Nothing

    # zExit:

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : N_Undefine
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Clear traces of dead Proc
# '---------------------------------------------------------------------------------------
def n_undefine():

    if ClientDsc Is Nothing Then:
    # GoTo zExit
    if LenB(Key) = 0 Then:
    # Key = ClientDsc.Key
    # D_ErrInterface.Remove Key
    # Set ClientDsc.ErrActive = Nothing
    # Set ClientDsc = Nothing
    print(Debug.Print "Proc " & LString(Key, 2 * lDbgM) & " has been undefined, #" & i)

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : ProcCall
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Define or use Proc with atDsc and Err at time of Entry
# '---------------------------------------------------------------------------------------
def proccall():

    # Dim qMsg As String
    # Dim ClientDsc As cProcItem

    if isEmpty(QModeNames) Then                    ' basic inits for BugHelp required:
    # Call Z_StartUp(False)

    if FastMode Then:
    # ExplainS = Trim(ExplainS & b & "FastMode using Y_Type")
    # Call DoCall(ClientKey, CallType, eQyMode, ClientDsc, ExplainS)
    # Set ClientErr = ClientDsc.ErrActive
    # ClientErr.atRecursionOK = Recursive
    if aNameSpace Is Nothing _:
    # Or aRDOSession Is Nothing Then
    # Call N_ConDscErr(ClientDsc, ClientKey, CallType, Qmode, Nothing)
    # ClientDsc.CallCounter = 0
    # LogZProcs = True
    # GoTo FuncExit
    else:
    if Not ClientDsc Is Nothing Then:
    if ClientDsc.CallMode <> Qmode _:
    # And ClientDsc.CallMode > eQnoDef Then
    if ClientDsc.CallCounter > 0 Then:
    # aBugTxt = "** Change in Qmode is improbable, but allowed"
    # DoVerify ClientDsc.CallMode = Qmode
    if StackDebug > 5 Then:
    # qMsg = String(20, b) _
    # & "Changing Mode for " & ClientDsc.DbgId _
    # & ", CallCounter=" & Right(String(5, b) _
    # & ClientDsc.CallCounter, 5) & String(31, b) _
    # & "from " & ClientDsc.ModeName _
    # & "(" & ClientDsc.CallMode _
    # & ") to M=" & QModeNames(Qmode) & "(" & Qmode & ")"
    print(Debug.Print qMsg)
    else:
    # ClientDsc.CallCounter = 0      ' start a new call count after improbable mode change
    # ClientDsc.CallMode = Qmode             ' always set the (probably unchanged, unless new) CallMode
    # Call DoCall(ClientKey, CallType, Qmode, ClientDsc, ExplainS)
    # Set ClientErr = ClientDsc.ErrActive


    if Qmode = eQArMode Or Qmode = eQrMode Then:
    # Recursive = True
    # ClientErr.atRecursionOK = Recursive
    # ExplainS = ExplainS & ", Recursive"
    if Qmode = eQEPMode Then:
    # Call Z_EntryPoint(ClientDsc)

    if Qmode > eQuMode Then                        ' Application Levels:
    # E_Active.EventBlock = True                 ' Default for Applications during ProcCall
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call BugEval
    # ClientErr.atCallState = eCpaused

    if ClientDsc.ProcIndex = 0 And InStr(CallType, " EP") = 0 Then ' OK for Macro EP:
    # DoVerify ClientKey <> ExternCaller.Key, _
    # ClientKey & " used as dummy for Extern.Caller only!!"
    # GoTo FuncExit                              ' end of all inits. Stacks set up, no further actions

    # FuncExit:
    # SuppressStatusFormUpdate = False
    if Qmode > eQuMode Then                        ' Application Levels:
    # E_Active.EventBlock = False                ' for new Applications
    # Set E_AppErr = E_Active
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # frmErrStatus.fCurrAppl = E_Active.atDsc.DbgId
    # Call BugEval


# '---------------------------------------------------------------------------------------
# ' Method : ProcExit
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Reduce the .RecursionLvl and set the CallState when leaving an instance
# '---------------------------------------------------------------------------------------
def procexit():

    # Call DoExit(fromErr.atKey, DisplayValue)
    if fromErr.atFastMode Then:
    # GoTo ProcRet

    # S_DbgId = fromErr.atDsc.DbgId

    if fromErr.atDsc.CallMode = eQEPMode Then:
    # Call ReturnEP

    if ErrStatusFormUsable And Not SuppressStatusFormUpdate Then:
    if frmErrStatus.Visible Then:
    # Call QueryErrStatusChange(True)        ' .f values are used to modify globals
    # Call frmErrStatus.UpdInfo

    # With E_Active                                  ' Restore the values from caller's env
    # ZErrSnoCatch = .ErrSnoCatch               ' No ErrHandler Recursion;  NO N_PublishBugState
    # ZErrNoRec = .ErrNoRec                     ' No ErrHandler Recursion; use N_PublishBugState
    # .atFuncResult = DisplayValue
    # End With                                       ' E_Active

    # FuncExit:
    # Call ErrReset(4)                               ' keep caller's Try setting

    if E_Active.atCallDepth < 2 Then:
    if AppStartComplete Then:
    print(Debug.Print)
    # Call LogEvent("* " & String(20, "-") _
    # & " Outlook waiting for Events or Macro Calls", eLSome)
    # Call N_ShowHeader("BugHelp Log " & TimerNow)
    # ZAppStart.ErrActive.EventBlock = False
    # Call SetOnline(olCachedConnectedFull)
    # E_Active.EventBlock = False
    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # frmErrStatus.fOnline.Caption = "Online"
    # Call BugEval
    # NoEventOnAddItem = False
    if E_Active.atDsc.CallMode = eQAsMode Then:
    # Set E_AppErr = E_Active

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub QueryErrStatusChange
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def queryerrstatuschange():
    # Const zKey As String = "BugHelp.QueryErrStatusChange"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    if Recursive Then:
    print(Debug.Print String(OffCal, b) & "ignored recursion .from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet

    if ErrStatusFormUsable Then:
    # ErrDisplayModify = False                   ' must check this if ErrStatusForm is usable
    else:
    # ErrDisplayModify = True                    ' will need to update if/when ErrStatusForm becomes usable
    # GoTo ProcRet
    # Recursive = True

    # Call frmErrStatus.ReEvaluate(Not Reversed)     ' set the .f values from globals or vice versa

    # FuncExit:
    # Recursive = False

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub ReturnEP
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def returnep():

    # Const zKey As String = "BugHelp.ReturnEP"
    # Call DoCall(zKey, tSub, eQzMode)

    if EPCalled Then:
    # Call FldActions2Do                         ' if we have open items, do them now

    # EPCalled = False
    # StopRecursionNonLogged = False
    # NoEventOnAddItem = False

    if Not P_EntryPoint.ErrActive Is Nothing Then:
    # P_EntryPoint.ErrActive.Explanations = "(-- waiting for next Event --)"

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : ShowCallTrace
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show the Call Stack in Reverse Order
# '---------------------------------------------------------------------------------------
def showcalltrace():

    # Dim TEntry As cTraceEntry
    # Dim ErrClient As cErr

    # Dim i As Long
    # Dim j As Long
    # Dim SD As Long
    # Dim aTCal As String
    # Dim logLine As String
    # Dim saveDebug As Long
    # Dim saveStatusFormState As Boolean

    # saveDebug = StackDebug
    # StackDebug = 0                                 ' this causes output ONLY to file
    # saveStatusFormState = SuppressStatusFormUpdate
    # SuppressStatusFormUpdate = True

    if C_CallTrace.Count = 0 Then:
    # GoTo FuncExit

    # Call CloseLog                                  ' use new logfile for CallTrace
    # Call N_ShowHeader("C-CallTrace", force:=True, ModelType:=3)

    # i = TraceTop Mod ErrLifeTime
    if limitCount <= 0 Then:
    # j = C_CallTrace.Count
    else:
    # j = limitCount
    # Do
    # T_DC.LogFileLen = 0                        ' no size limit here
    # Set TEntry = C_CallTrace(i)
    # Set ErrClient = TEntry.TErr
    # With ErrClient
    if .atCallState = eCExited Then:
    # aTCal = "E"
    elif .atCallState = eCpaused Then:
    # aTCal = "P"
    elif .atCallState = eCActive Then:
    # aTCal = "A"
    else:
    # DoVerify Not DebugMode, "** undefined never happens in C_CallTrace?"
    # aTCal = "W2"
    # End With                                   ' ErrClient

    if ErrClient.atCallState = eCUndef Then:
    # logLine = RString(TEntry.Tinx, 5) & b _
    # & ErrClient.NrMT & aTCal & String(OffCal - OffLvl + 2, b)
    else:
    # SD = ErrClient.atCallDepth
    # logLine = RString(TEntry.Tinx, 5) & b _
    # & ErrClient.NrMT & aTCal & b & RString(SD, 3) & b
    # logLine = logLine _
    # & TEntry.TLog _
    # & " RL=" & TEntry.TRL _
    # & " Pre=" & TEntry.TPre & " Suc=" & TEntry.TSuc
    print(' Debug.Print logLine)

    # Call LogEvent(logLine, eLall)
    # j = j - 1
    if j < 1 Then:
    # Exit Do
    if i = 1 Then:
    # i = C_CallTrace.Count
    else:
    # i = i - 1
    # Loop

    # Call N_ShowHeader("C-CallTrace", StartOrEnd:=True, ModelType:=3)
    # Call ShowLogWait(False)
    # Call CloseLog(KeepName:=False)

    # FuncExit:
    # Set TEntry = Nothing
    # Set ErrClient = Nothing
    # StackDebug = saveDebug
    # SuppressStatusFormUpdate = saveStatusFormState


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowDbgStatus
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showdbgstatus():

    # Call N_Suppress(Push, "BugHelp.ShowDbgStatus")
    # SuppressStatusFormUpdate = True

    # Dim sTime As Variant

    if MayChangeErr Then:
    try:

        if ErrStatusFormUsable Then:
        # GoTo useTermForm
        else:
        # DebugControlsUsable = DebugControlsWanted
        if DebugControlsUsable Then                ' ignore the default:
        # useTermForm:
        # DebugControlsUsable = True

        if ErrStatusFormUsable Then:
        # GoTo ShowIt

        if aNonModalForm Is Nothing Then:
        # Set aNonModalForm = frmErrStatus           ' calls QueryErrStatusChange, but not showing yet
        # ' frmErrStatus.showmodal = False
        # '  vbNullString             can not set here, do this in form's properties!!!
        else:
        # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
        if LenB(Prompt) = 0 Then:
        # frmErrStatus.fTerminationFlag.Caption _
        # = "Termination Flag = " & GetTerminationState
        else:
        # frmErrStatus.fTerminationFlag.Caption = Prompt
        # ShowIt:
        # Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
        # doMyEvents                                     ' allow interaction, delay and wait
        if DebugMode Then:
        if Wait(5, trueStart:=sTime, DebugOutput:=False) Then:
        # aBugTxt = "Press Enter then Confirm continue on Debug for ShowDbgStatus"
        elif LenB(Testvar) = 0 Then                  ' no need to Show this form:
        if Not IgnoreUnhandledError Then:
        # frmErrStatus.Hide
        # GoTo FuncExit
        # errCode:
        if aNonModalForm Is Nothing Then:
        # ErrStatusFormUsable = False

        # FuncExit:
        # Call ErrReset(0)
        # Call N_Suppress(Pop, "BugHelp.ShowDbgStatus")
        # SuppressStatusFormUpdate = False


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowDefProcs
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showdefprocs():
    # Const zKey As String = "BugHelp.ShowDefProcs"

    # Call N_Suppress(Push, zKey)
    # SuppressStatusFormUpdate = True

    # Dim i As Long
    # Dim aDsc As cProcItem

    # Dim CallNumSave As Long
    # Dim keyV As Variant
    # Dim AllOrNot As String
    # Dim aDict As Dictionary
    # Dim Both As Boolean

    # CallNumSave = CallNr
    # CallNr = 0
    # Both = Full And ErrInterface

    if Full Then:
    # AllOrNot = " (all)"
    else:
    # AllOrNot = " (only active)"

    # doBoth:
    if Both Then:
    # Call N_ShowHeader("D_ErrInterface in Creation Order" & AllOrNot, force:=True)
    # Set aDict = D_ErrInterface
    else:
    # Call N_ShowHeader("D_ErrInterface in Creation Order" & AllOrNot, force:=True)
    # Set aDict = D_ErrInterface
    # keyV = aDict.Keys(i)
    if isEmpty(aDict.Items(i)) Then:
    print(Debug.Print LString(i, 5) & String(20, b) & LString(keyV, lDbgM) _)
    # & "??? is empty, is removed"
    # aDict.Remove keyV                      ' INTRINSIC proc
    else:
    # Set aDsc = aDict.Items(i)
    if aDsc Is Nothing Then:
    print(Debug.Print "**** error in D_ErrInterface: Null Entry at pos=" & i)
    else:
    if aDsc.ErrActive.atLastInSec > 0 _:
    # Or Full Then
    # Call N_ShowProcDsc(aDsc, (i), WithInstances:=WithInstances, _
    # Consumption:=True)
    # Call N_ShowHeader("Procs in Creation Order", StartOrEnd:=True)
    if Both Then:
    # Both = False
    # GoTo doBoth

    # FuncExit:
    # Set aDsc = Nothing
    # Set aDict = Nothing
    # Call N_Suppress(Pop, zKey)
    # CallNr = CallNumSave
    # SuppressStatusFormUpdate = False


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowErr
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showerr():

    # Const zKey As String = "BugHelp.BugEval"

    # '------------------- gated Entry -------------------------------------------------------
    if AppStartComplete Then:

    # DebugControlsUsable = True
    # Call ShowOrHideForm(frmErrStatus, True)
    # ErrStatusHide = False



# '---------------------------------------------------------------------------------------
# ' Method : ShowErrInterface
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show the Data in D_ErrInterface
# '---------------------------------------------------------------------------------------
def showerrinterface():

    # Dim i As Long
    # Dim TName As String
    # Dim TKey As String
    # Dim ClientDsc As cProcItem
    # Dim ErrClient As cErr
    # Dim TDetail As String
    # Dim TdbgId As String
    # Dim saveStatusFormState As Boolean

    # saveStatusFormState = SuppressStatusFormUpdate
    # SuppressStatusFormUpdate = True
    # Call Try(testAll)                                 ' Try anything, autocatch, no resetting

    if Not Predefined Then:
    if D_ErrInterface.Count > 0 Then:
    # Call ShowDefProcs(ErrInterface:=True, Full:=Full)
    # GoTo FuncExit

    print(Debug.Print LString("Num", 4) & LString("TypeName", 16) _)
    # & RString("Has/at Key", lKeyM) & b _
    # & LString("Wanted Key", lKeyM) & "Detail"
    # TDetail = vbNullString
    # Catch
    # TName = TypeName(D_ErrInterface.Items(i))
    # TKey = D_ErrInterface.Keys(i)
    if WithDetails = 0 Then:
    # GoTo nullDetail

    match TName:
        case "cProcItem":
    # Set ClientDsc = D_ErrInterface.Items(i)
    # Set ErrClient = ClientDsc.ErrActive
    # isDsc:
    if ClientDsc Is Nothing Then:
    # GoTo noDsc
    if ErrorCaught <> 0 Then:
    # TDetail = Err.Description
    else:
    # TdbgId = Mid(D_ErrInterface.Keys(i), 3)
    if InStr(ClientDsc.DbgId, TdbgId) = 0 Then:
    if ClientDsc.DbgId <> "Extern" Then:
    # TDetail = "InF=" & TdbgId & " <> Dsc=" & ClientDsc.Key
    elif LenB(ErrClient.atKey) = 0 Then:
    # noDsc:
    if WithDetails > 2 Then:
    # GoTo nextInLoop
    if LenB(ClientDsc.CallType) > 0 Then:
    # GoTo badCall
    else:
    # badCall:
    if ClientDsc.Key <> ErrClient.atKey Then:
    # TDetail = "Dsc=" & ClientDsc.Key & " <> Err=" & ErrClient.atKey
    if LenB(TDetail) > 0 Then:
    # TDetail = "IdsOk=F " & TDetail
    if LenB(ClientDsc.CallType) > 0 Then:
    if WithDetails = 0 Then:
    # GoTo nullDetail
    else:
    if WithDetails = 1 Then:
    # GoTo nextInLoop
    else:
    if WithDetails > 2 Then:
    # Call Z_CheckLinkage(ClientDsc, zeroClientIsOk:=True, Result:=TDetail)
    # Call N_ShowProcDsc(ClientDsc, i, WithInstances:=(WithDetails > 3), _
    # ExplainS:=TDetail, ErrClient:=ErrClient, Consumption:=False)
        case "cErr":
    # Set ErrClient = D_ErrInterface.Items(i)
    # Set ClientDsc = ErrClient.atDsc
    if WithDetails > 1 Then:
    # GoTo isDsc
    if ClientDsc.Key <> ErrClient.atKey Then:
    # TDetail = "atK=" & ErrClient.atKey & " <> Dsc=" & ClientDsc.Key
        case _:
    if WithDetails > 0 Then:
    # TDetail = D_ErrInterface.Items(i).Count
    if LenB(TDetail) > 0 Then:
    # TDetail = "Count=" & TDetail
    # nullDetail:
    print(Debug.Print LString(i, 4) & LString(TName, 16) _)
    # & RString(D_ErrInterface.Keys(i), lKeyM) & b _
    # & LString(TKey, lKeyM) & TDetail

    # FuncExit:
    # Call ErrReset(0)
    # Set ErrClient = Nothing
    # SuppressStatusFormUpdate = saveStatusFormState


# '---------------------------------------------------------------------------------------
# ' Method : ShowErrorStatus
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show a form that allows displaying error information, debug controls, allows remedy
# '---------------------------------------------------------------------------------------
def showerrorstatus():

    # Const zKey As String = "BugHelp.ShowErrorStatus"
    # ' Dim ZShowErrorStatus As cProcItem                                      ' Predefined!

    # Call N_Suppress(Push, zKey)

    # With T_DC
    if isEmpty(aNonModalForm) Then:
    # Set aNonModalForm = Nothing
    if aNonModalForm Is Nothing Then:
    # Call ShowDbgStatus
    if aNonModalForm Is Nothing Then:
    # DoVerify "* Unable to Show aNonModalForm!"
    # ErrStatusFormUsable = False
    # GoTo ProcReturn
    else:
    # ErrStatusFormUsable = True

    # With aNonModalForm
    # .fLastErr = T_DC.DCerrNum
    # .fLastErrSource = T_DC.DCerrSource
    # .fLastErrMsg = T_DC.DCerrMsg
    # ' get information from E_AppErr (probably set by ErrTry)
    # .fLastErrExplanations = E_AppErr.Explanations
    # .fLastErrReasoning = E_AppErr.Reasoning
    if .fLastErrIndications.Enabled Then:
    # .fModifications = True

    if ZShowErrorStatus.CallCounter = 1 Then:
    # .Top = 245
    # .Left = 1041
    # .fLastErrExplanations = "Manueller StartUp mit Stop: " _
    # & "Alternative Werte fr Debugoptionen jetzt whlen! " & Time
    print(Debug.Print .fLastErrExplanations)
    # .Show                              ' Modal
    # Call BugEval
    else:
    # .Show vbModeless
    # End With                                   ' aNonModalForm
    # End With                                       ' T_DC

    # ProcReturn:
    # Call N_Suppress(Pop, zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowErrStack
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showerrstack():
    # Const zKey As String = "BugHelp.ShowErrStack"

    # Call N_Suppress(Push, zKey, Value:=False)

    # Dim ErrClient As cErr
    # Dim i As Long
    # Dim uTop As Long
    # Dim TheErrStackName As String

    # uTop = E_Active.atCallDepth
    if uTop < 2 Then                               ' do not show external caller:
    print(Debug.Print "- The current active Error Stack is empty. Gimme sompn to do.")
    # GoTo FuncExit

    # TheErrStackName = "Following the CalledBy"
    # Call N_ShowHeader(TheErrStackName, force:=True)

    # Set ErrClient = E_Active
    # Call N_ShowErrInstance(ErrClient, i)
    # Set ErrClient = E_Active.atCalledBy

    # Call N_ShowHeader(TheErrStackName, StartOrEnd:=True)

    # FuncExit:
    # Set ErrClient = Nothing
    # Call N_Suppress(Pop, zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowFunctionValue
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showfunctionvalue():

    # Const zKey As String = "BugHelp.ShowFunctionValue"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim tInfo As cInfo
    # Dim i As Long
    # Dim printVal As String
    # Dim FunctionName As String

    if VarType(FunctionRef) = vbObject Then:
    if TestSafe Then:
    # Call getInfo(tInfo, FuncValue, Assign:=False)
    else:
    # GoTo zExit

    if tInfo.iAssignmentMode = 1 Then:
    # printVal = tInfo.iValue
    if NoMultiLine Then:
    # i = InStr(tInfo.iValue, vbCr)
    if i > 0 Then:
    # printVal = Left(printVal, i - 1) & "..."
    if InStr(FunctionName, " of ") = 0 Then:
    print(Debug.Print "value of " & FunctionName & "=" & Quote(printVal))
    else:
    print(Debug.Print FunctionName & b & Quote(printVal))

    # FuncExit:
    # Set tInfo = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : ShowLiveStack
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Show the live Call Stack
# '---------------------------------------------------------------------------------------
def showlivestack():
    # Const zKey As String = "BugHelp.ShowLiveStack"

    # Dim i As Long
    # Dim liveCount As Long
    # Dim LCI As cCallEnv

    if Not ErrExActive Then:
    print(Debug.Print "* ErrEx is not active, can't Show Live Stack")
    # GoTo zExit

    # Call N_Suppress(Push, zKey)

    # Set D_LiveStack = N_GetLiveStack
    # liveCount = D_LiveStack.Count - 1

    if doPrint Then:
    # Call N_ShowHeader("LiveCallstack, Time=" & Now(), force:=True, ModelType:=2)
    if doPrint Then:
    # Set LCI = D_LiveStack.Items(i)
    if LCI.CallerErr Is Nothing Then:
    if Full Then:
    # GoTo ShowIt
    else:
    # ShowIt:
    # Call N_PrintNameInfo(i, LCI)
    # Call N_ShowHeader("LiveCallstack", StartOrEnd:=True)

    # FuncExit:
    # Call N_Suppress(Pop, zKey)
    # Set LCI = Nothing

    # zExit:


# '---------------------------------------------------------------------------------------
# ' Method : N_GetLiveStack
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get an Extract of the live Call Stack into a dictionary
# '---------------------------------------------------------------------------------------
def n_getlivestack():

    # Dim i As Long
    # Dim LCI As cCallEnv

    # Set N_GetLiveStack = New Dictionary
    # Set LCS = ErrEx.LiveCallstack
    # LCS.FirstLevel                                 ' just back to the first one
    # Do
    # Set LCI = New cCallEnv
    # With LCI
    # .StackDepth = i
    # .ModuleName = LCS.ModuleName
    # .ProcedureName = LCS.ProcedureName
    # .LineNumber = LCS.LineNumber
    # .LineCode = LCS.LineCode
    # .ModProc = .ModuleName & "." & .ProcedureName
    # .DscKind = ExStackProcNames(LCS.ProcedureKind)
    # N_GetLiveStack.Add i, LCI
    # End With                                   ' LineInfo.LCI
    # i = i + 1
    # Loop While LCS.NextLevel

    # Set LCI = Nothing


# '---------------------------------------------------------------------------------------
# ' Method : Sub ShowStacks
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def showstacks():
    # Call N_ShowStacks(Full:=Full)                  ' alias for Entry use

# '---------------------------------------------------------------------------------------
# ' Method : Sub SimulateAnError
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub SimulateAnError()

# Dim msg As String

print(Debug.Print "Errex Enabled=" & ErrEx.IsEnabled)

print(Debug.Print 1 / 0                              ' Simulate division by zero error (error 11))
print(Debug.Print "A" / 1                            ' Simulate type mismatch error (error 13))
# Err.Raise &H123, , "This should be caught by the catch &H123 block..."
# Err.Raise &H456, , "This should be caught in the catch-all block..."

# ' Program flow now transfers to Finally...

# ErrEx.Catch 11, 13                             ' you can set up a constant/enumeration if you prefer
print(Debug.Print "Catch: #" & ErrEx.Number & "/" & Err.Number _)
# & " is Handled with resume next"
print('Error: either division by zero, or type mismatch!')
pass  # resume next

# ErrEx.Catch &H123                              ' you can set up a constant/enumeration if you prefer
print(Debug.Print "Catch: #" & ErrEx.Number & "/" & Err.Number _)
# & " is Handled with resume next"
print('Error: caught error 291 (&H123)!')
pass  # resume next

# ErrEx.CatchAll
print(Debug.Print "Catch All: #" & ErrEx.Number & "/" & Err.Number _)
# & " is Handled with resume next"
print('Error (catch-all):')
pass  # resume next

# ErrEx.Finally
print(Debug.Print "Catch Finally: #" & ErrEx.Number & "/" & Err.Number _)
# & " is Handled, then causes 1/0 "
print('Finally')
print(' Debug.Print 1 / 0    ' Errors here automatically ignored due to implicit OnErrorResumeNext)
if Err.Number <> 0 Then:
# msg = "Error # " & str(Err.Number) & " was generated by " _
# & Err.Source & Chr(13) & Err.Description
# Err.Clear
# Call ErrEx.DoFinally

print(Debug.Print 1 / 0                              ' Simulate division by zero error (error 11))
print(Debug.Print "A" / 1                            ' Simulate type mismatch error (error 13))
# Err.Raise &H123, , "This should be caught by the catch &H123 block..."
# Err.Raise &H456, , "This should be caught in the catch-all block..."

# ErrEx.Catch 11                                 ' you can set up a constant/enumeration if you prefer
print(Debug.Print "Catch2: #" & ErrEx.Number & "/" & Err.Number _)
# & " is Handled with resume next"
print('Error: division by zero!')
pass  # resume next

# zExit:


# '---------------------------------------------------------------------------------------
# ' Method : StartEP
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Init when an external event happens.
# '          This is a wrapper for the ProcCall, in that it uses Z_StartUp if we never inited.
# '---------------------------------------------------------------------------------------
def startep():
    # DoVerify Qmode = eQEPMode And CallType = tSubEP, _
    # "EP's must be EP Application and CallType must be tSubEP"
    if Not DidStop Then:
    if UseStartUp = 0 Then:
    # Call Z_StartUp                         ' detour for test purposes, continue
    # UseStartUp = 1
    elif UseStartUp < 2 Then:
    # DoVerify False
    # UseStartUp = 0
    # Call Z_StartUp                         ' call N_PreDefined test, skip

    if AppStartComplete And ItemsToDoCount + Deferred.Count > 0 Then:
    # Call FldActions2Do                         ' (must) have (at least 1) open items

    # Call ProcCall(ErrClient, ClientKey, Qmode:=Qmode, CallType:=CallType, ExplainS:=ExplainS) ' the corresponding exit happens in ProcExit->ReturnEP


# '---------------------------------------------------------------------------------------
# ' Method : Sub StartUp
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Entry from External Caller sets up Application Environment in Z_StartUp.
# ' Note   : To support debugging in BugHelp, it calls N_DebugStart. Happens only here.
# '          If called more than once, it will reset the Application Environment
# ' Typical Use: when debugging and a re-run is intended.
# '---------------------------------------------------------------------------------------
def startup():

    # Call N_DeInit(False)
    # FastMode = Fast
    # LogZProcs = LogAny
    # Call N_CheckStartSession(doStop)


# '---------------------------------------------------------------------------------------
# ' Method : Sub TerminateApp
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def terminateapp():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "BugHelp.TerminateApp"
    # Call DoCall(zKey, tSub, eQzMode)

    if T_DC Is Nothing Then:
    # DoVerify False, " if tester continues ..."
    else:
    # Call T_DC.Terminate
    # '                                 ... we will reach this
    if IsMissing(newEP) Then:
    # EPCalled = False
    else:
    if newEP = True Then:
    # EPCalled = newEP
    elif newEP = False Then:
    # EPCalled = newEP

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Try
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Announce which errors are acceptable
# '---------------------------------------------------------------------------------------
def try():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "BugHelp.Try"

    # Dim logVal As String

    # '------------------- gated Entry -------------------------------------------------------
    # logVal = Left(WhatEver, 8)
    if logVal = "52" Then                          ' exclude log of Try in LogEvent:
    # logVal = vbNullString

    # E_Active.Permit = WhatEver
    # Call ErrReset(4)


# '---------------------------------------------------------------------------------------
# ' Method : Z_CheckLinkage
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Check if the linkage from/to ProcDsc to ProcErr is intact
# '---------------------------------------------------------------------------------------
def z_checklinkage():

    # Const zKey As String = "BugHelp.Z_CheckLinkage"
    # Call DoCall(zKey, tSub, eQzMode)

    if StackDebug < 8 Then:
    # GoTo zExit

    # Dim stopHere As Boolean
    # Dim ErrClient As cErr

    # stopHere = isEmpty(Result)

    if ClientDsc Is Nothing Then:
    # Result = "NoClient"
    else:
    # Set ErrClient = ClientDsc.ErrActive
    if zeroClientIsOk Then:
    if ErrClient Is Nothing Then:
    # Result = "NoErrClient"
    else:
    if LenB(ErrClient.atKey) = 0 Then:
    # Result = "NoAtKeyErrClient"
    else:
    if Not ClientDsc Is ErrClient.atDsc Then:
    # Result = Result & " BadParentDsc"

    if LenB(Result) > 0 Then:
    if stopHere Then:
    # DoVerify False
    # Set ErrClient = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_CheckLiveStack
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Find if Proc is Live
# '---------------------------------------------------------------------------------------
def z_checklivestack():

    # Const zKey As String = "BugHelp.Z_CheckLiveStack"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim LiveStack As Collection
    # Dim LiveKey As String

    # Call N_GetLive(LiveStack, Filtered:=myFilter, Logging:=False)

    if LiveStack.Count > 0 Then:
    # LiveKey = LiveStack.Item(LiveStack.Count)
    if myFilter Then                           ' if not filtered, do some checking:
    if this <> LiveKey Then:
    # Warn = True
    # msg = msg & " - unchecked, found " & LiveKey
    else:
    if LiveStack.Count < 2 Then            ' there must be a caller and the active one!:
    # Warn = True
    else:
    if Not myFilter Then:
    # Warn = True

    # Set LiveStack = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_EntryPoint
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Application-Level Entry Housekeeping
# '---------------------------------------------------------------------------------------
def z_entrypoint():

    # Const zKey As String = "BugHelp.Z_EntryPoint"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)

    # aDebugState = E_AppErr.DebugState
    # Call ShowStatusUpdate

    if ClientDsc.ErrActive.atRecursionLvl > 1 Then:
    # DoVerify False, " Applications should not be recursive "
    # ' GoTo ProcReturn

    if ClientDsc.CallMode >= eQuMode Then          ' Application level definitions:
    if ErrStatusFormUsable Then:
    # frmErrStatus.fCurrEP = ClientDsc.DbgId ' allow display of current Entry point

    # FuncExit:                                          ' ends always when S_AppIndex < 2 in Z_AppExit

    # ProcReturn:
    # Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Z_GetProcDsc
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Get ClientDsc from Key in D_ErrInterface and test on relevant CallMode
# '---------------------------------------------------------------------------------------
def z_getprocdsc():
    # '''' Proc Should ONLY CALL Z_Type PROCS                       ' trivial proc
    # Const zKey As String = "BugHelp.Z_GetProcDsc"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim DbgId As String

    # DbgId = Mid(Key, InStr(Key, ".") + 1)

    if D_ErrInterface.Exists(Key) Then             ' do search for it:
    # Set ClientDsc = D_ErrInterface.Item(Key)
    if ClientDsc.CallMode = eQnoDef Then       ' must be on D_ErrInterface!!!:
    # msg = msg & "dummy proc " & zKey _
    # & " M=" & ClientDsc.ModeLetter
    # makeDummy:
    # aBugTxt = "assuming no Z_Type procs are off-stack: not needed any longer" ' ??? ??? ???
    # DoVerify False
    # msg = msg & "dummy: " & Key & vbCrLf
    # Call DoCall(Key, tSub, eQyMode, ExplainS:=", as Dummy")
    # Call DoExit(Key)
    # GoTo zExit
    else:
    # msg = msg & "name not defined yet: "
    # GoTo makeDummy

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_InitBugHelp
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Initialize BugHelp
# '---------------------------------------------------------------------------------------
def z_initbughelp():

    # Const zKey As String = "BugHelp.Z_InitBugHelp"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim AddMsg As String

    # Dim ErrClient As cErr

    # ' Initialize Management Variables for BugHelp
    # S_AppIndex = -1                                ' = D_AppStack.Count - 1

    # Call N_SetErrExHdl

    # S_ActKey = ExternCaller.Key

    if CallLogging Then:
    # AddMsg = "* Z_InitBugHelp  has been successfully completed: BugHelp is operational"
    if LogPerformance Then:
    # AddMsg = AddMsg & ", Performance Data are collected"
    print(Debug.Print AddMsg)

    # DidItAlready = True
    # MayChangeErr = True

    # FuncExit:
    # Set ErrClient = Nothing

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_IsUnacceptable
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Returns True on non-acceptable errors.
# '---------------------------------------------------------------------------------------
def z_isunacceptable():
    # '''' Proc Should ONLY CALL Z_Type PROCS                       ' trivial proc
    # Const zKey As String = "BugHelp.Z_IsUnacceptable"
    # ' do not Call DoCall(zKey, tFunction, eQzMode)

    # With E_Active
    # Z_IsUnacceptable = True                    ' with allowed exceptions, see OK

    if T_DC.DCerrNum = Hell Then:
    # T_DC.TermRQ = True
    # Call TerminateApp
    # GoTo zExit                             ' should never be reached

    # ' if a message is passed instead of error numbers, compare that

    if Left(MatchAllow, 1) = "*" Then:
    # Z_IsUnacceptable = False               ' should not be reached
    # GoTo SelfHandledError

    if IsNumeric(MatchAllow) Then:
    # noPrint = True                         ' do not print if acceptable
    else:
    if isEmpty(MatchAllow) Then:
    # GoTo IsErr                         ' no message to compare, disallowed
    elif Left(Trim(MatchAllow), 1) = "-" Then ' Unacceptable message (complete):
    if InStr(.Description, Mid(MatchAllow, 2)) > 0 Then:
    # GoTo IsErr
    elif Left(Trim(MatchAllow), 1) = "+" Then ' acceptable error msg (complete):
    if InStr(.Description, Mid(MatchAllow, 2)) > 0 Then:
    # noPrint = True
    # GoTo Allowed
    elif Left(Trim(MatchAllow), 1) = "%" Then ' expected error msg (partial):
    if InStr(.Description, Mid(MatchAllow, 2)) > 0 Then:
    # noPrint = True                 ' NO reset of error state
    # GoTo Allowed                   ' unless by optional parameter
    else:
    # GoTo IsErr
    else:
    if .Description = MatchAllow Then  ' exactly this message (complete):
    # Z_IsUnacceptable = False
    # noPrint = True                 ' do not print if acceptable
    # GoTo zExit

    if .errNumber = MatchAllow Then:
    # Allowed:
    if LogAllErrors _:
    # Or ((DebugMode Or DebugLogging) And Not noPrint) Then
    if LogAllErrors Then:
    # Call LogEvent("Error: " & .Description, eLall)
    else:
    print(Debug.Print .Description)
    if DebugMode Then:
    # ' Show the aNonModalForm with error information and
    # ' allow user to ignore the error
    if DebugControlsWanted Then:
    # Call ShowErrorStatus
    else:
    # DoVerify False

    # Z_IsUnacceptable = False
    # .FoundBadErrorNr = 0
    # GoTo zExit
    else:
    # IsErr:
    if Z_IsUnacceptable Then:
    if Not noPrint Then                ' Print more later:
    print(Debug.Print .Description)
    if T_DC.DCAllowedMatch = 0 Then:
    # T_DC.TermRQ = True
    if Err.Number > 0 Then:
    if (DebugMode Or DebugLogging) Then:
    if Err.Number > 0 Then:
    if noPrint And DebugMode Then:
    print(Debug.Print String(80, b) & vbCrLf _)
    # & "Error " & Err.Number _
    # & ":" & Err.Description
    else:
    # Z_IsUnacceptable = False
    # T_DC.TermRQ = False
    # GoTo zExit
    if Left(MatchAllow, 1) = "*" Then:
    # SelfHandledError:
    if DebugMode And Not noPrint Then:
    print(Debug.Print String(80, b) & vbCrLf _)
    # & "!!! Error in " & S_DbgId & b & .errNumber _
    # & " (&H" & Hex8(.errNumber) & "): " _
    # & .Description
    # Z_IsUnacceptable = False           ' E_Active is not changed!
    # End With                                       ' E_Active

    # zExit:
    # ' do not Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_LogApp_Exit
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Print Application level Call (do not Z_AppExit)
# '---------------------------------------------------------------------------------------
def z_logapp_exit():

    # Const zKey As String = "BugHelp.Z_LogApp_Exit"
    # Call DoCall(zKey, tSub, eQzMode)

    # Static Recursion As Boolean                    ' special: check disallowed Recursion

    if Recursion Then:
    if ClientDsc.ErrActive.atCallState < eCpaused Then:
    # DoVerify False, "??? Recursion not needed if no hit"
    # GoTo zExit                                 ' omit recursion on self (yesss, it is!!)

    # Recursion = True

    # Dim noPrint As Boolean
    # Dim Caller As String
    # Dim addInfo As String

    if StackDebug <= 9 Then:
    if StackDebug > 8 Then:
    if ClientDsc.CallMode <= eQxMode Then:
    # GoTo testHidden
    if ClientDsc.CallMode = eQzMode Then       ' covers  .., Z_.., and O_Goodies, Classes:
    # testHidden:
    if StackDebug > 4 Then:
    # GoTo FuncExit

    if Left(moreEE, 1) = "!" Then                  ' do not print:
    # noPrint = True
    # addInfo = Mid(moreEE, 2)
    else:
    # addInfo = moreEE

    # Caller = ObjStr

    # GenOut:

    # Call Z_Protocol(ClientDsc, CallNr, Caller, "<==", ClientDsc.DbgId, addInfo)

    # FuncExit:
    # Recursion = False

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : Z_Protocol
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Output an Application-level Protocol line
# '---------------------------------------------------------------------------------------
def z_protocol():

    # Const zKey As String = "BugHelp.Z_Protocol"
    # Call DoCall(zKey, tSub, eQzMode)

    # Dim Lvl As Long
    # Dim Mlen As Long
    # Dim CallS As String

    # Lvl = Abs(ClientDsc.ErrActive.atCallDepth)
    # CallS = Caller & Io & ObjInfo
    # Mlen = Lvl + Len(CallS) + Len(Io) + 3

    if Mlen > lKeyM Then:
    # CallS = CallS & ">>" & Lvl
    else:
    # CallS = LString(String(Lvl, b) & CallS, lKeyM)
    # addInfo = Trim(addInfo & b & Replace(S_AppKey, dModuleWithP, vbNullString))
    # Call N_ShowProcDsc(ClientDsc, CallNr, WithInstances:=False, ExplainS:=addInfo, Caller:=CallS)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : N_SetErrExHdl
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Define UseErrExOn as default error handler for ErrExHandler
# '---------------------------------------------------------------------------------------
def n_seterrexhdl():

    # Dim msg As String

    if LenB(UseErrExOn) = 0 Then:
    if LenB(LastErrExOn) = 0 Then:
    # UseErrExOn = "N_OnError"               ' define the Proc that implements the global error events

    # msg = "* ErrEx" & String(13, b) & RString("Global error handler", lDbgM) & b
    if ErrEx.IsEnabled Then:
    # ErrExConstructed = True
    if ErrExConstructed Then:
    # msg = msg & "was inactive, "
    else:
    # msg = msg & "was already active, "
    # ErrExConstructed = False

    if LenB(UseErrExOn) > 0 Then:
    if LenB(LastErrExOn) = 0 Then:
    if doPrint And StackDebug > 5 Then:
    # msg = msg & "unknown previous handler, "
    print(Debug.Print msg & "new global error handler '" & UseErrExOn & "'")
    # Call ErrEx.Enable(UseErrExOn)
    else:
    if LastErrExOn = UseErrExOn Then:
    if doPrint And StackDebug > 5 Then:
    # msg = msg & " left unchanged '" & LastErrExOn & "'"
    else:
    if doPrint And StackDebug > 5 Then:
    # msg = msg & " changing from '" & LastErrExOn & "'"
    # Call ErrEx.Enable(UseErrExOn)
    if doPrint And StackDebug > 5 Then:
    print(Debug.Print msg & " to '" & UseErrExOn & "'")

    # LastErrExOn = UseErrExOn

    elif ErrExActive Then:
    if CallLogging Then:
    if LenB(LastErrExOn) = 0 Then:
    print(Debug.Print msg & "remains disabled")
    else:
    print(Debug.Print msg & "disabled from '" & LastErrExOn & "'")

    # Call ErrEx.Disable                         ' this sets ErrExActive = False, but keeps the LastErrExOn unchanged
    if ErrStatusFormUsable Then:
    # frmErrStatus.fErrAppl = vbNullString

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Z_SetupDialogs
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub Z_SetupDialogs()
# '--- Proc MAY ONLY CALL Z_Type PROCS                            ' Simple proc
# Const zKey As String = "BugHelp.Z_SetupDialogs"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

# With ErrEx.DialogOptions
if isEmpty(T_DC.DCAllowedMatch) Then:
# .HTML_MainBody = "<font face=Arial size=13pt color=#4040DF><b>" _
# & "A runtime error has occurred:</b></font><br><br><b>" _
# & "<ERRDESC></b><br><br>Source:|<SOURCEPROJ>.<SOURCEMOD>." _
# & "<SOURCEPROC><br>Filename:|<SOURCEFILENAME><br>Number:" _
# & "|&H<ERRNUMBERHEX> (<ERRNUMBER>)<br>Source Line:|" _
# & "<font bgcolor=#FFD8AF> #<SOURCELINENUMBER>.        <SOURCELINECODE>   </font><br>" _
# & "<br>No Acceptable Error(s)"
else:
# .HTML_MainBody = "<font face=Arial size=13pt color=#4040DF><b>" _
# & "A runtime error has occurred:</b></font><br><br><b><ERRDESC>" _
# & "</b><br><br>Source:|<SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC>" _
# & "<br>Filename:|<SOURCEFILENAME><br>Number:|&H<ERRNUMBERHEX> " _
# & "(<ERRNUMBER>)<br>Source Line:|<font bgcolor=#FFD8AF> " _
# & "#<SOURCELINENUMBER>.        <SOURCELINECODE>   </font><br>" _
# & "<br>Acceptable Error(s): " & T_DC.DCAllowedMatch
# .HTML_MainBody = .HTML_MainBody & "<br>Date/Time:|<ERRDATETIME><br>" _
# & "<br><b><font size=12pt color=#4040DF>What do you want to do?</font></b>"
# .HTML_MoreInfoBody = "<br><b><font color=#40408F bgcolor=#C8D8FF>                                                     ?    VBA CALL STACK    ?                                                     </font></b><br><CALLSTACK>"
# .HTML_CallStackItem = "  <b><SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC></b>" _
# & "<br> | #<SOURCELINENUMBER>.        <SOURCELINECODE>   <br>"
# .HTML_VariableItem = "(<VARSCOPE>)|<VARNAME> As <VARTYPE>| = <VARVALUE><br>"
# .WindowCaption = "YourApplicationName - runtime error"
# .MinimumWindowWidth = 600
# .MoreInfoCaption = "More info"
# .LessInfoCaption = "Less info"
# .ButtonPaddingH = 10
# .ButtonPaddingV = 5
# .ButtonSpacingH = 5
# .ButtonSpacingV = 7
# .PaddingH = 15
# .PaddingV = 15
# .ScreenBorderPaddingV = 50
# .ColumnPaddingH = 20
# .LineSpacingV = 2
# .MainBackColor = 16777215
# .MainBackColor2 = 13693168
# .MainBackFillType = 8
# .MoreInfoBackColor = 15724768
# .MoreInfoBackColor2 = 16774642
# .MoreInfoBackFillType = 0
# .ButtonBarBackColor = 16443364
# .ButtonBarBackColor2 = 14337988
# .ButtonBarBackFillType = 1
# .MaxNumCallStackItems = 10
# .MaxNumVariablesItems = inv
# .DefaultButtonID = 5
# .DefaultButtonIsBold = True
# .ShowMoreInfoButton = True
# .AllowEnterKey = True
# .AllowTabKey = True
# .AllowArrowKeys = False
# .Timeout = 0
# .FlashWindowOnOpen = True
# .CustomImageTransparentColor = -1
# Dim TempImageData As String
# TempImageData = vbNullString
# .CustomImageData = TempImageData
# Call .RemoveAllButtons
# Call .AddCustomButton("Search internet", "OnSearchInternet")
# Call .AddButton("Show Variables", BUTTONACTION_SHOWVARIABLES)
# Call .AddButton("Debug sourcecode", BUTTONACTION_ONERRORDEBUG)
# Call .AddButton("Ignore and continue", BUTTONACTION_ONERRORRESUMENEXT)
# Call .AddButton("Help", BUTTONACTION_SHOWHELP)
# Call .AddButton("Close", BUTTONACTION_ONERROREND)
# End With                                       ' ErrEx.DialogOptions

# With ErrEx.VariablesDialogOptions
# .HTML_MainBody = "<CALLSTACK>" & vbCrLf & "<Accept>"
# .HTML_MoreInfoBody = vbNullString
# .HTML_CallStackItem = "<b><SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC>" _
# & "</b><br><br><VARIABLES><br>"
# .HTML_VariableItem = "   <font color=#808080>(<VARSCOPE>)</font>" _
# & "|<VARNAME> As <VARTYPE>| = <VARVALUE><br>"
# .WindowCaption = "Microsoft Visual Basic"
# .MinimumWindowWidth = 600
# .MoreInfoCaption = "More info"
# .LessInfoCaption = "Less info"
# .ButtonPaddingH = 10
# .ButtonPaddingV = 5
# .ButtonSpacingH = 5
# .ButtonSpacingV = 7
# .PaddingH = 15
# .PaddingV = 15
# .ScreenBorderPaddingV = 50
# .ColumnPaddingH = 20
# .LineSpacingV = 2
# .MainBackColor = 16777215
# .MainBackColor2 = 13693168
# .MainBackFillType = 8
# .MoreInfoBackColor = 15724768
# .MoreInfoBackColor2 = 16774642
# .MoreInfoBackFillType = 0
# .ButtonBarBackColor = 16443364
# .ButtonBarBackColor2 = 14337988
# .ButtonBarBackFillType = 1
# .MaxNumCallStackItems = 10
# .MaxNumVariablesItems = inv
# .DefaultButtonID = 3
# .DefaultButtonIsBold = True
# .ShowMoreInfoButton = False
# .AllowEnterKey = True
# .AllowTabKey = True
# .AllowArrowKeys = False
# .Timeout = 0
# .FlashWindowOnOpen = True
# .CustomImageTransparentColor = -1
# Dim TempImageData2 As String
# TempImageData2 = vbNullString
# .CustomImageData = TempImageData2
# Call .RemoveAllButtons
# Call .AddButton("Close", BUTTONACTION_VARIABLES_CLOSE)
# End With                                       ' ErrEx.VariablesDialogOptions

# ProcReturn:
# Call ProcExit(zErr)

# '---------------------------------------------------------------------------------------
# ' Method : ShowLiveNameCount
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Count the defined names in the D_LiveStack Dictionary.
# '          If Filter=True, eliminate the undefined ones
# '---------------------------------------------------------------------------------------
def showlivenamecount():
    if Not ErrExActive Then:
    # GoTo ProcRet

    # Dim i As Long
    # Dim LCI As cCallEnv
    # Dim LiveStackCount As Long

    # Call ShowLiveStack(doPrint:=doPrint, _
    # getNewStack:=getNew, Full:=Full)

    # LiveStackCount = D_LiveStack.Count
    # Set LCI = D_LiveStack.Items(i)
    if LenB(LCI.CallerInfo) > 0 Then ' caller is not defined:
    # ShowLiveNameCount = ShowLiveNameCount + 1

    if doPrint Then:
    if ShowLiveNameCount = 0 Then:
    print(Debug.Print "there are no relevant entries in D_LiveStack")
    else:
    print(Debug.Print " there are " & ShowLiveNameCount _)
    # & " relevant entries in D_LiveStack, other=" _
    # & D_LiveStack.Count - ShowLiveNameCount

    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Z_SourceAnalyse
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Test if a SourceLine contains a reference to sub or function
# '---------------------------------------------------------------------------------------
def z_sourceanalyse():
    # Const zKey As String = "BugHelp.Z_SourceAnalyse"
    # Call DoCall(zKey, tFunction, eQzMode)

    # Dim i As Long

    # sLine = Trim(Trunc(1, sLine, "'"))             ' drop comment part
    # i = InStr(sLine, ProcName)                     ' is ref contained at all?
    if i > 1 Then:
    if Mid(sLine, i - 1, 1) <> b Then          ' delimiters before=blank, (, or .:
    if Mid(sLine, i - 1, 1) <> "(" Then:
    if Mid(sLine, i - 1, 1) <> "." Then:
    # i = 0                          ' none: not valid reference to proc
    # sLine = Mid(sLine, i)                      ' if ref terminated by params?
    # i = InStr(sLine, "(")

    # Z_SourceAnalyse = (i > 0)

    # zExit:
    # Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : N_CheckStartSession
# ' Author : rgbig
# ' Date   : 15.07.2020
# ' Purpose: Check if OlSession has been started and set up everything if not
# '---------------------------------------------------------------------------------------
def n_checkstartsession():

    if OlSession Is Nothing Then:
    # Debug.Assert False                             ' currently doing test, remove         Stop
    # ' no need to Call N_DeInit
    # Set OlSession = New cOutlookSession
    # UseTestStart = UseTestStartDft                 ' using Default constant
    # Call N_DebugStart(stopRQ)
    # Call Z_StartUp(Not DidStop)
    # Call BugTimer.BugState_UnPause


# '---------------------------------------------------------------------------------------
# ' Method : Sub Z_StartUp
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Set or reset the main Application for Start
# '---------------------------------------------------------------------------------------
def z_startup():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "BugHelp.Z_StartUp"

    # '------------------- gated Entry -------------------------------------------------------
    # Static Recursive As Boolean

    # Dim zErr As cErr
    # Dim msg As String

    if Recursive Then:
    # ' choose Ignored or Forbidden and dependence on StackDebug
    if LogZProcs And Not P_Active Is Nothing Then:
    print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
    # & P_Active.DbgId & " => " & zKey
    # GoTo ProcRet
    # Recursive = True                                ' restored by    Recursive = False ProcRet:

    # Set olApp = Outlook.Application

    if OlSession Is Nothing Then:
    # Set OlSession = New cOutlookSession

    # Set OlExplorer = New cOlExplorer
    # Call N_Prepare

    if aNameSpace Is Nothing _:
    # Or aRDOSession Is Nothing Then
    # aBugTxt = "get Namespace/ActiveExplorer"
    # Call Try(allowAll)
    # Set aNameSpace = olApp.GetNamespace("MAPI")
    # Catch
    # Set ActiveExplorerItem(1) = ActiveExplorer
    # Catch
    # aBugTxt = "get RDOSession"
    # Call Try(allowNew)
    # Set aRDOSession = CreateObject("Redemption.RDOSession")
    # Catch

    # aBugTxt = "RDO-Logon"
    # Call Try(allowNew)
    # aRDOSession.Logon                           'no param == an empty string to use the default MAPI profile
    # Catch
    # '---------------- Start of normal Entry ------------------------------------------------
    # Call DoCall(zKey, "Sub", eQzMode, ZAppStart)   ' Z_Startup replaces ThisOutlookSession.Application_Startup

    # Set zErr = ZAppStart.ErrActive                 ' ZAppStart has been replaced...
    # ZAppStart.CallCounter = 1                      ' ... always called only once
    # zErr.EventBlock = True                          ' no Events until we leave Setups
    # zErr.atCallState = eCActive
    # Set zErr.atCalledBy = ExternCaller.ErrActive
    # Set E_Active = zErr
    # Set P_EntryPoint = ZAppStart

    if ExternalEntryCount > 0 Then:
    if UseStartUp = 2 Then:
    # doStop = False
    if UseStartUp = 0 Then:
    # ProtectStackState = inv
    # UseStartUp = 2                              ' continue open items after inits (in StartMainApp)
    elif UseStartUp = 1 Then:
    print(Debug.Print "---------- Z_StartUp recognized a re-start --------------")
    if doStop And Not DidStop Then:
    # DoVerify False, "Application Restart programmed stop"
    elif UseStartUp = 2 Then:
    # UseStartUp = 1

    # NoPrintLog = False
    # DebugMode = False
    # LogImmediate = True

    # Call Z_InitBugHelp(ZProcStart)

    if TopFolders Is Nothing Then                  ' Inits for Outlook Item Classes:
    # Call Z_olInits

    # Call N_ShowProgress(CallNr, ZAppStart, _
    # ZAppStart.Key, "Ready for Application", vbNullString)
    # Set E_AppErr = ZAppStart.ErrActive

    if LenB(UseErrExOn) > 0 Then:
    # WithLiveCheck = True

    if UseStartUp = 0 Then                     ' directly called from an entry point:
    # UseStartUp = 1                         ' resume right there
    elif UseStartUp = 2 Then                 ' use the EP under Test:
    # Call OlSession.StartMainApp            ' this is any App under test

    # Call getDebugMode(ExternalEntryCount = 1)  ' fetch from profile on first call

    if StackDebug > 4 Then:
    # Set aNonModalForm = frmErrStatus       ' must NOT use New!
    # Call ShowErrorStatus

    if ErrStatusFormUsable Then:
    # frmErrStatus.fNoEvents = E_Active.EventBlock
    # Call Start_BugTimer

    print(Debug.Print String(OffCal, b) & _)
    # LString("Error Module 'BugHelp': Initiation completed by " _
    # & Application.Name, lKeyM) _
    # & "#P=" & LString(0, 5) & ExternCaller.Key

    # Set zErr = Nothing

    # ZShowErrorStatus.CallCounter = 1          ' force Position of window
    # Call ShowErrorStatus
    # AppStartComplete = True
    # Call FldActions2Do                         ' open items

    # zExit:
    # Call DoExit(zKey)
    if E_Active.atCallDepth < 1 Then           ' 0 is ExternCaller:
    # Call BugTimerDeActivate
    print(Debug.Print "Timer DeActivated")
    # Call N_ShowHeader("BugHelp Log " & TimerNow)
    # msg = "* " & String(OffCal - 2, "-") _
    # & " Outlook waiting for Events or Macro Calls " _
    # & String(OffCal - 2, "-") & " *"
    # Call LogEvent(String(Len(msg), "-") _
    # & vbCrLf & msg & vbCrLf _
    # & String(Len(msg), "-"), eLSome)
    # ZAppStart.ErrActive.EventBlock = False

    # E_AppErr.EventBlock = False
    # Call SetOnline(olCachedConnectedFull)
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Z_StateToTestVar
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def z_statetotestvar():
    # '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
    # Const zKey As String = "BugHelp.Z_StateToTestVar"
    # Call DoCall(zKey, "Sub", eQzMode)

    if InStr(1, Testvar, "OFF", vbTextCompare) > 0 Then:
    # DebugMode = False
    # StackDebug = 0
    # TraceMode = False
    # LogAppStack = False
    # LogPerformance = False
    # Testvar = Trim("OFF")
    else:
    # Testvar = vbNullString                     ' build new from State
    if LogAllErrors Then:
    # Testvar = "ERR " & Testvar
    if DebugLogging Then:
    # Testvar = "LOG " & Testvar
    # Testvar = "StackDebug=" & StackDebug & b & Testvar
    if LogPerformance Then:
    # Testvar = "LogPerformance " & Testvar
    if ShowFunctionValues Then:
    # Testvar = "ShowFunctionValues " & Testvar
    if TraceMode Then:
    # Testvar = "TraceMode " & Testvar
    if DebugMode Then:
    # Testvar = "DebugMode " & Testvar

    # Testvar = Trim(Testvar)

    # zExit:
    # Call DoExit(zKey)

# '---------------------------------------------------------------------------------------
# ' Method : Z_UsedThisCall
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Compute Performance Data, getting Z_UsedThisCall
# '---------------------------------------------------------------------------------------
def z_usedthiscall():
    # Static Recursive As Boolean

    # '------------------- gated Entry -------------------------------------------------------
    # Z_UsedThisCall = 0

    if Recursive Then                              ' no Message, because recursive happens but does not matter:
    print(' Do not set message: Debug.Print String(OffCal, b) & "Forbidden recursion from ...")
    # GoTo ProcRet

    if ErrClient.atProcIndex = inv Then:
    # GoTo ProcRet                               ' before it is defined, proc can't have UsedThisCall
    if ErrClient.atLastInSec = 0& Then             ' Proc without time, e.g. ProcExit:
    # GoTo ProcRet
    if ErrClient.atThisEntrySec = 0& Then:
    # GoTo ProcRet                               ' unnecessary to compute TimeUsed when exited

    # Recursive = True

    # Dim DiffDate As Double

    # With ErrClient
    # DiffDate = DateDiff("d", Date, .atLastInDate) * 86400#
    # Z_UsedThisCall = TimeNow - .atThisEntrySec + DiffDate
    if .atDsc Is Nothing Then:
    # Z_UsedThisCall = TimeNow - .atThisEntrySec
    else:
    # With .atDsc
    # .TotalProcTime = .TotalProcTime + Z_UsedThisCall
    # End With                               ' .atDsc
    # .atPrevEntrySec = .atThisEntrySec
    # .atThisEntrySec = 0
    # End With                                       ' ErrClient

    # FuncExit:
    # Recursive = False
    # ProcRet:

