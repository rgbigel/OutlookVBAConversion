Attribute VB_Name = "BugHelp"
Option Explicit
                                    
' Name Conventions: Procs in this module should be simple to recognize on Stacks
'        names N_ are not visible on any stacks, do not use Z_Entry/DoExit etc.
'                 these should (must ???) never call non-N_Type Procs
'        names Z_ are defined but not usually visible
'                   Mode must be <=eQyMode<=eQxDMode
'              Proc/Live/Show/Query equivalent "Z_", e.g. ProcCall, ProcReturn etc.
'        Z_    Procs which Show/manipulate stacks
'              Z_ = x/y/z_Type procs visible on stacks, iff they call DoCall/DoExit or found on LiveStack
'              Z_   not normally using D_ErrInterface
'              ALL  Procs can use ProcCall and ProcExit, because they use DoCall / DoExit
'        Z_Types are used for all Procs when FastMode = True (changing Qmode to y)
'        All other Names can be of any Qmode-Type, most of these use:
' DoCall      Mark entry to  EVERY Proc (Sub, Function, ... )
' ProcCall      "  as above, but with special rules for error handling
' DoExit      Mark exit from EVERY Proc (Sub, Function, ... )
' ProcReturn    "  as above, but with special rules for error handling
' StartEP     like ProcCall, but used for external entry points/event handlers (EP-Procs)
' ReturnEP    like ProcReturn, but used when Ending EP-Procs

' Note      ' DscMode Determines the eQmode which sets the rules for ProcCall
'               all eQmodes>eQzMode are in D_ErrInterface Collection and do the following:
'               Watches for Err.Number<>0 and handles Errors according to eQmode
'               Eliminate ClientDsc's marked invalid(???)
'               Completes all values in ClientDsc and adds new ones to D_ErrInterfaceM
'               Calls N_ConDscErr to define consistent cErr/cProcDsc
'               for non-predefining calls, use N_StackDef to put on any Stacks,
'                  with N_CallEnv to define Call Environment (like back ref to Caller etc),
'                  Set Date/Time of ClientDsc's call
'               Adds to D_AppStack Stack (if QMode >= eQASMode)
'               Optionally Logs using N_LogEntry-->LogGen

' If ClientDsc not defined when using LiveStack:
'                  create cProcDsc and add to D_ErrInterface (Dictionary)
'                  create cErr    (as ClientDsc.ErrActive, using N_ConDscErr)
' ProcCall:  Notes the Call of a procedure of any QMode using DoCall
'                  add 1 to atDsc's CallCounter
'                  inserts atCalledBy from D_AppStack
'                 Checks if recursion is allowed.
'                    If so and recursion level>1, create cErr Instance and chain it
'                 Pauses Callers
'                 If QMode >= eQASMode adds to D_AppStack by Call Z_ToAppStack

' Hierarchy of Calls of Procedures in this Module

' General:      use the following general parameter list:
'               (ClientDsc As cProcItem, DscMode As eQmode, tSub/tFunction As String, Mode As String, _
'                Optional ClassInstance As Long, Optional RecursionRequested As Boolean)

'---------------------------------------------------------------------------------------
' Method : BugEval
' Author : Rolf G. Berchtocall
' Date   : 20211108@11_47
' Purpose: Re-Evaluate settings in aNonModalForm
'---------------------------------------------------------------------------------------
Sub BugEval()

Const zKey As String = "BugHelp.BugEval"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug > 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                          & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcRet
    End If
    Recursive = True

    Call DoCall(zKey, tSub, eQzMode)

    If ErrStatusFormUsable Then
        frmErrStatus.fModifications.Enabled = True
        frmErrStatus.frmIgnoreErrStatusChange = True
        frmErrStatus.fModifications = True
        Call frmErrStatus.ReEvaluate
        StackDebugOverride = StackDebug            ' this value is always >= 0
        Call ShowOrHideForm(frmErrStatus, Not frmErrStatus.fHideMe)
    End If
    
FuncExit:
    Recursive = False
    
zExit:
    Call DoExit(zKey)
ProcRet:
End Sub                                            ' BugHelp.BugEval

'---------------------------------------------------------------------------------------
' Method : BugSet
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Update settings in aNonModalForm from Program vars, espc. from E_Active
'---------------------------------------------------------------------------------------
Sub BugSet()

Const zKey As String = "BugHelp.BugEval"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If StackDebug >= 8 Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                          & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcRet
    End If
    Recursive = True

    If ErrStatusFormUsable Then
        ErrDisplayModify = True
        Call frmErrStatus.UpdInfo
    End If
    
FuncExit:
    Recursive = False

ProcRet:
End Sub                                            ' BugHelp.BugSet

'---------------------------------------------------------------------------------------
' Method : Function Catch
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: After Try, if any error was found, check the details and Clear if acceptable
' Note   : the Global Error handler has analized if the error was acceptable (by Try-Rules)
'          if an error did occur, but was acceptable, the Error is resetted
'                           unless DoClear=False is specified
'          The Returnvalue of Catch indicates if there was an Error at all.
'---------------------------------------------------------------------------------------
Function Catch(Optional DoClear As Boolean = True, Optional DoMessage As Boolean = True, Optional AddMsg As String, Optional HandleErr As Variant) As Boolean

    If T_DC.TermRQ Then
        Call TerminateRun                          ' aborting Entry if not handled
    End If

Dim msg As String
        
    With E_Active
        If LenB(.Explanations) > 0 Then
            msg = .Explanations
        ElseIf LenB(aBugTxt) > 0 Then
            msg = aBugTxt
        ElseIf LenB(T_DC.DCerrMsg) > 0 Then
            msg = T_DC.DCerrMsg
        End If
        If LenB(msg) = 0 Then
            msg = "Catch"
        End If
        If .errNumber <> 0 Then
            If IsMissing(HandleErr) Then
                Catch = True
            ElseIf T_DC.DCerrNum <> HandleErr Then
                Catch = True
            Else
                GoTo ErrOk
            End If
        End If
        If .FoundBadErrorNr = 0 And T_DC.DCerrNum = 0 Then
ErrOk:
            If DoMessage And (DebugMode Or DebugLogging Or InStr(msg, "***") > 0) Then
                Call LogEvent("OK:" & msg)
            End If
            If DoClear Then
                Call ErrReset(4)
            End If
        Else
            If .FoundBadErrorNr <> 0 Then
                If LenB(AddMsg) > 0 Then
                    msg = msg & vbCrLf & AddMsg
                End If
                Call LogEvent("!!! Failed: " & msg)
                Debug.Assert False
            End If
            If DoClear Then
                Call ErrReset(4)
            End If
        End If
    End With                                       ' E_Active
    aBugTxt = vbNullString
    
ProcRet:
End Function                                       ' BugHelp.Catch

' ------ Variation of Catch with different parms ---------------------------------------
Function CatchNC(Optional AddMsg As String, Optional DoMessage As Boolean = True, Optional DoClear As Boolean, Optional HandleErr As Variant) As Boolean
    Call Catch(DoClear, DoMessage, AddMsg, HandleErr)
End Function                                       ' CatchNC

' Global Interface variables of BugHelp see Module Z_ErrIf, generated by ZZIfGen (but long outdated)

'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Method : DoCall
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: If #MoreDiagnostics, maintain entries in D_Errinterface
'---------------------------------------------------------------------------------------
Sub DoCall(ClientKey As String, CallType As String, Mode As eQMode, Optional CalledDsc As cProcItem, Optional ExplainS As String)
Const zKey As String = "BugHelp.DoCall"

Dim aKey As String
Dim aCallType As String
Dim aMode As eQMode
Dim aDsc As cProcItem                              ' that is <Self>, not Client (CalledDsc, aDsc.ErrActive)
Dim aErr As cErr
Dim TimeIn As Double
Dim isNewDsc As Boolean

    If QuitStarted Then
        Exit Sub
    End If
    
    TimeIn = Timer
    
    Call BugTimer.BugState_SetPause             ' do not allow any timer and other events
    
    If D_ErrInterface.Count = 0 Then            ' first definition, come back for ActiveProc below
        aCallType = tSub
        aMode = eQnoDef
        Call N_ConDscErr(aDsc, "Extern.Caller", tSub, aMode, aErr)
        Set aErr.atCalledBy = aErr              ' self reference, the only allowed case
        Call N_StackDef(aDsc, aErr, False)      ' External Caller is the Stack root, but has no at_Live
        Set ExternCaller = aDsc
        Set P_Active = aDsc
        Set E_Active = aErr
        
        aKey = zKey                             ' use this to define atDsc / Err for DoCall
        Set aDsc = Nothing
        Set aErr = Nothing                      ' continue for really called Prog later!
    Else
        aKey = ClientKey
        aCallType = CallType
        aMode = Mode
ActiveProc:
        If LenB(aKey) = 0 Then
            Stop                                ' is this ever possible ???
            GoTo isNew
        End If
        If CalledDsc Is Nothing Then
            GoTo unknown
        ElseIf CalledDsc.Key = vbNullString Then
            GoTo unknown
        ElseIf CalledDsc.ErrActive Is Nothing Then
unknown:
            Set aDsc = Nothing
            Set aErr = Nothing
        End If
    End If
    
    If aDsc Is Nothing Then
        If Not aErr Is Nothing Then             ' should not ever happen ???
            DoVerify aErr.atKey = aKey, _
                    "** Key mismatch at definition of cErr ???"
            Set aDsc = aErr.atDsc
            DoVerify Not aDsc Is Nothing, _
                    "inconsistent cErr ???"
            aDsc.Key = aKey
            aCallType = CallType
            aMode = Mode
            GoTo isNew
        End If
    Else
        If Not dontIncrementCallDepth Then
            DoVerify aDsc.Key = aKey, _
                    "** Key mismatch at definition of aDsc=" _
                     & aDsc.Key & "<>" & aKey
        End If
    End If
        
    If D_ErrInterface.Exists(aKey) Then
        If isEmpty(D_ErrInterface.Item(aKey)) Then
            DoVerify False, _
                "*** VBA Bug: how can D_ErrInterface.Item be Empty ???"
            Call N_ConDscErr(aDsc, aKey, tSub, aMode, aErr)
            Set D_ErrInterface.Item(aKey) = aDsc
            If aErr Is Nothing Then
                Set aErr = aDsc.ErrActive
            End If
            Set aDsc.ErrActive = aErr
        Else
            Set aDsc = D_ErrInterface.Item(aKey)
            If aErr Is Nothing Then
                Set aErr = aDsc.ErrActive
            Else
                Stop ' ??? need if ???
            End If
        End If
        
        Set aDsc.ErrActive = aErr
        If aDsc Is Nothing Or aErr Is Nothing Then
            DoVerify False, _
                "D_ErrInterface.Item atDsc/Err is Nothing ???"
            GoTo isNew
        Else
            aBugTxt = "D_ErrInterface.Item Key in atDsc <> Err ???"
            If DoVerify(aErr.atKey = aKey) Then
                GoTo isNew
            End If
            GoTo isKnown
        End If
    Else
isNew:
        Call N_ConDscErr(aDsc, aKey, CallType, aMode, aErr)
        If aErr.atCallDepth < 0 Then
            aErr.atCallDepth = 0                            ' show it is new in DoCall
        End If
        aErr.atMessage = ExplainS
        isNewDsc = True
isKnown:
        Set CalledDsc = aDsc                                ' CalledDsc: delivering values
        If N_StackDef(aDsc, aErr, False) Then               ' False:: always GenCallData
            If aKey = zKey Then
                Set N§Call = aDsc                           ' self-Defining has been done
                N§Call.CallCounter = 0                      ' Inc happens below
                If ClientKey <> D_ErrInterface.Keys(0) Then ' not again for Extern.Caller
                    aKey = ClientKey
                    Set aDsc = Nothing
                    Set aErr = Nothing
                    GoTo ActiveProc                         ' now define true Caller
                Else
                    Set P_Active = D_ErrInterface.Item(ClientKey)
                    Set E_Active = P_Active.ErrActive
                End If
            Else
                Set aDsc.ErrActive = aErr
            End If
        End If
    End If
    
FuncExit:
    N§Call.CallCounter = N§Call.CallCounter + 1
    N§Call.TotalProcTime = N§Call.TotalProcTime + Timer - TimeIn
    
    If ExLiveDscGen And Not dontIncrementCallDepth Then    ' not predefining; if Live Stack use requested
        aErr.atLiveLevel = ExLiveCheck(ForcePrint:=ExLiveCheckLog, LocalStack:=D_LiveStack)
    End If
    
    If Not aErr.atCalledBy Is Nothing Then
        aErr.atCalledBy.atCallState = eCpaused
    End If
    Set aDsc = Nothing
    Set aErr = Nothing
    Set CalledDsc = P_Active
    If LenB(ExplainS) > 0 And CallLogging Then
        Call LogEvent("Calling " & P_Active.DbgId & b & ExplainS)
    End If
    Call BugTimer.BugState_UnPause                 ' do restore timer and other events
    
ProcRet:
End Sub                                            ' BugHelp.DoCall

'---------------------------------------------------------------------------------------
' Method : ExLiveCheck
' Author : rgbig
' Date   : 20211108@11_47
' Purpose: Get Call Level of running VBA Procs and the current top running proc
' Note   : ignores itself, Main result is in the Public LCA (cCallEnv)
'---------------------------------------------------------------------------------------
Function ExLiveCheck(Optional ForcePrint As Boolean, Optional ClearLCS As Boolean = True, Optional MustMatch As String, Optional LookUpDsc As Boolean, Optional SkipRemainingStack As Boolean, Optional LocalStack As Dictionary) As Long
    
Dim LTimer As Double
Dim msg As String
Dim ModProc As String
Dim FoundModProc As String
Dim i As Long
Dim MatchFullQualified As Boolean
Dim ActiveProc As String
Dim HaveMatch As Boolean
Dim IsRelevant As Boolean

Dim aProcDsc As cProcItem
Dim nErr As cErr
Dim aErr As cErr

    If ExLiveCheck Then
        SkipRemainingStack = False
        MustMatch = vbNullString
        MatchFullQualified = True
        LookUpDsc = True
    ElseIf MustMatch <> vbNullString Then
        MatchFullQualified = InStr(MustMatch, ".")
        If LookUpDsc And Not ExLiveCheck Then
            If DoVerify(Not SkipRemainingStack, "if we have MustMatch, must SkipRemainingStack") Then
                SkipRemainingStack = True
            End If
        End If
    Else
        MatchFullQualified = True
        If SkipRemainingStack Then
            If DoVerify(Not LookUpDsc, _
                        "if we have LookUpDsc, can't SkipRemainingStack") Then
                SkipRemainingStack = False
            End If
        End If
    End If
    LTimer = Timer
    Set LocalStack = N_GetLiveStack
    msg = "Live Stack Root"
        
    For i = LocalStack.Count - 1 To 1 Step -1
        Set LCA = LocalStack.Items(i)
        ExLiveCheck = LocalStack.Count - i
        IsRelevant = N_RelevantProc(LCA, ForcePrint, ModProc)
        IsEntryPoint = InStr(LCA.ProcedureName, "Appl") > 0
        If IsRelevant Then
FoundPreCall:
            ModProc = Left(LCA.ModProc & String(30, b), 30)
            If i = 1 Then
                If ForcePrint Then
                    msg = "This is the active Proc, called by " & ActiveProc
                End If
            Else
                If ForcePrint And i < LocalStack.Count - 1 Then
                    msg = "Root at - " & i
                End If
                If i = 2 Then
                    ActiveProc = ModProc
                End If
            End If
            
            If MustMatch = IIf(MatchFullQualified, LCA.ModProc, LCA.ProcedureName) Then
                HaveMatch = True
                If ForcePrint Then
                    msg = Trim(msg & " found the MustMatch " & MustMatch)
                End If
                FoundModProc = LCA.ModProc
                If ForcePrint Then
                    Debug.Print ExLiveCheck & " / " & LSD, _
                                FoundModProc, "Used=" & Timer - LTimer, _
                                "Line " & LCA.LineNumber & " Code: " & LCA.LineCode _
                                & vbCrLf & String(80, "-")
                End If
                GoTo FoundIt
            End If
            
            If ForcePrint And Not ShutUpMode Then
                Debug.Print i, ModProc, _
                        LString(msg, 20) & "--->", LCA.LineCode
            End If
            
            If LenB(FoundModProc) = 0 Then
                If LSD > 0 Then
                    If LSD > ExLiveCheck Then
                        GoTo PrChange
                    End If
                ElseIf MustMatch = vbNullString Then
ChangeIt:
                    If ForcePrint And Not ShutUpMode Then
PrChange:
                        Debug.Print i, ModProc, _
                                    LString("LSD has been set", 20), _
                                    "to Call Depth=" & ExLiveCheck
                    End If
                    LSD = ExLiveCheck
                End If
            End If                                 ' IsRelevant
FoundIt:
            If ExLiveCheck Or IsEntryPoint _
            Or (LookUpDsc And LCA.CallerErr Is Nothing) Then
                If D_ErrInterface.Exists(LCA.ModProc) Then              ' N_ Procs never exist
                    If isEmpty(D_ErrInterface.Item(LCA.ModProc)) Then
                        Set D_ErrInterface.Item(LCA.ModProc) = aProcDsc ' should be Nothing
                        GoTo mustCorrErr
                    End If
                    Set aProcDsc = D_ErrInterface.Item(LCA.ModProc)
                    Set nErr = aProcDsc.ErrActive
                Else
                    If ExLiveCheck Then               ' generate a description for all procs
mustCorrErr:
                        Set nErr = New cErr
                        Call N_ConDscErr(aProcDsc, LCA.ModProc, LCA.DscKind, eQnoDef, nErr)
                        nErr.Explanations = "! ExLiveCheck=True !"
                        Set LCA.CallerErr = nErr
                        If dontIncrementCallDepth Then
                            Set aProcDsc = nErr.atDsc
                        Else
                            nErr.atCallDepth = ExLiveCheck
                            dontIncrementCallDepth = True   ' no push because False at start
                            Call N_StackDef(aProcDsc, nErr, IsRelevant)
                            dontIncrementCallDepth = False  ' no pop necessary
                            If LCA.CallerErr Is Nothing Then
                                Set LCA.CallerErr = aProcDsc.ErrActive
                            End If
                        End If
                    End If
                End If
            Else
            End If
            If nErr Is Nothing Then
                Stop ' ???
            Else
                Set LCA.CallerErr = nErr
                Set aErr = nErr
                Set aErr.at_Live = LocalStack       ' maybe we do not need all of that ???
                nErr.atCatDict = Trim(nErr.atCatDict & b & nErr.atKey)
                If Not aErr.atCalledBy Is Nothing Then
                    If DoVerify(nErr.atCalledBy Is aErr.atCalledBy, _
                            "??? CalledBy correct? is " _
                            & nErr.atCalledBy.atKey & " to " _
                            & aErr.atCalledBy.atKey) Then
                        Set nErr.atCalledBy = aErr.atCalledBy
                    End If
                End If
                
                If i > 1 Then
                    nErr.atCallState = eCpaused             ' all procs below active are paused
                End If
            End If
                        
            If SkipRemainingStack And FoundModProc <> vbNullString Then
                GoTo FuncExit
            End If
        End If
    Next i
    
FuncExit:
    Set aProcDsc = Nothing
    Set aErr = Nothing
    Set nErr = Nothing
    Set LocalStack = Nothing
    
    If ClearLCS Then
        Set LCS = Nothing
    End If
   
ProcRet:
End Function                                       ' BugHelp.ExLiveCheck

'---------------------------------------------------------------------------------------
' Method : N_RelevantProc
' Author : rgbig
' Date   : 13.03.2020
' Purpose: Determine if Proc is relevant based on Name
'---------------------------------------------------------------------------------------
Function N_RelevantProc(LCA As cCallEnv, ForcePrint As Boolean, ModProc As String) As Boolean

    N_RelevantProc = True                          ' all the rest is relevant

    If Left(LCA.ProcedureName, 1) = "<" Then       ' immediate and external code (eg. events)
        GoTo IsIrrelevant
    ElseIf InStr("ExLi Live DoCa DoEx Proc Show Query ", _
                 Left(LCA.ProcedureName, 4)) > 0 Then
IsIrrelevant:
        N_RelevantProc = False                     ' DoCall etc. are irrelevant
        If ForcePrint Then
            ModProc = Left(LCA.ModProc & String(30, b), 30)
            Debug.Print LCA.StackDepth, ModProc, _
                        LString("Irrelevant Proc", 20) _
                        & "--->", LCA.LineCode
        End If
    End If
         
End Function                                       ' BugHelp.N_RelevantProc

'---------------------------------------------------------------------------------------
' Method : DoExit
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Exit required after for DoCall. Used in ProcExit
'---------------------------------------------------------------------------------------
Sub DoExit(zKey As String, Optional DisplayValue As String)
Dim fromDsc As cProcItem
Dim fromErr As cErr
Dim toErr As cErr
Dim Consumed As String
Dim FuncVal As String

    If QuitStarted Then
        Exit Sub
    End If
    
    Call BugTimer.BugState_SetPause                ' do not allow any timer and other events

    aBugTxt = "Return without a call definition ???"
    If DoVerify(D_ErrInterface.Exists(zKey)) Then
        GoTo ERET
    End If
    aBugVer = isEmpty(D_ErrInterface.Item(zKey))
    If aBugVer Then
        If StackDebug >= 8 Then
            Debug.Print zKey & "has empty cErr ???"
            Debug.Assert False
        End If
        GoTo ERET
    End If
    
    If E_Active.errNumber <> 0 Then                ' any pending errors?
        Call N_CaptureNewErr                       ' handle error as defined (likely will pause there)
    End If
    If T_DC.TermRQ Then                            ' aborting Entry if not handled by user
        Call TerminateRun                          ' never returns unless user says so
    End If
    If E_Active.errNumber = 0 Then
        T_DC.N_ClearTermination
    End If

    Set fromDsc = D_ErrInterface.Item(zKey)
    Set fromErr = fromDsc.ErrActive
    Set toErr = E_Active.atCalledBy
    If toErr Is Nothing Then                       ' external Caller
        Set toErr = D_ErrInterface.Items(0)
    End If
    DoVerify toErr.atCallDepth <= 1 Or Not fromErr Is toErr, _
                                        "??? returning to itself=" & toErr.atKey
    DoVerify fromErr Is fromDsc.ErrActive, "atDsc and Err not linked ???"
    If DoVerify(fromDsc.Key = fromErr.atKey, "Dsc.Key <> Err.atKey ??? in " & fromDsc.Key) Then
        fromErr.atKey = fromDsc.Key
    End If
    If DoVerify(Not E_Active.atDsc Is Nothing, "E_Active.atDsc is Nothing") Then
        Set E_Active = fromErr
    End If
    
    If LogPerformance Then
        If fromErr.atThisEntrySec > 0 Then
            Call Z_UsedThisCall(fromErr, Timer)
            Consumed = " Consumed " & ElapsedTime & " sec "
        End If
    Else
        Consumed = vbNullString
    End If
    
    fromErr.atCallState = eCExited
        
    If LogZProcs Or Not AppStartComplete Then
        If LenB(DisplayValue) > 0 Then
            FuncVal = " >" & DisplayValue
        Else
            FuncVal = vbNullString
        End If
        Call N_ShowProgress(CallNr, fromDsc, "-Z", _
                            "zD=" & toErr.atCallDepth & "«" & fromErr.atCallDepth _
                            & IIf(E_Active.EventBlock, " NoEvents", vbNullString), _
                            Consumed & FuncVal, ErrClient:=fromErr)
    End If
    
    Set fromErr.atErrPrev = Nothing                ' no previous error instance
    Set fromErr.at_Live = Nothing                  ' no call stack after exit
    fromErr.atCatDict = vbNullString
    
    If toErr.atCallDepth <> 0 Then                 ' only extern caller has =0
        If DoVerify(toErr.atRecursionLvl > 0, _
                    "error in recursion level, too many returns ???") Then
            fromErr.atRecursionLvl = 0             ' this does not really fix the problem!
        Else
            fromErr.atRecursionLvl = fromErr.atRecursionLvl - 1
        End If
        If toErr.atDsc.CallMode > eQnoDef Then
            DoVerify toErr.atCallState = eCpaused, _
                 "exiting to caller: " & toErr.atKey _
                 & "=" & CStateNames(toErr.atCallState) _
                 ' unlikely call state, it should be Paused
        End If
    End If
    
    If ExLiveDscGen Then
        If ExLiveDscGen = toErr.atLiveLevel Then
            If DebugMode And LogZProcs Then
                Debug.Print "DoExit setting CallState to Active for " & toErr.atKey
            End If
            toErr.atCallState = eCActive
        End If
    End If
    
    Set P_Active = toErr.atDsc
    Set E_Active = toErr
    Call BugTimer.BugState_UnPause                 ' restore any timer and other events
    
ERET:
    Set fromDsc = Nothing
    Set fromErr = Nothing
    Set toErr = Nothing
    aBugTxt = vbNullString                         ' no active DoVerify
    aBugVer = True
End Sub                                            ' BugHelp.DoExit

'---------------------------------------------------------------------------------------
' Method : DoVerify
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Stop if Debug Condition false (with optional message)
'---------------------------------------------------------------------------------------
Function DoVerify(Optional NoStop As Variant, Optional Message As String, Optional ShowMsgBox As Boolean) As Boolean

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Long

Dim DebugHalt As Boolean

    If Recursive > 0 Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If Recursive > 1 Then
            GoTo blockit
        End If
        If NoStop Then                             ' not an error within error
            If DebugLogging Then
blockit:
                Debug.Print String(OffCal, b) & "Ignored recursion from DoVerify"
                If Not NoStop Then
                    Debug.Print String(OffCal, b) & "Message: " & Message
                End If
                If DebugLogging Then
                    DebugHalt = True
                End If
            End If
        End If
        GoTo FuncExit
    End If
    Recursive = Recursive + 1                      ' restored by    Recursive = False ProcRet:
    
    If IsMissing(NoStop) Then
        NoStop = aBugVer
    End If
    If LenB(Message) = 0 Then
        Message = aBugTxt
    End If
    If NoStop Then
        DoVerify = False
        Message = "Verified OK:" & vbCrLf & Message
        If LogAllErrors Then
            If InStr(Message, testAll) > 0 Then
                Call LogEvent(Message)
            End If
            If InStr(Message, "***") > 0 And DebugMode Then
                DebugHalt = True
            End If
        End If
    Else
        DoVerify = True
        E_AppErr.Reasoning = Message
        Call BugTimerDeActivate                    ' Function must not Recurse! (Gated)
        If LenB(Message) = 0 Then
            If E_Active Is Nothing Then
                Message = "*** Debug Stop requested in uninited state "
            Else
                Message = "*** Debug Stop requested in " & E_Active.atKey
            End If
        End If
        If Recursive = 1 Then
            If ShowMsgBox Then
                Call MsgBox(Message, vbOKOnly)
            Else
                Debug.Print Message
                If InStr(Message, "***") > 0 Or DebugMode Then
                    DebugHalt = True
                End If
            End If
        End If
    End If

FuncExit:
    E_AppErr.Reasoning = Message
    If DebugHalt Then
        Debug.Assert False
    End If

zExit:
    Recursive = Recursive - 1
ProcRet:
End Function                                       ' BugHelp.DoVerify

'---------------------------------------------------------------------------------------
' Method : N_PrintNameInfo
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_PrintNameInfo(i As Long, CallEnv As cCallEnv)

Dim msg As String

    With CallEnv
        msg = LString(Right(String(5, b) & i, 5) _
                      & b & .ProcedureName, OffObj - lCallInfo) _
                      & b & LString(.CallerInfo, lCallInfo) _
                      & RString(.StackDepth, 3) _
                      & RString("L" & .LineNumber, 5) _
                      & b & Trim(.LineCode)
        Call LogEvent(msg, eLall)
    End With                                       ' CallEnv
        
End Sub                                            ' cNameInfo.N_PrintNameInfo

'---------------------------------------------------------------------------------------
' Method : ShowLiveStack
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show all info contained in D_LiveStack Dictionary
'---------------------------------------------------------------------------------------
Sub N_ShowLiveStack()
Dim i As Long
Dim CallEnv As cCallEnv

    For i = 0 To D_LiveStack.Count - 1
        Set CallEnv = D_LiveStack.Items(i)
        Call N_PrintNameInfo(i + 1, CallEnv)
    Next i
    
    Set CallEnv = Nothing
    
End Sub                                            ' cNameInfo.N_ShowLiveStack

'---------------------------------------------------------------------------------------
' Method : Sub getDebugMode
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub getDebugMode(Optional forceGet As Boolean)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "BugHelp.getDebugMode"
Dim zErr As cErr

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If StackDebug > 8 Then
        GoTo ProcRet
    End If
    
    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    
Dim originalTestVar As String

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, Recursive:=False)

    If forceGet Then
        Testvar = GetEnvironmentVar("Test")        ' Variable "Test" im "1. sichtbaren" Environment
    End If
    
    originalTestVar = Testvar
 
    If LenB(Testvar) > 0 Then
        Call N_InterpretTestVar
    Else
        Call Z_StateToTestVar
    End If
    If StackDebugOverride > 0 Then
        If StackDebugOverride <> StackDebug Then
            Call Z_StateToTestVar
        End If
        StackDebugOverride = -StackDebug           ' do not override again
    End If
        
    If originalTestVar <> Testvar Then
        Debug.Print "Changing Variable 'Test' to: " & Quote(Testvar)
        Debug.Print "Previously was " & String(14, b) & Quote(originalTestVar)
        
        Call SetEnvironmentVariable("Test", Testvar)
        Call SetGlobal("Test", Testvar)
    End If

ProcReturn:
    Call ProcExit(zErr)
    
ProcRet:
End Sub                                            ' BugHelp.getDebugMode

' Method : N_CaptureNewErr
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Capture change of Err.Number and pass to proc that analyzes it
' Note   : when called via ErrEx, use ErrEx.Number instead
'---------------------------------------------------------------------------------------
Sub N_CaptureNewErr(Optional ErrExEvent As Boolean)
Const zKey As String = "BugHelp.N_CaptureNewErr"

    If E_Active Is Nothing Then
        Set E_Active = New cErr
        E_Active.atKey = "** undefined proc entering " & zKey & " **"
    End If
    
    With E_Active
        If ErrExEvent Then
            .errNumber = T_DC.DCerrNum
            .Description = T_DC.DCerrMsg
            .Source = T_DC.DCerrSource
        Else
            .errNumber = Err.Number
            .Description = Err.Description
            .Source = Err.Source
        End If
        
        If .errNumber = 0 Then                     ' shunt if no error. So the Entry is not logged in that case
            DoVerify False, "don't call N_CaptureNewErr if no error has occurred"
            Call ErrReset(4)
            GoTo ProcRet
        End If
        
        If Z§ErrSnoCatch Then                      ' N_PublishBugState Recursion not allowed
            aBugTxt = "**** previous error handling not complete when a new error occurred"
            DoVerify .FoundBadErrorNr = 0
            Call ErrReset(4)                       ' escape from this invalid state
            StackDebug = 9                         ' Trace this!
            GoTo ProcRet
        End If
        If MayChangeErr Then
            Err.Clear                              ' this error now cleared to get subsequent
        End If
        
        Z§ErrSnoCatch = True                       ' No ErrHandler Recursions for:
        .ErrSnoCatch = Z§ErrSnoCatch               ' NO N_PublishBugState
        If isEmpty(T_DC.DCAllowedMatch) Then       ' simple version testing acceptable errors
            GoTo trapIt
        ElseIf Left(T_DC.DCAllowedMatch, 1) = "*" Then
            GoTo ProcRet
        ElseIf T_DC.DCAllowedMatch = T_DC.DCerrNum Then
            GoTo ProcRet
        ElseIf InStr(T_DC.DCerrMsg, T_DC.DCAllowedMatch) > 0 Then
            GoTo ProcRet
        End If
        
trapIt:
        .FoundBadErrorNr = .errNumber
        StackDebug = 9                             ' Trace this!
    End With                                       ' E_Active
    
    If AppStartComplete Then
        Call ShowErr                               ' for testing only: forces frmErrStatus to show
        frmErrStatus.fErrNumber = T_DC.DCerrNum
        frmErrStatus.Top = 245
        frmErrStatus.Left = 1041
        Set aNonModalForm = frmErrStatus
        ErrStatusFormUsable = True
    End If
    
ProcRet:
End Sub                                            ' BugHelp.N_CaptureNewErr

'---------------------------------------------------------------------------------------
' Method : N_ChkRC
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Test on Err. Exits false if err.number is 0, sets ErrorCaught always
' Note   : if Try allows several errors as defined by E_Active
'        : Multiple N_ChkRC calls may follow, with ErrorCaught match.
'        : All but the last one must then code DoClear:=False or ErrorCaught missing
'---------------------------------------------------------------------------------------
Function N_ChkRC(Optional TryingThisCode As Variant, Optional WhatThisMeans As Variant, Optional MatchAllow As Variant, Optional DoClear As Boolean = True) As Boolean

Const zKey As String = "BugHelp.N_ChkRC"

Dim Putlead As String
    
    If T_DC.TermRQ Then
        Call TerminateRun                          ' aborting Entry if not handled
    End If

    If E_Active Is Nothing Then
        N_ChkRC = True
        GoTo ProcRet
    End If
        
    With E_Active
        If .FoundBadErrorNr = 0 And Err.Number <> 0 Then
            .errNumber = Err.Number
            .Description = Err.Description
            .FoundBadErrorNr = Err.Number
        End If
        If .errNumber = 0 Then
            GoTo FuncExit                          ' no problem!
        End If
        If LenB(.atKey) = 0 Then
            .atKey = "** undefined proc entering " & zKey
        End If
    
        If Not IsMissing(TryingThisCode) Then
            .Explanations = TryingThisCode
        End If
        If Not IsMissing(WhatThisMeans) Then
            .Reasoning = WhatThisMeans
        End If
        If IsMissing(MatchAllow) Then
            MatchAllow = 0                         ' .Permitted is not changed!
            If Left(T_DC.DCAllowedMatch, 1) = "*" Then ' ignore any error
                N_ChkRC = True
                .Permitted = Empty                 ' but now it is!
                MatchAllow = T_DC.DCerrNum
                .FoundBadErrorNr = 0
                GoTo FuncExit                      ' so, this will just finish Error Try
            Else
                N_ChkRC = Z_IsUnacceptable(.Permitted) ' is it bad or not?
            End If
        Else
            .Permitted = MatchAllow
            N_ChkRC = Z_IsUnacceptable(.Permitted) ' is it bad or not?
        End If
        
        Putlead = String(OffCal, b)
        If Catch Then
            If DebugMode Or DebugLogging Then
                Debug.Print Putlead _
                            & " *#'* " & .atKey _
                            & " has caused Error " & .FoundBadErrorNr _
                            & " '" & .atMessage & "'"
                If LenB(TryingThisCode) > 0 Then
                    Debug.Print Putlead & "Purpose:     " & TryingThisCode
                End If
                If LenB(WhatThisMeans) > 0 Then
                    Debug.Print Putlead & "Explanation: " & WhatThisMeans
                End If
                Debug.Print Putlead & "currently permittted: " & .Permitted
                Debug.Assert False
            End If
        End If
        
        If .Permitted = testOne Then        ' all Errors allowed, only once, err is returned
            .FoundBadErrorNr = 0
            .Permitted = Empty              ' error capure is now off again
        ElseIf .Permitted = testAll Then    ' all allowed, Permitted stays, err is returned
            .FoundBadErrorNr = 0
        ElseIf .Permitted = allowAll Then   ' all allowed, Permitted stays, err is cleared
            .FoundBadErrorNr = 0
        ElseIf .Permitted = allowNew Then   ' all Errors allowed, only once, err is cleared
            .FoundBadErrorNr = 0
            .Permitted = Empty              ' error capure is now off again
        Else                                ' leave .Explain set and ErrorTest=true for analysis
        End If
    End With                                ' E_Active
       
FuncExit:
    If IsMissing(MatchAllow) Then
        Call ErrReset(0)                           ' error Try block has ended, N_ChkRC NOT changing
    ElseIf MatchAllow = 0 And DoClear Then
        If T_DC.DCAllowedMatch = testAll Then         ' Keep Permitted ANYTHING
            Call ErrReset(4)                       ' error Try done, T_DC NOT changing
        Else
            Call ErrReset(0)                       ' error Try scope end, DCAllowedMatch := Empty
        End If
    ElseIf MatchAllow = 0 Then                     ' And NOT DoClear
        If T_DC.DCAllowedMatch = testAll Then         ' Keep Permitted ANYTHING
            Call ErrReset(4)                       ' error Try done, T_DC NOT changing
            ' ErrorCaught not changing
        Else
            Call ErrReset(0)                       ' error Try scope end, DCAllowedMatch := Empty
        End If
    ElseIf Not N_ChkRC And DoClear Then
        Call ErrReset(3)                           ' error Try block has ended and no app bugs found
        Call T_DC.N_ClearTermination               ' this ends the scope of Try
    End If
    
    Z§ErrSnoCatch = E_Active.ErrSnoCatch
    
ProcRet:
End Function                                       ' BugHelp.N_ChkRC

'---------------------------------------------------------------------------------------
' Method : Sub N_ClearAppErr
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub N_ClearAppErr()
Const zKey As String = "BugHelp.N_ClearAppErr"
    
    If MayChangeErr Then
        Err.Clear
    End If
    If E_AppErr Is Nothing Then
        DoVerify False, "Clear an App without it being started ???"
        Set E_AppErr = New cErr                    ' no need to init, just clear
    Else
        Call N_ErrClean(0)                         ' clear Err data in E_AppErr only
    End If
End Sub                                            ' BugHelp.N_ClearAppErr

'---------------------------------------------------------------------------------------
' Method : N_ConDscErr
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Set up a consistent atDsc / Err Pair. Uses no Stacks
'---------------------------------------------------------------------------------------
Sub N_ConDscErr(ClientDsc As cProcItem, ClientKey As String, CallType As String, Qmode As eQMode, ClientErr As cErr)

Const zKey As String = "BugHelp.N_ConDscErr"
Static zDsc As cProcItem

Dim msg As String
Dim ModuleName As String
Dim ModeNameN As String
Dim ModeLetterN As String

    ModeNameN = QModeNames(Qmode)
    ModeLetterN = UCase(Left(ModeNameN, 1))
       
    If LogZProcs And DebugLogging Then
        msg = zKey & " defines atDsc/Err for " & ClientKey _
              & " Qmode=" & ModeLetterN
        Call LogEvent(msg)
    End If
    
    DoVerify LenB(ClientKey) > 0, "ClientKey must not be empty string"
    If ClientDsc Is Nothing Then
        Set ClientDsc = New cProcItem
    End If
    With ClientDsc
        If Not .ErrActive Is Nothing Then          ' if previously chained, N_ConDscErr done before
            If .ErrActive.atKey <> .Key Then
                Stop                                ' how that????????
                .ErrActive.atKey = .Key
            End If
            If .CallMode < Qmode And .CallCounter <= 0 Then
                ClientDsc.CallMode = Qmode         ' allow mode upgrade z->y->x
                GoTo reEntry
            End If
            GoTo zExit
        End If
        .ProcIndex = 0                             ' more details added below iff CallType<>""
        .CallCounter = -1
        .Key = ClientKey
        .DbgId = RTail(ClientKey, ".", ModuleName)
        If LenB(ModuleName) > 0 Then
            .Module = ModuleName
        Else
            .Module = "Extern"
        End If
        
        ' Purpose: Construct a pair of linked ClientDsc/ClientErr
        ' Note   : in case of conflict, ClientDsc wins
        
        If ClientErr Is Nothing Then
            Set ClientErr = New cErr
implicit:
            Set ClientErr.atDsc = ClientDsc
            ClientErr.atKey = ClientDsc.Key
            Set ClientErr.atDsc = ClientDsc
            ClientErr.NrMT = "--- u" & ModeLetterN
            Set ClientDsc.ErrActive = ClientErr
        Else
            If ClientErr.atKey = vbNullString Then
                GoTo implicit
            End If
            DoVerify ClientErr.atDsc Is ClientDsc
            DoVerify ClientErr.atKey = ClientDsc.Key
        End If
            
        If LenB(CallType) > 0 Then
            If LenB(ClientDsc.CallType) > 0 Then
                DoVerify ClientDsc.CallType = CallType, "change in CallType ???"
            Else
                .CallType = CallType
            End If
            If .CallMode = Qmode Then
                .ModeLetter = ModeLetterN
            Else
                DoVerify .CallMode = eQnoDef, " ** analyze mode change"
                ClientDsc.CallMode = Qmode
            End If
reEntry:
            With .ErrActive
                If Qmode <= eQxMode Then           ' y or z Mode: do not autocheck
                    .atRecursionOK = True          ' recursion is checked by the proc using individual rules
                End If
                .atProcIndex = inv
                .atLastInDate = Date
                .atLastInSec = Timer
                .atThisEntrySec = Timer
                Set .at_Live = New Dictionary
                .EventBlock = E_Active.EventBlock
            End With                               ' .ErrActive
        End If
        Set ClientErr = .ErrActive
    End With                                       ' ClientDsc
    Call N_SetErrLvl(ClientDsc)

zExit:
End Sub                                            ' BugHelp.N_ConDscErr

'---------------------------------------------------------------------------------------
' Method : N_DebugStart
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Initialize debug start values
'---------------------------------------------------------------------------------------
Sub N_DebugStart(Optional doStop As Boolean = True)

    If StackDebug = 0 Then
        StackDebug = 8                             ' >=4 all, 3 inline, 2 verbose, 1 normal StackDebug
    End If
    StackDebugOverride = StackDebug
    If UseTestStart Then
        ExLiveCheckLog = True
        ExLiveDscGen = True
        TraceMode = True
        CallLogging = True
        LogZProcs = True                ' Due UseTestStart
        LogAppStack = True
        ShowFunctionValues = True
        StackDebugOverride = -8
    ElseIf StackDebug >= 8 Then
        TraceMode = True
        CallLogging = True
        LogZProcs = True                ' Due StackDebug >=8
        LogAppStack = True
        ShowFunctionValues = True
    ElseIf StackDebug = 7 Then
        CallLogging = True
        LogZProcs = False               ' Due StackDebug = 7
        ShowFunctionValues = ShowFunctionValues Or UseTestStart
        LogPerformance = False                     ' no performance data for 7 and up and 4 down
    ElseIf StackDebug = 6 Then
        CallLogging = False                        ' no CallLogging for 6 and up and 4 down
        LogZProcs = False               ' Due StackDebug = 6
        LogPerformance = False
    ElseIf StackDebug = 5 Then
        StackDebug = Abs(StackDebugOverride)
        ExLiveCheckLog = False
        ExLiveDscGen = False
        TraceMode = False
        CallLogging = False
        LogZProcs = False               ' Due StackDebug = 5
        LogAppStack = False
        ShowFunctionValues = False
        LogPerformance = False
    ElseIf StackDebug = 4 Then
        StackDebug = Abs(StackDebugOverride)
        ExLiveDscGen = True
        LogZProcs = True                ' Due StackDebug = 4
    ElseIf StackDebug = 3 Then
        StackDebug = Abs(StackDebugOverride)
        LogZProcs = True                ' Due StackDebug = 3
    ElseIf StackDebug = 2 Then
        StackDebug = Abs(StackDebugOverride)
        LogPerformance = True
    ElseIf StackDebug = 1 Then
        StackDebug = Abs(StackDebugOverride)
        CallLogging = True
    ElseIf StackDebug = -1 Then
        StackDebug = Abs(StackDebugOverride)
        FastMode = True                            ' true on -1 only
        ExLiveCheckLog = False                     ' true on UseTeststart and 5
        ExLiveDscGen = False                          ' true on UseTestStart and 4, 5
        TraceMode = False                          ' true on UseTestStart and 5, 8
        CallLogging = False                        ' true on UseTestStart and 1, 7, 8
        LogZProcs = False                          ' true on UseTestStart and 1, 3, 7, 8
        LogAppStack = False                        ' true on UseTestStart and 8
        ShowFunctionValues = False                 ' true on UseTestStart and 1, 8, kept on 7
        LogPerformance = False                     ' true on UseTestStart and 2
        UseTestStart = False                       ' false on -1
    End If
    If doStop And Not DidStop Then
        DidStop = True
        doStop = False
    End If
    Call BugSet

End Sub                                            ' BugHelp.N_DebugStart

'---------------------------------------------------------------------------------------
' Method : Sub N_ErrClean
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ErrClean(ForceLevel As Long)
Const zKey As String = "BugHelp.N_ErrClean"

Dim ExplainS As String

    If Not E_AppErr.errNumber = 0 And (DebugMode Or DebugLogging) Then
        ExplainS = "Err Nr.=" & Err.Number
        If E_AppErr.FoundBadErrorNr <> 0 Then
            ExplainS = ExplainS & " (bad)"
        End If
        ExplainS = "Reset Err, Lvl=" & ForceLevel & ", " & ExplainS
    End If
    
    If MayChangeErr Then
        If LenB(ExplainS) > 0 Then
            Debug.Print ExplainS
        End If
        Err.Clear
    Else
        If LenB(ExplainS) > 0 Then
            Debug.Print ExplainS & ", Err not cleared=" & Err.Number
        End If
    End If
    
    With E_AppErr
        If .errNumber = 0 Then
            If ForceLevel < 1 Then
                GoTo simple
            End If
        End If
        If ForceLevel = inv Then                   ' do not log at all and like ForceLevel = 0
simple:
            .errNumber = 0
            .Description = vbNullString
            .atFuncResult = vbNullString
            .FoundBadErrorNr = 0
            .Permitted = Empty
            .Reasoning = vbNullString
            GoTo ProcRet
        End If
        
        ' all cases 0-2
        .errNumber = 0
        .Description = vbNullString
        .atFuncResult = vbNullString
        .FoundBadErrorNr = 0
        .Permitted = Empty
        .Reasoning = vbNullString
        
        Select Case ForceLevel
            Case 0
                ' no further N_ErrCleaning in E_AppErr. (Keeps the ErrTry-State tErr way)
            Case 1
                .Explanations = vbNullString
                .Reasoning = vbNullString
            Case 2
                DoVerify False, " whatfor and whatelse-just use: New"
                .atProcIndex = inv
                .atCallState = eCUndef
                .atMessage = vbNullString
                .atKey = vbNullString
                .atShowStack = vbNullString
                .DebugState = False
                .Explanations = vbNullString
                .Reasoning = vbNullString
                .atRecursionOK = False
                .atPrevEntrySec = 0
                .atThisEntrySec = 0
                .atTraceStackPos = 0
                Set .atCalledBy = Nothing
                Set .atDsc = Nothing
                Set .atErrPrev = Nothing
                .atCallDepth = 0
                Set .atErrPrev = Nothing
            Case Else
                DoVerify False, " Invalid ForceLevel"
        End Select
    End With                                       ' E_AppErr
ProcRet:
End Sub                                            ' BugHelp.N_ErrClean

'---------------------------------------------------------------------------------------
' Method : Sub N_ErrClear
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ErrClear(Optional ForceLevel As Long)
Const zKey As String = "BugHelp.N_ErrClear"

    Call N_Suppress(Push, zKey)
    
    If T_DC Is Nothing Then
        rsp = MsgBox("new termination object?", vbOKCancel)
        If rsp <> vbCancel Then
            Call ShowDbgStatus
            If IgnoreUnhandledError Then
                GoTo FuncExit
            End If
            T_DC.N_ClearTermination
        End If
    End If
    Call N_ErrClean(ForceLevel)

FuncExit:
    Call N_Suppress(Pop, zKey)

End Sub                                            ' BugHelp.N_ErrClear

'---------------------------------------------------------------------------------------
' Method : Sub N_ErrInterfacePrint
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ErrInterfacePrint(sExplain As String, Optional ErrClient As cErr)

Const zKey As String = "BugHelp.N_ErrInterfacePrint"
    
Dim smi1 As Long
Dim ClientDsc As cProcItem
    
    Call N_Suppress(Push, zKey)
    If IsMissing(ErrClient) Then
        Set ErrClient = E_AppErr
    End If
    If ErrClient Is Nothing Then
        Debug.Print "can't N_ErrInterfacePrint, no Data!"
        GoTo FuncExit
    End If
    If ClientDsc.ProcIndex > inv Then
        Set ClientDsc = D_ErrInterface.Items(ErrClient.atProcIndex)
    Else
        Set ClientDsc = D_ErrInterface.Items(2 - ErrClient.atProcIndex)
    End If
    
    With ErrClient
        If .atCallState = eCActive Then
            If ClientDsc.CallMode = eQArMode Then
                smi1 = .atCallDepth - 1
                .atMessage = Right("    " & CStr(smi1), 5) & String(smi1, ">") _
                    & b & sExplain _
                    & vbCrLf & Right("    " & CStr(smi1), 5) & String(smi1, b) _
                    & " called by " _
                    & Replace(.atCalledBy.atKey, dModuleWithP, vbNullString) _
                    & " (using Stack, Call Depth=" & smi1 & ")"
            Else                                   ' .at-values are in place for this type of call.
                sExplain = "ProcCall/Exit, No ErrHandler: " & .atKey
                .atMessage = String(.atCallDepth, ">") _
                    & b & sExplain & b & vbCrLf & String(.atCallDepth, b) _
                    & " Caller: " & Replace(.atCalledBy.atKey, dModuleWithP, vbNullString)
                End If
            ElseIf .atCallState = eCExited Then
                sExplain = Replace("Exit to: " & .atCalledBy.atKey _
                    & " from: " & .atKey, dModuleWithP, vbNullString)
                .atMessage = String(.atCallDepth, "<") _
                    & b & sExplain & vbTab
            ElseIf .atCallState = eCpaused Then
                sExplain = Replace("Paused: " & .atKey _
                    & " from: " & .atCalledBy.atKey, dModuleWithP, vbNullString)
                .atMessage = String(.atCallDepth, "<") _
                    & b & sExplain & vbTab
            Else
                sExplain = "never called: " & .atKey
                .atMessage = sExplain
            End If

            Debug.Print .atMessage
    
        End With                                   ' ErrClient

FuncExit:
        Call N_Suppress(Pop, zKey)
  
End Sub                                        ' BugHelp.N_ErrInterfacePrint

'---------------------------------------------------------------------------------------
' Method : N_ErrStackLines
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Create a string to output for File or Immediate window
'---------------------------------------------------------------------------------------
Sub N_ErrStackLines(LiveErrors As String)
Const zKey As String = "BugHelp.N_ErrStackLines"

Dim MN As String
Dim PN As String
Dim extra As String

    LiveErrors = Now() _
        & " - " & CStr(ErrEx.Number) & " - " & CStr(ErrEx.Description)
    LiveErrors = LiveErrors & ", Saved=" & ErrEx.SourceProjectIsSaved _
                 & ", VBEVersion=" & ErrEx.VBEVersion
    ' separate the call stack to single lines in the log
    
    With ErrEx.LiveCallstack
        Do
            If .ModuleName = Mid(dModuleWithP, 2) Then ' Omitting default Module for readability
                MN = vbNullString
            Else
                MN = .ModuleName & "."
            End If
            If 1 = 0 Then                          ' Omitting ProjectName in Outlook
                PN = .ProjectName & "."
            Else
                PN = vbNullString
            End If
            If .ModuleName = dModule Then
                extra = String(4, b)
                MN = vbNullString
            Else
                extra = vbNullString
            End If
            LiveErrors = LiveErrors & vbCrLf _
                         & extra & "       --> " & PN _
                         & MN _
                         & .ProcedureName & ", " _
                         & "#" & .LineNumber & ", " _
                         & .LineCode
        Loop While .NextLevel
    End With                                       ' ErrEx.LiveCallstack

zExit:

End Sub                                            ' BugHelp.N_ErrStackLines

'---------------------------------------------------------------------------------------
' Method : N_GenCallData
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Set up the values for this Call
'---------------------------------------------------------------------------------------
Sub N_GenCallData(aDsc As cProcItem, aErr As cErr)
Dim Consumed As String
Dim TEntry As cTraceEntry

    aErr.atFastMode = FastMode                     ' set this for correct Exit Mode
    
    If LogZProcs Then
        If LogPerformance Then
            Call Z_UsedThisCall(aErr, aErr.atThisEntrySec)
            Consumed = " Consumed " & ElapsedTime & " sec "
        Else
            Consumed = vbNullString
        End If
    End If
    
    If Not dontIncrementCallDepth Then              ' (not predefining)
        aErr.atCallDepth = E_Active.atCallDepth + 1
    End If
    
    If TraceMode Or dontIncrementCallDepth Then
        Set TEntry = New cTraceEntry
        Set TEntry.TErr = aErr
        Call N_TraceEntry(TEntry)                  ' generate trace Entry for ExternCaller
        Set TEntry = Nothing
    End If
    
    If Not BugTimer Is Nothing And Not ShutUpMode Then
        If BugTimer.BugStateReCheck Then
            If Timer - BugTimer.BugStateLast _
               > 2 * BugTimer.BugStateTicks Then
                Call BugTimerEvent("Call")
            End If
        End If
    End If
    
    If Not dontIncrementCallDepth Then              ' Predefining does not inc:
        With aErr
            ' for each call we increment .atRecursionLvl until we DoExit
            If .atRecursionLvl > 0 Then
                Set aErr = aErr.Clone           ' new instance due to recursion
                aErr.atRecursionLvl = .atRecursionLvl + 1
            Else
                .atRecursionLvl = 1                 ' no new instance (no recursion)
            End If
            If aDsc.MaxRecursions < E_Active.atRecursionLvl Then
                aDsc.MaxRecursions = E_Active.atRecursionLvl
            End If
        End With                                    ' aErr
        
        With aErr
            Set .atCalledBy = E_Active
            .atLastInDate = Date
            .atThisEntrySec = Timer
            .atPrevEntrySec = .atThisEntrySec
            .atLastInSec = .atThisEntrySec
            If CallLogging Then
                Call N_ShowProgress(CallNr, aDsc, "+" & aDsc.ModeLetter, _
                                "zD=" & .atCallDepth - 1 & "»" & .atCallDepth, _
                                Consumed, ErrClient:=aErr)
            End If
            Set P_Active = aDsc                     ' replacing Caller with called Program
            Set P_Active.ErrActive = aErr
            Set E_Active = aErr
            E_Active.atCallState = eCActive
            P_Active.CallCounter = P_Active.CallCounter + 1
        End With                                    ' aErr
    End If                                          ' branch for not predefining (dontIncrementCallDepth)
ProcRet:
End Sub                                             ' BugHelp.N_GenCallData

'---------------------------------------------------------------------------------------
' Method : N_GetLive
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Use Live Stack info to get relevant LiveStack elements Needed
'---------------------------------------------------------------------------------------
Sub N_GetLive(LiveStack As Collection, Optional Needing As Long = 2, Optional Filtered As Boolean, Optional Logging As Boolean = True)

Const zKey As String = "BugHelp.N_GetLive"

Dim limit As Long
Dim PN As String
Dim StackIndex As Long
Dim msg As String
Dim ClientDsc As cProcItem
Dim gotThem As Boolean
Dim ModuleName As String

    Set LiveStack = New Collection
    If Needing = 0 Then
        limit = ErrLifeKept                        ' we never deliver more results than that
    Else
        limit = Needing
    End If
    
    msg = vbNullString
    StackIndex = 1                                 ' start counting real position in LiveStack
    With ErrEx.LiveCallstack
        Do
            Set ClientDsc = Nothing
            
            PN = .ProcedureName
            ModuleName = .ModuleName
            
            Call Z_GetProcDsc(ModuleName & "." & PN, ClientDsc, msg)
            If ClientDsc Is Nothing Then           ' msg given from Z_GetProcDsc
                GoTo noDsc
            End If
            
            If ClientDsc.CallMode = eQnoDef Then   ' it's a dummy
                If Filtered Then
                    GoTo DidStep                   ' do not Show in LiveStack !
                End If
            End If

            LiveStack.Add ClientDsc
            If LiveStack.Count >= limit Then
                gotThem = True
                GoTo FuncExit                      ' we only that many (max)
            End If
DidStep:
            StackIndex = StackIndex + 1
noDsc:
        Loop While .NextLevel
        
    End With                                       ' ErrEx.LiveCallstack
    
    If Not gotThem Then
        If Logging Then
            If CallLogging Then
                Debug.Print msg
                Debug.Print String(15, b) & "limit reached= " & limit _
                                          & ". found only " & LiveStack.Count _
                                          & " relevant entries. Full Stack: ";
                Call ShowLiveStack(doPrint:=True, _
                                   tSubFilter:=False, getNewStack:=True, Full:=Filtered)
                Debug.Assert False
            End If
        End If
    End If
    
FuncExit:
    Set ClientDsc = Nothing

zExit:

End Sub                                            ' BugHelp.N_GetLive

'---------------------------------------------------------------------------------------
' Method : Sub N_InterpretTestVar
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_InterpretTestVar()
Const zKey As String = "BugHelp.N_InterpretTestVar"

Dim i As Long

    If StackDebugOverride > 0 Then
        StackDebug = StackDebugOverride            ' until this is saved
        ' Override: TraceMode = InStr(1, Testvar, "TraceMode", vbTextCompare) > 0
        ' Override: LogPerformance = InStr(1, Testvar, "LogPerformance", vbTextCompare) > 0
        GoTo FromTestVar
    Else
        If InStr(1, Testvar, "StackDebug", vbTextCompare) > 0 Then
            i = InStr(1, Testvar, "StackDebug=", vbTextCompare)
            If i > 0 Then
                StackDebug = Mid(Testvar, i + Len("StackDebug="), 2)
            End If
        End If
        TraceMode = InStr(1, Testvar, "TraceMode", vbTextCompare) > 0
        LogPerformance = InStr(1, Testvar, "LogPerformance", vbTextCompare) > 0
FromTestVar:
        DebugMode = InStr(1, Testvar, "DebugMode", vbTextCompare) > 0
        DebugLogging = InStr(Testvar, "LOG") > 0
        LogAllErrors = InStr(Testvar, "ERR") > 0
        ShowFunctionValues = InStr(Testvar, "ShowFunctionValues") > 0
    End If

End Sub                                            ' BugHelp.N_InterpretTestVar

'---------------------------------------------------------------------------------------
' Method : N_LogEntry
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Print Application-Relevant Call
'---------------------------------------------------------------------------------------
Sub N_LogEntry(ClientDsc As cProcItem, ObjStr As String, moreEE As String)
Const zKey As String = "BugHelp.N_LogEntry"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True

Dim noPrint As Boolean
Dim Lvl As Long
Dim Caller As String
Dim ObjInfo As String
Dim addInfo As String
    
    Call N_ShowHeader("BugHelp Log " & TimerNow)
    
    If StackDebug <= 9 Then
        If StackDebug > 8 Then
            If ClientDsc.CallMode <= eQxMode Then
                GoTo testHidden
            End If
        End If
        If ClientDsc.CallMode = eQzMode Then       ' covers  .., Z_.., and O_Goodies, Classes
testHidden:
            If StackDebug > 4 Then
                GoTo FuncExit
            End If
        End If
    End If

    If Left(moreEE, 1) = "!" Then                  ' do not print
        noPrint = True
        addInfo = Mid(moreEE, 2)
    Else
        addInfo = moreEE
    End If
    
    Caller = Replace(ClientDsc.Key, dModuleWithP, vbNullString)
    ObjInfo = ObjStr

    Lvl = S_AppIndex + 1
           
GenOut:
     
    Call Z_Protocol(ClientDsc, CallNr, Caller, "==>", ObjInfo, addInfo)

FuncExit:
    Recursive = False
    
zExit:

ProcRet:
End Sub                                            ' BugHelp.N_LogEntry

'---------------------------------------------------------------------------------------
' Method : N_LogErrEx
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Log ErrEx Data to File
'---------------------------------------------------------------------------------------
Sub N_LogErrEx()
Const zKey As String = "BugHelp.N_LogErrEx"
Dim zErr As cErr

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Not ErrExActive Then
        GoTo ProcRet
    End If
    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True

Dim LogLines As Variant
Dim sL As Variant
Dim logLine As String
Dim InitialSkip As Boolean
    
    Call N_ErrStackLines(logLine)
    LogLines = split(logLine, vbCrLf)
    
    For Each sL In LogLines
        If Not InitialSkip Then                    ' skip lines at start of dump output
            If InStr(sL, "BugHelp.Z") > 0 Then
                GoTo SkipLine
            Else
                InitialSkip = True                 ' stop when useful info reached
            End If
        End If
        
        logLine = sL
        Debug.Print logLine
SkipLine:
    Next sL

    Recursive = False

ProcRet:
End Sub                                            ' BugHelp.N_LogErrEx

'---------------------------------------------------------------------------------------
' Method : N_OnError
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Fehler Abfangen mit ErrEx > Global Error Handler <
'---------------------------------------------------------------------------------------
Sub N_OnError()
Const zKey As String = "BugHelp.N_OnError"

Dim msg As String
Dim ForceDialog As Boolean
Dim Prefix As String

'------------------- gated Entry -------------------------------------------------------
    BugState = ErrEx.State
    BugStateAsStr = ErrEx.StateAsStr
    BugDlgRsp = BugState                           ' no Dialog yet, default as same
    BugDlgRspAsStr = BugStateAsStr
    ' all error data now in E_Active
    msg = ErrEx.Description
    If LenB(msg) > 0 Then
        T_DC.DCAppl = S_AppKey
        T_DC.DCerrMsg = msg
        T_DC.DCerrNum = ErrEx.Number
        ErrorCaught = T_DC.DCerrNum
        ErrEx.Number = 0
    End If
    
    msg = ErrEx.SourceModule & "." & ErrEx.SourceProcedure
    msg = msg & vbCrLf & "Line " & ErrEx.SourceLineNumber & ": " & ErrEx.SourceLineCode
    If isEmpty(T_DC.DCAllowedMatch) Then
        msg = msg & ", UnExpected"
    Else
        msg = msg & ", Accepting=" & T_DC.DCAllowedMatch _
              & " Error Number=" & T_DC.DCerrNum
        msg = msg & "=&H" & Hex8(T_DC.DCerrNum) _
        & vbCrLf & "   ErrorMessage: " & T_DC.DCerrMsg
        msg = msg & vbCrLf & "   ErrStatus=" & BugState _
              & "(" & BugStateAsStr & ")"
        If Left(T_DC.DCAllowedMatch, 1) = "*" _
        Or IsNumeric(T_DC.DCAllowedMatch) Then
            Prefix = String(3, b)
        Else
            Prefix = vbCrLf & String(3, "!! ")
        End If
    End If
    T_DC.DCerrSource = msg
        
    If DebugMode Or StackDebug > 7 Then
        msg = Prefix & "N_OnError called, " & msg
        
        If StackDebug > 8 _
        Or BugState > 3 Then                    ' not handled at caller's level
            Debug.Print String(20, "-") _
                        & vbCrLf _
                        & msg & "   must handle by Catch"
        ElseIf E_Active.atCallDepth > 0 Then
            If T_DC.DCAllowedMatch <> T_DC.DCerrNum _
            And Left(T_DC.DCAllowedMatch, 1) <> "*" Then
                msg = String(80, b) & vbCrLf & msg
                Call LogEvent(msg, eLSome)         ' not accepting
            End If
        End If
        msg = vbNullString                         ' do not repeat report later
    End If
    
    Call N_CaptureNewErr(True)                     ' handle error as defined
    If AppStartComplete Then
        Call N_PublishBugState                     ' also "locks" Err until exit
    End If
    
    ForceDialog = DebugMode Or DebugLogging
    
    With E_Active
        Select Case BugState
        
            Case 0, OnErrorGoto0, OnErrorCatch, OnErrorCatchAll ' 0, 1, 10, 11
                If Left(.Permitted, 1) = "*" Then
                    If .Permitted = "*" Then       ' all are allowed, err is returned
                        T_DC.DCAllowedMatch = Empty ' error capure is now off again
                    ElseIf .Permitted = testAll Then  ' all allowed, Permitted stays, err is returned
                    ElseIf .Permitted = allowAll Then ' all allowed, Permitted stays, err is cleared
                        Call ErrReset(4)           ' cleans up error trace!!
                    ElseIf .Permitted = allowNew Then
                        Call ErrReset(0)
                    Else
                        ForceDialog = ForceDialog Or (StackDebug > 8)
                    End If
                    .FoundBadErrorNr = 0
                    GoTo itsOK
                Else
                    If isEmpty(.Permitted) Then
                        ForceDialog = True
                    End If
itsOK:
                    .Permitted = T_DC.DCAllowedMatch
                    If Not ForceDialog Then
                        ErrEx.State = OnErrorResumeNext
                        GoTo zExit
                    End If
                End If
                
                ' ---------------------------------------------------------------
                ' Unhandled errors
                ' ---------------------------------------------------------------
                ForceDialog = True
                
            Case OnErrorResumeNext                 ' 2
                ' ---------------------------------------------------------------
                ' Ignore errors when On Error Resume Next is set
                ' ---------------------------------------------------------------
                
            Case OnErrorGotoLabel                  ' 3
                ' ---------------------------------------------------------------
                ' Ignore locally handled errors, so go where instructed
                ' ---------------------------------------------------------------
                
            Case CalledByLocalHandler              ' 6
                ' ---------------------------------------------------------------
                ' ErrEx.CallGlobalBugHelp was called
                
                ' This is a special case for when local error handling was in use
                ' but the local error handler has not dealt with the error and
                ' so has passed it on to the global error handler
                ' ---------------------------------------------------------------
                
            Case OnErrorPropagate                  ' 8
                ' ---------------------------------------------------------------
                ' Propagation would not cause ProcExit to handle stack!
                ' So Propagation must not be used: Only locally handled errors
                ' (otherwise handled in a previous routine in the call stack)
                ' ---------------------------------------------------------------
                If LenB(BugWillPropagateTo) > 0 Then
                    Stop                           ' error during Propagation
                    BugWillPropagateTo = vbNullString
                End If
                
                With ErrEx.Callstack
                    .FirstLevel
                    Do
                        If .HasActiveErrorHandler = True Then
                            BugWillPropagateTo = .ProjectName _
                                                 & "." & .ModuleName & "." & .ProcedureName
                            Exit Do
                        End If
                    Loop While .NextLevel
                End With                           ' ErrEx.Callstack
                
            Case OnErrorInsideFinally              ' 14 = &HE
            
                ' An error occurred inside the ErrEx.Finally block (typically for cleanup code).
                ' We will use OnErrorResumeNext to skip over these
                ErrEx.State = OnErrorResumeNext
                
            Case Else                              ' 4, 5, 7, 9, 12 = &HC, 13 = &HD, 15 = &HF, 17 = &H11, 18 = &H12
                Debug.Print "ErrEx.State " & BugState _
                            & "=" & BugStateAsStr _
                            & " is not handled in N_OnError"
                ForceDialog = True
                Debug.Assert False
        End Select                                 ' ErrEx.State
    End With                                       ' E_Active
    
    If ForceDialog Then
        If LenB(msg) > 0 Then
            Call LogEvent(msg, eLSome)
        End If
        Call N_LogErrEx
        
        BugDlgRsp = ErrEx.ShowErrorDialog
        
        If BugDlgRsp = OnErrorResumeNext Then
            Debug.Print "Request to clear Error: " & BugDlgRsp _
                        & " Debugging(" & ErrEx.StateAsStr & ")"
            Err.Clear
        ElseIf BugDlgRsp = OnErrorDebug Then
            Call ErrReset(4)                       ' clear this error logically (but recognizable)
            Debug.Print "T_DC Status not cleared: " & BugDlgRsp _
                        & " Debugging(" & ErrEx.StateAsStr & ")" & b;
            Debug.Print "retry the erroneous statement: " & ErrEx.StateAsStr
            GoTo zExit
        ElseIf BugDlgRsp = OnErrorEnd Then
            Debug.Print "ShowErrorDialog Choice is: " & BugDlgRsp & b & ErrEx.StateAsStr
            If Left(E_Active.Permitted, 1) <> "*" Then
                Stop                               ' continue possible: Press Ctrl-Shift-F8
            End If
            ' If the close button is pressed on the error dialog, we don't
            ' want to end abruptly (OnErrorEnd), but instead want the program
            ' flow to continue in our local error handler:
            If BugState = OnErrorEnd Then
                ErrEx.State = CalledByLocalHandler
                Debug.Print "Another End Request from ShowErrorDialog: " _
                            & BugDlgRsp & b & ErrEx.StateAsStr
            Else
                ErrEx.State = OnErrorResumeNext
                Debug.Print "Request to resume from ShowErrorDialog:" _
                            ; " & ErrEx.StateAsStr"
                Call ErrReset
            End If
                
            GoTo zExit
        Else
SetState:
            ErrEx.State = BugDlgRsp
            BugDlgRspAsStr = ErrEx.StateAsStr
            Debug.Print "ShowErrorDialog Choice is: " _
                        & BugDlgRsp & " (" & BugDlgRspAsStr & ")"
        End If
    End If
    
FuncExit:
    IgnoreUnhandledError = False                   ' now we care about Unhandled again
    MayChangeErr = True

zExit:
    BugStateAsStr = ErrEx.StateAsStr

End Sub                                            ' BugHelp.N_OnError

' ----------------------------------------------------------------
' Procedure Name: N_PopLog
' Purpose: Pop stack into LogAppStack
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Rolf G. Bercht
' Date: 29.06.2017
' ----------------------------------------------------------------
Sub N_PopLog()
'--- Proc MAY ONLY CALL Z_Type PROCS                            ' Simple proc
Const zKey As String = "BugHelp.N_PopLog"
    
Dim oldCount As Long

    oldCount = C_PushLogStack.Count
    DoVerify oldCount > 0, "** nothing on stack to pop"
    LogAppStack = C_PushLogStack.Item(oldCount)
    C_PushLogStack.Remove oldCount

End Sub                                            ' BugHelp.N_PopLog

' ----------------------------------------------------------------
' Procedure Name: N_PopSimple
' Purpose       : Pop stack state variant from top of the stack C_AllPurposeStack
' Author        : Rolf G. Bercht
' Date   : 20211108@11_47
' ----------------------------------------------------------------
Sub N_PopSimple(aVarO As Variant, Optional VariantObject As Boolean, Optional useStack As Collection)
Const zKey As String = "BugHelp.N_PopSimple"

Dim oldCount As Long

    If useStack Is Nothing Then
        Set useStack = C_AllPurposeStack
    End If
    oldCount = useStack.Count
    
    If VariantObject Then
        Set aVarO = useStack.Item(oldCount)
    Else
        aVarO = useStack.Item(oldCount)
    End If
    useStack.Remove oldCount

zExit:

End Sub                                            ' BugHelp.N_PopSimple

'---------------------------------------------------------------------------------------
' Method : N_Prepare
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Prepare indispensable variables used during initializations
'---------------------------------------------------------------------------------------
Sub N_Prepare(Optional WannaKnow As Boolean)

Dim nextField As String
Dim i As Long
    
    lHeadLine = 156
    
    testNonValueProperties = "Nothing Empty"
    ScalarTypes = split(ScalarTypeNames)
    ScalarTypeV = Array(vbInteger, vbLong, vbSingle, vbDouble, _
                        vbDate, vbString, vbBoolean, 20&)
    Set dSType = New Dictionary
    For i = 0 To UBound(ScalarTypeV)
        dSType.Add InStr(ScalarTypeNames & b, _
                         ScalarTypes(i) & b), ScalarTypeV(i)
    Next i
    
    dModuleWithP = dModule & "."
    QModeNames = split(QModeString)
    For i = 0 To UBound(QModeNames)
        QModeNames(i) = LString(QModeNames(i), 7)  ' 6+1 b
    Next i
    OkValueNames = split(OkValueString)
    
    CStateNames = split(CStateString)
    For i = 0 To UBound(CStateNames)
        CStateNames(i) = LString(CStateNames(i), 8)
    Next i
    
    LogLevelNames = split(LogLevelString)
    AccountTypeNames = split(AccountTypeString)
    PushTypes = split(PushTypeString)
    ExStackProcNames = split(ExStackProcString)
    ExModeNames = split(ExModeNamesString)

    Call N_SetErrExHdl
    
    ModelLine(1) = "Call# Pr# MTS Lvl "
    OffPrN = InStr(ModelLine(1), "Pr#")
    OffMTS = InStr(ModelLine(1), "MTS")
    OffLvl = InStr(ModelLine(1), "Lvl")
    OffCal = Len(ModelLine(1))                     ' offset of Head Line Call-Column
    
    ModelLine(3) = ModelLine(1) & LString("Caller", lKeyM - 15)
    ModelLine(1) = ModelLine(1) & LString("Caller", lKeyM)
    OffObj = Len(ModelLine(1))
    
    ModelLine(2) = Left(LString("Call#", 5) _
                        & b & LString("Caller", OffObj - 6 - lCallInfo) _
                        & b & LString("CallerInfo", lCallInfo) _
                        & "sD=  -- Code " _
                        & String(lDbgM + lKeyM, "-"), lHeadLine)
    
    
    ModelLine(1) = ModelLine(1) & Left("----- Object " & String(lDbgM, "-"), lDbgM + 7) & b
    OffTim = Len(ModelLine(1))
    
    nextField = String(4, "-") & " Time " & String(4, "-") & b
    ModelLine(1) = ModelLine(1) & nextField
    OffAdI = Len(ModelLine(1))
        
    ModelLine(3) = ModelLine(3) & nextField
    ModelLine(3) = ModelLine(3) & "Lne -- Code " & String(lDbgM, "-")
    
    nextField = LString("--- additional Information " & String(lKeyM, "-"), lKeyM)
    ModelLine(1) = LString(ModelLine(1) & nextField, lHeadLine)
    
    ModelLine(3) = LString(ModelLine(3) & String(9, "-") & nextField, lHeadLine)
    Call N_ShowHeader("BugHelp Log " & TimerNow)
    
    Call N_PreDefine
    
    If WannaKnow Then
        Debug.Print "OffPrN=" & OffPrN, "OffMTS=" & OffMTS, "OffLvl=" & OffLvl, _
                    "OffCal=" & OffCal, "OffObj=" & OffObj, _
                    "OffTim=" & OffTim, "OffAdI=" & OffAdI, _
                    "lHeadLine=" & lHeadLine
        Debug.Print ModelLine(1)
        Debug.Print ModelLine(2)
        Debug.Print ModelLine(3)
    End If
    
End Sub                                            ' BugHelp.N_Prepare

'---------------------------------------------------------------------------------------
' Method : N_PublishBugState
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Interprets Error information from E_Active and then
'          Decide what to do (using Z_IsUnacceptable)
'             sets this into T_DC and E_AppErr
'             if acceptable error, proceed,
'             else, or if debugmode, Show Error Status form
' Note   : Data from N_OnError->N_CaptureNewErr, which puts Err->E_Active
'---------------------------------------------------------------------------------------
Sub N_PublishBugState()
Const zKey As String = "BugHelp.N_PublishBugState"
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    
    If E_Active Is Nothing Then
        Debug.Print "* Error before error-setup complete"
        Debug.Print RetLead & zKey & " -> " & S_DbgId
        Call TerminateRun
    End If
    
    With E_Active
        If IgnoreUnhandledError Then
            If DebugMode Then
                Debug.Print "N_PublishBugState called when IgnoreUnhandledError = True"
                DoVerify False
            End If
            GoTo FuncExit
        End If
        If Err.Number <> 0 Then
            Debug.Print "* unhandled error between error handling calls. Manual intervention?"
            DoVerify False
            GoTo FuncExit
            Resume Next                            ' on manual intervention only: analyze problem area
        End If
        Recursive = True
    
        
        E_AppErr.errNumber = .errNumber            ' to inform outside world (App lvl)
        E_AppErr.Description = .Description
        E_AppErr.Source = .Source
        
        IgnoreUnhandledError = True                ' anything called until Exit will not check Unhandled
                
    End With                                       ' E_Active

FuncExit:
    Recursive = False

ProcRet:
End Sub                                            ' BugHelp.N_PublishBugState

' ----------------------------------------------------------------
' Procedure Name: N_PushLog
' Purpose: Push state of LogAppStack
' Author: Rolf G. Bercht
' Date: 29.06.2017
' ----------------------------------------------------------------
Sub N_PushLog(NewState As Boolean)
Const zKey As String = "BugHelp.N_PushLog"
    
    C_PushLogStack.Add LogAppStack
    LogAppStack = NewState

End Sub                                            ' BugHelp.N_PushLog

' ----------------------------------------------------------------
' Procedure Name: N_PushSimple
' Purpose: Push state variant on top of the stack C_AllPurposeStack
' Author: Rolf G. Bercht
' Date: 29.06.2017
' ----------------------------------------------------------------
Sub N_PushSimple(OldState As Variant, NewState As Variant, Optional VariantObject As Boolean, Optional useStack As Collection)

Const zKey As String = "BugHelp.N_PushSimple"
    
    If useStack Is Nothing Then
        Set useStack = C_AllPurposeStack
    End If
    If (DebugMode Or aDebugState) And useStack.Count > 0 Then
        If VariantObject Then
            DoVerify useStack(useStack.Count) Is OldState
        Else
            DoVerify useStack(useStack.Count) = OldState
        End If
    End If
    
    useStack.Add OldState
    If VariantObject Then
        Set OldState = NewState
    Else
        OldState = NewState
    End If

zExit:

End Sub                                            ' BugHelp.N_PushSimple

'---------------------------------------------------------------------------------------
' Method : Sub N_SetErrLvl
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_SetErrLvl(ClientDsc As cProcItem)
Const zKey As String = "BugHelp.N_SetErrLvl"
    
    With ClientDsc
        Select Case .CallMode
            Case eQnoDef                           '  = 0
                .ErrLevel = eLdebug                ' = 0
            Case eQzMode                           '  = 1
                .ErrLevel = eLdebug                ' = 0
            Case eQyMode                           '  = 2
                If .ErrLevel = 0 Then
                    .ErrLevel = eLmin              ' = 3 + StackDebug > 9 only
                Else
                    .ErrLevel = eLmin
                End If
            Case eQxMode                           '  = 3
                .ErrLevel = eLmin                  ' = 3 + StackDebug > 9 only
            Case eQrMode                           '  = 4
                .ErrLevel = eLmin                  ' = 3
            Case eQuMode                           '  = 5
                .ErrLevel = eLmin                  ' = 3
            Case eQEPMode                          '  = 6
                .ErrLevel = eLall                  ' = 1
            Case eQAsMode                          '  = 7
                .ErrLevel = eLmin                  ' = 3
            Case eQArMode                          '  = 8
                .ErrLevel = eLSome                 ' = 2
            Case Else
                DoVerify False, " this is an incorrect Qmode"
        End Select                                 ' ClientDsc.CallMode
    End With                                       ' ClientDsc

End Sub                                            ' BugHelp.N_SetErrLvl

'---------------------------------------------------------------------------------------
' Method : Sub N_ShowErrInstance
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ShowErrInstance(ErrClient As cErr, IndexNr As Long)
Dim cPos As Long
Dim CommentPart As String
Dim ActionPart As String
Dim ExplainS As String
    
    With ErrClient
        
        cPos = InStr(.atShowStack, "#")
        If cPos > 0 Then
            CommentPart = Trim(Mid(.atShowStack, cPos + 1))
            ActionPart = Trim(Left(.atShowStack, cPos - 1))
        Else
            If LenB(.atShowStack) > 0 Then
                CommentPart = Trim(.atShowStack)
            Else
                CommentPart = Trim(.Reasoning)
            End If
        End If
        
        ExplainS = CommentPart
            
        Call N_ShowProcDsc(ErrClient.atDsc, IndexNr, WithInstances:=False, _
                           ExplainS:=ExplainS, ErrClient:=ErrClient)
    End With                                       ' ErrClient

End Sub                                            ' BugHelp.N_ShowErrInstance

'---------------------------------------------------------------------------------------
' Method : Sub N_ShowHeader
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ShowHeader(ListName As String, Optional force As Boolean, Optional ModelType As Long = 1, Optional StartOrEnd As Boolean)
Static lListName As Long
Static midOff As Long
Static modelNr As Long
Dim EndLine As String

    If force Or lListName = 0 Then
        lHeadLine = 0
        MinusLine = vbNullString
        HeadlineName = vbNullString
        lListName = Len(ListName)
        midOff = Round((lListName + 1) / 2)
        If modelNr = 0 Then
            modelNr = ModelType
        End If
        lHeadLine = Len(ModelLine(modelNr))
        lMinus = (lHeadLine - midOff) / 2 - 1
    End If
    
    If StartOrEnd Then                             ' StartOrEnd=True  ==> (=End) Force
        EndLine = Left(MinusLine, lMinus + lListName * 2 + 1) _
        & " -- End " & String(lMinus, "-")
        Call LogEvent(Left(EndLine, lHeadLine), eLall)
        HeadlineName = vbNullString                ' force a new value for Headline next call
    ElseIf HeadlineName <> ListName Then
        If lHeadLine > midOff Then
            lListName = Len(ListName)
            midOff = Round((lListName + 1) / 2)
            lMinus = (lHeadLine - midOff) / 2 - 1
            MinusLine = Left(String(lMinus, "-") & b & ListName & b _
                             & String(lMinus, "-"), lHeadLine)
            
            Call LogEvent(MinusLine, eLall)        ' printing with LogEvent
            Call LogEvent(ModelLine(ModelType), eLall)
            HeadlineName = ListName                ' remember for change detection
            modelNr = ModelType
        End If
    End If                                         ' if no change in Headline, don't print it
ProcRet:
End Sub                                            ' BugHelp.N_ShowHeader

'---------------------------------------------------------------------------------------
' Method : N_ShowInstances
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show the instances of default or specified cProcItem
'---------------------------------------------------------------------------------------
Sub N_ShowInstances(ClientDsc As cProcItem, Optional Ordinal As Long, Optional withItemDsc As Boolean)
Dim ErrClient As cErr

Dim instCount As Long
Dim msg As String
Dim ThisHeadLine As String

    If ClientDsc Is Nothing Then
        Debug.Print "no ProcItem"
        GoTo FuncExit
    End If
    
    If withItemDsc Then
        Call N_ShowProcDsc(ClientDsc, Ordinal)
    End If
    Set ErrClient = ClientDsc.ErrActive
    
    If ErrClient Is Nothing Then
        DoVerify False
        GoTo FuncExit
    End If
    
    If ErrClient.atTraceStackPos > 0 Then
        msg = ClientDsc.DbgId & " beginning at TracePos=" _
              & ErrClient.atTraceStackPos _
                
    Else
        msg = LString(ClientDsc.DbgId, lDbgM + 5)
    End If
    
    ThisHeadLine = " Call Chain Instances belonging to " & msg
    
    Call N_ShowHeader(ThisHeadLine, force:=True)

    For instCount = 0 To ErrLifeTime
        If ErrClient Is Nothing Then
            Exit For
        End If
        Call N_ShowErrInstance(ErrClient, ErrClient.atCallDepth)
        If ErrClient Is ErrClient.atErrPrev Then
            Exit For
        End If
        Set ErrClient = ErrClient.atErrPrev
    Next instCount
    
    Call N_ShowHeader(ThisHeadLine, StartOrEnd:=True)
    
FuncExit:
    Set ErrClient = Nothing
End Sub                                            ' BugHelp.N_ShowInstances

'---------------------------------------------------------------------------------------
' Method : Sub N_ShowProcDsc
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ShowProcDsc(ClientDsc As cProcItem, ByVal i As Long, Optional WithInstances As Boolean, Optional ExplainS As String, Optional ErrClient As cErr, Optional Caller As String, Optional Consumption As Boolean)
Const zKey As String = "BugHelp.N_ShowProcDsc"
        
Dim sProc As String
Dim ObjInfo As String
Dim Consumed As String
Dim ObjStatus As String
Dim Tmsg As String
Dim rL As Long
Dim StackState As String
    
    If ClientDsc Is Nothing Then
        If ErrClient Is Nothing Then
            Tmsg = "! ! ! There is an empty element on the stack in position " & i
            GoTo FuncExit
        Else
            Set ClientDsc = ErrClient.atDsc
            If ClientDsc Is Nothing Then
                Tmsg = "! ! ! unspecified Client in position " & i
                GoTo FuncExit
            End If
        End If
    Else
        If ErrClient Is Nothing Then
            Set ErrClient = ClientDsc.ErrActive
        End If
        If ErrClient Is Nothing Then
            ObjInfo = " E"
        Else
            If Not ClientDsc Is ErrClient.atDsc Then
                ObjInfo = ObjInfo & "!"
                DoVerify False
            End If
            If Not ErrClient Is ClientDsc.ErrActive Then
                If ErrClient.atRecursionOK Then
                    ObjInfo = ObjInfo & " ®"
                Else
                    ObjInfo = ObjInfo & " ®??"     ' recursion instance call
                    ' DoVerify False
                End If
followChain:
                If ErrClient.atErrPrev Is Nothing Then
                    If ErrClient.atRecursionLvl > 1 Then
                        ObjInfo = ObjInfo & "Chain?"
                        DoVerify False, "ErrPrev-Chain leaves recursion"
                    Else
                        ObjInfo = ObjInfo & vbCrLf & vbTab & vbTab & "recursion from " _
                                  & ErrClient.atErrPrev.atCalledBy.atKey
                    End If
                Else
                    Set ErrClient = ErrClient.atErrPrev ' follow chain
                End If
            End If
            If LenB(ErrClient.Reasoning) > 0 Then
                If InStr(" LsT", Left(ErrClient.Reasoning, 1)) = 0 Then
                    ObjInfo = ObjInfo & "X"
                End If
            End If
        End If
    End If
    
    If ClientDsc.ErrActive Is Nothing Then
        Tmsg = LString(i, 5) & String(OffCal - 5, b) _
        & ClientDsc.DbgId & " is an invalid Entry: .ErrActive Is Nothing."
        GoTo FuncExit
    End If
            
    If LenB(ExplainS) = 0 Then
        If E_AppErr Is ErrClient Then              ' ErrClient is the last in the loop
            ObjStatus = " current App"
        Else
            ObjStatus = vbNullString
        End If
    Else
        If InStr(ExplainS, "IdsOK") > 0 Then       ' mode is Dump D_ErrInterface
            Tmsg = ExplainS & b
        Else
            ObjStatus = ExplainS
        End If
        ExplainS = vbNullString
    End If
    
    If ClientDsc.TotalRunTime > 0 Then
        Consumed = "rT=" & RString(ClientDsc.TotalRunTime, 6)
    End If
    If Not ErrClient.atCalledBy Is Nothing Then
        If ErrClient.atCalledBy.atDsc Is Nothing Then
            Consumed = Trim(Consumed & " C: NoCaller")
        Else
            Consumed = Trim(Consumed & " C: " & ErrClient.atCalledBy.atKey)
        End If
    End If
    
    If LenB(Caller) = 0 Then
        Caller = ClientDsc.DbgId
    End If

    Call N_Suppress(Push, zKey, Value:=False)
    If Consumption Then
        ObjStatus = Consumed
        Call N_ShowProgress(i, ClientDsc, Caller, _
                            ObjectMsg:=StackState & ObjInfo & ObjStatus, _
                            ExplainS:=sProc & b & ExplainS, _
                            ErrClient:=ErrClient, _
                            doPrint:=True)
    Else
        Call N_ShowProgress(i, ClientDsc, Caller, _
                            ObjectMsg:=StackState & ObjInfo & ObjStatus, _
                            ExplainS:=sProc & Trim(b & ExplainS) & ": " & Consumed, _
                            ErrClient:=ErrClient, _
                            doPrint:=True)
    End If
    Call N_Suppress(Pop, zKey)
    If WithInstances Then
        Call N_ShowInstances(ClientDsc)
    End If

FuncExit:
    If LenB(Tmsg) > 0 Then
        Debug.Print Tmsg
    End If

End Sub                                            ' BugHelp.N_ShowProcDsc

'---------------------------------------------------------------------------------------
' Method : Sub N_ShowProgress
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_ShowProgress(Line As Long, ClientDsc As cProcItem, ByVal ObjKey As String, ByVal ObjectMsg As String, ByVal ExplainS As String, Optional ErrClient As cErr, Optional Result As String, Optional doPrint As Boolean = True)
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        GoTo ProcRet                               ' simply ignore recursion
    End If
    If ShutUpMode Then
        GoTo ProcRet
    End If
    Recursive = True

Dim P As Long
Dim px As String
Dim Target As String

    Output = vbNullString
    
    If ErrClient Is Nothing Then
        Set ErrClient = ClientDsc.ErrActive
    End If
        
    If InStr(ObjectMsg, ".") > 0 Then
        GoTo PutDsc
    End If
    
    P = InStr("+-", Left(ObjKey, 1))
    If P > 0 Then                                   ' only 0, 1, 2 possible
        px = Mid(ObjKey, 2, 1)
        ObjKey = Mid(ObjKey, 3)
        
        If px = "D" Then                            ' +D_DefProc ...
            ObjectMsg = "#" & px & "=" & LString(ClientDsc.ProcIndex, 5) & ObjectMsg
            P = 1                                   ' never removed
        ElseIf px = "T" Then                        ' +C_CallTrace
            ObjectMsg = "#" & px & "=" & LString(TraceTop + 1, 5) & ObjectMsg
            P = 1                                   ' never removed
        ElseIf InStr("NZYXRQESR", px) > 0 Then      ' calls/exits
Proc:
            If P = 2 Then                           ' "-": do not double info for exit
                Target = ErrClient.atCalledBy.atDsc.DbgId     ' ... Show only on exit
                DoVerify LenB(Target) <> 0, _
                                      "on exit, Target must be defined ???"
                ObjKey = LRString(Target, "<= " & ErrClient.atDsc.DbgId, _
                                  ErrClient.atCallDepth, lKeyM, rCutL:=Len(Target) + 3)
                P = P + 1                           ' indent out for exits
                GoTo PutDsc
            End If
        End If
        
        Output = ClientDsc.DbgId & b & Quote(ClientDsc.Module, Bracket)
        ObjKey = LRString(ObjKey & b & Output, _
                          Target, E_Active.atCallDepth, lKeyM)
    Else
        ObjKey = String(ErrClient.atCallDepth, b) & ObjKey
    End If
    
PutDsc:
    Output = RString(Line, 5) & b _
                              & ErrClient.NrMT _
                              & Left(CStateNames(ErrClient.atCallState), 1) & b _
                              & RString(ErrClient.atCallDepth _
                                        + ErrClient.atLiveLevel - LSD - P + 1, 3) & b _
                                        & LString(ObjKey, lKeyM) _
                                        & ObjectMsg ' Indent-value includes Live Stack depth with Offset LSD
     
    Output = LString(Output, OffTim) & b & LString(ErrClient.atLastInSec, 14) _
        & " rL=" & ErrClient.atRecursionLvl _
          & " cc=" & ClientDsc.CallCounter
    
    If InStr(Result, "IdsOK") = 0 Then
        Result = Output & b & ExplainS
    Else
        Result = Result & b & Output
    End If
    
    If LenB(Result) > 0 Then
        Line = Line + 1
        If doPrint Then
            Call LogEvent(Result, eLSome)          ' this saves/restores .Permitted so it is not changed
        Else
            Debug.Print Result
        End If
    End If
    
    Recursive = False
    
ProcRet:
End Sub                                            ' BugHelp.N_ShowProgress

' a=all, b=both AllErr, AppStack and CallTrace,
' c=CallTrace, e=Defproc + AllErr
' IMPLEMENT: p=Perf
Sub N_ShowStacks(Optional what As String = "e", Optional Full As Boolean)
Const zKey As String = "BugHelp.N_ShowStacks"
    
    Call N_Suppress(Push, zKey)
    
    If what = "a" Then
        Call ShowDefProcs(WithInstances:=True, Full:=Full)
        what = "b"
    End If

    If what = "b" Then
        what = "p"
    End If
    
    If what = "e" Then
        If Full Then                               ' not full version shows the same as ShowErrStack
            Call ShowDefProcs(False, Full:=True)   ' so omit
        End If
        Call ShowErrStack
    End If
    
    If what = "b" Then
        Call ShowErrStack
        Call ShowCallTrace
    Else
        If what = "c" Then
            Call ShowCallTrace
        End If
    End If
    
    Call N_Suppress(Pop, zKey)

zExit:

End Sub                                            ' BugHelp.N_ShowStacks

'---------------------------------------------------------------------------------------
' Method : Function N_Stackdef
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Set up the atDsc and Err for new items and stack these. Returns True for FirstItem called.
'---------------------------------------------------------------------------------------
Function N_StackDef(sDsc As cProcItem, sErr As cErr, IgnoreCallData As Boolean) As Boolean
Const zKey As String = "BugHelp.N_StackDef"
Const MyId As String = "N_StackDef"
Const MyClass As String = "BugHelp"

    If sDsc.ProcIndex = 0 Or sErr.atProcIndex = 0 Then
        If D_ErrInterface.Exists(sDsc.Key) Then
            If isEmpty(D_ErrInterface.Item(sDsc.Key)) Then
                DoVerify sDsc.ProcIndex = 0, sDsc.Key & " already defined at Position " & sDsc.ProcIndex
                Set D_ErrInterface.Item(sDsc.Key) = sDsc
                GoTo corrItem
            End If
            If sErr.atKey = vbNullString Then
                Set sErr = sDsc.ErrActive
            End If
            If DoVerify(D_ErrInterface.Item(sDsc.Key).ErrActive Is sErr, _
                "D_ErrInterface item(" & sDsc.Key & ") is incorrect") Then
                Set D_ErrInterface.Items(sDsc.Key) = sErr.atKey ' conflict, sErr wins
            End If
        Else
            D_ErrInterface.Add sDsc.Key, sDsc
corrItem:
            sDsc.ProcIndex = D_ErrInterface.Count
            If sErr.atProcIndex <> sDsc.ProcIndex Then  ' this includes atProcIndex = Inval
                sErr.atProcIndex = sDsc.ProcIndex       ' conflict, atDsc wins
            End If
            If CallLogging And (LogZProcs Or dontIncrementCallDepth) Then
                If dontIncrementCallDepth Then
                    sErr.Explanations = "Predefining"
                End If
                Call N_ShowProgress(CallNr, sDsc, "+D", _
                                    "creator=" & P_Active.DbgId, _
                                    sErr.Explanations, ErrClient:=sErr)
            End If
        End If
    End If
    
    If IgnoreCallData Then
        N_StackDef = True
        sDsc.CallCounter = sDsc.CallCounter + 1
    Else
        If D_ErrInterface.Count <= 2 Then          ' real calls follow Extern.Caller and DoCall
            sDsc.CallCounter = sDsc.CallCounter + 1
            sErr.atCallState = eCExited
            N_StackDef = True                      ' Omitting N_GenCallData
        Else
            Call N_GenCallData(sDsc, sErr)
        End If
    End If
    
ProcRet:
End Function                                       ' BugHelp.N_Stackdef

'---------------------------------------------------------------------------------------
' Method : N_Suppress
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Push or Pop Shutupmode and SuppressStatusFormUpdate (if allowed)
'---------------------------------------------------------------------------------------
Sub N_Suppress(Push As Boolean, Caller As String, Optional Value As Boolean = True)
Dim IdentifiedVal As cPair
Dim LastEntry As Long
Dim PrevState As Boolean
Static LogMe As Boolean
    
    If ProtectStackState = inv Then
        Set C_ProtectedStack = New Collection
    End If
    If C_ProtectedStack.Count = 0 Then
        ProtectStackState = StackDebug
    End If

    If ProtectStackState < 1 Then                  ' do not ShutUp
        ShutUpMode = False                         ' Parameter "Value" is ignored
        GoTo msgout                                ' Not on Entry Nor Exit
    End If
    If ProtectStackState > 8 Then                  ' do not change Shutupmode at all
msgout:
        If LogMe Then
            Debug.Print "Ignored! Caller=" & Caller, "ShutUpValue==" & ShutUpMode
        End If
        GoTo FuncExit                              ' Not on Entry Nor Exit
    End If

DoStacking:

    If Push Then
        If LogMe Then
            Debug.Print "Pushing Caller=" & Caller, "Count=" _
                        & C_ProtectedStack.Count, "ShutUpValue=" _
                        & ShutUpMode & "->" & Value
        End If
        Set IdentifiedVal = New cPair
        IdentifiedVal.pValue = ShutUpMode
        SuppressStatusFormUpdate = ShutUpMode
        IdentifiedVal.pId = Caller
        C_ProtectedStack.Add IdentifiedVal
        ShutUpMode = Value
    Else
        LastEntry = C_ProtectedStack.Count
        If LastEntry = 0 Then
            Debug.Print "* Pop by Caller " _
                        & Caller & " impossible, no items to pop: ", _
                        "Count=" & C_ProtectedStack.Count, _
                        "ShutUpValue(unch.)=" & ShutUpMode
            DoVerify False
        Else
            Set IdentifiedVal = C_ProtectedStack.Item(LastEntry)
            DoVerify IdentifiedVal.pId = Caller, " popper must be pusher: " & Caller
            PrevState = ShutUpMode
            ShutUpMode = IdentifiedVal.pValue
            SuppressStatusFormUpdate = ShutUpMode
            If LogMe Then
                Debug.Print "Pop by  Caller=" & Caller, "Count=" _
                        & C_ProtectedStack.Count, "ShutUpValue=" _
                        & ShutUpMode & "«" & PrevState
            End If
            If LastEntry > 0 Then
                C_ProtectedStack.Remove LastEntry
            End If
        End If
    End If
    
FuncExit:
    Set IdentifiedVal = Nothing
    SuppressStatusFormUpdate = Not ShutUpMode
    
End Sub                                            ' BugHelp.N_Suppress

'---------------------------------------------------------------------------------------
' Method : N_TraceCallReduce
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: At end of lifetime, drop elements from CallTrac
'---------------------------------------------------------------------------------------
Sub N_TraceCallReduce()

Const zKey As String = "BugHelp.N_TraceCallReduce"
    Call DoCall(zKey, tSub, eQzMode)
    
Dim i As Long
Dim j As Long
Dim StackErr As cErr
Dim Removed As Long
Dim stopRemove As Boolean
Dim TEntry As cTraceEntry
Dim TSucc As Long

    If C_CallTrace.Count > ErrLifeTime Then        ' clear if too much history
        If StackDebug > 8 Then
            Debug.Print "* Cleaning C_CallTrace because LifeTime reached: " & ErrLifeTime
        End If
        Removed = 0                                ' for debug restart
        For i = 2 To C_CallTrace.Count             ' keeping buffer ErrLifeKept to ErrLifeTime
            j = i - Removed
            Set TEntry = C_CallTrace.Item(j)
            Set StackErr = TEntry.TErr
            If ErrLifeKept >= Removed And Not stopRemove Then
                If TEntry.TRL = 0 Then
                    C_CallTrace.Remove j           ' clean up outdated, leaving #1 unchanged
                    j = inv                        ' mark the Instance as outdated, removing from C_CallTrace
                    Removed = Removed + 1
                Else
                    If TEntry.TLne > 0 Then
                        'stopRemove = True                  ' leave intact otherwise *** ???
                    End If
                End If
            End If
            StackErr.atTraceStackPos = j           ' correct this position
            TSucc = TEntry.TSuc - Removed
            If TSucc > 0 Then
                C_CallTrace.Item(TSucc).TPre = Removed - i ' used to be there before...
            End If
        Next i
    End If                                         ' clear if too much history
    
    If TraceMode Then
        If C_CallTrace.Count - ErrLifeKept - 1 > 0 Then
            Debug.Print "* N_TraceCallReduce was not able to remove " _
                        & C_CallTrace.Count - ErrLifeKept - 1 _
                        & " Items, CallTrace.Count=" _
                        & C_CallTrace.Count
        End If
    End If
    DoVerify C_CallTrace.Count < ErrLifeTime, " REALLY bad!! ErrLifeTime may be too small"
    Call ShowStatusUpdate

    Set StackErr = Nothing
    Set TEntry = Nothing
    
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.N_TraceCallReduce

'---------------------------------------------------------------------------------------
' Method : Sub N_TraceDef
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub N_TraceDef(TEntry As cTraceEntry, Optional WithIndent As Boolean)

Dim ClientDsc As cProcItem
Dim ClientErr As cErr
Dim PreEntry As cTraceEntry
Dim Caller As String
Dim Source As String

    With TEntry
        .TLog = vbNullString
        Set ClientErr = .TErr
        Set ClientDsc = ClientErr.atDsc
        .TLD = ClientErr.atLastInDate              ' trace<>non-instance vals in cErr
        .TLS = ClientErr.atLastInSec
        .TPS = ClientErr.atPrevEntrySec
        .TES = ClientErr.atThisEntrySec
    
        .TSrc = ClientDsc.DbgId
        Source = .TSrc & b & Quote(ClientDsc.Module, Bracket)
        .TLog = LRString(Caller, LString(.TES, 14), _
                         ClientErr.atCallDepth, lKeyM - 1) & b _
                         & LString(.TLne, 3) & b _
                         & LString(Source, lHeadLine - OffObj) & b _
                         & ClientErr.Explanations
    End With                                       ' TEntry
    
    Set ClientDsc = Nothing
    Set ClientErr = Nothing
    Set PreEntry = Nothing

zExit:

End Sub                                            ' BugHelp.N_TraceDef

'---------------------------------------------------------------------------------------
' Method : N_TraceEntry
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Trace the current stack top. Obviously can't trace itself or recurse
'---------------------------------------------------------------------------------------
Sub N_TraceEntry(Optional TEntry As cTraceEntry, Optional ExplainS As String)

Const zKey As String = "BugHelp.N_TraceEntry"
Const MyId As String = "N_TraceEntry"

Dim ClientDsc As cProcItem
Dim ClientErr As cErr
Dim PN As String
Dim MN As String
Dim LineNo As Long
Dim Line As String

Dim sLiveName As String                            ' if data is from LiveStack, use source Line
Dim sNamPos As Long                                ' position from Proc Name in sLiveName
Dim sType As String                                ' sub or function from sLiveName
Dim TdbgId As String
Dim Key As String
Dim sKey As String
Dim putS As Boolean
Dim iPat As Long
Dim HasMoreLevels As Boolean
Dim HaveCalledProc As Boolean
Dim IsInLiveStack As Long
Dim Lvl As Long

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then                       ' do not trace the Trace
        ' Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True
    
    Lvl = 0
    HaveCalledProc = False
    IsInLiveStack = False
    Line = vbNullString
    sLiveName = vbNullString
    sKey = vbNullString
    TdbgId = vbNullString
    putS = False
    If TEntry Is Nothing Then
        Set TEntry = New cTraceEntry
    Else
        If LenB(TEntry.TSrc) > 0 Then              ' this is derived from LiveStack at some earlier time
            putS = True                            ' Put on stack!
            sLiveName = Trunc(1, TEntry.TSrc, "(") ' cut off Call Parameters
            
            Set ClientErr = TEntry.TErr
            If ClientErr Is Nothing Then
                sNamPos = InStr(sLiveName, "Call ")
                If sNamPos > 0 Then
                    sType = tSub
                    sLiveName = Mid(sLiveName, sNamPos + 5)
                Else
                    sNamPos = InStr(sLiveName, " = ")
                    If sNamPos > 0 Then
                        sType = tFunction
                        sLiveName = Mid(sLiveName, sNamPos + 3)
                    Else
                        sNamPos = InStr(sLiveName, b)
                        sType = tSub               ' wild assumption
                        sLiveName = Mid(sLiveName, sNamPos + 1)
                    End If
                    sLiveName = "Live." & sLiveName
                    GoTo dummyGen
                End If
            Else
                sLiveName = ClientErr.atKey
                If DoVerify(D_ErrInterface.Exists(sLiveName), "no entry in D_ErrInterface ???") Then
                    sLiveName = "Missing." & sLiveName
dummyGen:
                    Call N_ConDscErr(ClientDsc, sLiveName, sType, eQnoDef, ClientErr)
                    Set TEntry.TErr = ClientErr
                    GoTo LessDetails
                End If
            End If
            
            Set ClientDsc = ClientErr.atDsc
            TdbgId = ClientDsc.DbgId
            If InStr(sLiveName, TdbgId) > 0 Then   ' got a reference
                sKey = ClientErr.atKey
            End If
        Else                                       ' have TEntry.Tsrc<>""
            Set ClientErr = TEntry.TErr
            Set ClientDsc = ClientErr.atDsc
            sKey = ClientErr.atKey
            putS = True
            If ClientErr.atCallDepth = 0 Then      ' must be Extern.Caller
                HaveCalledProc = True
                GoTo LessDetails
            End If
        End If
    End If                                         ' new or existing TEntry
    
    If LenB(ExplainS) = 0 Then
        ExplainS = ClientErr.atMessage
        ClientErr.atMessage = vbNullString
    End If
    If Not ErrExActive Then
        If DebugMode Then
            Debug.Print "* ErrEx is not active, can't use Live Stack"
        End If
        TEntry.TLne = 0
        TEntry.TSrc = vbNullString
        HaveCalledProc = Not ClientErr Is Nothing
        GoTo LessDetails
    End If
    
    If Not (DebugMode Or DebugLogging) Then        ' determine if we want to use source line info
        HaveCalledProc = True
        GoTo LessDetails                           ' no source line in trace printout
    End If
    
    ' look for Caller match (not Proc reference here)
    With ErrEx.LiveCallstack
        Do
            Lvl = Lvl + 1
            Set ClientDsc = Nothing
            ' get all data from ErrEx.LiveCallStack line
            PN = .ProcedureName
            MN = .ModuleName
            LineNo = .LineNumber
            Line = .LineCode
            HasMoreLevels = .NextLevel
            
            '    Debug.Print Lvl, "More: " & HasMoreLevels, LString(TdbgId, lDbgM), LString(Key, lDbgM), LString(PN, lDbgM), Line
            
            ' position to next ErrEx.LiveCallStack line if any, no more old line info avail.
            If PN = MyId Then
                GoTo noDsc
            End If
            If HasMoreLevels Then
                If Z_SourceAnalyse(sLiveName, TdbgId) Then
                    HaveCalledProc = True
                    Key = sKey
                    GoTo UseLastKey
                End If
            End If
            
            Key = MN & "." & PN
            If LenB(sKey) > 0 Then                 ' want only this Proc
                If iPat > 1 Then
                    If Not IsSimilar(sKey, Key) Then
                        GoTo noDsc
                    End If
                ElseIf Key <> sKey Then            ' refuse others
                    GoTo noDsc
                End If
            End If
            If IsInLiveStack = 0 Then
                IsInLiveStack = Lvl                ' youngest proc reference only
                GoTo noDsc                         ' skip for potential later reference
            End If
UseLastKey:
            If D_ErrInterface.Exists(Key) Then     ' do search for it
                Set ClientDsc = D_ErrInterface.Item(Key)
                Set ClientErr = ClientDsc.ErrActive
            Else
                GoTo noDsc
            End If
            
LessDetails:
            If Not HaveCalledProc Then
                GoTo noDsc
            End If
            DoVerify TEntry.TErr Is ClientErr, "design check ???"
                            
            If putS And Not dontIncrementCallDepth Then
                With TEntry
                    .TDet = ExplainS
                    .TLne = LineNo
                    If LenB(.TSrc) > 0 Then
                        .TSrc = Trim(Trim(Trunc(1, Line, "'", sLiveName)) _
                                     & " '" & sLiveName)
                        If .TSrc = "'" Then
                            .TSrc = vbNullString
                        End If
                    End If
                    .TRL = ClientErr.atRecursionLvl
                End With                           ' TEntry
                Call TEntry.TraceAdd(ExplainS)
                Call N_TraceDef(TEntry)
                If HaveCalledProc Then
                    GoTo FuncExit
                End If
            End If
noDsc:
        Loop While HasMoreLevels
    End With                                       ' ErrEx.LiveCallStack
    
    If IsInLiveStack > 0 Then                      ' did not find caller line
        Lvl = 0
        With ErrEx.LiveCallstack
            Do
                Lvl = Lvl + 1
                If Lvl = IsInLiveStack + 1 Then    ' skip to item after ProcCall
addLast:
                    Line = .LineCode
                    LineNo = .LineNumber
                    With TEntry
                        .TDet = ExplainS
                        .TLne = LineNo
                        .TSrc = Trim(Trim(Trunc(1, Line, "'", sLiveName)) _
                                     & " '" & sLiveName)
                        .TRL = ClientErr.atRecursionLvl
                    End With                       ' TEntry
                    Call TEntry.TraceAdd(ExplainS)
                    Call N_TraceDef(TEntry)
                    
                    GoTo FuncExit
                End If
                HasMoreLevels = .NextLevel
            Loop While HasMoreLevels
            GoTo addLast
        End With                                   ' ErrEx.LiveCallStack
    End If

    If putS And (DebugMode Or DebugLogging) Then
        Debug.Print sKey & " is not on live stack, can not trace"
        DoVerify Not DebugMode, " instructed to trace, but not on LiveStack"
    End If
    Set TEntry.TErr = Nothing                      ' assume no success

FuncExit:
    Recursive = False
    Set ClientDsc = Nothing
    Set ClientErr = Nothing
    
zExit:

ProcRet:
End Sub                                            ' BugHelp.N_TraceEntry

'---------------------------------------------------------------------------------------
' Method : N_Undefine
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Clear traces of dead Proc
'---------------------------------------------------------------------------------------
Sub N_Undefine(ClientDsc As cProcItem, i As Long, Optional Key As String)

    If ClientDsc Is Nothing Then
        GoTo zExit
    End If
    If LenB(Key) = 0 Then
        Key = ClientDsc.Key
    End If
    D_ErrInterface.Remove Key
    Set ClientDsc.ErrActive = Nothing
    Set ClientDsc = Nothing
    Debug.Print "Proc " & LString(Key, 2 * lDbgM) & " has been undefined, #" & i
       
zExit:

End Sub                                            ' BugHelp.N_Undefine

'---------------------------------------------------------------------------------------
' Method : ProcCall
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Define or use Proc with atDsc and Err at time of Entry
'---------------------------------------------------------------------------------------
Sub ProcCall(ClientErr As cErr, ByVal ClientKey As String, Qmode As eQMode, ByVal CallType As String, Optional ByVal ExplainS As String, Optional Recursive As Boolean)

Dim qMsg As String
Dim ClientDsc As cProcItem

    If isEmpty(QModeNames) Then                    ' basic inits for BugHelp required
        Call Z_StartUp(False)
    End If

    If FastMode Then
        ExplainS = Trim(ExplainS & b & "FastMode using Y_Type")
        Call DoCall(ClientKey, CallType, eQyMode, ClientDsc, ExplainS)
        Set ClientErr = ClientDsc.ErrActive
        ClientErr.atRecursionOK = Recursive
        If aNameSpace Is Nothing _
        Or aRDOSession Is Nothing Then
            Call N_ConDscErr(ClientDsc, ClientKey, CallType, Qmode, Nothing)
            ClientDsc.CallCounter = 0
        End If
        LogZProcs = True
        GoTo FuncExit
    Else
        If Not ClientDsc Is Nothing Then
            If ClientDsc.CallMode <> Qmode _
            And ClientDsc.CallMode > eQnoDef Then
                If ClientDsc.CallCounter > 0 Then
                    aBugTxt = "** Change in Qmode is improbable, but allowed"
                    DoVerify ClientDsc.CallMode = Qmode
                    If StackDebug > 5 Then
                        qMsg = String(20, b) _
                                & "Changing Mode for " & ClientDsc.DbgId _
                                & ", CallCounter=" & Right(String(5, b) _
                                & ClientDsc.CallCounter, 5) & String(31, b) _
                                & "from " & ClientDsc.ModeName _
                                & "(" & ClientDsc.CallMode _
                                & ") to M=" & QModeNames(Qmode) & "(" & Qmode & ")"
                        Debug.Print qMsg
                    End If
                Else
                    ClientDsc.CallCounter = 0      ' start a new call count after improbable mode change
                End If
            End If
            ClientDsc.CallMode = Qmode             ' always set the (probably unchanged, unless new) CallMode
        End If
        Call DoCall(ClientKey, CallType, Qmode, ClientDsc, ExplainS)
        Set ClientErr = ClientDsc.ErrActive

    End If
        
    If Qmode = eQArMode Or Qmode = eQrMode Then
        Recursive = True
        ClientErr.atRecursionOK = Recursive
        ExplainS = ExplainS & ", Recursive"
    End If
    If Qmode = eQEPMode Then
        Call Z_EntryPoint(ClientDsc)
    End If
    
    If Qmode > eQuMode Then                        ' Application Levels
        E_Active.EventBlock = True                 ' Default for Applications during ProcCall
        If ErrStatusFormUsable Then
            frmErrStatus.fNoEvents = E_Active.EventBlock
            Call BugEval
        End If
        ClientErr.atCallState = eCpaused
    End If
    
    If ClientDsc.ProcIndex = 0 And InStr(CallType, " EP") = 0 Then ' OK for Macro EP
        DoVerify ClientKey <> ExternCaller.Key, _
                ClientKey & " used as dummy for Extern.Caller only!!"
        GoTo FuncExit                              ' end of all inits. Stacks set up, no further actions
    End If
    
FuncExit:
    SuppressStatusFormUpdate = False
    If Qmode > eQuMode Then                        ' Application Levels
        E_Active.EventBlock = False                ' for new Applications
        Set E_AppErr = E_Active
        If ErrStatusFormUsable Then
            frmErrStatus.fNoEvents = E_Active.EventBlock
            frmErrStatus.fCurrAppl = E_Active.atDsc.DbgId
            Call BugEval
        End If
    End If

End Sub                                            ' BugHelp.ProcCall

'---------------------------------------------------------------------------------------
' Method : ProcExit
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Reduce the .RecursionLvl and set the CallState when leaving an instance
'---------------------------------------------------------------------------------------
Sub ProcExit(fromErr As cErr, Optional DisplayValue As String)

    Call DoExit(fromErr.atKey, DisplayValue)
    If fromErr.atFastMode Then
        GoTo ProcRet
    End If
            
    S_DbgId = fromErr.atDsc.DbgId
                
    If fromErr.atDsc.CallMode = eQEPMode Then
        Call ReturnEP
    End If
        
    If ErrStatusFormUsable And Not SuppressStatusFormUpdate Then
        If frmErrStatus.Visible Then
            Call QueryErrStatusChange(True)        ' .f values are used to modify globals
            Call frmErrStatus.UpdInfo
        End If
    End If
    
    With E_Active                                  ' Restore the values from caller's env
        Z§ErrSnoCatch = .ErrSnoCatch               ' No ErrHandler Recursion;  NO N_PublishBugState
        Z§ErrNoRec = .ErrNoRec                     ' No ErrHandler Recursion; use N_PublishBugState
        .atFuncResult = DisplayValue
    End With                                       ' E_Active
    
FuncExit:
    Call ErrReset(4)                               ' keep caller's Try setting
    
    If E_Active.atCallDepth < 2 Then
        If AppStartComplete Then
            Debug.Print
            Call LogEvent("* " & String(20, "-") _
                          & " Outlook waiting for Events or Macro Calls", eLSome)
            Call N_ShowHeader("BugHelp Log " & TimerNow)
            Z§AppStart.ErrActive.EventBlock = False
            Call SetOnline(olCachedConnectedFull)
        End If
        E_Active.EventBlock = False
        If ErrStatusFormUsable Then
            frmErrStatus.fNoEvents = E_Active.EventBlock
            frmErrStatus.fOnline.Caption = "Online"
            Call BugEval
        End If
        NoEventOnAddItem = False
    End If
    If E_Active.atDsc.CallMode = eQAsMode Then
        Set E_AppErr = E_Active
    End If
    
ProcRet:
End Sub                                            ' BugHelp.ProcExit

'---------------------------------------------------------------------------------------
' Method : Sub QueryErrStatusChange
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub QueryErrStatusChange(Reversed As Boolean)      ' set or clear ErrDisplayModify
Const zKey As String = "BugHelp.QueryErrStatusChange"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        Debug.Print String(OffCal, b) & "ignored recursion .from " _
                                      & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    
    If ErrStatusFormUsable Then
        ErrDisplayModify = False                   ' must check this if ErrStatusForm is usable
    Else
        ErrDisplayModify = True                    ' will need to update if/when ErrStatusForm becomes usable
        GoTo ProcRet
    End If
    Recursive = True
     
    Call frmErrStatus.ReEvaluate(Not Reversed)     ' set the .f values from globals or vice versa

FuncExit:
    Recursive = False
     
ProcRet:
End Sub                                            ' BugHelp.Call QueryErrStatusChange

'---------------------------------------------------------------------------------------
' Method : Sub ReturnEP
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ReturnEP()

Const zKey As String = "BugHelp.ReturnEP"
    Call DoCall(zKey, tSub, eQzMode)
    
    If EPCalled Then
        Call FldActions2Do                         ' if we have open items, do them now
    End If
    
    EPCalled = False
    StopRecursionNonLogged = False
    NoEventOnAddItem = False
    
    If Not P_EntryPoint.ErrActive Is Nothing Then
        P_EntryPoint.ErrActive.Explanations = "(-- waiting for next Event --)"
    End If

zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.ReturnEP

'---------------------------------------------------------------------------------------
' Method : ShowCallTrace
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show the Call Stack in Reverse Order
'---------------------------------------------------------------------------------------
Sub ShowCallTrace(Optional limitCount As Long = -1) ' *** Entry Point ***

Dim TEntry As cTraceEntry
Dim ErrClient As cErr

Dim i As Long
Dim j As Long
Dim SD As Long
Dim aTCal As String
Dim logLine As String
Dim saveDebug As Long
Dim saveStatusFormState As Boolean
   
    saveDebug = StackDebug
    StackDebug = 0                                 ' this causes output ONLY to file
    saveStatusFormState = SuppressStatusFormUpdate
    SuppressStatusFormUpdate = True
    
    If C_CallTrace.Count = 0 Then
        GoTo FuncExit
    End If
                   
    Call CloseLog                                  ' use new logfile for CallTrace
    Call N_ShowHeader("C-CallTrace", force:=True, ModelType:=3)
    
    i = TraceTop Mod ErrLifeTime
    If limitCount <= 0 Then
        j = C_CallTrace.Count
    Else
        j = limitCount
    End If
    Do
        T_DC.LogFileLen = 0                        ' no size limit here
        Set TEntry = C_CallTrace(i)
        Set ErrClient = TEntry.TErr
        With ErrClient
            If .atCallState = eCExited Then
                aTCal = "E"
            ElseIf .atCallState = eCpaused Then
                aTCal = "P"
            ElseIf .atCallState = eCActive Then
                aTCal = "A"
            Else
                DoVerify Not DebugMode, "** undefined never happens in C_CallTrace?"
                aTCal = "W2"
            End If
        End With                                   ' ErrClient
        
        If ErrClient.atCallState = eCUndef Then
            logLine = RString(TEntry.Tinx, 5) & b _
                                              & ErrClient.NrMT & aTCal & String(OffCal - OffLvl + 2, b)
        Else
            SD = ErrClient.atCallDepth
            logLine = RString(TEntry.Tinx, 5) & b _
                                              & ErrClient.NrMT & aTCal & b & RString(SD, 3) & b
        End If
        logLine = logLine _
                  & TEntry.TLog _
                  & " RL=" & TEntry.TRL _
                  & " Pre=" & TEntry.TPre & " Suc=" & TEntry.TSuc
        ' Debug.Print logLine
        
        Call LogEvent(logLine, eLall)
        j = j - 1
        If j < 1 Then
            Exit Do
        End If
        If i = 1 Then
            i = C_CallTrace.Count
        Else
            i = i - 1
        End If
    Loop
    
    Call N_ShowHeader("C-CallTrace", StartOrEnd:=True, ModelType:=3)
    Call ShowLogWait(False)
    Call CloseLog(KeepName:=False)
    
FuncExit:
    Set TEntry = Nothing
    Set ErrClient = Nothing
    StackDebug = saveDebug
    SuppressStatusFormUpdate = saveStatusFormState

End Sub                                            ' BugHelp.ShowCallTrace

'---------------------------------------------------------------------------------------
' Method : Sub ShowDbgStatus
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowDbgStatus(Optional Prompt As String)       ' *** Entry Point ***
    
    Call N_Suppress(Push, "BugHelp.ShowDbgStatus")
    SuppressStatusFormUpdate = True

Dim sTime As Variant
    
    If MayChangeErr Then
        On Error GoTo errCode                      ' clears the err!!
    End If
    
    If ErrStatusFormUsable Then
        GoTo useTermForm
    Else
        DebugControlsUsable = DebugControlsWanted
        If DebugControlsUsable Then                ' ignore the default
useTermForm:
            DebugControlsUsable = True
        End If
    End If
    
    If ErrStatusFormUsable Then
        GoTo ShowIt
    End If
    
    If aNonModalForm Is Nothing Then
        Set aNonModalForm = frmErrStatus           ' calls QueryErrStatusChange, but not showing yet
        ' frmErrStatus.showmodal = False
        '  vbNullString             can not set here, do this in form's properties!!!
    Else
        Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
    End If
    If LenB(Prompt) = 0 Then
        frmErrStatus.fTerminationFlag.Caption _
        = "Termination Flag = " & GetTerminationState
    Else
        frmErrStatus.fTerminationFlag.Caption = Prompt
    End If
ShowIt:
    Call ShowOrHideForm(frmErrStatus, ShowIt:=True)
    doMyEvents                                     ' allow interaction, delay and wait
    If DebugMode Then
        If Wait(5, trueStart:=sTime, DebugOutput:=False) Then
            aBugTxt = "Press Enter then Confirm continue on Debug for ShowDbgStatus"
            DoVerify False, ShowMsgBox:=True
        End If
    ElseIf LenB(Testvar) = 0 Then                  ' no need to Show this form
        If Not IgnoreUnhandledError Then
            frmErrStatus.Hide
        End If
    End If
    GoTo FuncExit
errCode:
    If aNonModalForm Is Nothing Then
        ErrStatusFormUsable = False
    End If

FuncExit:
    Call ErrReset(0)
    Call N_Suppress(Pop, "BugHelp.ShowDbgStatus")
    SuppressStatusFormUpdate = False

End Sub                                            ' BugHelp.ShowDbgStatus

'---------------------------------------------------------------------------------------
' Method : Sub ShowDefProcs
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowDefProcs(Optional WithInstances As Boolean, Optional Full As Boolean, Optional ErrInterface As Boolean) ' *** Entry Point ***
Const zKey As String = "BugHelp.ShowDefProcs"
    
    Call N_Suppress(Push, zKey)
    SuppressStatusFormUpdate = True
    
Dim i As Long
Dim aDsc As cProcItem
            
Dim CallNumSave As Long
Dim keyV As Variant
Dim AllOrNot As String
Dim aDict As Dictionary
Dim Both As Boolean
    
    CallNumSave = CallNr
    CallNr = 0
    Both = Full And ErrInterface
    
    If Full Then
        AllOrNot = " (all)"
    Else
        AllOrNot = " (only active)"
    End If
    
doBoth:
    If Both Then
        Call N_ShowHeader("D_ErrInterface in Creation Order" & AllOrNot, force:=True)
        Set aDict = D_ErrInterface
    Else
        Call N_ShowHeader("D_ErrInterface in Creation Order" & AllOrNot, force:=True)
        Set aDict = D_ErrInterface
    End If
    For i = 0 To aDict.Count - 1
        keyV = aDict.Keys(i)
        If isEmpty(aDict.Items(i)) Then
            Debug.Print LString(i, 5) & String(20, b) & LString(keyV, lDbgM) _
        & "??? is empty, is removed"
            aDict.Remove keyV                      ' INTRINSIC proc
        Else
            Set aDsc = aDict.Items(i)
            If aDsc Is Nothing Then
                Debug.Print "**** error in D_ErrInterface: Null Entry at pos=" & i
            Else
                If aDsc.ErrActive.atLastInSec > 0 _
                Or Full Then
                    Call N_ShowProcDsc(aDsc, (i), WithInstances:=WithInstances, _
                                       Consumption:=True)
                End If
            End If
        End If
    Next i
    Call N_ShowHeader("Procs in Creation Order", StartOrEnd:=True)
    If Both Then
        Both = False
        GoTo doBoth
    End If

FuncExit:
    Set aDsc = Nothing
    Set aDict = Nothing
    Call N_Suppress(Pop, zKey)
    CallNr = CallNumSave
    SuppressStatusFormUpdate = False

End Sub                                            ' BugHelp.ShowDefProcs

'---------------------------------------------------------------------------------------
' Method : Sub ShowErr
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowErr()                                      ' *** Entry Point ***

Const zKey As String = "BugHelp.BugEval"

'------------------- gated Entry -------------------------------------------------------
    If AppStartComplete Then
        
        DebugControlsUsable = True
        Call ShowOrHideForm(frmErrStatus, True)
        ErrStatusHide = False
        
    End If                                         ' Appstart complete only
    
End Sub                                            ' BugHelp.ShowErr

'---------------------------------------------------------------------------------------
' Method : ShowErrInterface
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show the Data in D_ErrInterface
'---------------------------------------------------------------------------------------
Sub ShowErrInterface(Optional Full As Boolean = True, Optional WithDetails As Long, Optional Predefined As Boolean = False) ' *** Entry Point ***

Dim i As Long
Dim TName As String
Dim TKey As String
Dim ClientDsc As cProcItem
Dim ErrClient As cErr
Dim TDetail As String
Dim TdbgId As String
Dim saveStatusFormState As Boolean

    saveStatusFormState = SuppressStatusFormUpdate
    SuppressStatusFormUpdate = True
    Call Try(testAll)                                 ' Try anything, autocatch, no resetting

    If Not Predefined Then
        If D_ErrInterface.Count > 0 Then
            Call ShowDefProcs(ErrInterface:=True, Full:=Full)
            GoTo FuncExit
        End If
    End If
    
    Debug.Print LString("Num", 4) & LString("TypeName", 16) _
        & RString("Has/at Key", lKeyM) & b _
                                       & LString("Wanted Key", lKeyM) & "Detail"
    For i = 0 To D_ErrInterface.Count - 1
        TDetail = vbNullString
        Catch
        TName = TypeName(D_ErrInterface.Items(i))
        TKey = D_ErrInterface.Keys(i)
        If WithDetails = 0 Then
            GoTo nullDetail
        End If
        
        Select Case TName
            Case "cProcItem"
                Set ClientDsc = D_ErrInterface.Items(i)
                Set ErrClient = ClientDsc.ErrActive
isDsc:
                If ClientDsc Is Nothing Then
                    GoTo noDsc
                End If
                If ErrorCaught <> 0 Then
                    TDetail = Err.Description
                Else
                    TdbgId = Mid(D_ErrInterface.Keys(i), 3)
                    If InStr(ClientDsc.DbgId, TdbgId) = 0 Then
                        If ClientDsc.DbgId <> "Extern" Then
                            TDetail = "InF=" & TdbgId & " <> Dsc=" & ClientDsc.Key
                        End If
                    ElseIf LenB(ErrClient.atKey) = 0 Then
noDsc:
                        If WithDetails > 2 Then
                            GoTo nextInLoop
                        End If
                        If LenB(ClientDsc.CallType) > 0 Then
                            GoTo badCall
                        End If
                    Else
badCall:
                        If ClientDsc.Key <> ErrClient.atKey Then
                            TDetail = "Dsc=" & ClientDsc.Key & " <> Err=" & ErrClient.atKey
                        End If
                    End If
                    If LenB(TDetail) > 0 Then
                        TDetail = "IdsOk=F " & TDetail
                        If LenB(ClientDsc.CallType) > 0 Then
                            If WithDetails = 0 Then
                                GoTo nullDetail
                            End If
                        Else
                            If WithDetails = 1 Then
                                GoTo nextInLoop
                            End If
                        End If
                    Else
                        If WithDetails > 2 Then
                            Call Z_CheckLinkage(ClientDsc, zeroClientIsOk:=True, Result:=TDetail)
                        End If
                    End If
                    Call N_ShowProcDsc(ClientDsc, i, WithInstances:=(WithDetails > 3), _
                                       ExplainS:=TDetail, ErrClient:=ErrClient, Consumption:=False)
                End If
            Case "cErr"
                Set ErrClient = D_ErrInterface.Items(i)
                Set ClientDsc = ErrClient.atDsc
                If WithDetails > 1 Then
                    GoTo isDsc
                End If
                If ClientDsc.Key <> ErrClient.atKey Then
                    TDetail = "atK=" & ErrClient.atKey & " <> Dsc=" & ClientDsc.Key
                End If
            Case Else
                If WithDetails > 0 Then
                    TDetail = D_ErrInterface.Items(i).Count
                    If LenB(TDetail) > 0 Then
                        TDetail = "Count=" & TDetail
                    End If
                End If
nullDetail:
                Debug.Print LString(i, 4) & LString(TName, 16) _
        & RString(D_ErrInterface.Keys(i), lKeyM) & b _
                                                 & LString(TKey, lKeyM) & TDetail
        End Select
nextInLoop:
    Next i
    
FuncExit:
    Call ErrReset(0)
    Set ErrClient = Nothing
    SuppressStatusFormUpdate = saveStatusFormState

End Sub                                            ' BugHelp.ShowErrInterface

'---------------------------------------------------------------------------------------
' Method : ShowErrorStatus
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show a form that allows displaying error information, debug controls, allows remedy
'---------------------------------------------------------------------------------------
Sub ShowErrorStatus()                              ' *** Entry Point ***

Const zKey As String = "BugHelp.ShowErrorStatus"
' Dim Z§ShowErrorStatus As cProcItem                                      ' Predefined!
    
    Call N_Suppress(Push, zKey)
    
    With T_DC
        If isEmpty(aNonModalForm) Then
            Set aNonModalForm = Nothing
        End If
        If aNonModalForm Is Nothing Then
            Call ShowDbgStatus
        End If
        If aNonModalForm Is Nothing Then
            DoVerify "* Unable to Show aNonModalForm!"
            ErrStatusFormUsable = False
            GoTo ProcReturn
        Else
            ErrStatusFormUsable = True
        End If
        
        With aNonModalForm
            .fLastErr = T_DC.DCerrNum
            .fLastErrSource = T_DC.DCerrSource
            .fLastErrMsg = T_DC.DCerrMsg
            ' get information from E_AppErr (probably set by ErrTry)
            .fLastErrExplanations = E_AppErr.Explanations
            .fLastErrReasoning = E_AppErr.Reasoning
            If .fLastErrIndications.Enabled Then
                .fModifications = True
            End If
            
            If Z§ShowErrorStatus.CallCounter = 1 Then
                .Top = 245
                .Left = 1041
                .fLastErrExplanations = "Manueller StartUp mit Stop: " _
                                        & "Alternative Werte für Debugoptionen jetzt wählen! " & Time
                Debug.Print .fLastErrExplanations
                .Show                              ' Modal
                Call BugEval
            Else
                .Show vbModeless
            End If
        End With                                   ' aNonModalForm
    End With                                       ' T_DC

ProcReturn:
    Call N_Suppress(Pop, zKey)
    
End Sub                                            ' BugHelp.ShowErrorStatus

'---------------------------------------------------------------------------------------
' Method : Sub ShowErrStack
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowErrStack()                                 ' *** Entry Point ***
Const zKey As String = "BugHelp.ShowErrStack"
    
    Call N_Suppress(Push, zKey, Value:=False)

Dim ErrClient As cErr
Dim i As Long
Dim uTop As Long
Dim TheErrStackName As String
    
    uTop = E_Active.atCallDepth
    If uTop < 2 Then                               ' do not show external caller
        Debug.Print "- The current active Error Stack is empty. Gimme sompn to do."
        GoTo FuncExit
    End If
    
    TheErrStackName = "Following the CalledBy"
    Call N_ShowHeader(TheErrStackName, force:=True)
    
    Set ErrClient = E_Active
    For i = uTop To 1 Step -1
        Call N_ShowErrInstance(ErrClient, i)
        Set ErrClient = E_Active.atCalledBy
    Next i
    
    Call N_ShowHeader(TheErrStackName, StartOrEnd:=True)

FuncExit:
    Set ErrClient = Nothing
    Call N_Suppress(Pop, zKey)

End Sub                                            ' BugHelp.ShowErrStack

'---------------------------------------------------------------------------------------
' Method : Sub ShowFunctionValue
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowFunctionValue(FunctionRef As Variant, FuncValue As Variant, Optional TestSafe As Boolean = True, Optional NoMultiLine As Boolean = True)

Const zKey As String = "BugHelp.ShowFunctionValue"
    Call DoCall(zKey, tSub, eQzMode)

Dim tInfo As cInfo
Dim i As Long
Dim printVal As String
Dim FunctionName As String

    If VarType(FunctionRef) = vbObject Then
    End If
    If TestSafe Then
        Call getInfo(tInfo, FuncValue, Assign:=False)
    Else
        GoTo zExit
    End If
    
    If tInfo.iAssignmentMode = 1 Then
        printVal = tInfo.iValue
        If NoMultiLine Then
            i = InStr(tInfo.iValue, vbCr)
            If i > 0 Then
                printVal = Left(printVal, i - 1) & "..."
            End If
        End If
        If InStr(FunctionName, " of ") = 0 Then
            Debug.Print "value of " & FunctionName & "=" & Quote(printVal)
        Else
            Debug.Print FunctionName & b & Quote(printVal)
        End If
    End If
    
FuncExit:
    Set tInfo = Nothing
    
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.ShowFunctionValue

'---------------------------------------------------------------------------------------
' Method : ShowLiveStack
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Show the live Call Stack
'---------------------------------------------------------------------------------------
Sub ShowLiveStack(Optional Full As Boolean, Optional doPrint As Boolean = True, Optional tSubFilter As Boolean, Optional getNewStack As Boolean = True) ' *** Entry Point ***
Const zKey As String = "BugHelp.ShowLiveStack"
    
Dim i As Long
Dim liveCount As Long
Dim LCI As cCallEnv
    
    If Not ErrExActive Then
        Debug.Print "* ErrEx is not active, can't Show Live Stack"
        GoTo zExit
    End If
    
    Call N_Suppress(Push, zKey)

    Set D_LiveStack = N_GetLiveStack
    liveCount = D_LiveStack.Count - 1
    
    If doPrint Then
        Call N_ShowHeader("LiveCallstack, Time=" & Now(), force:=True, ModelType:=2)
        For i = 1 To liveCount                     ' not starting with 0-Element (that would be "N_GetLiveStack")
            If doPrint Then
                Set LCI = D_LiveStack.Items(i)
                If LCI.CallerErr Is Nothing Then
                    If Full Then
                        GoTo ShowIt
                    End If
                Else
ShowIt:
                    Call N_PrintNameInfo(i, LCI)
                End If
            End If
        Next i
        Call N_ShowHeader("LiveCallstack", StartOrEnd:=True)
    End If
    
FuncExit:
    Call N_Suppress(Pop, zKey)
    Set LCI = Nothing
    
zExit:

End Sub                                            ' BugHelp.ShowLiveStack

'---------------------------------------------------------------------------------------
' Method : N_GetLiveStack
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Get an Extract of the live Call Stack into a dictionary
'---------------------------------------------------------------------------------------
Function N_GetLiveStack() As Dictionary

Dim i As Long
Dim LCI As cCallEnv

    Set N_GetLiveStack = New Dictionary
    Set LCS = ErrEx.LiveCallstack
    LCS.FirstLevel                                 ' just back to the first one
    Do
        Set LCI = New cCallEnv
        With LCI
            .StackDepth = i
            .ModuleName = LCS.ModuleName
            .ProcedureName = LCS.ProcedureName
            .LineNumber = LCS.LineNumber
            .LineCode = LCS.LineCode
            .ModProc = .ModuleName & "." & .ProcedureName
            .DscKind = ExStackProcNames(LCS.ProcedureKind)
            N_GetLiveStack.Add i, LCI
        End With                                   ' LineInfo.LCI
        i = i + 1
    Loop While LCS.NextLevel

    Set LCI = Nothing

End Function                                       ' BugHelp.N_GetLiveStack

'---------------------------------------------------------------------------------------
' Method : Sub ShowStacks
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub ShowStacks(Optional Full As Boolean)           ' *** Entry Point ***
    Call N_ShowStacks(Full:=Full)                  ' alias for Entry use
End Sub                                            ' BugHelp.ShowStacks

'---------------------------------------------------------------------------------------
' Method : Sub SimulateAnError
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SimulateAnError()

Dim msg As String

    Debug.Print "Errex Enabled=" & ErrEx.IsEnabled

    Debug.Print 1 / 0                              ' Simulate division by zero error (error 11)
    Debug.Print "A" / 1                            ' Simulate type mismatch error (error 13)
    Err.Raise &H123, , "This should be caught by the catch &H123 block..."
    Err.Raise &H456, , "This should be caught in the catch-all block..."
    
    ' Program flow now transfers to Finally...
        
    ErrEx.Catch 11, 13                             ' you can set up a constant/enumeration if you prefer
    Debug.Print "Catch: #" & ErrEx.Number & "/" & Err.Number _
                & " is Handled with resume next"
    MsgBox "Error: either division by zero, or type mismatch!"
    Resume Next

    ErrEx.Catch &H123                              ' you can set up a constant/enumeration if you prefer
    Debug.Print "Catch: #" & ErrEx.Number & "/" & Err.Number _
                & " is Handled with resume next"
    MsgBox "Error: caught error 291 (&H123)!"
    Resume Next
            
    ErrEx.CatchAll
    Debug.Print "Catch All: #" & ErrEx.Number & "/" & Err.Number _
                & " is Handled with resume next"
    MsgBox "Error (catch-all):" & vbCrLf & vbCrLf & Err.Description
    Resume Next
    
    ErrEx.Finally
    Debug.Print "Catch Finally: #" & ErrEx.Number & "/" & Err.Number _
                & " is Handled, then causes 1/0 "
    MsgBox "Finally"
    ' Debug.Print 1 / 0    ' Errors here automatically ignored due to implicit OnErrorResumeNext
    If Err.Number <> 0 Then
        msg = "Error # " & str(Err.Number) & " was generated by " _
                                           & Err.Source & Chr(13) & Err.Description
        Call MsgBox(msg, , "Error", Err.HelpFile, Err.HelpContext)
        Err.Clear
        Call ErrEx.DoFinally
    End If

    Debug.Print 1 / 0                              ' Simulate division by zero error (error 11)
    Debug.Print "A" / 1                            ' Simulate type mismatch error (error 13)
    Err.Raise &H123, , "This should be caught by the catch &H123 block..."
    Err.Raise &H456, , "This should be caught in the catch-all block..."
        
    ErrEx.Catch 11                                 ' you can set up a constant/enumeration if you prefer
    Debug.Print "Catch2: #" & ErrEx.Number & "/" & Err.Number _
                & " is Handled with resume next"
    MsgBox "Error: division by zero!"
    Resume Next

zExit:

End Sub                                            ' BugHelp.SimulateAnError

'---------------------------------------------------------------------------------------
' Method : StartEP
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Init when an external event happens.
'          This is a wrapper for the ProcCall, in that it uses Z_StartUp if we never inited.
'---------------------------------------------------------------------------------------
Sub StartEP(ErrClient As cErr, ClientKey As String, CallType As String, Optional Qmode As eQMode, Optional ExplainS As String)
    DoVerify Qmode = eQEPMode And CallType = tSubEP, _
             "EP's must be EP Application and CallType must be tSubEP"
    If Not DidStop Then
        If UseStartUp = 0 Then
            Call Z_StartUp                         ' detour for test purposes, continue
            UseStartUp = 1
        ElseIf UseStartUp < 2 Then
            DoVerify False
            UseStartUp = 0
            Call Z_StartUp                         ' call N_PreDefined test, skip
        End If
    End If
            
    If AppStartComplete And ItemsToDoCount + Deferred.Count > 0 Then
        Call FldActions2Do                         ' (must) have (at least 1) open items
    End If
    
    Call ProcCall(ErrClient, ClientKey, Qmode:=Qmode, CallType:=CallType, ExplainS:=ExplainS) ' the corresponding exit happens in ProcExit->ReturnEP
            
End Sub                                            ' BugHelp.StartEP

'---------------------------------------------------------------------------------------
' Method : Sub StartUp
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Entry from External Caller sets up Application Environment in Z_StartUp.
' Note   : To support debugging in BugHelp, it calls N_DebugStart. Happens only here.
'          If called more than once, it will reset the Application Environment
' Typical Use: when debugging and a re-run is intended.
'---------------------------------------------------------------------------------------
Sub StartUp(Optional doStop As Boolean = True, Optional Fast As Boolean = False, Optional LogAny As Boolean = True) ' *** Entry Point ***
    
    Call N_DeInit(False)
    FastMode = Fast
    LogZProcs = LogAny
    Call N_CheckStartSession(doStop)

End Sub                                            ' BugHelp.StartUp

'---------------------------------------------------------------------------------------
' Method : Sub TerminateApp
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub TerminateApp(Optional newEP As Variant)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "BugHelp.TerminateApp"
    Call DoCall(zKey, tSub, eQzMode)

    If T_DC Is Nothing Then
        DoVerify False, " if tester continues ..."
    Else
        Call T_DC.Terminate
    End If
    '                                 ... we will reach this
    If IsMissing(newEP) Then
        EPCalled = False
    Else
        If newEP = True Then
            EPCalled = newEP
        ElseIf newEP = False Then
            EPCalled = newEP
        End If
    End If

zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.TerminateApp

'---------------------------------------------------------------------------------------
' Method : Try
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Announce which errors are acceptable
'---------------------------------------------------------------------------------------
Sub Try(Optional WhatEver As Variant = "*")
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "BugHelp.Try"

Dim logVal As String

'------------------- gated Entry -------------------------------------------------------
    logVal = Left(WhatEver, 8)
    If logVal = "52" Then                          ' exclude log of Try in LogEvent
        logVal = vbNullString
    End If
    
    E_Active.Permit = WhatEver
    Call ErrReset(4)
   
End Sub                                            ' BugHelp.Try

'---------------------------------------------------------------------------------------
' Method : Z_CheckLinkage
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Check if the linkage from/to ProcDsc to ProcErr is intact
'---------------------------------------------------------------------------------------
Sub Z_CheckLinkage(ClientDsc As cProcItem, Optional zeroClientIsOk As Boolean, Optional Result As String)

Const zKey As String = "BugHelp.Z_CheckLinkage"
    Call DoCall(zKey, tSub, eQzMode)
        
    If StackDebug < 8 Then
        GoTo zExit
    End If
    
Dim stopHere As Boolean
Dim ErrClient As cErr

    stopHere = isEmpty(Result)
    
    If ClientDsc Is Nothing Then
        Result = "NoClient"
    Else
        Set ErrClient = ClientDsc.ErrActive
        If zeroClientIsOk Then
            If ErrClient Is Nothing Then
                Result = "NoErrClient"
            Else
                If LenB(ErrClient.atKey) = 0 Then
                    Result = "NoAtKeyErrClient"
                End If
            End If
        Else
            If Not ClientDsc Is ErrClient.atDsc Then
                Result = Result & " BadParentDsc"
            End If
        End If
    End If
    
    If LenB(Result) > 0 Then
        If stopHere Then
            DoVerify False
        End If
    End If
    Set ErrClient = Nothing
    
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_CheckLinkage

'---------------------------------------------------------------------------------------
' Method : Z_CheckLiveStack
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Find if Proc is Live
'---------------------------------------------------------------------------------------
Sub Z_CheckLiveStack(this As String, myFilter As Boolean, msg As String, Warn As Boolean)

Const zKey As String = "BugHelp.Z_CheckLiveStack"
    Call DoCall(zKey, tSub, eQzMode)

Dim LiveStack As Collection
Dim LiveKey As String
    
    Call N_GetLive(LiveStack, Filtered:=myFilter, Logging:=False)
    
    If LiveStack.Count > 0 Then
        LiveKey = LiveStack.Item(LiveStack.Count)
        If myFilter Then                           ' if not filtered, do some checking
            If this <> LiveKey Then
                Warn = True
                msg = msg & " - unchecked, found " & LiveKey
            End If
        Else
            If LiveStack.Count < 2 Then            ' there must be a caller and the active one!
                Warn = True
            End If
        End If
    Else
        If Not myFilter Then
            Warn = True
        End If
    End If
    
    Set LiveStack = Nothing
    
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_CheckLiveStack

'---------------------------------------------------------------------------------------
' Method : Z_EntryPoint
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Application-Level Entry Housekeeping
'---------------------------------------------------------------------------------------
Sub Z_EntryPoint(ClientDsc As cProcItem)

Const zKey As String = "BugHelp.Z_EntryPoint"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)
    
    aDebugState = E_AppErr.DebugState
    Call ShowStatusUpdate
       
    If ClientDsc.ErrActive.atRecursionLvl > 1 Then
        DoVerify False, " Applications should not be recursive "
        ' GoTo ProcReturn
    End If

    If ClientDsc.CallMode >= eQuMode Then          ' Application level definitions
        If ErrStatusFormUsable Then
            frmErrStatus.fCurrEP = ClientDsc.DbgId ' allow display of current Entry point
        End If
    End If
    
FuncExit:                                          ' ends always when S_AppIndex < 2 in Z_AppExit
    
ProcReturn:
    Call ProcExit(zErr)

End Sub                                            ' BugHelp.Z_EntryPoint

'---------------------------------------------------------------------------------------
' Method : Z_GetProcDsc
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Get ClientDsc from Key in D_ErrInterface and test on relevant CallMode
'---------------------------------------------------------------------------------------
Sub Z_GetProcDsc(Key As String, ClientDsc As cProcItem, msg As String)
'''' Proc Should ONLY CALL Z_Type PROCS                       ' trivial proc
Const zKey As String = "BugHelp.Z_GetProcDsc"
    Call DoCall(zKey, tSub, eQzMode)

Dim DbgId As String

    DbgId = Mid(Key, InStr(Key, ".") + 1)
       
    If D_ErrInterface.Exists(Key) Then             ' do search for it
        Set ClientDsc = D_ErrInterface.Item(Key)
        If ClientDsc.CallMode = eQnoDef Then       ' must be on D_ErrInterface!!!
            msg = msg & "dummy proc " & zKey _
                  & " M=" & ClientDsc.ModeLetter
makeDummy:
            aBugTxt = "assuming no Z_Type procs are off-stack: not needed any longer" ' ??? ??? ???
            DoVerify False
            msg = msg & "dummy: " & Key & vbCrLf
            Call DoCall(Key, tSub, eQyMode, ExplainS:=", as Dummy")
            Call DoExit(Key)
            GoTo zExit
        End If
    Else
        msg = msg & "name not defined yet: "
        GoTo makeDummy
    End If
        
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_GetProcDsc

'---------------------------------------------------------------------------------------
' Method : Z_InitBugHelp
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Initialize BugHelp
'---------------------------------------------------------------------------------------
Sub Z_InitBugHelp(ClientDsc As cProcItem)          ' ClientDsc is Out, not In

Const zKey As String = "BugHelp.Z_InitBugHelp"
    Call DoCall(zKey, tSub, eQzMode)

Dim AddMsg As String

Dim ErrClient As cErr
    
    ' Initialize Management Variables for BugHelp
    S_AppIndex = -1                                ' = D_AppStack.Count - 1
    
    Call N_SetErrExHdl
    
    S_ActKey = ExternCaller.Key
    
    If CallLogging Then
        AddMsg = "* Z_InitBugHelp  has been successfully completed: BugHelp is operational"
        If LogPerformance Then
            AddMsg = AddMsg & ", Performance Data are collected"
        End If
        Debug.Print AddMsg
    End If
    
    DidItAlready = True
    MayChangeErr = True

FuncExit:
    Set ErrClient = Nothing

zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_InitBugHelp

'---------------------------------------------------------------------------------------
' Method : Z_IsUnacceptable
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Returns True on non-acceptable errors.
'---------------------------------------------------------------------------------------
Function Z_IsUnacceptable(MatchAllow As Variant, Optional noPrint As Boolean) As Boolean
'''' Proc Should ONLY CALL Z_Type PROCS                       ' trivial proc
Const zKey As String = "BugHelp.Z_IsUnacceptable"
' do not Call DoCall(zKey, tFunction, eQzMode)

    With E_Active
        Z_IsUnacceptable = True                    ' with allowed exceptions, see OK
        
        If T_DC.DCerrNum = Hell Then
            T_DC.TermRQ = True
            Call TerminateApp
            GoTo zExit                             ' should never be reached
        End If
                
        ' if a message is passed instead of error numbers, compare that
        
        If Left(MatchAllow, 1) = "*" Then
            Z_IsUnacceptable = False               ' should not be reached
            GoTo SelfHandledError
        End If
        
        If IsNumeric(MatchAllow) Then
            noPrint = True                         ' do not print if acceptable
        Else
            If isEmpty(MatchAllow) Then
                GoTo IsErr                         ' no message to compare, disallowed
            ElseIf Left(Trim(MatchAllow), 1) = "-" Then ' Unacceptable message (complete)
                If InStr(.Description, Mid(MatchAllow, 2)) > 0 Then
                    GoTo IsErr
                End If
            ElseIf Left(Trim(MatchAllow), 1) = "+" Then ' acceptable error msg (complete)
                If InStr(.Description, Mid(MatchAllow, 2)) > 0 Then
                    noPrint = True
                    GoTo Allowed
                End If
            ElseIf Left(Trim(MatchAllow), 1) = "%" Then ' expected error msg (partial)
                If InStr(.Description, Mid(MatchAllow, 2)) > 0 Then
                    noPrint = True                 ' NO reset of error state
                    GoTo Allowed                   ' unless by optional parameter
                Else
                    GoTo IsErr
                End If
            Else
                If .Description = MatchAllow Then  ' exactly this message (complete)
                    Z_IsUnacceptable = False
                    noPrint = True                 ' do not print if acceptable
                    GoTo zExit
                End If
            End If
        End If
        
        If .errNumber = MatchAllow Then
Allowed:
            If LogAllErrors _
            Or ((DebugMode Or DebugLogging) And Not noPrint) Then
                If LogAllErrors Then
                    Call LogEvent("Error: " & .Description, eLall)
                Else
                    Debug.Print .Description
                End If
            End If
            If DebugMode Then
                ' Show the aNonModalForm with error information and
                ' allow user to ignore the error
                If DebugControlsWanted Then
                    Call ShowErrorStatus
                Else
                    DoVerify False
                End If
            End If
                
            Z_IsUnacceptable = False
            .FoundBadErrorNr = 0
            GoTo zExit
        Else
IsErr:
            If Z_IsUnacceptable Then
                If Not noPrint Then                ' Print more later
                    Debug.Print .Description
                End If
                If T_DC.DCAllowedMatch = 0 Then
                    T_DC.TermRQ = True
                End If
                If Err.Number > 0 Then
                    If (DebugMode Or DebugLogging) Then
                        If Err.Number > 0 Then
                            If noPrint And DebugMode Then
                                Debug.Print String(80, b) & vbCrLf _
                                    & "Error " & Err.Number _
                                    & ":" & Err.Description
                            End If
                        End If
                    End If
                Else
                    Z_IsUnacceptable = False
                    T_DC.TermRQ = False
                    GoTo zExit
                End If                             ' and no plans to check outside
            End If
            If Left(MatchAllow, 1) = "*" Then
SelfHandledError:
                If DebugMode And Not noPrint Then
                    Debug.Print String(80, b) & vbCrLf _
                        & "!!! Error in " & S_DbgId & b & .errNumber _
                        & " (&H" & Hex8(.errNumber) & "): " _
                                & .Description
                End If
                Z_IsUnacceptable = False           ' E_Active is not changed!
            End If
        End If
    End With                                       ' E_Active
    
zExit:
    ' do not Call DoExit(zKey)

End Function                                       ' BugHelp.Z_IsUnacceptable

'---------------------------------------------------------------------------------------
' Method : Z_LogApp_Exit
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Print Application level Call (do not Z_AppExit)
'---------------------------------------------------------------------------------------
Sub Z_LogApp_Exit(ClientDsc As cProcItem, ObjStr As String, moreEE As String)

Const zKey As String = "BugHelp.Z_LogApp_Exit"
    Call DoCall(zKey, tSub, eQzMode)
       
    Static Recursion As Boolean                    ' special: check disallowed Recursion

    If Recursion Then
        If ClientDsc.ErrActive.atCallState < eCpaused Then
            DoVerify False, "??? Recursion not needed if no hit"
        End If
        GoTo zExit                                 ' omit recursion on self (yesss, it is!!)
    End If
    
    Recursion = True

Dim noPrint As Boolean
Dim Caller As String
Dim addInfo As String
    
    If StackDebug <= 9 Then
        If StackDebug > 8 Then
            If ClientDsc.CallMode <= eQxMode Then
                GoTo testHidden
            End If
        End If
        If ClientDsc.CallMode = eQzMode Then       ' covers  .., Z_.., and O_Goodies, Classes
testHidden:
            If StackDebug > 4 Then
                GoTo FuncExit
            End If
        End If
    End If

    If Left(moreEE, 1) = "!" Then                  ' do not print
        noPrint = True
        addInfo = Mid(moreEE, 2)
    Else
        addInfo = moreEE
    End If
    
    Caller = ObjStr
           
GenOut:
    
    Call Z_Protocol(ClientDsc, CallNr, Caller, "<==", ClientDsc.DbgId, addInfo)
        
FuncExit:
    Recursion = False
    
zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_LogApp_Exit

'---------------------------------------------------------------------------------------
' Method : Z_Protocol
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Output an Application-level Protocol line
'---------------------------------------------------------------------------------------
Sub Z_Protocol(ClientDsc As cProcItem, CallNr As Long, Caller As String, Io As String, ObjInfo As String, addInfo As String)

Const zKey As String = "BugHelp.Z_Protocol"
    Call DoCall(zKey, tSub, eQzMode)
    
Dim Lvl As Long
Dim Mlen As Long
Dim CallS As String
    
    Lvl = Abs(ClientDsc.ErrActive.atCallDepth)
    CallS = Caller & Io & ObjInfo
    Mlen = Lvl + Len(CallS) + Len(Io) + 3
    
    If Mlen > lKeyM Then
        CallS = CallS & ">>" & Lvl
    Else
        CallS = LString(String(Lvl, b) & CallS, lKeyM)
    End If
    addInfo = Trim(addInfo & b & Replace(S_AppKey, dModuleWithP, vbNullString))
    Call N_ShowProcDsc(ClientDsc, CallNr, WithInstances:=False, ExplainS:=addInfo, Caller:=CallS)

zExit:
    Call DoExit(zKey)

End Sub                                            ' BugHelp.Z_Protocol

'---------------------------------------------------------------------------------------
' Method : N_SetErrExHdl
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Define UseErrExOn as default error handler for ErrExHandler
'---------------------------------------------------------------------------------------
Sub N_SetErrExHdl(Optional doPrint As Boolean = True)

Dim msg As String
    
    If LenB(UseErrExOn) = 0 Then
        If LenB(LastErrExOn) = 0 Then
            UseErrExOn = "N_OnError"               ' define the Proc that implements the global error events
        End If
    End If
    
    msg = "* ErrEx" & String(13, b) & RString("Global error handler", lDbgM) & b
    If ErrEx.IsEnabled Then
        ErrExConstructed = True
    End If
    If ErrExConstructed Then
        msg = msg & "was inactive, "
    Else
        msg = msg & "was already active, "
        ErrExConstructed = False
    End If
    
    If LenB(UseErrExOn) > 0 Then
        If LenB(LastErrExOn) = 0 Then
            If doPrint And StackDebug > 5 Then
                msg = msg & "unknown previous handler, "
                Debug.Print msg & "new global error handler '" & UseErrExOn & "'"
            End If
            Call ErrEx.Enable(UseErrExOn)
        Else
            If LastErrExOn = UseErrExOn Then
                If doPrint And StackDebug > 5 Then
                    msg = msg & " left unchanged '" & LastErrExOn & "'"
                End If
            Else
                If doPrint And StackDebug > 5 Then
                    msg = msg & " changing from '" & LastErrExOn & "'"
                End If
                Call ErrEx.Enable(UseErrExOn)
                If doPrint And StackDebug > 5 Then
                    Debug.Print msg & " to '" & UseErrExOn & "'"
                End If
            End If
        End If
        
        LastErrExOn = UseErrExOn
        
    ElseIf ErrExActive Then
        If CallLogging Then
            If LenB(LastErrExOn) = 0 Then
                Debug.Print msg & "remains disabled"
            Else
                Debug.Print msg & "disabled from '" & LastErrExOn & "'"
            End If
        End If
    
        Call ErrEx.Disable                         ' this sets ErrExActive = False, but keeps the LastErrExOn unchanged
        If ErrStatusFormUsable Then
            frmErrStatus.fErrAppl = vbNullString
        End If
    End If

ProcRet:
End Sub                                            ' BugHelp.N_SetErrExHdl

'---------------------------------------------------------------------------------------
' Method : Sub Z_SetupDialogs
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub Z_SetupDialogs()
'--- Proc MAY ONLY CALL Z_Type PROCS                            ' Simple proc
Const zKey As String = "BugHelp.Z_SetupDialogs"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    With ErrEx.DialogOptions
        If isEmpty(T_DC.DCAllowedMatch) Then
            .HTML_MainBody = "<font face=Arial size=13pt color=#4040DF><b>" _
                             & "A runtime error has occurred:</b></font><br><br><b>" _
                             & "<ERRDESC></b><br><br>Source:|<SOURCEPROJ>.<SOURCEMOD>." _
                             & "<SOURCEPROC><br>Filename:|<SOURCEFILENAME><br>Number:" _
                             & "|&H<ERRNUMBERHEX> (<ERRNUMBER>)<br>Source Line:|" _
                             & "<font bgcolor=#FFD8AF> #<SOURCELINENUMBER>.        <SOURCELINECODE>   </font><br>" _
                             & "<br>No Acceptable Error(s)"
        Else
            .HTML_MainBody = "<font face=Arial size=13pt color=#4040DF><b>" _
                             & "A runtime error has occurred:</b></font><br><br><b><ERRDESC>" _
                             & "</b><br><br>Source:|<SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC>" _
                             & "<br>Filename:|<SOURCEFILENAME><br>Number:|&H<ERRNUMBERHEX> " _
                             & "(<ERRNUMBER>)<br>Source Line:|<font bgcolor=#FFD8AF> " _
                             & "#<SOURCELINENUMBER>.        <SOURCELINECODE>   </font><br>" _
                             & "<br>Acceptable Error(s): " & T_DC.DCAllowedMatch
        End If
        .HTML_MainBody = .HTML_MainBody & "<br>Date/Time:|<ERRDATETIME><br>" _
                         & "<br><b><font size=12pt color=#4040DF>What do you want to do?</font></b>"
        .HTML_MoreInfoBody = "<br><b><font color=#40408F bgcolor=#C8D8FF>                                                     ?    VBA CALL STACK    ?                                                     </font></b><br><CALLSTACK>"
        .HTML_CallStackItem = "  <b><SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC></b>" _
                              & "<br> | #<SOURCELINENUMBER>.        <SOURCELINECODE>   <br>"
        .HTML_VariableItem = "(<VARSCOPE>)|<VARNAME> As <VARTYPE>| = <VARVALUE><br>"
        .WindowCaption = "YourApplicationName - runtime error"
        .MinimumWindowWidth = 600
        .MoreInfoCaption = "More info"
        .LessInfoCaption = "Less info"
        .ButtonPaddingH = 10
        .ButtonPaddingV = 5
        .ButtonSpacingH = 5
        .ButtonSpacingV = 7
        .PaddingH = 15
        .PaddingV = 15
        .ScreenBorderPaddingV = 50
        .ColumnPaddingH = 20
        .LineSpacingV = 2
        .MainBackColor = 16777215
        .MainBackColor2 = 13693168
        .MainBackFillType = 8
        .MoreInfoBackColor = 15724768
        .MoreInfoBackColor2 = 16774642
        .MoreInfoBackFillType = 0
        .ButtonBarBackColor = 16443364
        .ButtonBarBackColor2 = 14337988
        .ButtonBarBackFillType = 1
        .MaxNumCallStackItems = 10
        .MaxNumVariablesItems = inv
        .DefaultButtonID = 5
        .DefaultButtonIsBold = True
        .ShowMoreInfoButton = True
        .AllowEnterKey = True
        .AllowTabKey = True
        .AllowArrowKeys = False
        .Timeout = 0
        .FlashWindowOnOpen = True
        .CustomImageTransparentColor = -1
    Dim TempImageData As String
        TempImageData = vbNullString
        .CustomImageData = TempImageData
        Call .RemoveAllButtons
        Call .AddCustomButton("Search internet", "OnSearchInternet")
        Call .AddButton("Show Variables", BUTTONACTION_SHOWVARIABLES)
        Call .AddButton("Debug sourcecode", BUTTONACTION_ONERRORDEBUG)
        Call .AddButton("Ignore and continue", BUTTONACTION_ONERRORRESUMENEXT)
        Call .AddButton("Help", BUTTONACTION_SHOWHELP)
        Call .AddButton("Close", BUTTONACTION_ONERROREND)
    End With                                       ' ErrEx.DialogOptions

    With ErrEx.VariablesDialogOptions
        .HTML_MainBody = "<CALLSTACK>" & vbCrLf & "<Accept>"
        .HTML_MoreInfoBody = vbNullString
        .HTML_CallStackItem = "<b><SOURCEPROJ>.<SOURCEMOD>.<SOURCEPROC>" _
                              & "</b><br><br><VARIABLES><br>"
        .HTML_VariableItem = "   <font color=#808080>(<VARSCOPE>)</font>" _
                             & "|<VARNAME> As <VARTYPE>| = <VARVALUE><br>"
        .WindowCaption = "Microsoft Visual Basic"
        .MinimumWindowWidth = 600
        .MoreInfoCaption = "More info"
        .LessInfoCaption = "Less info"
        .ButtonPaddingH = 10
        .ButtonPaddingV = 5
        .ButtonSpacingH = 5
        .ButtonSpacingV = 7
        .PaddingH = 15
        .PaddingV = 15
        .ScreenBorderPaddingV = 50
        .ColumnPaddingH = 20
        .LineSpacingV = 2
        .MainBackColor = 16777215
        .MainBackColor2 = 13693168
        .MainBackFillType = 8
        .MoreInfoBackColor = 15724768
        .MoreInfoBackColor2 = 16774642
        .MoreInfoBackFillType = 0
        .ButtonBarBackColor = 16443364
        .ButtonBarBackColor2 = 14337988
        .ButtonBarBackFillType = 1
        .MaxNumCallStackItems = 10
        .MaxNumVariablesItems = inv
        .DefaultButtonID = 3
        .DefaultButtonIsBold = True
        .ShowMoreInfoButton = False
        .AllowEnterKey = True
        .AllowTabKey = True
        .AllowArrowKeys = False
        .Timeout = 0
        .FlashWindowOnOpen = True
        .CustomImageTransparentColor = -1
    Dim TempImageData2 As String
        TempImageData2 = vbNullString
        .CustomImageData = TempImageData2
        Call .RemoveAllButtons
        Call .AddButton("Close", BUTTONACTION_VARIABLES_CLOSE)
    End With                                       ' ErrEx.VariablesDialogOptions

ProcReturn:
    Call ProcExit(zErr)
End Sub                                            ' BugHelp.Z_SetupDialogs

'---------------------------------------------------------------------------------------
' Method : ShowLiveNameCount
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Count the defined names in the D_LiveStack Dictionary.
'          If Filter=True, eliminate the undefined ones
'---------------------------------------------------------------------------------------
Function ShowLiveNameCount(Optional Full As Boolean, Optional tSubFilter As Boolean, Optional doPrint As Boolean, Optional getNew As Boolean) As Long
    If Not ErrExActive Then
        GoTo ProcRet
    End If

Dim i As Long
Dim LCI As cCallEnv
Dim LiveStackCount As Long

    Call ShowLiveStack(doPrint:=doPrint, _
                       getNewStack:=getNew, Full:=Full)
    
    LiveStackCount = D_LiveStack.Count
    For i = LiveStackCount - 1 To 0 Step -1        ' print in reverse order
        Set LCI = D_LiveStack.Items(i)
        If LenB(LCI.CallerInfo) > 0 Then ' caller is not defined
            ShowLiveNameCount = ShowLiveNameCount + 1
        End If
    Next i
    
    If doPrint Then
        If ShowLiveNameCount = 0 Then
            Debug.Print "there are no relevant entries in D_LiveStack"
        Else
            Debug.Print " there are " & ShowLiveNameCount _
                        & " relevant entries in D_LiveStack, other=" _
                        & D_LiveStack.Count - ShowLiveNameCount
        End If
    End If
    
ProcRet:
End Function                                       ' BugHelp.ShowLiveNameCount

'---------------------------------------------------------------------------------------
' Method : Z_SourceAnalyse
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Test if a SourceLine contains a reference to sub or function
'---------------------------------------------------------------------------------------
Function Z_SourceAnalyse(ByVal sLine As String, ProcName As String) As Boolean
Const zKey As String = "BugHelp.Z_SourceAnalyse"
    Call DoCall(zKey, tFunction, eQzMode)
    
Dim i As Long
    
    sLine = Trim(Trunc(1, sLine, "'"))             ' drop comment part
    i = InStr(sLine, ProcName)                     ' is ref contained at all?
    If i > 1 Then
        If Mid(sLine, i - 1, 1) <> b Then          ' delimiters before=blank, (, or .
            If Mid(sLine, i - 1, 1) <> "(" Then
                If Mid(sLine, i - 1, 1) <> "." Then
                    i = 0                          ' none: not valid reference to proc
                End If
            End If
        End If
        sLine = Mid(sLine, i)                      ' if ref terminated by params?
        i = InStr(sLine, "(")
    End If
    
    Z_SourceAnalyse = (i > 0)

zExit:
    Call DoExit(zKey)

End Function                                       ' BugHelp.Z_SourceAnalyse

'---------------------------------------------------------------------------------------
' Method : N_CheckStartSession
' Author : rgbig
' Date   : 15.07.2020
' Purpose: Check if OlSession has been started and set up everything if not
'---------------------------------------------------------------------------------------
Sub N_CheckStartSession(Optional stopRQ As Boolean = True)
        
    If OlSession Is Nothing Then
        Debug.Assert False                             ' currently doing test, remove         Stop
        ' no need to Call N_DeInit
        Set OlSession = New cOutlookSession
        UseTestStart = UseTestStartDft                 ' using Default constant
        Call N_DebugStart(stopRQ)
        Call Z_StartUp(Not DidStop)
        Call BugTimer.BugState_UnPause
    End If

End Sub ' BugHelp.N_CheckStartSession

'---------------------------------------------------------------------------------------
' Method : Sub Z_StartUp
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Set or reset the main Application for Start
'---------------------------------------------------------------------------------------
Sub Z_StartUp(Optional doStop As Boolean = True)
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "BugHelp.Z_StartUp"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean
        
Dim zErr As cErr
Dim msg As String
    
    If Recursive Then
        ' choose Ignored or Forbidden and dependence on StackDebug
        If LogZProcs And Not P_Active Is Nothing Then
            Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                          & P_Active.DbgId & " => " & zKey
        End If
        GoTo ProcRet
    End If
    Recursive = True                                ' restored by    Recursive = False ProcRet:

    Set olApp = Outlook.Application

    If OlSession Is Nothing Then
        Set OlSession = New cOutlookSession
    End If
    
    Set OlExplorer = New cOlExplorer
    Call N_Prepare
    
    If aNameSpace Is Nothing _
    Or aRDOSession Is Nothing Then
        aBugTxt = "get Namespace/ActiveExplorer"
        Call Try(allowAll)
        Set aNameSpace = olApp.GetNamespace("MAPI")
        Catch
        Set ActiveExplorerItem(1) = ActiveExplorer
        Catch
        aBugTxt = "get RDOSession"
        Call Try(allowNew)
        Set aRDOSession = CreateObject("Redemption.RDOSession")
        Catch
        
        aBugTxt = "RDO-Logon"
        Call Try(allowNew)
        aRDOSession.Logon                           'no param == an empty string to use the default MAPI profile
        Catch
    End If
    '---------------- Start of normal Entry ------------------------------------------------
    Call DoCall(zKey, "Sub", eQzMode, Z§AppStart)   ' Z_Startup replaces ThisOutlookSession.Application_Startup
    
    Set zErr = Z§AppStart.ErrActive                 ' Z§AppStart has been replaced...
    Z§AppStart.CallCounter = 1                      ' ... always called only once
    zErr.EventBlock = True                          ' no Events until we leave Setups
    zErr.atCallState = eCActive
    Set zErr.atCalledBy = ExternCaller.ErrActive
    Set E_Active = zErr
    Set P_EntryPoint = Z§AppStart
    
    If ExternalEntryCount > 0 Then
        If UseStartUp = 2 Then
            doStop = False
        End If
    End If
    If UseStartUp = 0 Then
        ProtectStackState = inv
        UseStartUp = 2                              ' continue open items after inits (in StartMainApp)
    ElseIf UseStartUp = 1 Then
        Debug.Print "---------- Z_StartUp recognized a re-start --------------"
        If doStop And Not DidStop Then
            DoVerify False, "Application Restart programmed stop"
        End If
    ElseIf UseStartUp = 2 Then
        UseStartUp = 1
    End If
    
    NoPrintLog = False
    DebugMode = False
    LogImmediate = True
    
    Call Z_InitBugHelp(Z§ProcStart)
           
    If TopFolders Is Nothing Then                  ' Inits for Outlook Item Classes
        Call Z_olInits
    End If
  
    Call N_ShowProgress(CallNr, Z§AppStart, _
                        Z§AppStart.Key, "Ready for Application", vbNullString)
    Set E_AppErr = Z§AppStart.ErrActive
    
    If LenB(UseErrExOn) > 0 Then
        WithLiveCheck = True
    End If

    If UseStartUp = 0 Then                     ' directly called from an entry point
        UseStartUp = 1                         ' resume right there
    ElseIf UseStartUp = 2 Then                 ' use the EP under Test
        Call OlSession.StartMainApp            ' this is any App under test
    End If

    Call getDebugMode(ExternalEntryCount = 1)  ' fetch from profile on first call

    If StackDebug > 4 Then
        Set aNonModalForm = frmErrStatus       ' must NOT use New!
        Call ShowErrorStatus
    End If

    If ErrStatusFormUsable Then
        frmErrStatus.fNoEvents = E_Active.EventBlock
        Call Start_BugTimer
    End If

    Debug.Print String(OffCal, b) & _
                                  LString("Error Module 'BugHelp': Initiation completed by " _
                                          & Application.Name, lKeyM) _
                                          & "#P=" & LString(0, 5) & ExternCaller.Key

    Set zErr = Nothing

    Z§ShowErrorStatus.CallCounter = 1          ' force Position of window
    Call ShowErrorStatus
    AppStartComplete = True
    Call FldActions2Do                         ' open items

zExit:
    Call DoExit(zKey)
    If E_Active.atCallDepth < 1 Then           ' 0 is ExternCaller
        Call BugTimerDeActivate
        Debug.Print "Timer DeActivated"
        Call N_ShowHeader("BugHelp Log " & TimerNow)
        msg = "* " & String(OffCal - 2, "-") _
            & " Outlook waiting for Events or Macro Calls " _
            & String(OffCal - 2, "-") & " *"
        Call LogEvent(String(Len(msg), "-") _
                      & vbCrLf & msg & vbCrLf _
                      & String(Len(msg), "-"), eLSome)
        Z§AppStart.ErrActive.EventBlock = False
    End If
    
    E_AppErr.EventBlock = False
    Call SetOnline(olCachedConnectedFull)
ProcRet:
End Sub                                        ' BugHelp.Z_StartUp

'---------------------------------------------------------------------------------------
    ' Method : Z_StateToTestVar
    ' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
    ' Purpose:
'---------------------------------------------------------------------------------------
Sub Z_StateToTestVar()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "BugHelp.Z_StateToTestVar"
    Call DoCall(zKey, "Sub", eQzMode)

    If InStr(1, Testvar, "OFF", vbTextCompare) > 0 Then
        DebugMode = False
        StackDebug = 0
        TraceMode = False
        LogAppStack = False
        LogPerformance = False
        Testvar = Trim("OFF")
    Else
        Testvar = vbNullString                     ' build new from State
        If LogAllErrors Then
            Testvar = "ERR " & Testvar
        End If
        If DebugLogging Then
            Testvar = "LOG " & Testvar
        End If
        Testvar = "StackDebug=" & StackDebug & b & Testvar
        If LogPerformance Then
            Testvar = "LogPerformance " & Testvar
        End If
        If ShowFunctionValues Then
            Testvar = "ShowFunctionValues " & Testvar
        End If
        If TraceMode Then
            Testvar = "TraceMode " & Testvar
        End If
        If DebugMode Then
            Testvar = "DebugMode " & Testvar
        End If
        
        Testvar = Trim(Testvar)
    End If

zExit:
    Call DoExit(zKey)
End Sub                                            ' BugHelp.Z_StateToTestVar

'---------------------------------------------------------------------------------------
' Method : Z_UsedThisCall
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Compute Performance Data, getting Z_UsedThisCall
'---------------------------------------------------------------------------------------
Function Z_UsedThisCall(ErrClient As cErr, TimeNow As Double) As Double
Static Recursive As Boolean

'------------------- gated Entry -------------------------------------------------------
    Z_UsedThisCall = 0
    
    If Recursive Then                              ' no Message, because recursive happens but does not matter
        ' Do not set message: Debug.Print String(OffCal, b) & "Forbidden recursion from ..."
        GoTo ProcRet
    End If

    If ErrClient.atProcIndex = inv Then
        GoTo ProcRet                               ' before it is defined, proc can't have UsedThisCall
    End If
    If ErrClient.atLastInSec = 0& Then             ' Proc without time, e.g. ProcExit
        GoTo ProcRet
    End If
    If ErrClient.atThisEntrySec = 0& Then
        GoTo ProcRet                               ' unnecessary to compute TimeUsed when exited
    End If

    Recursive = True
    
Dim DiffDate As Double
    
    With ErrClient
        DiffDate = DateDiff("d", Date, .atLastInDate) * 86400#
        Z_UsedThisCall = TimeNow - .atThisEntrySec + DiffDate
        If .atDsc Is Nothing Then
            Z_UsedThisCall = TimeNow - .atThisEntrySec
        Else
            With .atDsc
                .TotalProcTime = .TotalProcTime + Z_UsedThisCall
            End With                               ' .atDsc
        End If
        .atPrevEntrySec = .atThisEntrySec
        .atThisEntrySec = 0
    End With                                       ' ErrClient

FuncExit:
    Recursive = False
ProcRet:
End Function                                       ' BugHelp.Z_UsedThisCall


