VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmErrStatus 
   Caption         =   "Error and Debug Status"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   14460
   ClientWidth     =   7350
   OleObjectBlob   =   "frmErrStatus.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmErrStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmIgnoreErrStatusChange As Boolean       ' to avoid recursion if EventRoutine fModifications_Change

'---------------------------------------------------------------------------------------
' Method : Sub Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub Activate()

Const zKey As String = "frmErrStatus.Activate"
Static zDsc As cProcItem

    '------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    aBugVer = zDsc.CallCounter < 2
    If DoVerify(aBugVer, "frmErrStatus must never have 2 Window Instances") Then
        GoTo ProcRet
    End If

    If Recursive Then
        Debug.Print String(OffCal, b) & "Forbidden recursion from " _
                                    & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    Recursive = True
    
    Call UserForm_Activate
    
    Show
    Call WindowSetForeground(Me.Caption, Nothing)
    
    Recursive = False

ProcRet:
End Sub                                          ' frmErrStatus.Activate

'---------------------------------------------------------------------------------------
' Method : ReEvaluate
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: When in debug window, evaluate state of Settings, or if Reverse: global values
'---------------------------------------------------------------------------------------
Public Sub ReEvaluate(Optional Reverse As Boolean = False)
Const zKey As String = "frmErrStatus.ReEvaluate"
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Recursive Then
        GoTo ProcRet
    End If
    Recursive = True
    
    If E_AppErr Is Nothing Then
        fLastErrReasoning = vbNullString
    ElseIf E_AppErr.Reasoning <> vbNullString Then
        fLastErrReasoning = E_AppErr.Reasoning
    End If
    
    If ItemsToDoCount + Deferred.Count <> 0 Then
        fDeferredCount = ItemsToDoCount
        fCurFolder = curFolderPath
    Else
        fDeferredCount = vbNullString
    End If
     
    frmIgnoreErrStatusChange = False            ' not restored on exit!
    Call N_Suppress(Push, zKey)                 ' ShutupMode on
    ChangeAssignReverse = Reverse
    SuppressStatusFormUpdate = True
    fModifications.Visible = True
    If fModifications Then                      ' reset fModifications
        fModifications.Enabled = False          ' without  action here
        fModifications = False
    End If
    fModifications.Enabled = True               ' with actions now
    fModifications = True                       ' will Call fModifications_Change
    
    fLastErrExplanations = "Werte für Debugoptionen aus frmErrStatus ausgelesen " & Time
    ChangeAssignReverse = False
    Call N_Suppress(Pop, zKey)                  ' ShutupMode restored
    Recursive = False

ProcRet:
End Sub                                          ' frmErrStatus.ReEvaluate

'---------------------------------------------------------------------------------------
' Method : Sub UpdInfo
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub UpdInfo()
'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    If Not Visible Then                          ' invisible windows can't come to foreground
        GoTo ProcRet
    End If
    If Recursive Then
        GoTo ProcRet
    End If
    Recursive = True

    If ErrDisplayModify And Not fHideMe Then     ' should frmErrStatus show (with modif. values)?
        On Error GoTo ignoreAll
        Call WindowSetForeground(Me.Caption, Nothing)
ignoreAll:
        ErrDisplayModify = False
    End If
    
FuncExit:
    Recursive = False
ProcRet:
End Sub                                          ' frmErrStatus.UpdInfo

'---------------------------------------------------------------------------------------
' Method : Sub Form_Load
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Form_Load()

Const zKey As String = "frmErrStatus.Form_Load"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)
    
    Call getDebugMode                            ' don't force: using debug as/if defined
    
    If LenB(Testvar) = 0 Then
        fToggleDebug.Caption = "Toggle Debugmode:=ON"
    Else
        fToggleDebug.Caption = "Debuging Options: " & Quote(Testvar)
    End If
    If DebugMode Then
        fToggleDebug.BackColor = &H80FFFF
    Else
        fToggleDebug.BackColor = &H8000000F
    End If
    Repaint
    doMyEvents

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmErrStatus.Form_Load

'---------------------------------------------------------------------------------------
' Method : Sub fAllOfThese_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fAllOfThese_Change()

    If frmIgnoreErrStatusChange Then
        GoTo ProcRet
    End If
    
FuncExit:
    frmIgnoreErrStatusChange = False
    fModifications = True                        ' raises event on change
    
ProcRet:
End Sub                                          ' frmErrStatus.fAllOfThese_Change

'---------------------------------------------------------------------------------------
' Method : Sub fBreak_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fBreak_Click()

    Debug.Print "Break Key pressed in frmErrStatus"
    Call BugTimerDeActivate
    Debug.Assert False                           ' always as trivial function
    Call BugTimerActivate(0)
    
End Sub                                          ' frmErrStatus.fBreak_Click

'---------------------------------------------------------------------------------------
' Method : Sub fCallLogging_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fCallLogging_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fCallLogging_Click

'---------------------------------------------------------------------------------------
' Method : Sub fLogZProcs_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fLogZProcs_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fLogZProcs_Click

'---------------------------------------------------------------------------------------
' Method : Sub fLImmediate_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fLImmediate_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fLogZProcs_Click

'---------------------------------------------------------------------------------------
' Method : Sub fCancelTermination_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fCancelTermination_Click()

Const zKey As String = "frmErrStatus.fCancelTermination_Click"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

    Call N_ErrClear
    
    E_AppErr.errNumber = vbObjectError + 101     ' user defined this to be an acceptable error
    E_AppErr.FoundBadErrorNr = 0

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmErrStatus.fCancelTermination_Click

'---------------------------------------------------------------------------------------
' Method : Sub fEditWatchState_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fEditWatchState_Click()

Const zKey As String = "frmErrStatus.fEditWatchState_Click"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

    If Not MayChangeErr Then
        rsp = MsgBox("editing of watch state is not allowed during N_PublishBugState. " _
                   & vbCrLf & "You can OK to override and set the Test variable now, " _
                   & "but this will cause an error on Resume from N_PublishBugState.", _
                     vbOKCancel)
        If rsp = vbCancel Then
            GoTo ProcReturn                      ' user cancelled the Edit
        End If
    End If

    Call EditWatch
    
ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmErrStatus.fEditWatchState_Click

'---------------------------------------------------------------------------------------
' Method : Sub fDebugLogging_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fDebugLogging_Change()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fDebugLogging_Change

'---------------------------------------------------------------------------------------
' Method : Sub fHideMe_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fHideMe_Click()
    fModifications = True
    If fHideMe Then
        Me.Hide
    Else
        Me.Show vbModeless
    End If
End Sub                                          ' frmErrStatus.fHideMe_Click

'---------------------------------------------------------------------------------------
' Method : Sub fLogAllErrors_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fLogAllErrors_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fLogAllErrors_Click

'---------------------------------------------------------------------------------------
' Method : Sub fLogPerformance_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fLogPerformance_Change()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fLogPerformance_Change

'---------------------------------------------------------------------------------------
' Method : Sub fNoTimerEvent_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fNoTimerEvent_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub

Private Sub fOnline_Click()
    If fOnline.Caption = "Offline" Then
        Call SetOnline(olCachedConnectedFull)
        fOnline.Caption = "Online"
    Else
        Call SetOffline
        fOnline.Caption = "Offline"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : Sub fShowFunVal_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fShowFunVal_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fShowFunVal_Click

'---------------------------------------------------------------------------------------
' Method : Sub fShowLog_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fShowLog_Click()
    Call ShowLog
End Sub                                          ' frmErrStatus.fShowLog_Click

'---------------------------------------------------------------------------------------
' Method : Sub fExLiveCheck_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fExLiveCheck_Click()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fExLiveCheck_Click

'---------------------------------------------------------------------------------------
' Method : Sub fTraceMode_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fTraceMode_Change()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fTraceMode_Change

'---------------------------------------------------------------------------------------
' Method : Sub fDebugMode_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fDebugMode_Change()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True
    End If
End Sub                                          ' frmErrStatus.fDebugMode_Change

'---------------------------------------------------------------------------------------
' Method : Sub fStackDebug_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fStackDebug_AfterUpdate()
    If Not frmIgnoreErrStatusChange Then
        If Not IsNumeric(fStackDebug) Then       ' ignore if not numeric
            fStackDebug = StackDebug
            fIncDecDebug = StackDebug
        Else
            fIncDecDebug = fStackDebug
            If fAllOfThese Then
                Call N_DebugStart
            End If
            fModifications = True                ' Raise that _Change
        End If
    End If
End Sub                                          ' frmErrStatus.fStackDebug_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub fModifications_Change
' Author : rgbig
' Date   : 20211109@17_00
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fModifications_Change()
Const zKey As String = "frmErrStatus.fModifications_Change"

    '------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean
Dim ModString As String
Dim chgMode As Boolean
Dim chgMsg As String

    If frmIgnoreErrStatusChange Or Not fModifications.Enabled Then    ' also protects against recursion
        GoTo ProcRet
    End If

    If Recursive Then
        Debug.Print String(OffCal, b) & "Ignored recursion from " _
                                        & P_Active.DbgId & " => " & zKey
        GoTo ProcRet
    End If
    
    fModifications.Enabled = False              ' active in Change: do not trigger
    Recursive = True                            ' restored by    Recursive = False ProcRet:
    Call BugTimer.BugState_SetPause

    If fIncDecDebug = 0 Then                    ' consistency forced
        fIncDecDebug = StackDebug               ' not gated by frmIgnoreErrStatusChange
    End If
    If fModifications Then
        LogImmediate = fLImmediate              ' form always wins
        If fIncDecDebug = 0 Then
            fIncDecDebug = StackDebug           ' not gated by frmIgnoreErrStatusChange
        Else
            StackDebug = fIncDecDebug
        End If
        If CLng(fIncDecDebug) <> StackDebug Then
            fStackDebug = CStr(fIncDecDebug.Value)
        End If
        If AssignIfChanged(StackDebug, fStackDebug) Then
            If ChangeAssignReverse Then
                StackDebug = ModThisTo
                fIncDecDebug = StackDebug
            Else
                fStackDebug = ModThisTo
                fIncDecDebug.Value = CLng(0 & fStackDebug.Value)
            End If
            StackDebugOverride = -StackDebug
            ModString = vbCrLf & "StackDebug = " & StackDebug
        Else
            StackDebugOverride = StackDebug     ' New setting affecting ErrReset
        End If
        If AssignIfChanged(DebugMode, fDebugMode) Then
            If ChangeAssignReverse Then
                fDebugMode = ModThisTo
            Else
                DebugMode = ModThisTo
            End If
            ModString = ModString & vbCrLf & "DebugMode = " & CStr(DebugMode)
        End If
        If AssignIfChanged(DebugLogging, fDebugLogging) Then
            If ChangeAssignReverse Then
                fDebugLogging = ModThisTo
            Else
                DebugLogging = ModThisTo
            End If
            ModString = ModString & vbCrLf & "DebugLogging = " & CStr(DebugLogging)
        End If
        If AssignIfChanged(LogAllErrors, fLogAllErrors) Then
            If ChangeAssignReverse Then
                fLogAllErrors = ModThisTo
            Else
                LogAllErrors = ModThisTo
            End If
            ModString = ModString & vbCrLf & "LogAllErrors = " & CStr(LogAllErrors)
        End If
        If fNoTimerEvent Then
            If Not NoTimerEvent Then
                Call BugTimerDeActivate         ' turn off Timer
                NoTimerEvent = True
                ModString = ModString & vbCrLf & "NoTimerEvent = " & CStr(NoTimerEvent)
            End If
        Else                                    ' want timer if possible
            If NoTimerEvent Then                ' check if possible here
                If ErrTimerEventNotReady Then   ' no it isn't possible
                    NoTimerEvent = False
                    ModString = ModString & vbCrLf & "NoTimerEvent = " & CStr(NoTimerEvent)
                End If
            Else                                ' timer off already
                If DebugMode Then
                    DoVerify ErrTimerEventNotReady, "*** Timer should be off"
                End If
            End If
        End If
        If AssignIfChanged(LogPerformance, fLogPerformance) Then
            If ChangeAssignReverse Then
                fLogPerformance = ModThisTo
            Else
                LogPerformance = ModThisTo
            End If
            ModString = ModString & vbCrLf & "LogPerformance = " & CStr(LogPerformance)
        End If
        If AssignIfChanged(TraceMode, fTraceMode) Then
            If ChangeAssignReverse Then
                fTraceMode = ModThisTo
            Else
                TraceMode = ModThisTo
            End If
            ModString = ModString & vbCrLf & "TraceMode = " & CStr(TraceMode)
        End If
        If AssignIfChanged(ExLiveDscGen, fExLiveCheck) Then
            If ChangeAssignReverse Then
                fExLiveCheck = ModThisTo
            Else
                ExLiveDscGen = ModThisTo
            End If
            ModString = ModString & vbCrLf & "ExLiveDscGen = " & CStr(ExLiveDscGen)
        End If
        If AssignIfChanged(ShowFunctionValues, fShowFunVal) Then
            If ChangeAssignReverse Then
                fShowFunVal = ModThisTo
            Else
                ShowFunctionValues = ModThisTo
            End If
            ModString = ModString & vbCrLf & "ShowFunctionValues = " & CStr(ShowFunctionValues)
        End If
        If AssignIfChanged(CallLogging, fCallLogging) Then
            If ChangeAssignReverse Then
                fCallLogging = ModThisTo
            Else
                CallLogging = ModThisTo
            End If
            ModString = ModString & vbCrLf & "CallLogging = " & CStr(CallLogging)
        End If
        If AssignIfChanged(LogZProcs, fLogZProcs) Then
            If ChangeAssignReverse Then
                fLogZProcs = ModThisTo
            Else
                LogZProcs = ModThisTo
            End If
            ModString = ModString & vbCrLf & "LogZProcs = " & CStr(LogZProcs)
        End If
        If AssignIfChanged(E_Active.EventBlock, fNoEvents) Then
            If ChangeAssignReverse Then
                fNoEvents = ModThisTo
            Else
                E_Active.EventBlock = ModThisTo
            End If
            If Not E_Active.EventBlock Then ' no need to block Events -> online
                chgMode = SetOnline(olCachedConnectedFull)
                If chgMode Then
                    chgMsg = " changed to Online"
                Else
                    chgMsg = " was Online"
                End If
                If DoVerify(fOnline = actOnlineStatus, _
                    "*** Eventblock: " & E_Active.EventBlock _
                    & " showing " & fOnline & " actual OnlineStatus is " _
                    & actOnlineStatus & " correcting caption!" _
                    & chgMsg) Then
                    fOnline.Caption = "Online"
                End If
            Else                        ' need to block Events -> offline
                chgMode = SetOffline()
                If chgMode Then
                    chgMsg = " changed to Offline"
                Else
                    chgMsg = " was Offline"
                End If
                If DoVerify(fOnline = actOnlineStatus, _
                    "*** Eventblock: " & E_Active.EventBlock _
                    & " showing " & fOnline & " but OnlineStatus is " _
                    & actOnlineStatus & " correcting caption!" _
                    & chgMsg) Then
                    fOnline.Caption = "Offline"
                End If
            End If
            If chgMode And DebugMode Then
                Call LogEvent("Outlook changed to " _
                            & IIf(E_Active.EventBlock, "Offline", "Online"))
            End If
            ModString = ModString & vbCrLf & "EventBlock(" & E_Active.atKey _
                      & ") = " & CStr(E_Active.EventBlock)
        End If
        If AssignIfChanged(T_DC.DCerrSource, fLastErrSource) Then
            If ChangeAssignReverse Then
                fLastErrSource = ModThisTo
            Else
                T_DC.DCerrSource = ModThisTo
            End If
            ' not a normal change, dont log ModString = ModString & vbCrLf & "LastErrSource = " & T_DC.DCerrSource
        End If
        If AssignIfChanged(UseErrExOn, fUseErrExOn) Then
            If ChangeAssignReverse Then
                fUseErrExOn = ModThisTo
            Else
                UseErrExOn = ModThisTo
            End If
            fUseErrExOn = ModThisTo             ' special logic! NOT changing UseErrExOn !!!
            Call N_SetErrExHdl(doPrint:=False)
            ModString = ModString & vbCrLf & "UseErrExOn = " & fUseErrExOn & b & " ErrExActive=" & ErrExActive
        End If
        If AssignIfChanged(T_DC.DCUseErrExOn, fUseErrExOn) Then ' special logic! Irrelevant if reversed
            If LenB(ModThisTo) = 0 Or LenB(UseErrExOn) = 0 Then ' LastErrExOn is never changed
                fUseErrExOn = vbNullString
                fErrorHandler = "Not using Global Error Handler"
                fUseErrExOn.BackColor = 188
                fUseErrExOn.BackStyle = fmBackStyleOpaque
            Else
                fUseErrExOn = UseErrExOn
                fErrorHandler = "Module to use for " & vbCrLf & "  Global Error Handler"
                fUseErrExOn.BackColor = 255
                fUseErrExOn.BackStyle = fmBackStyleTransparent
            End If
            T_DC.DCUseErrExOn = fUseErrExOn
            UseErrExOn = fUseErrExOn
            ModString = ModString & vbCrLf & "DCUseErrExOn = " & fUseErrExOn
        End If
        If AssignIfChanged(T_DC.DCAppl, fErrAppl) Then ' special logic!
            If LenB(ModThisTo) = 0 Then
                ModThisTo = S_AppKey
                fErrAppl = ModThisTo
                T_DC.DCAppl = ModThisTo
            Else
                If ChangeAssignReverse Then
                    fErrAppl = ModThisTo
                Else
                    T_DC.DCAppl = ModThisTo
                End If
            End If
            If DebugMode And DebugLogging Then   ' normally no log because normal change
                ModString = ModString & vbCrLf & "DCAppAct = " & ModThisTo
            End If
        End If
        If ChangeAssignReverse Then              ' limited asssignments: not editable
            If isEmpty(T_DC.DCAllowedMatch) Then ' T_DC.DCAllowedMatch is not editable string
                If fAcceptableErrors <> "No Acceptable Errors" Then
unset:
                    fAcceptableErrors = "No Acceptable Errors"
                    fAcceptableErrors.Enabled = False
                    If DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then
                        ErrDisplayModify = True
                    End If
                End If
                fAcceptableErrors.SpecialEffect = fmSpecialEffectFlat
            ElseIf LenB(T_DC.DCAllowedMatch) = 0 Then
                GoTo unset
            Else
                If fAcceptableErrors <> "Accept Errors=" & CStr(T_DC.DCAllowedMatch) Then
                    If DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then
                        ErrDisplayModify = True
                    End If
                    fAcceptableErrors = "Accept Errors=" & CStr(T_DC.DCAllowedMatch)
                End If
                fAcceptableErrors.SpecialEffect = fmSpecialEffectBump
            End If
            If fLastErrExplanations <> E_Active.Explanations Then ' E_Active.Explanations is not editable string
                If LenB(E_Active.Explanations) > 2 Then
                    fLastErrExplanations = E_Active.Explanations
                Else
                    fLastErrExplanations = vbNullString
                End If
                If DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then
                    ErrDisplayModify = True
                End If
            End If
            If fLastErrReasoning <> E_Active.Reasoning Then ' E_Active.Reasoning is not editable string
                If LenB(E_Active.Reasoning) > 2 Then
                    fLastErrReasoning = E_Active.Reasoning
                Else
                    fLastErrReasoning = vbNullString
                End If
                If DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then
                    ErrDisplayModify = True
                End If
            End If
            If T_DC.DCerrNum = 0 Then            ' T_DC.DCerrNum is not editable
                If LenB(fLastErr) > 0 Then
                    fLastErr = vbNullString
                    fLastErrHex = vbNullString
                    If DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then
                        ErrDisplayModify = True
                    End If
                End If
            ElseIf fLastErr <> CStr(T_DC.DCerrNum) Then
                fErrAppl = T_DC.DCAppl
                fErrNumber = T_DC.DCerrNum
                fLastErr = CStr(T_DC.DCerrNum)
                fLastErrHex = Hex8(T_DC.DCerrNum)
                ErrDisplayModify = True
            End If
        End If
        
        ModString = Trim(ModString)
        If LenB(ModString) > 0 And Not frmIgnoreErrStatusChange Then
            Debug.Print "* Settings changed: " & ModString
            ErrDisplayModify = True
            Call Z_StateToTestVar
        End If
    End If
    
    fModifications = False                      ' causes _Change recursion, but immediately exits
    Recursive = False
    Call BugTimer.BugState_UnPause
    frmIgnoreErrStatusChange = False
    fModifications.Enabled = True               ' enable the change trigger

ProcRet:
End Sub                                          ' frmErrStatusfModifications_Change

'---------------------------------------------------------------------------------------
' Method : Sub fNoEvents_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fNoEvents_Click()
    If Not frmIgnoreErrStatusChange Then
        E_Active.EventBlock = fNoEvents
        If E_AppErr.EventBlock <> fNoEvents Then
            E_AppErr.EventBlock = fNoEvents
            DoVerify True, "*** changed EventBlock to " _
                & E_AppErr.EventBlock & " for " & E_AppErr.atKey
        End If
        fModifications = True                    ' Raise that _Change
    End If
End Sub                                          ' frmErrStatus.fNoEvents_Click

'---------------------------------------------------------------------------------------
' Method : Sub fTerminationFlag_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fTerminationFlag_Click()
    Me.Hide                                      ' not gated by frmIgnoreErrStatusChange
    T_DC.TermRQ = True
End Sub                                          ' frmErrStatus.fTerminationFlag_Click

'---------------------------------------------------------------------------------------
' Method : Sub fToggleDebug_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fToggleDebug_Click()
    Call SetDebugMode(noMsg:=True)               ' not gated by frmIgnoreErrStatusChange
End Sub                                          ' frmErrStatus.fToggleDebug_Click

'---------------------------------------------------------------------------------------
' Method : Sub fUseErrExOn_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fUseErrExOn_AfterUpdate()
    If Not frmIgnoreErrStatusChange Then
        fModifications = True                    ' Raise that _Change
    End If
End Sub                                          ' frmErrStatus.fUseErrExOn_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub fIncDecDebug_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub fIncDecDebug_Change()
    If fStackDebug <> fIncDecDebug Then
        fStackDebug = fIncDecDebug               ' not gated by frmIgnoreErrStatusChange
        fModifications = True                    ' Raise that _Change
    End If
End Sub                                          ' frmErrStatus.fIncDecDebug_Change

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    Call QueryErrStatusChange(False)             ' .f values are changeable
End Sub                                          ' frmErrStatus.UserForm_Activate


'---------------------------------------------------------------------------------------
' Method : Sub UserForm_DblClick
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ReEvaluate
    ErrDisplayModify = True
End Sub                                          ' frmErrStatus.UserForm_DblClick

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Deactivate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Deactivate()
    If fModifications Then
        Call getDebugMode
        fModifications = False
    End If
    fLastErrIndications.Enabled = True
End Sub                                          ' frmErrStatus.UserForm_Deactivate

Private Sub UserForm_Initialize()
'--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
Const zKey As String = "frmErrStatus.UserForm_Initialize"
Dim zErr As cErr
    
    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

    frmIgnoreErrStatusChange = True
    fAcceptableErrors = "Acceptable Errors"
    fAcceptableErrors.Enabled = False
    fAcceptableErrors.SpecialEffect = fmSpecialEffectFlat
    fCallLogging = CallLogging
    fDebugLogging = DebugLogging
    fDebugMode = DebugMode
    fUseErrExOn = LastErrExOn
    If LenB(T_DC.DCUseErrExOn) = 0 Then
        fUseErrExOn = LastErrExOn
    Else
        fUseErrExOn = T_DC.DCUseErrExOn
    End If
    fErrNumber = T_DC.DCerrNum
    fLastErr = CStr(T_DC.DCerrNum)
    fLastErrExplanations = E_Active.Explanations
    fLastErrHex = Hex8(T_DC.DCerrNum)
    fLastErrMsg = T_DC.DCerrMsg
    fLastErrReasoning = E_Active.Reasoning
    fLastErrSource = T_DC.DCerrSource
    fLImmediate = True
    fLogAllErrors = LogAllErrors
    fLogPerformance = LogPerformance
    fLogZProcs = LogZProcs
    fNoEvents = E_Active.EventBlock
    fNoTimerEvent = NoTimerEvent
    fShowFunVal = ShowFunctionValues
    fExLiveCheck = ExLiveDscGen
    fStackDebug = StackDebug
    fTraceMode = TraceMode
    frmErrStatus.Top = 245
    frmErrStatus.Left = 1041
       
    Call QueryErrStatusChange(False)             ' .f values are changable
    If Visible Then
        Call WindowSetForeground(Me.Caption, Nothing)
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' frmErrStatus.UserForm_Initialize

'---------------------------------------------------------------------------------------
' Method : Sub fBeginTermination
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Ask User about Termination request (may abort Termination)
'---------------------------------------------------------------------------------------
Public Sub fBeginTermination(Optional fTerminationFlag As Boolean = True)
    T_DC.TermRQ = fTerminationFlag
    If DebugMode Then
        Me.Show vbModeless
    End If
End Sub                                          ' frmErrStatus.fBeginTermination


'---------------------------------------------------------------------------------------
' Method : Sub UserForm_QueryClose
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' this is the X button, do not normally use it
    Me.Hide
    Set aNonModalForm = Nothing
    ErrStatusFormUsable = False
End Sub


