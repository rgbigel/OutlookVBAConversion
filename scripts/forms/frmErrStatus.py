# Converted from frmErrStatus.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmErrStatus
# Caption         =   "Error and Debug Status"
# ClientHeight    =   10125
# ClientLeft      =   120
# ClientTop       =   14460
# ClientWidth     =   7350
# OleObjectBlob   =   "frmErrStatus.frx":0000
# ShowModal       =   0   'False
# End
# Attribute VB_Name = "frmErrStatus"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# Public frmIgnoreErrStatusChange As Boolean       ' to avoid recursion if EventRoutine fModifications_Change

# '---------------------------------------------------------------------------------------
# ' Method : Sub Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub Activate()

# Const zKey As String = "frmErrStatus.Activate"
# Static zDsc As cProcItem

# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean

# aBugVer = zDsc.CallCounter < 2
if DoVerify(aBugVer, "frmErrStatus must never have 2 Window Instances") Then:
# GoTo ProcRet

if Recursive Then:
print(Debug.Print String(OffCal, b) & "Forbidden recursion from " _)
# & P_Active.DbgId & " => " & zKey
# GoTo ProcRet
# Recursive = True

# Call UserForm_Activate

# Show
# Call WindowSetForeground(Me.Caption, Nothing)

# Recursive = False

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : ReEvaluate
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: When in debug window, evaluate state of Settings, or if Reverse: global values
# '---------------------------------------------------------------------------------------
# Public Sub ReEvaluate(Optional Reverse As Boolean = False)
# Const zKey As String = "frmErrStatus.ReEvaluate"
# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean

if Recursive Then:
# GoTo ProcRet
# Recursive = True

if E_AppErr Is Nothing Then:
# fLastErrReasoning = vbNullString
elif E_AppErr.Reasoning <> vbNullString Then:
# fLastErrReasoning = E_AppErr.Reasoning

if ItemsToDoCount + Deferred.Count <> 0 Then:
# fDeferredCount = ItemsToDoCount
# fCurFolder = curFolderPath
else:
# fDeferredCount = vbNullString

# frmIgnoreErrStatusChange = False            ' not restored on exit!
# Call N_Suppress(Push, zKey)                 ' ShutupMode on
# ChangeAssignReverse = Reverse
# SuppressStatusFormUpdate = True
# fModifications.Visible = True
if fModifications Then                      ' reset fModifications:
# fModifications.Enabled = False          ' without  action here
# fModifications = False
# fModifications.Enabled = True               ' with actions now
# fModifications = True                       ' will Call fModifications_Change

# fLastErrExplanations = "Werte fr Debugoptionen aus frmErrStatus ausgelesen " & Time
# ChangeAssignReverse = False
# Call N_Suppress(Pop, zKey)                  ' ShutupMode restored
# Recursive = False

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub UpdInfo
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub UpdInfo()
# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean

if Not Visible Then                          ' invisible windows can't come to foreground:
# GoTo ProcRet
if Recursive Then:
# GoTo ProcRet
# Recursive = True

if ErrDisplayModify And Not fHideMe Then     ' should frmErrStatus show (with modif. values)?:
try:
    # Call WindowSetForeground(Me.Caption, Nothing)
    # ignoreAll:
    # ErrDisplayModify = False

    # FuncExit:
    # Recursive = False
    # ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub Form_Load
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Form_Load()

# Const zKey As String = "frmErrStatus.Form_Load"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub)

# Call getDebugMode                            ' don't force: using debug as/if defined

if LenB(Testvar) = 0 Then:
# fToggleDebug.Caption = "Toggle Debugmode:=ON"
else:
# fToggleDebug.Caption = "Debuging Options: " & Quote(Testvar)
if DebugMode Then:
# fToggleDebug.BackColor = &H80FFFF
else:
# fToggleDebug.BackColor = &H8000000F
# Repaint
# doMyEvents

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub fAllOfThese_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fAllOfThese_Change()

if frmIgnoreErrStatusChange Then:
# GoTo ProcRet

# FuncExit:
# frmIgnoreErrStatusChange = False
# fModifications = True                        ' raises event on change

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub fBreak_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fBreak_Click()

print(Debug.Print "Break Key pressed in frmErrStatus")
# Call BugTimerDeActivate
# Debug.Assert False                           ' always as trivial function
# Call BugTimerActivate(0)


# '---------------------------------------------------------------------------------------
# ' Method : Sub fCallLogging_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fCallLogging_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fLogZProcs_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fLogZProcs_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fLImmediate_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fLImmediate_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fCancelTermination_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fCancelTermination_Click()

# Const zKey As String = "frmErrStatus.fCancelTermination_Click"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

# Call N_ErrClear

# E_AppErr.errNumber = vbObjectError + 101     ' user defined this to be an acceptable error
# E_AppErr.FoundBadErrorNr = 0

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub fEditWatchState_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fEditWatchState_Click()

# Const zKey As String = "frmErrStatus.fEditWatchState_Click"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

if Not MayChangeErr Then:
# & vbCrLf & "You can OK to override and set the Test variable now, " _
# & "but this will cause an error on Resume from N_PublishBugState.", _
# vbOKCancel)
if rsp = vbCancel Then:
# GoTo ProcReturn                      ' user cancelled the Edit

# Call EditWatch

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub fDebugLogging_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fDebugLogging_Change()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fHideMe_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fHideMe_Click()
# fModifications = True
if fHideMe Then:
# Me.Hide
else:
# Me.Show vbModeless

# '---------------------------------------------------------------------------------------
# ' Method : Sub fLogAllErrors_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fLogAllErrors_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fLogPerformance_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fLogPerformance_Change()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fNoTimerEvent_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fNoTimerEvent_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# Private Sub fOnline_Click()
if fOnline.Caption = "Offline" Then:
# Call SetOnline(olCachedConnectedFull)
# fOnline.Caption = "Online"
else:
# Call SetOffline
# fOnline.Caption = "Offline"

# '---------------------------------------------------------------------------------------
# ' Method : Sub fShowFunVal_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fShowFunVal_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fShowLog_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fShowLog_Click()
# Call ShowLog

# '---------------------------------------------------------------------------------------
# ' Method : Sub fExLiveCheck_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fExLiveCheck_Click()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fTraceMode_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fTraceMode_Change()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fDebugMode_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fDebugMode_Change()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fStackDebug_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fStackDebug_AfterUpdate()
if Not frmIgnoreErrStatusChange Then:
if Not IsNumeric(fStackDebug) Then       ' ignore if not numeric:
# fStackDebug = StackDebug
# fIncDecDebug = StackDebug
else:
# fIncDecDebug = fStackDebug
if fAllOfThese Then:
# Call N_DebugStart
# fModifications = True                ' Raise that _Change

# '---------------------------------------------------------------------------------------
# ' Method : Sub fModifications_Change
# ' Author : rgbig
# ' Date   : 20211109@17_00
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fModifications_Change()
# Const zKey As String = "frmErrStatus.fModifications_Change"

# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean
# Dim ModString As String
# Dim chgMode As Boolean
# Dim chgMsg As String

if frmIgnoreErrStatusChange Or Not fModifications.Enabled Then    ' also protects against recursion:
# GoTo ProcRet

if Recursive Then:
print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
# & P_Active.DbgId & " => " & zKey
# GoTo ProcRet

# fModifications.Enabled = False              ' active in Change: do not trigger
# Recursive = True                            ' restored by    Recursive = False ProcRet:
# Call BugTimer.BugState_SetPause

if fIncDecDebug = 0 Then                    ' consistency forced:
# fIncDecDebug = StackDebug               ' not gated by frmIgnoreErrStatusChange
if fModifications Then:
# LogImmediate = fLImmediate              ' form always wins
if fIncDecDebug = 0 Then:
# fIncDecDebug = StackDebug           ' not gated by frmIgnoreErrStatusChange
else:
# StackDebug = fIncDecDebug
if CLng(fIncDecDebug) <> StackDebug Then:
# fStackDebug = CStr(fIncDecDebug.Value)
if AssignIfChanged(StackDebug, fStackDebug) Then:
if ChangeAssignReverse Then:
# StackDebug = ModThisTo
# fIncDecDebug = StackDebug
else:
# fStackDebug = ModThisTo
# fIncDecDebug.Value = CLng(0 & fStackDebug.Value)
# StackDebugOverride = -StackDebug
# ModString = vbCrLf & "StackDebug = " & StackDebug
else:
# StackDebugOverride = StackDebug     ' New setting affecting ErrReset
if AssignIfChanged(DebugMode, fDebugMode) Then:
if ChangeAssignReverse Then:
# fDebugMode = ModThisTo
else:
# DebugMode = ModThisTo
# ModString = ModString & vbCrLf & "DebugMode = " & CStr(DebugMode)
if AssignIfChanged(DebugLogging, fDebugLogging) Then:
if ChangeAssignReverse Then:
# fDebugLogging = ModThisTo
else:
# DebugLogging = ModThisTo
# ModString = ModString & vbCrLf & "DebugLogging = " & CStr(DebugLogging)
if AssignIfChanged(LogAllErrors, fLogAllErrors) Then:
if ChangeAssignReverse Then:
# fLogAllErrors = ModThisTo
else:
# LogAllErrors = ModThisTo
# ModString = ModString & vbCrLf & "LogAllErrors = " & CStr(LogAllErrors)
if fNoTimerEvent Then:
if Not NoTimerEvent Then:
# Call BugTimerDeActivate         ' turn off Timer
# NoTimerEvent = True
# ModString = ModString & vbCrLf & "NoTimerEvent = " & CStr(NoTimerEvent)
else:
if NoTimerEvent Then                ' check if possible here:
if ErrTimerEventNotReady Then   ' no it isn't possible:
# NoTimerEvent = False
# ModString = ModString & vbCrLf & "NoTimerEvent = " & CStr(NoTimerEvent)
else:
if DebugMode Then:
# DoVerify ErrTimerEventNotReady, "*** Timer should be off"
if AssignIfChanged(LogPerformance, fLogPerformance) Then:
if ChangeAssignReverse Then:
# fLogPerformance = ModThisTo
else:
# LogPerformance = ModThisTo
# ModString = ModString & vbCrLf & "LogPerformance = " & CStr(LogPerformance)
if AssignIfChanged(TraceMode, fTraceMode) Then:
if ChangeAssignReverse Then:
# fTraceMode = ModThisTo
else:
# TraceMode = ModThisTo
# ModString = ModString & vbCrLf & "TraceMode = " & CStr(TraceMode)
if AssignIfChanged(ExLiveDscGen, fExLiveCheck) Then:
if ChangeAssignReverse Then:
# fExLiveCheck = ModThisTo
else:
# ExLiveDscGen = ModThisTo
# ModString = ModString & vbCrLf & "ExLiveDscGen = " & CStr(ExLiveDscGen)
if AssignIfChanged(ShowFunctionValues, fShowFunVal) Then:
if ChangeAssignReverse Then:
# fShowFunVal = ModThisTo
else:
# ShowFunctionValues = ModThisTo
# ModString = ModString & vbCrLf & "ShowFunctionValues = " & CStr(ShowFunctionValues)
if AssignIfChanged(CallLogging, fCallLogging) Then:
if ChangeAssignReverse Then:
# fCallLogging = ModThisTo
else:
# CallLogging = ModThisTo
# ModString = ModString & vbCrLf & "CallLogging = " & CStr(CallLogging)
if AssignIfChanged(LogZProcs, fLogZProcs) Then:
if ChangeAssignReverse Then:
# fLogZProcs = ModThisTo
else:
# LogZProcs = ModThisTo
# ModString = ModString & vbCrLf & "LogZProcs = " & CStr(LogZProcs)
if AssignIfChanged(E_Active.EventBlock, fNoEvents) Then:
if ChangeAssignReverse Then:
# fNoEvents = ModThisTo
else:
# E_Active.EventBlock = ModThisTo
if Not E_Active.EventBlock Then ' no need to block Events -> online:
# chgMode = SetOnline(olCachedConnectedFull)
if chgMode Then:
# chgMsg = " changed to Online"
else:
# chgMsg = " was Online"
if DoVerify(fOnline = actOnlineStatus, _:
# "*** Eventblock: " & E_Active.EventBlock _
# & " showing " & fOnline & " actual OnlineStatus is " _
# & actOnlineStatus & " correcting caption!" _
# & chgMsg) Then
# fOnline.Caption = "Online"
else:
# chgMode = SetOffline()
if chgMode Then:
# chgMsg = " changed to Offline"
else:
# chgMsg = " was Offline"
if DoVerify(fOnline = actOnlineStatus, _:
# "*** Eventblock: " & E_Active.EventBlock _
# & " showing " & fOnline & " but OnlineStatus is " _
# & actOnlineStatus & " correcting caption!" _
# & chgMsg) Then
# fOnline.Caption = "Offline"
if chgMode And DebugMode Then:
# Call LogEvent("Outlook changed to " _
# & IIf(E_Active.EventBlock, "Offline", "Online"))
# ModString = ModString & vbCrLf & "EventBlock(" & E_Active.atKey _
# & ") = " & CStr(E_Active.EventBlock)
if AssignIfChanged(T_DC.DCerrSource, fLastErrSource) Then:
if ChangeAssignReverse Then:
# fLastErrSource = ModThisTo
else:
# T_DC.DCerrSource = ModThisTo
# ' not a normal change, dont log ModString = ModString & vbCrLf & "LastErrSource = " & T_DC.DCerrSource
if AssignIfChanged(UseErrExOn, fUseErrExOn) Then:
if ChangeAssignReverse Then:
# fUseErrExOn = ModThisTo
else:
# UseErrExOn = ModThisTo
# fUseErrExOn = ModThisTo             ' special logic! NOT changing UseErrExOn !!!
# Call N_SetErrExHdl(doPrint:=False)
# ModString = ModString & vbCrLf & "UseErrExOn = " & fUseErrExOn & b & " ErrExActive=" & ErrExActive
if AssignIfChanged(T_DC.DCUseErrExOn, fUseErrExOn) Then ' special logic! Irrelevant if reversed:
if LenB(ModThisTo) = 0 Or LenB(UseErrExOn) = 0 Then ' LastErrExOn is never changed:
# fUseErrExOn = vbNullString
# fErrorHandler = "Not using Global Error Handler"
# fUseErrExOn.BackColor = 188
# fUseErrExOn.BackStyle = fmBackStyleOpaque
else:
# fUseErrExOn = UseErrExOn
# fErrorHandler = "Module to use for " & vbCrLf & "  Global Error Handler"
# fUseErrExOn.BackColor = 255
# fUseErrExOn.BackStyle = fmBackStyleTransparent
# T_DC.DCUseErrExOn = fUseErrExOn
# UseErrExOn = fUseErrExOn
# ModString = ModString & vbCrLf & "DCUseErrExOn = " & fUseErrExOn
if AssignIfChanged(T_DC.DCAppl, fErrAppl) Then ' special logic!:
if LenB(ModThisTo) = 0 Then:
# ModThisTo = S_AppKey
# fErrAppl = ModThisTo
# T_DC.DCAppl = ModThisTo
else:
if ChangeAssignReverse Then:
# fErrAppl = ModThisTo
else:
# T_DC.DCAppl = ModThisTo
if DebugMode And DebugLogging Then   ' normally no log because normal change:
# ModString = ModString & vbCrLf & "DCAppAct = " & ModThisTo
if ChangeAssignReverse Then              ' limited asssignments: not editable:
if isEmpty(T_DC.DCAllowedMatch) Then ' T_DC.DCAllowedMatch is not editable string:
if fAcceptableErrors <> "No Acceptable Errors" Then:
# unset:
# fAcceptableErrors = "No Acceptable Errors"
# fAcceptableErrors.Enabled = False
if DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then:
# ErrDisplayModify = True
# fAcceptableErrors.SpecialEffect = fmSpecialEffectFlat
elif LenB(T_DC.DCAllowedMatch) = 0 Then:
# GoTo unset
else:
if fAcceptableErrors <> "Accept Errors=" & CStr(T_DC.DCAllowedMatch) Then:
if DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then:
# ErrDisplayModify = True
# fAcceptableErrors = "Accept Errors=" & CStr(T_DC.DCAllowedMatch)
# fAcceptableErrors.SpecialEffect = fmSpecialEffectBump
if fLastErrExplanations <> E_Active.Explanations Then ' E_Active.Explanations is not editable string:
if LenB(E_Active.Explanations) > 2 Then:
# fLastErrExplanations = E_Active.Explanations
else:
# fLastErrExplanations = vbNullString
if DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then:
# ErrDisplayModify = True
if fLastErrReasoning <> E_Active.Reasoning Then ' E_Active.Reasoning is not editable string:
if LenB(E_Active.Reasoning) > 2 Then:
# fLastErrReasoning = E_Active.Reasoning
else:
# fLastErrReasoning = vbNullString
if DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then:
# ErrDisplayModify = True
if T_DC.DCerrNum = 0 Then            ' T_DC.DCerrNum is not editable:
if LenB(fLastErr) > 0 Then:
# fLastErr = vbNullString
# fLastErrHex = vbNullString
if DebugMode Or DebugLogging Or T_DC.DCerrNum <> 0 Then:
# ErrDisplayModify = True
elif fLastErr <> CStr(T_DC.DCerrNum) Then:
# fErrAppl = T_DC.DCAppl
# fErrNumber = T_DC.DCerrNum
# fLastErr = CStr(T_DC.DCerrNum)
# fLastErrHex = Hex8(T_DC.DCerrNum)
# ErrDisplayModify = True

# ModString = Trim(ModString)
if LenB(ModString) > 0 And Not frmIgnoreErrStatusChange Then:
print(Debug.Print "* Settings changed: " & ModString)
# ErrDisplayModify = True
# Call Z_StateToTestVar

# fModifications = False                      ' causes _Change recursion, but immediately exits
# Recursive = False
# Call BugTimer.BugState_UnPause
# frmIgnoreErrStatusChange = False
# fModifications.Enabled = True               ' enable the change trigger

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Sub fNoEvents_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fNoEvents_Click()
if Not frmIgnoreErrStatusChange Then:
# E_Active.EventBlock = fNoEvents
if E_AppErr.EventBlock <> fNoEvents Then:
# E_AppErr.EventBlock = fNoEvents
# DoVerify True, "*** changed EventBlock to " _
# & E_AppErr.EventBlock & " for " & E_AppErr.atKey
# fModifications = True                    ' Raise that _Change

# '---------------------------------------------------------------------------------------
# ' Method : Sub fTerminationFlag_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fTerminationFlag_Click()
# Me.Hide                                      ' not gated by frmIgnoreErrStatusChange
# T_DC.TermRQ = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub fToggleDebug_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fToggleDebug_Click()
# Call SetDebugMode(noMsg:=True)               ' not gated by frmIgnoreErrStatusChange

# '---------------------------------------------------------------------------------------
# ' Method : Sub fUseErrExOn_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fUseErrExOn_AfterUpdate()
if Not frmIgnoreErrStatusChange Then:
# fModifications = True                    ' Raise that _Change

# '---------------------------------------------------------------------------------------
# ' Method : Sub fIncDecDebug_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub fIncDecDebug_Change()
if fStackDebug <> fIncDecDebug Then:
# fStackDebug = fIncDecDebug               ' not gated by frmIgnoreErrStatusChange
# fModifications = True                    ' Raise that _Change

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# Call QueryErrStatusChange(False)             ' .f values are changeable


# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_DblClick
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
# Call ReEvaluate
# ErrDisplayModify = True

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Deactivate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Deactivate()
if fModifications Then:
# Call getDebugMode
# fModifications = False
# fLastErrIndications.Enabled = True

# Private Sub UserForm_Initialize()
# '--- Proc MAY ONLY CALL Z_Type OR Z_Type PROCS
# Const zKey As String = "frmErrStatus.UserForm_Initialize"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmErrStatus")

# frmIgnoreErrStatusChange = True
# fAcceptableErrors = "Acceptable Errors"
# fAcceptableErrors.Enabled = False
# fAcceptableErrors.SpecialEffect = fmSpecialEffectFlat
# fCallLogging = CallLogging
# fDebugLogging = DebugLogging
# fDebugMode = DebugMode
# fUseErrExOn = LastErrExOn
if LenB(T_DC.DCUseErrExOn) = 0 Then:
# fUseErrExOn = LastErrExOn
else:
# fUseErrExOn = T_DC.DCUseErrExOn
# fErrNumber = T_DC.DCerrNum
# fLastErr = CStr(T_DC.DCerrNum)
# fLastErrExplanations = E_Active.Explanations
# fLastErrHex = Hex8(T_DC.DCerrNum)
# fLastErrMsg = T_DC.DCerrMsg
# fLastErrReasoning = E_Active.Reasoning
# fLastErrSource = T_DC.DCerrSource
# fLImmediate = True
# fLogAllErrors = LogAllErrors
# fLogPerformance = LogPerformance
# fLogZProcs = LogZProcs
# fNoEvents = E_Active.EventBlock
# fNoTimerEvent = NoTimerEvent
# fShowFunVal = ShowFunctionValues
# fExLiveCheck = ExLiveDscGen
# fStackDebug = StackDebug
# fTraceMode = TraceMode
# frmErrStatus.Top = 245
# frmErrStatus.Left = 1041

# Call QueryErrStatusChange(False)             ' .f values are changable
if Visible Then:
# Call WindowSetForeground(Me.Caption, Nothing)

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub fBeginTermination
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Ask User about Termination request (may abort Termination)
# '---------------------------------------------------------------------------------------
# Public Sub fBeginTermination(Optional fTerminationFlag As Boolean = True)
# T_DC.TermRQ = fTerminationFlag
if DebugMode Then:
# Me.Show vbModeless


# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_QueryClose
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
# ' this is the X button, do not normally use it
# Me.Hide
# Set aNonModalForm = Nothing
# ErrStatusFormUsable = False

