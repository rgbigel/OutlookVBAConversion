# Converted from cErr.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cErr"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public atKey As String                             ' this is the string value of the cErr
# Attribute atKey.VB_VarUserMemId = 0
# Attribute atKey.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute atKey.VB_VarUserMemId = 0
# ' Attribute atKey.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public atFastMode As Boolean                       ' the real value at Entry to the Proc, used for exit
# Public atCallState As eCState                      ' if <= eCPaused, the atDsc did not Exit

# Public atCalledBy As cErr                          ' immediate Caller
# Public atDsc As cProcItem                          ' cProcDsc attached to cErr, unique for all instances
# Public at_Live As Dictionary                       ' DoCall defined Items: cCallEnv from LiveCallStack
# Public atCatDict As String                         ' list of Procs from at_Live

# Public atLastInDate As String                      ' date function
# Public atLastInSec As Double                       ' timer [sec]
# Public atPrevEntrySec As Double                    ' always >0, time of last entry
# Public atThisEntrySec As Double                    ' may be 0 if called another entry, activated
# ' again if called entry exits.
# ' Time consumed is accumulated in atDsc.TotalProcTime

# Public atMessage As String
# Public atProcIndex As Long

# Public atErrPrev As cErr                           ' back chain of same proc ???
# Public atTraceStackPos As Long                     ' Position of the ProcErr in C_TraceStack

# Public atRecursionLvl As Long                      ' incremented on every DoCall until DoExit
# Public atCallDepth As Long                         ' position on the active call stack
# Public atRecursionOK As Boolean
# Public atShowStack As String

# Public atFuncResult As String

# Public atLiveLevel As Long                         ' Live Stack level of call

# ' Error data used for controlling the ErrrorHandlerModule

# Public DebugState As Boolean
# Public EventBlock As Boolean                       ' Events are blocked During current environment

# Public NrMT As String                              ' short info for log purposes
# Public ErrSnoCatch As Boolean                      ' No ErrHandler Recursion;  NO N_PublishBugState
# Public ErrNoRec As Boolean                         ' No ErrHandler Recursion; use N_PublishBugState
# Public errNumber As Long
# Public Description As String
# Public Source As String

# Public FoundBadErrorNr As Long

# Public Permitted As Variant
# Public Explanations As String
# Public Reasoning As String

# Property Let Permit(allow As Variant)
# Permitted = allow
# T_DC.DCAllowedMatch = allow
if isEmpty(allow) Then                         ' end of handled "Try":
# MayChangeErr = True
# Permitted = Empty
# errNumber = 0
# FoundBadErrorNr = 0
# Description = vbNullString
# Explanations = vbNullString
# Reasoning = vbNullString
if E_Active Is Me Then:
# ErrSnoCatch = False
# ZErrSnoCatch = False
else:
# ' ErrSnoCatch   remains
# ' ZErrSnoCatch remains
# ErrorCaught = 0                            ' raw error last set by N_OnError

if Not E_Active Is E_AppErr Then:
# With E_AppErr
# .Permitted = Empty
# .errNumber = 0
# .FoundBadErrorNr = 0
# .Description = vbNullString
# .Explanations = vbNullString
# End With                               ' E_AppErr
if ErrStatusFormUsable Then:
# frmErrStatus.fAcceptableErrors = "No acceptable Errors"
elif ErrStatusFormUsable Then:
# frmErrStatus.fAcceptableErrors = "Acceptable Errors: " & allow
if MayChangeErr Then:
# Err.Clear
# End Property                                       ' cErr Permit Let

# '---------------------------------------------------------------------------------------
# ' Method : Sub cPrint
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def cprint():

    # Dim SbS As Boolean                                 ' Side-by-Side used if compare to otherErr
    # Dim oLine As String
    # Dim hLine As String
    # Dim lLine As Long

    if Not otherErr Is Nothing Then:
    if Not otherErr Is Me Then:
    # SbS = True                             ' Side-by-Side
    if DoDebug Then:
    # GoTo doPrint
    if DebugMode Or DebugLogging Then:
    # doPrint:
    # lLine = lKeyM + 20
    # oLine = String(20, "-") & LString(" Error Data ", lKeyM)
    if SbS Then:
    # hLine = LString(oLine, lLine) & " | " & String(lLine, "-")
    else:
    # hLine = oLine
    print(Debug.Print hLine)

    # oLine = LString("Key", 20) & atKey
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atKey
    print(Debug.Print oLine)

    # oLine = LString("CallDepth", 20) & atCallDepth
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atCallDepth
    print(Debug.Print oLine)

    # oLine = LString("LiveLevel", 20) & atLiveLevel
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atLiveLevel
    print(Debug.Print oLine)

    # oLine = LString("EventBlock", 20) & EventBlock
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.EventBlock
    print(Debug.Print oLine)

    # oLine = LString("LiveLevel", 20) & atLiveLevel
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atLiveLevel
    print(Debug.Print oLine)

    # oLine = LString("CallState", 20) & atCallState & " = " & CStateNames(atCallState)
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atCallState & " = " & CStateNames(otherErr.atCallState)
    print(Debug.Print oLine)

    # oLine = LString("CalledBy", 20) & atCalledBy.atKey
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atCalledBy.atKey
    print(Debug.Print oLine)

    # oLine = LString("TraceStackPos", 20) & atTraceStackPos
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atTraceStackPos
    print(Debug.Print oLine)

    # oLine = LString("RecursionLvl", 20) & atRecursionLvl
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atRecursionLvl
    print(Debug.Print oLine)

    # oLine = LString("ProcIndex", 20) & atProcIndex
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atProcIndex
    print(Debug.Print oLine)

    if atProcIndex >= 0 Then:
    if D_ErrInterface.Items(atProcIndex).Key = atDsc.Key Then:
    # oLine = vbNullString
    else:
    # oLine = LString("Proc/Index mismatch", 20) _
    # & D_ErrInterface.Items(atProcIndex).Key
    elif atProcIndex < -1 Then:
    if D_ErrInterface.Items(2 - atProcIndex).Key = atDsc.Key Then:
    # oLine = vbNullString
    else:
    # oLine = LString("Proc/Index mismatch", 20) _
    # & D_ErrInterface.Items(2 - atProcIndex).Key

    if SbS Then:
    if otherErr.atProcIndex > inv Then:
    if D_ErrInterface.Items(otherErr.atProcIndex).Key = otherErr.atDsc.Key Then:
    if LenB(oLine) > 0 Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & LString("Proc/Index match", 20) _
    # & D_ErrInterface.Items(otherErr.atProcIndex).Key
    else:
    if LenB(oLine) = 0 Then:
    # oLine = LString("Proc/Index match", 20) _
    # & D_ErrInterface.Items(atProcIndex).Key
    # oLine = LString(oLine, lLine) _
    # & " | " & D_ErrInterface.Items(otherErr.atProcIndex).Key
    if otherErr.atProcIndex < -1 Then:
    if D_ErrInterface.Items(2 - otherErr.atProcIndex).Key = otherErr.atDsc.Key Then:
    if LenB(oLine) > 0 Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & LString("Proc/Index match", 20) _
    # & D_ErrInterface.Items(2 - otherErr.atProcIndex).Key
    else:
    if LenB(oLine) = 0 Then:
    # oLine = LString("Proc/Index match", 20) _
    # & D_ErrInterface.Items(2 - atProcIndex).Key
    # oLine = LString(oLine, lLine) _
    # & " | " & D_ErrInterface.Items(2 - otherErr.atProcIndex).Key
    if LenB(oLine) > 0 Then:
    print(Debug.Print oLine)

    # oLine = LString("Parent Mode", 20) _
    # & atDsc.CallMode & " = " & atDsc.ModeName
    if SbS Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & otherErr.atDsc.CallMode _
    # & " = " & otherErr.atDsc.ModeName
    print(Debug.Print oLine)

    # oLine = LString("NrMT", 20) & NrMT
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.NrMT
    print(Debug.Print oLine)

    # oLine = LString("Last In Date", 20) & atLastInDate
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atLastInDate
    print(Debug.Print oLine)

    # oLine = LString("Last In Sec", 20) & FormatRight(atLastInSec, 14, 8)
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atLastInSec
    print(Debug.Print oLine)

    # oLine = LString("Time Prev Entry", 20) & FormatRight(atPrevEntrySec, 14, 8)
    if SbS Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & FormatRight(otherErr.atPrevEntrySec, 14, 8)
    print(Debug.Print oLine)

    # oLine = LString("Time This Entry", 20) & FormatRight(atThisEntrySec, 14, 8)
    if SbS Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & FormatRight(otherErr.atThisEntrySec, 14, 8)
    print(Debug.Print oLine)

    # oLine = LString("Live Dictionary", 20) & LString(atCatDict, 30)
    if SbS Then:
    # oLine = LString(oLine, lLine) _
    # & " | " & Left(otherErr.atCatDict, 30)
    print(Debug.Print oLine)

    # oLine = LString("Prev Instance", 20)
    if atErrPrev Is Nothing Then:
    # oLine = oLine & "is Nothing"
    else:
    # oLine = oLine & " | " & atErrPrev.atKey
    if SbS Then:
    if otherErr.atErrPrev Is Nothing Then:
    # oLine = LString(oLine, lLine) & " | " & "is Nothing"
    else:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atErrPrev.atKey
    print(Debug.Print oLine)

    # oLine = LString("FuncResult", 20)
    if LenB(atFuncResult) > 0 Then:
    # oLine = oLine & atFuncResult
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atFuncResult
    print(Debug.Print oLine)

    if atDsc.Key <> atKey Or otherErr.atDsc.Key <> otherErr.atKey Then:
    # oLine = LString("ProcKey mismatch", 20) & "atDsc = " & atDsc.Key
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atDsc.Key
    print(Debug.Print oLine)

    if All Or DebugMode Or errNumber <> 0 Or otherErr.errNumber <> 0 _:
    # Or FoundBadErrorNr <> 0 Or otherErr.FoundBadErrorNr <> 0 Then
    # oLine = LString("errNumber", 20) & errNumber
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.errNumber
    print(Debug.Print oLine)

    # oLine = LString("Description", 20) & Description
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.Description
    print(Debug.Print oLine)

    # oLine = LString("Bad Error", 20) & FoundBadErrorNr
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.FoundBadErrorNr
    print(Debug.Print oLine)

    # oLine = LString("Permit", 20) & Permitted
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.Permitted
    print(Debug.Print oLine)

    # oLine = LString("ErrSnoCatch", 20) & ErrSnoCatch
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.ErrSnoCatch
    print(Debug.Print oLine)

    # oLine = LString("DebugState", 20) & DebugState
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.DebugState
    print(Debug.Print oLine)

    # oLine = LString("Explanations", 20) & Explanations
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.Explanations
    print(Debug.Print oLine)

    # oLine = LString("Reasoning", 20) & Reasoning
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.Reasoning
    print(Debug.Print oLine)

    # oLine = LString("atShowStack", 20) & atShowStack
    if SbS Then:
    # oLine = LString(oLine, lLine) & " | " & otherErr.atShowStack
    print(Debug.Print oLine)

    print(Debug.Print hLine & " End")



# Property Get TraceEntry() As cTraceEntry
# Dim TP As Long

# TP = atTraceStackPos Mod ErrLifeTime

if TP > 0 Then:
# Set TraceEntry = C_CallTrace(TP)
# End Property                                       ' cErr.TraceEntry Get

# Property Let TraceSucc(TraceSucc As Long)
# Dim TP As Long

# TP = atTraceStackPos Mod ErrLifeTime

if TP > 0 Then:
# C_CallTrace.Item(TP).TSuc = TraceSucc
# End Property                                       ' cErr.TraceSucc Let

# Property Get TraceSucc() As Long
# Dim TP As Long

# TP = atTraceStackPos Mod ErrLifeTime

if TP > 0 Then:
# TraceSucc = C_CallTrace.Item(TP).TSuc
# End Property                                       ' cErr.TraceSucc Get

# Property Let TracePreDec(Pre As Long)
# Dim TP As Long

# TP = atTraceStackPos Mod ErrLifeTime

if TP > 0 Then:
# C_CallTrace.Item(TP).TSuc = Pre
# End Property                                       ' cErr.TracePreDec Let

# Property Get TracePreDec() As Long
# Dim TP As Long

# TP = atTraceStackPos Mod ErrLifeTime

if TP > 0 Then:
# TracePreDec = C_CallTrace.Item(TP).TPre
# End Property                                       ' cErr.TracePreDec Get

# '---------------------------------------------------------------------------------------
# ' Method : Clone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose: Replicates a cErr (for Recursion, Application start etc.)
# '           Used until return from stack.
# ' Note   : the optional False Parm Exact=True will copy everything without exception.)
# '---------------------------------------------------------------------------------------
def clone():

    # Set Clone = New cErr

    # Clone.atFastMode = atFastMode
    # Clone.atCallState = atCallState
    # Clone.atMessage = atMessage
    # Set Clone.atDsc = atDsc
    # Clone.atKey = atKey
    # Clone.atProcIndex = atProcIndex
    # Clone.atRecursionLvl = atRecursionLvl
    # Clone.atCallDepth = atCallDepth
    # Clone.atRecursionOK = atRecursionOK
    # Clone.atShowStack = atShowStack
    # Clone.DebugState = DebugState
    # Clone.EventBlock = EventBlock
    # Clone.NrMT = NrMT
    # Clone.errNumber = errNumber
    # Clone.Description = Description
    # Clone.Source = Source
    # Clone.FoundBadErrorNr = FoundBadErrorNr
    # Clone.Permitted = Permitted
    # Clone.Explanations = Explanations
    # Clone.Reasoning = Reasoning
    # Set Clone.atErrPrev = atErrPrev                ' links Clone to current always
    if Exact Then:
    # Clone.atTraceStackPos = atTraceStackPos
    # Set Clone.atCalledBy = atCalledBy
    # Set Clone.at_Live = at_Live
    # Clone.atCatDict = atCatDict
    # Clone.atLastInDate = atLastInDate
    # Clone.atLastInSec = atLastInSec
    # Clone.atPrevEntrySec = atPrevEntrySec
    # Clone.atThisEntrySec = atThisEntrySec
    # Clone.atFuncResult = atFuncResult
    # Clone.ErrSnoCatch = ErrSnoCatch
    # Clone.ErrNoRec = ErrNoRec
    else:
    # Clone.ErrSnoCatch = ZErrSnoCatch          ' No ErrHandler Recursion;  NO N_PublishBugState
    # Clone.ErrNoRec = ZErrNoRec                ' No ErrHandler Recursion; use N_PublishBugState
    # 'Clone.atCatDict = vbNullString               '
    # 'Set Clone.at_Live = nothing                  ' no LiveStack exists as yet for this instance
    # 'Clone.atCatDict = vbNullString               '
    # 'Clone.atPrevEntrySec, Last, This, Date = 0   ' instance can not share time values
    # 'Set Clone.atCalledBy = Nothing               '    "     and no caller yet
    # 'Clone.atFuncResult = vbNullString            '    "     and no result yet
    # CloneCounter = CloneCounter + 1


