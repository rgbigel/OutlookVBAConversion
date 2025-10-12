# Converted from cTraceEntry.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cTraceEntry"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Private TKey As String
# Attribute TKey.VB_VarUserMemId = 0
# Attribute TKey.VB_VarDescription = "Display Key of traced Err"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute TKey.VB_VarUserMemId = 0
# ' Attribute TKey.VB_VarDescription = "Display Key of traced Err"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public TErr As cErr
# Public Tinx As Long
# Public TPre As Long
# Public TSuc As Long

# Public TSrc As String
# Public TLne As Long
# Public TDet As String
# Public TLD As String                               ' atLastInDate   : trace<>non-instance vals in cErr
# Public TLS As Double                               ' atLastInSec
# Public TPS As Double                               ' atPrevEntrySec
# Public TES As Double                               ' atThisEntrySec
# Public TLog As String                              ' LogProgress Line
# Public TRL As Long                                 ' recursion Level (instance)
# Public TCal As String                              ' Stack call situation: A if top of stack and active
# ' P if recursive and not yet inactive (one other instance=A)
# ' E exited because not (A or P)

# '---------------------------------------------------------------------------------------
# ' Method : Function TCallState
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def tcallstate():
    # TCallState = TCal
    match TErr.atCallState:
        case eCExited:
    if TCal <> "E" Then                    ' debug.assert not working??:
    print(Debug.Print Tinx, TCal, TLog)
    # DoVerify False
    # TCal = "E"
        case _:
    match TCal:
        case "A":
    # DoVerify TErr.atCallState <> eCUndef
        case "E":
    # DoVerify TErr.atCallState <> eCpaused
    # DoVerify TErr.atCallState <> eCUndef
        case "P":
    # DoVerify TErr.atCallState = eCpaused
        case "u":
    # DoVerify TErr.atCallState = eCUndef
        case _:
    # DoVerify False

    # aCallState = TCal
    # DoVerify aCallState = TCallState


# Property Get TraceSucc(Optional TT As Long) As cTraceEntry

# Dim TP As Long
# Dim ps As Long
# Dim aPredTE As cTraceEntry

if TT = 0 Then:
# TP = TraceTop
else:
# TP = TT
# TP = TP Mod ErrLifeTime

# Set aPredTE = C_CallTrace(TP)
# ps = aPredTE.TSuc Mod ErrLifeTime

if ps > 0 And TErr.atCallDepth > 0 Then:
# Set TraceSucc = C_CallTrace(ps)
# DoVerify TraceSucc.TErr.atKey = TErr.atKey, "** check integrity"
if DebugMode Then:
# DoVerify aPredTE.TErr.atErrPrev Is TraceSucc.TErr, "check integrity"

# End Property                                       ' cTraceEntry.TraceSucc Get

# Property Get TracePred(Optional TT As Long) As cTraceEntry

# Dim TP As Long
# Dim ps As Long
# Dim aPredTE As cTraceEntry

if TT = 0 Then:
# TP = TraceTop
else:
# TP = TT
if C_CallTrace.Count >= ErrLifeTime - 1 Then:
# TP = TP Mod ErrLifeTime
if TP < 3 Then                             ' skipping 200..202 (modulus 200):
# TP = TP + 3

# Set aPredTE = C_CallTrace(TP)
# ps = aPredTE.TPre Mod ErrLifeTime

if ps > 0 Then:
# Set TracePred = C_CallTrace(ps)

# aCallState = TCal

# End Property                                       ' cTraceEntry.TracePred Get

# '---------------------------------------------------------------------------------------
# ' Method : TraceAdd
# ' Author : Rolf G. Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Put Trace Entry onto Stack setting Tinx
# '---------------------------------------------------------------------------------------
def traceadd():
    # Dim TP As Long
    # Dim aPredTE As cTraceEntry

    # Tinx = TraceTop + 1
    # TKey = TErr.atKey

    if TraceTop = 0 Or C_CallTrace.Count < ErrLifeTime - 1 Then:
    # C_CallTrace.Add Me                         ' creating element up to ErrLifeTime-1
    # TErr.atTraceStackPos = Tinx
    else:
    # TP = Tinx
    # TP = TP Mod ErrLifeTime
    if TP < 3 Then                             ' skipping 200..202 (modulus 200):
    # TP = TP + 3
    # Tinx = Tinx + 3
    # Set aPredTE = C_CallTrace.Item(TP)
    if DebugMode Then:
    # DoVerify aPredTE.TErr.atCallState = eCExited, _
    # "entry #" & Tinx & b & aPredTE.TErr.atKey _
    # & " is " & CStateNames(aPredTE.TErr.atCallState) _
    # & " ??? It will be kicked from CallTrace!"
    # C_CallTrace.Remove TP                      ' "replace" position Tp
    if TP > C_CallTrace.Count Then:
    # C_CallTrace.Add Me                     ' creating element up to ErrLifeTime-1
    else:
    # C_CallTrace.Add Me, Before:=TP
    # TErr.atTraceStackPos = Tinx

    if Not TErr.atCalledBy Is Nothing Then:
    # TPre = TErr.atCalledBy.atTraceStackPos

    if TPre > 0 Then                               ' check if TracePreDec is still ok:
    # TP = TPre Mod ErrLifeTime
    # Set aPredTE = TracePred(Tinx)              ' who is the TracePreDec
    if TP <= Tinx Then                         ' probably has become invalid:
    if aPredTE Is Nothing Then:
    # TPre = 0
    # GoTo noPre
    else:
    if aPredTE.TErr.atCallState = eCExited Then:
    # aPredTE.TPre = 0               ' caller exited ==> not in use, can overwrite
    # GoTo noPre
    elif P_Active.CallMode <> eQnoDef And aPredTE.TErr.atCallState = eCActive Then:
    # DoVerify aPredTE.TErr Is E_Active, _
    # "Caller " & aPredTE.TErr.atKey _
    # & " is active (and not Extern.Caller)"
    else:
    # DoVerify 1 = 0, "TPre > 0 and alive"
    # aPredTE.TSuc = Tinx
    # noPre:
    if TSuc > 1 Then                               ' check if TraceSucc is still ok, ignore Extern.Caller (TSuc=1):
    # TP = TPre Mod ErrLifeTime
    # Set aPredTE = TraceSucc(Tinx)              ' who is the TraceSucc
    if TP <= aPredTE.TErr.atTraceStackPos Mod ErrLifeTime Then ' probably has become invalid:
    if aPredTE Is Nothing Then:
    # TSuc = 0
    else:
    if aPredTE.TErr.atCallState = eCExited Then ' this is ok, item not in use, can overwrite:
    # TSuc = 0
    else:
    # DoVerify Not DebugMode, "killing active TraceSucc??"
    # TSuc = 0
    else:
    # ' TSuc > 0 and alive
    if CallLogging And LogZProcs Then:
    # Call N_ShowProgress(CallNr, TErr.atDsc, "+T", vbNullString, ExplainS)

    # FuncExit:
    # TraceTop = Tinx
    # DoVerify aCallState = TCal, "design check on CallState ???"

