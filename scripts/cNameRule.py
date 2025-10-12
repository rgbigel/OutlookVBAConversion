# Converted from cNameRule.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cNameRule"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public RuleMatches As Boolean
# Public MatchOn As String                           ' last PropertyNameX we Matched or vbNullString if no Match
# Public PatFound As String                          ' explain pattern Match rule
# Public bConsistent As Boolean                      ' change indicator

# Public aRuleString As String                       ' Treat private, use GetRuleString or ChangeTo
# Public CritRestrictString As String                ' Syntax is for Restrict-operation
# Public CritFilterString As String                  ' Syntax is for Filter-Operation
# Public MatchesList As Variant
# Public CleanMatchesString As String
# Public CleanMatches As Variant

# Public PropWildcard As Boolean
# Public PropTailWild As Boolean
# Public PropFrontWild As Boolean
# Public PropMustLog As Boolean
# Public PropVisibility As Boolean

# Public PropAllRules As cAllNameRules

# '---------------------------------------------------------------------------------------
# ' Method : Sub RuleClear
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub RuleClear()                             ' Parent is NOT RuleCleared
# Dim zErr As cErr
# Const zKey As String = "cNameRule.RuleClear"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cNameRule")

# RuleMatches = False
# MatchOn = vbNullString
# PatFound = vbNullString
# aRuleString = vbNullString
# MatchesList = Empty
# CleanMatchesString = vbNullString
# CleanMatches = Empty

if Parent Is Nothing Then:
if bConsistent Then:
# DoVerify False

# bConsistent = False
# PropWildcard = False
# PropTailWild = False
# PropFrontWild = False
# PropMustLog = False
# PropVisibility = False

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub RuleCopy
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub RuleCopy(ByRef S As cNameRule, withMatchBits As Boolean)

# Const zKey As String = "cNameRule.RuleCopy"

# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean

# aBugVer = Not Recursive
if DoVerify(aBugVer, "Forbidden recursion from " _:
# & P_Active.DbgId & " => " & zKey) Then
# GoTo ProcRet
# Recursive = True                               ' restored by    Recursive = False ProcRet:

# Call DoCall(zKey, tSub, eQzMode)

if S.PropAllRules Is Nothing Then:
if DebugLogging And S.bConsistent And Not withMatchBits Then:
if DebugMode Then DoVerify False:

if withMatchBits Then:
# Me.RuleMatches = S.RuleMatches
# Me.bConsistent = S.bConsistent
else:
# Me.RuleMatches = False
# Me.bConsistent = False
# Me.MatchOn = S.MatchOn
# Me.PatFound = S.PatFound
# Me.aRuleString = S.aRuleString
# Me.MatchesList = S.MatchesList
# Me.CleanMatchesString = S.CleanMatchesString
# Me.CleanMatches = S.CleanMatches
# Me.PropWildcard = S.PropWildcard
# Me.PropTailWild = S.PropTailWild
# Me.PropFrontWild = S.PropFrontWild
# Me.PropMustLog = S.PropMustLog
# Me.PropVisibility = S.PropVisibility
# Set Me.PropAllRules = S.PropAllRules

# zExit:
# Call DoExit(zKey)
# Recursive = False

# ProcRet:

# '---------------------------------------------------------------------------------------
# ' Method : Clone
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def clone():

    # Dim nCloneNameRule As cNameRule

    # Set nCloneNameRule = New cNameRule
    # Call nCloneNameRule.RuleCopy(Me, withMatchBits)
    # Set Clone = nCloneNameRule


# Public Property Get GetMatchString() As String

# GetMatchString = Trim(Me.aRuleString)

# End Property                                       ' cNameRule.GetMatchString Get

# Public Property Let ChangeTo(S As String)

# Dim HasChanged As Boolean

if LenB(S) = 0 Then:
if aRuleString = S Then:
# HasChanged = False
else:
# HasChanged = True
# SelectedAttributes = vbNullString      ' append changes later
# Me.RuleClear
else:
# HasChanged = Me.DecodeThisMatch(S)
# Me.RuleMatches = HasChanged
if Not iRules Is Nothing Then:
if HasChanged Then                     ' something has changed:
# iRules.RuleInstanceValid = False
# SelectedAttributes = vbNullString  ' append changes later

# End Property                                       ' cNameRule.ChangeTo Let

# '---------------------------------------------------------------------------------------
# ' Method : Function DecodeThisMatch
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function DecodeThisMatch(ByVal RuleString As String) As Boolean
# '--- Proc MAY ONLY CALL Z_Type PROCS
# Const zKey As String = "cNameRule.DecodeThisMatch"

# Call DoCall(zKey, "Function", eQzMode)

if aRuleString <> RuleString Then:
# Me.bConsistent = False
# ' remove Line feeds etc.
# RuleString = Replace(RuleString, vbCr, b)
# RuleString = Replace(RuleString, vbLf, b)
# RuleString = ReplaceAll(RuleString, "% ", "%")

# ' eliminate double blanks in RuleString-String
# RuleString = RemoveDoubleBlanks(RuleString)
if RuleString <> aRuleString Then:
# DecodeThisMatch = True                 ' it has changed
if Not aID(0) Is Nothing Then:
# aOD(0).objDumpMade = -1            ' it is impossible that we did output this
# Me.bConsistent = False                 ' means: Attribute not tested against this Rule
# Me.MatchesList = split(Trim(RuleString), b)
# aRuleString = RuleString               ' with all relevant operator chars inside
# ' split semi clean List: includes operators
# ' remove invalid chars from RuleString-String
# RuleString = RemoveChars(RuleString, "*!%-+|:^()")
# RuleString = Remove(RuleString, "or", b, vbTextCompare)
# RuleString = Remove(RuleString, "and", b, vbTextCompare)
# RuleString = Remove(RuleString, "not", b, vbTextCompare)
# RuleString = RemoveOpAndParm(RuleString, "toendtime")
# RuleString = Trim(RemoveDoubleBlanks(RuleString)) ' reduce again
# ' if change has occurred (without in special chars)
if Me.CleanMatchesString <> RuleString Then:
# ' and Prop*-values are not defined yet, important change
# Me.CleanMatchesString = RuleString     ' all special chars were dropped
# Me.CleanMatches = split(Me.CleanMatchesString, b)
if Me.PropAllRules Is Nothing Then:
# Me.bConsistent = False
elif Me.PropAllRules.RuleType <> InstanceRule Then:
# Me.bConsistent = False                     ' only InstanceRule s can be consistent

# zExit:
# Call DoExit(zKey)
# ProcRet:

# ' remove two words (e.g. Op and its Parm), blank seperated
def removeopandparm():
    # Const zKey As String = "cNameRule.RemoveOpAndParm"

    # Dim i As Long
    # Dim OpParm As String

    # RemoveOpAndParm = FromThis                     ' no change
    # i = InStr(1, FromThis, Op & b, vbTextCompare)
    if i > 0 Then:
    # RemoveOpAndParm = Remove(RemoveOpAndParm, Op, b, vbTextCompare)
    # ' I now is position of parameter!
    # OpParm = Mid(RemoveOpAndParm, i)
    # OpParm = Trunc(1, OpParm, b)
    # RemoveOpAndParm = Remove(RemoveOpAndParm, OpParm, b, vbTextCompare)

    # ProcRet:

# ' Checking a Pattern Match String against ADName to find potential match
def checkpatterninstance():
    # Dim zErr As cErr
    # Const zKey As String = "cNameRule.CheckPatternInstance"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cNameRule")

    # Dim pTail As Boolean
    # Dim pFront As Boolean
    # Dim pLog As Boolean
    # Dim PropPat As String
    # Dim i As Long
    # Dim j As Long
    # Dim Ci As Long

    if Me.MatchOn <> adName Then:
    if LenB(Me.MatchOn) > 0 Then:
    if Me.PatFound <> adName Then:
    # DoVerify False, " classRules??? *** bad!!!"
    # Me.bConsistent = False
    if aTD Is Nothing Then:
    # DoVerify False, " nie!! oder???"
    # Me.bConsistent = False
    if LenB(TrueCritList) = 0 Then:
    if InStr(1, Explain, "Mandatory", vbTextCompare) > 0 _:
    # Or InStr(1, Explain, "Similarity", vbTextCompare) > 0 Then
    # GoTo noMatch

    if Me.bConsistent Then:
    if aTD.adName <> adName Then:
    # Set aTD = GetAttrDsc(aTD.adKey)
    if aTD Is Nothing Then:
    # Set iRules = Nothing                   ' this item does not have such a property (OK)
    # Explain = "not seen: "
    # Me.bConsistent = False
    # GoTo fastexit
    # CheckPatternInstance = Me.RuleMatches
    # Call Get_iRules(aTD)
    # GoTo fastexit

    if isEmpty(Me.CleanMatches) Then               ' operators not relevant for this:
    # Me.MatchOn = vbNullString
    if DebugLogging Then:
    print(Debug.Print "CheckPatternInstance skipped, nothing in CleanMatches list")
    # GoTo FunExit
    if DebugLogging Then:
    print(Debug.Print Format(Timer, "0#####.##"), _)
    # "CheckPatternInstance for ", adName, _
    # "(Prop. " & apropTrueIndex & ")", _
    # " is " & Explain & "?", aOD(aPindex).objItemClassName
    # ' Process operators and criteria
    if isEmpty(MatchesList) Then:
    # GoTo FunExit
    # Ci = LBound(Me.MatchesList)
    # PatFound = RemoveChars(Me.MatchesList(i), Bracket)
    if LenB(PatFound) = 0 Then:
    # GoTo skipMatch
    if InStr(" and or not ", LCase(PatFound)) > 0 Then:
    # GoTo skipMatch
    # ' check Op and Param case(s): list of Ops
    if InStr(" toendtime ", LCase(PatFound)) > 0 Then:
    # i = i + 1
    # GoTo skipMatch
    # PropPat = PatFound
    # PatFound = RemoveChars(PatFound, "+-%|!:") ' remove all operators, () already removed
    if LenB(PatFound) = 0 Then:
    # GoTo skipMatch

    if Not (pFront Or pTail) And Me.CleanMatches(Ci) <> PatFound Then:
    # GoTo skipMatch                         ' nothing compares...:)
    # ' CheckPatternInstance results stored now:
    if adName = PropPat Then:
    # GoTo FoundIt                           ' this is no rule with operator/wildcards
    if adName = PatFound Then:
    # GoTo FoundIt                           ' this is rule with irrelevant operator
    # ' here are complex rules only:
    # j = InStr(adName, PatFound)
    if Not (pFront Or pTail) Then:
    # j = 0                                  ' no pattern to match, and they are not fully matching: NONO
    if j > 0 Then                              ' sting does occur:
    if pFront And pTail Then               ' it is a *middle* (unlikely):
    # GoTo FoundIt
    elif pFront Or (j = 1 And Not pTail) Then ' it starts front or is *front:
    # GoTo FoundIt
    elif pTail And j = 1 Then            ' it is end*:
    # FoundIt:
    # Me.PropWildcard = pTail Or pFront
    # Me.PropTailWild = pTail
    # Me.PropFrontWild = pFront
    # Me.PropMustLog = pLog
    # Me.MatchOn = PropPat
    # Me.RuleMatches = True
    # CheckPatternInstance = True
    # GoTo FunExit
    # Ci = Ci + 1
    # skipMatch:
    # noMatch:
    # PatFound = vbNullString                        ' No matching pattern

    # FunExit:
    # Me.bConsistent = True
    # Set iRules = aTD.adRules
    # fastexit:
    if DebugLogging Then:
    print(Debug.Print Format(Timer, "0#####.##"), _)
    # "CheckPatternInstance out ", adName, _
    # aOD(aPindex).objTypeName, " RuleMatch=" & _
    # CheckPatternInstance;
    if CheckPatternInstance Then:
    print(Debug.Print , " Match auf " & PatFound)
    else:
    print(Debug.Print , " kein Match")
    # anymatch = anymatch Or Me.RuleMatches

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

