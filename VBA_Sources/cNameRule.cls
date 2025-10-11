VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RuleMatches As Boolean
Public MatchOn As String                           ' last PropertyNameX we Matched or vbNullString if no Match
Public PatFound As String                          ' explain pattern Match rule
Public bConsistent As Boolean                      ' change indicator

Public aRuleString As String                       ' Treat private, use GetRuleString or ChangeTo
Public CritRestrictString As String                ' Syntax is for Restrict-operation
Public CritFilterString As String                  ' Syntax is for Filter-Operation
Public MatchesList As Variant
Public CleanMatchesString As String
Public CleanMatches As Variant

Public PropWildcard As Boolean
Public PropTailWild As Boolean
Public PropFrontWild As Boolean
Public PropMustLog As Boolean
Public PropVisibility As Boolean

Public PropAllRules As cAllNameRules

'---------------------------------------------------------------------------------------
' Method : Sub RuleClear
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub RuleClear()                             ' Parent is NOT RuleCleared
Dim zErr As cErr
Const zKey As String = "cNameRule.RuleClear"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cNameRule")

    RuleMatches = False
    MatchOn = vbNullString
    PatFound = vbNullString
    aRuleString = vbNullString
    MatchesList = Empty
    CleanMatchesString = vbNullString
    CleanMatches = Empty
    
    If Parent Is Nothing Then
        If bConsistent Then
            DoVerify False
        End If
    End If
    
    bConsistent = False
    PropWildcard = False
    PropTailWild = False
    PropFrontWild = False
    PropMustLog = False
    PropVisibility = False

ProcReturn:
    Call ProcExit(zErr)

End Sub                                            ' cNameRule.RuleClear

'---------------------------------------------------------------------------------------
' Method : Sub RuleCopy
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub RuleCopy(ByRef S As cNameRule, withMatchBits As Boolean)

Const zKey As String = "cNameRule.RuleCopy"

'------------------- gated Entry -------------------------------------------------------
Static Recursive As Boolean

    aBugVer = Not Recursive
    If DoVerify(aBugVer, "Forbidden recursion from " _
                        & P_Active.DbgId & " => " & zKey) Then
        GoTo ProcRet
    End If
    Recursive = True                               ' restored by    Recursive = False ProcRet:

    Call DoCall(zKey, tSub, eQzMode)

    If S.PropAllRules Is Nothing Then
        If DebugLogging And S.bConsistent And Not withMatchBits Then
            If DebugMode Then DoVerify False
        End If
    End If
    
    If withMatchBits Then
        Me.RuleMatches = S.RuleMatches
        Me.bConsistent = S.bConsistent
    Else
        Me.RuleMatches = False
        Me.bConsistent = False
    End If
    Me.MatchOn = S.MatchOn
    Me.PatFound = S.PatFound
    Me.aRuleString = S.aRuleString
    Me.MatchesList = S.MatchesList
    Me.CleanMatchesString = S.CleanMatchesString
    Me.CleanMatches = S.CleanMatches
    Me.PropWildcard = S.PropWildcard
    Me.PropTailWild = S.PropTailWild
    Me.PropFrontWild = S.PropFrontWild
    Me.PropMustLog = S.PropMustLog
    Me.PropVisibility = S.PropVisibility
    Set Me.PropAllRules = S.PropAllRules

zExit:
    Call DoExit(zKey)
    Recursive = False
    
ProcRet:
End Sub                                            ' cNameRule.RuleCopy

'---------------------------------------------------------------------------------------
' Method : Clone
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function Clone(withMatchBits As Boolean) As cNameRule

Dim nCloneNameRule As cNameRule

    Set nCloneNameRule = New cNameRule
    Call nCloneNameRule.RuleCopy(Me, withMatchBits)
    Set Clone = nCloneNameRule

End Function                                       ' cNameRule.Clone

Public Property Get GetMatchString() As String
    
    GetMatchString = Trim(Me.aRuleString)

End Property                                       ' cNameRule.GetMatchString Get

Public Property Let ChangeTo(S As String)

Dim HasChanged As Boolean

    If LenB(S) = 0 Then
        If aRuleString = S Then
            HasChanged = False
        Else
            HasChanged = True
            SelectedAttributes = vbNullString      ' append changes later
        End If
        Me.RuleClear
    Else
        HasChanged = Me.DecodeThisMatch(S)
        Me.RuleMatches = HasChanged
        If Not iRules Is Nothing Then
            If HasChanged Then                     ' something has changed
                iRules.RuleInstanceValid = False
                SelectedAttributes = vbNullString  ' append changes later
            End If
        End If
    End If

End Property                                       ' cNameRule.ChangeTo Let

'---------------------------------------------------------------------------------------
' Method : Function DecodeThisMatch
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Function DecodeThisMatch(ByVal RuleString As String) As Boolean
'--- Proc MAY ONLY CALL Z_Type PROCS
Const zKey As String = "cNameRule.DecodeThisMatch"

    Call DoCall(zKey, "Function", eQzMode)

    If aRuleString <> RuleString Then
        Me.bConsistent = False
        ' remove Line feeds etc.
        RuleString = Replace(RuleString, vbCr, b)
        RuleString = Replace(RuleString, vbLf, b)
        RuleString = ReplaceAll(RuleString, "% ", "%")
        
        ' eliminate double blanks in RuleString-String
        RuleString = RemoveDoubleBlanks(RuleString)
        If RuleString <> aRuleString Then
            DecodeThisMatch = True                 ' it has changed
            If Not aID(0) Is Nothing Then
                aOD(0).objDumpMade = -1            ' it is impossible that we did output this
            End If
            Me.bConsistent = False                 ' means: Attribute not tested against this Rule
            Me.MatchesList = split(Trim(RuleString), b)
            aRuleString = RuleString               ' with all relevant operator chars inside
            ' split semi clean List: includes operators
        End If
        ' remove invalid chars from RuleString-String
        RuleString = RemoveChars(RuleString, "*!%-+|:^()")
        RuleString = Remove(RuleString, "or", b, vbTextCompare)
        RuleString = Remove(RuleString, "and", b, vbTextCompare)
        RuleString = Remove(RuleString, "not", b, vbTextCompare)
        RuleString = RemoveOpAndParm(RuleString, "toendtime")
        RuleString = Trim(RemoveDoubleBlanks(RuleString)) ' reduce again
        ' if change has occurred (without in special chars)
        If Me.CleanMatchesString <> RuleString Then
            ' and Prop*-values are not defined yet, important change
            Me.CleanMatchesString = RuleString     ' all special chars were dropped
            Me.CleanMatches = split(Me.CleanMatchesString, b)
        End If
    End If
    If Me.PropAllRules Is Nothing Then
        Me.bConsistent = False
    ElseIf Me.PropAllRules.RuleType <> InstanceRule Then
        Me.bConsistent = False                     ' only InstanceRule s can be consistent
    End If

zExit:
    Call DoExit(zKey)
ProcRet:
End Function                                       ' cNameRule.DecodeThisMatch

' remove two words (e.g. Op and its Parm), blank seperated
Function RemoveOpAndParm(FromThis As String, Op As String) As String
Const zKey As String = "cNameRule.RemoveOpAndParm"

Dim i As Long
Dim OpParm As String
    
    RemoveOpAndParm = FromThis                     ' no change
    i = InStr(1, FromThis, Op & b, vbTextCompare)
    If i > 0 Then
        RemoveOpAndParm = Remove(RemoveOpAndParm, Op, b, vbTextCompare)
        ' I now is position of parameter!
        OpParm = Mid(RemoveOpAndParm, i)
        OpParm = Trunc(1, OpParm, b)
        RemoveOpAndParm = Remove(RemoveOpAndParm, OpParm, b, vbTextCompare)
    End If

ProcRet:
End Function                                       ' cNameRule.RemoveOpAndParm

' Checking a Pattern Match String against ADName to find potential match
Function CheckPatternInstance(adName As Variant, anymatch As Boolean, Explain As String) As Boolean
Dim zErr As cErr
Const zKey As String = "cNameRule.CheckPatternInstance"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cNameRule")

Dim pTail As Boolean
Dim pFront As Boolean
Dim pLog As Boolean
Dim PropPat As String
Dim i As Long
Dim j As Long
Dim Ci As Long

    If Me.MatchOn <> adName Then
        If LenB(Me.MatchOn) > 0 Then
            If Me.PatFound <> adName Then
                DoVerify False, " classRules??? *** bad!!!"
            End If
        End If
        Me.bConsistent = False
    End If
    If aTD Is Nothing Then
        DoVerify False, " nie!! oder???"
        Me.bConsistent = False
    End If
    If LenB(TrueCritList) = 0 Then
        If InStr(1, Explain, "Mandatory", vbTextCompare) > 0 _
        Or InStr(1, Explain, "Similarity", vbTextCompare) > 0 Then
            GoTo noMatch
        End If
    End If
    
    If Me.bConsistent Then
        If aTD.adName <> adName Then
            Set aTD = GetAttrDsc(aTD.adKey)
        End If
        If aTD Is Nothing Then
            Set iRules = Nothing                   ' this item does not have such a property (OK)
            Explain = "not seen: "
            Me.bConsistent = False
            GoTo fastexit
        End If
        CheckPatternInstance = Me.RuleMatches
        Call Get_iRules(aTD)
        GoTo fastexit
    End If
    
    If isEmpty(Me.CleanMatches) Then               ' operators not relevant for this
        Me.MatchOn = vbNullString
        If DebugLogging Then
            Debug.Print "CheckPatternInstance skipped, nothing in CleanMatches list"
        End If
        GoTo FunExit
    End If
    If DebugLogging Then
        Debug.Print Format(Timer, "0#####.##"), _
        "CheckPatternInstance for ", adName, _
        "(Prop. " & apropTrueIndex & ")", _
                  " is " & Explain & "?", aOD(aPindex).objItemClassName
    End If
    ' Process operators and criteria
    If isEmpty(MatchesList) Then
        GoTo FunExit
    End If
    Ci = LBound(Me.MatchesList)
    For i = LBound(Me.MatchesList) To UBound(Me.MatchesList)
        PatFound = RemoveChars(Me.MatchesList(i), Bracket)
        If LenB(PatFound) = 0 Then
            GoTo skipMatch
        End If
        If InStr(" and or not ", LCase(PatFound)) > 0 Then
            GoTo skipMatch
        End If
        ' check Op and Param case(s): list of Ops
        If InStr(" toendtime ", LCase(PatFound)) > 0 Then
            i = i + 1
            GoTo skipMatch
        End If
        PropPat = PatFound
        PatFound = RemoveChars(PatFound, "+-%|!:") ' remove all operators, () already removed
        If LenB(PatFound) = 0 Then
            GoTo skipMatch
        End If
        
        If Not (pFront Or pTail) And Me.CleanMatches(Ci) <> PatFound Then
            GoTo skipMatch                         ' nothing compares...:)
        End If
        ' CheckPatternInstance results stored now:
        If adName = PropPat Then
            GoTo FoundIt                           ' this is no rule with operator/wildcards
        End If
        If adName = PatFound Then
            GoTo FoundIt                           ' this is rule with irrelevant operator
        End If
        ' here are complex rules only:
        j = InStr(adName, PatFound)
        If Not (pFront Or pTail) Then
            j = 0                                  ' no pattern to match, and they are not fully matching: NONO
        End If
        If j > 0 Then                              ' sting does occur
            If pFront And pTail Then               ' it is a *middle* (unlikely)
                GoTo FoundIt
            ElseIf pFront Or (j = 1 And Not pTail) Then ' it starts front or is *front
                GoTo FoundIt
            ElseIf pTail And j = 1 Then            ' it is end*
FoundIt:
                Me.PropWildcard = pTail Or pFront
                Me.PropTailWild = pTail
                Me.PropFrontWild = pFront
                Me.PropMustLog = pLog
                Me.MatchOn = PropPat
                Me.RuleMatches = True
                CheckPatternInstance = True
                GoTo FunExit
            End If
        End If
        Ci = Ci + 1
skipMatch:
    Next i
noMatch:
    PatFound = vbNullString                        ' No matching pattern
    
FunExit:
    Me.bConsistent = True
    Set iRules = aTD.adRules
fastexit:
    If DebugLogging Then
        Debug.Print Format(Timer, "0#####.##"), _
        "CheckPatternInstance out ", adName, _
        aOD(aPindex).objTypeName, " RuleMatch=" & _
                                 CheckPatternInstance;
        If CheckPatternInstance Then
            Debug.Print , " Match auf " & PatFound
        Else
            Debug.Print , " kein Match"
        End If
    End If
    anymatch = anymatch Or Me.RuleMatches

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cNameRule.CheckPatternInstance


