# Converted from cRuleFilter.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cRuleFilter"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# ' Rule Filter settings
# Public RulesToFilter As String
# Public RuleUbound As Long
# Public RulePart As Variant
# Public RulesIgnoreFront As Variant                 ' Boolean
# Public RulesIgnoreTail As Variant                  ' Boolean
# Public RuleLogic As Variant                        ' Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Function RuleFilter
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def rulefilter():
    # Dim zErr As cErr
    # Const zKey As String = "cRuleFilter.RuleFilter"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cRuleFilter")

    # Dim starter As Long
    # Dim aLogic As Boolean
    # Dim i As Long
    if LenB(RulesToFilter) = 0 Then:
    # RuleFilter = False                         ' nothing is filtered out
    else:
    # starter = InStr(1, RulePart(i), matchname, vbTextCompare)
    # aLogic = RuleLogic(i)
    if starter = 0 And i <= RuleUbound Then:
    if aLogic Then                     ' we can't determine as true yet:
    # GoTo nextPart
    # RuleFilter = aLogic                ' rule not/is filtered out
    if aLogic Then:
    # GoTo ProcReturn                ' this saves us a lot of else parts
    if RulesIgnoreFront(i) Then:
    # RuleFilter = aLogic
    else:
    if starter = 1 Then:
    # RuleFilter = aLogic
    if RulesIgnoreTail(i) Then:
    # RuleFilter = aLogic                ' it occurs so drop/don,,3,3,3 drop it
    else:
    if Len(matchname) - starter + 1 = Len(RulePart(i)) Then:
    # RuleFilter = aLogic
    else:
    # RuleFilter = Not aLogic
    if RuleFilter Then:
    # GoTo ProcReturn                    ' this saves us a lot of else parts

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub setRuleFilter
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def setrulefilter():
    # '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
    # Const zKey As String = "cRuleFilter.setRuleFilter"
    # Dim zErr As cErr

    # Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cRuleFilter")

    # Dim i As Long
    # RulesToFilter = Filter
    if LenB(Filter) = 0 Then:
    # RulesToFilter = vbNullString
    # RulePart = Array(vbNullString)
    # RulesIgnoreFront = True
    # RulesIgnoreTail = True
    # RuleUbound = UBound(RulePart)
    # GoTo ProcReturn
    else:
    # RulePart = split(Filter, "|")
    # RuleUbound = UBound(RulePart)
    # ReDim RulesIgnoreFront(0 To RuleUbound) As Boolean
    # ReDim RulesIgnoreTail(0 To RuleUbound) As Boolean
    # ReDim RuleLogic(0 To RuleUbound) As Boolean
    if Left(RulePart(i), 1) = "*" Then:
    # RulesIgnoreFront(i) = True
    # RulePart(i) = Mid(RulePart(i), 2)
    else:
    # RulesIgnoreFront(i) = False
    if Right(RulePart(i), 1) = "*" Then:
    # RulesIgnoreTail(i) = True
    # RulePart(i) = Left(RulePart(i), Len(RulePart(i)) - 1)
    else:
    # RulesIgnoreTail(i) = False
    if Left(RulePart(i), 1) = "^" Then:
    # RuleLogic(i) = False
    else:
    # RuleLogic(i) = True

    # ProcReturn:
    # Call ProcExit(zErr)

