# Converted from cAllNameRules.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cAllNameRules"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public clsObligMatches As cNameRule
# Public clsNeverCompare As cNameRule
# Public clsNotDecodable As cNameRule                ' do not try to decode
# Public clsSimilarities As cNameRule                ' Properties compared, but not counted as (super-)relevant

# Public ARName As String
# Attribute ARName.VB_VarUserMemId = 0
# Attribute ARName.VB_VarDescription = "Rule Name"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute ARName.VB_VarUserMemId = 0
# ' Attribute ARName.VB_VarDescription = "Rule Name"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public RuleInstanceValid As Boolean
# Public RuleIsSpecific As Boolean
# Public RuleType As String                          ' Defaultrule, ClassRules, Itemrule
# Public RuleObjDsc As cObjDsc

# Private Sub Class_Initialize()
if clsObligMatches Is Nothing Then:
# Set clsObligMatches = New cNameRule
# Set clsNeverCompare = New cNameRule
# Set clsNotDecodable = New cNameRule
# Set clsSimilarities = New cNameRule

# Private Sub Class_Terminate()

# Const zKey As String = "cAllNameRules.Class_Terminate"
# Call DoCall(zKey, tSub, eQzMode)

# Set clsObligMatches = Nothing
# Set clsNeverCompare = Nothing
# Set clsNotDecodable = Nothing
# Set clsSimilarities = Nothing

# zExit:
# Call DoExit(zKey)


# '---------------------------------------------------------------------------------------
# ' Method : CheckAllRules
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: checks all the pattern instance types (mandatory, nocompare, nodecode, similar)
# '---------------------------------------------------------------------------------------
def checkallrules():
    # Dim zErr As cErr
    # Const zKey As String = "cAllNameRules.CheckAllRules"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")

    # Dim anyruleMatch As Boolean

    # Call clsObligMatches.CheckPatternInstance(SubListName, anyruleMatch, sDC & "Obligatory")
    # Call clsNeverCompare.CheckPatternInstance(SubListName, anyruleMatch, sDC & "No Compare")
    # Call clsNotDecodable.CheckPatternInstance(SubListName, anyruleMatch, sDC & "can't Dec.")
    # Call clsSimilarities.CheckPatternInstance(SubListName, anyruleMatch, sDC & "Similarity")
    # RuleIsSpecific = anyruleMatch
    # aTD.adRules.RuleInstanceValid = True
    # IgString = B2

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : CheckAllRulesInList
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: .
# '---------------------------------------------------------------------------------------
def checkallrulesinlist():
    # Dim zErr As cErr
    # Const zKey As String = "cAllNameRules.CheckAllRulesInList"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")

    # Dim SubListName As Variant

    for sublistname in alist:
    # Call CheckAllRules(SubListName, sDC)

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : AllRulesCopy
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Deep AllRulesCopy
# '---------------------------------------------------------------------------------------
# '    Destination normally Is iRules, Source Is sRules
def allrulescopy():
    # Dim zErr As cErr
    # Const zKey As String = "cAllNameRules.AllRulesCopy"
    # Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")

    if LenB(dRuleType) = 0 Then:
    # Me.RuleType = S.RuleType
    else:
    # Me.RuleType = dRuleType
    # Me.ARName = S.ARName
    # Me.RuleInstanceValid = False
    # Me.RuleIsSpecific = S.RuleIsSpecific           ' for other than iRules = False
    # Set RuleObjDsc = S.RuleObjDsc

    # Call Me.clsObligMatches.RuleCopy(S.clsObligMatches, withMatchBits)
    # Call Me.clsNeverCompare.RuleCopy(S.clsNeverCompare, withMatchBits)
    # Call Me.clsNotDecodable.RuleCopy(S.clsNotDecodable, withMatchBits)
    # Call Me.clsSimilarities.RuleCopy(S.clsSimilarities, withMatchBits)
    if withMatchBits Then:
    # Me.RuleInstanceValid = S.RuleInstanceValid

    # ProcReturn:
    # Call ProcExit(zErr)

    # pExit:

# '---------------------------------------------------------------------------------------
# ' Method : AllRulesClone
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: AllRulesClone Class (deep AllRulesCopy)
# '---------------------------------------------------------------------------------------
def allrulesclone():
    # Dim nCloneAllNameRules As cAllNameRules

    # Set nCloneAllNameRules = New cAllNameRules     ' new Rules with sub classes
    # aBugVer = Not thisObjDsc Is Nothing
    if Not DoVerify(aBugVer, "AllRulesClone for undefined thisObjDsc???") Then:
    # Me.ARName = thisObjDsc.objTypeName         ' cloning for thisObjDsc as default
    # Call nCloneAllNameRules.AllRulesCopy(dRuleType, Me, withMatchBits)
    # Set AllRulesClone = nCloneAllNameRules
    # Set nCloneAllNameRules.RuleObjDsc = thisObjDsc ' parent link for Rules
    if dRuleType = "ClassRules" Then:
    if Not thisObjDsc Is Nothing Then:
    # Set thisObjDsc.objClsRules = nCloneAllNameRules

    # FuncExit:
    # Set nCloneAllNameRules = Nothing
