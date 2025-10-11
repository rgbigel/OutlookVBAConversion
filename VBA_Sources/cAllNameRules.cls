VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAllNameRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsObligMatches As cNameRule
Public clsNeverCompare As cNameRule
Public clsNotDecodable As cNameRule                ' do not try to decode
Public clsSimilarities As cNameRule                ' Properties compared, but not counted as (super-)relevant

Public ARName As String
Attribute ARName.VB_VarUserMemId = 0
Attribute ARName.VB_VarDescription = "Rule Name"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute ARName.VB_VarUserMemId = 0
' Attribute ARName.VB_VarDescription = "Rule Name"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public RuleInstanceValid As Boolean
Public RuleIsSpecific As Boolean
Public RuleType As String                          ' Defaultrule, ClassRules, Itemrule
Public RuleObjDsc As cObjDsc

Private Sub Class_Initialize()
    If clsObligMatches Is Nothing Then
        Set clsObligMatches = New cNameRule
        Set clsNeverCompare = New cNameRule
        Set clsNotDecodable = New cNameRule
        Set clsSimilarities = New cNameRule
    End If
End Sub                                            ' cAllNameRules.Class_Initialize

Private Sub Class_Terminate()

Const zKey As String = "cAllNameRules.Class_Terminate"
    Call DoCall(zKey, tSub, eQzMode)
   
    Set clsObligMatches = Nothing
    Set clsNeverCompare = Nothing
    Set clsNotDecodable = Nothing
    Set clsSimilarities = Nothing

zExit:
    Call DoExit(zKey)

End Sub                                            ' cAllNameRules.Class_Terminate

'---------------------------------------------------------------------------------------
' Method : CheckAllRules
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: checks all the pattern instance types (mandatory, nocompare, nodecode, similar)
'---------------------------------------------------------------------------------------
Sub CheckAllRules(SubListName As Variant, sDC As String)
Dim zErr As cErr
Const zKey As String = "cAllNameRules.CheckAllRules"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")

Dim anyruleMatch As Boolean

    Call clsObligMatches.CheckPatternInstance(SubListName, anyruleMatch, sDC & "Obligatory")
    Call clsNeverCompare.CheckPatternInstance(SubListName, anyruleMatch, sDC & "No Compare")
    Call clsNotDecodable.CheckPatternInstance(SubListName, anyruleMatch, sDC & "can't Dec.")
    Call clsSimilarities.CheckPatternInstance(SubListName, anyruleMatch, sDC & "Similarity")
    RuleIsSpecific = anyruleMatch
    aTD.adRules.RuleInstanceValid = True
    IgString = B2

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' cAllNameRules.CheckAllRules

'---------------------------------------------------------------------------------------
' Method : CheckAllRulesInList
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: .
'---------------------------------------------------------------------------------------
Sub CheckAllRulesInList(aList As Variant, Optional sDC As String)
Dim zErr As cErr
Const zKey As String = "cAllNameRules.CheckAllRulesInList"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")
    
Dim SubListName As Variant

    For Each SubListName In aList
        Call CheckAllRules(SubListName, sDC)
    Next SubListName
    
ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' cAllNameRules.CheckAllRulesInList

'---------------------------------------------------------------------------------------
' Method : AllRulesCopy
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: Deep AllRulesCopy
'---------------------------------------------------------------------------------------
'    Destination normally Is iRules, Source Is sRules
Sub AllRulesCopy(dRuleType As String, ByRef S As cAllNameRules, withMatchBits As Boolean)
Dim zErr As cErr
Const zKey As String = "cAllNameRules.AllRulesCopy"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="cAllNameRules")

    If LenB(dRuleType) = 0 Then
        Me.RuleType = S.RuleType
    Else
        Me.RuleType = dRuleType
    End If
    Me.ARName = S.ARName
    Me.RuleInstanceValid = False
    Me.RuleIsSpecific = S.RuleIsSpecific           ' for other than iRules = False
    Set RuleObjDsc = S.RuleObjDsc

    Call Me.clsObligMatches.RuleCopy(S.clsObligMatches, withMatchBits)
    Call Me.clsNeverCompare.RuleCopy(S.clsNeverCompare, withMatchBits)
    Call Me.clsNotDecodable.RuleCopy(S.clsNotDecodable, withMatchBits)
    Call Me.clsSimilarities.RuleCopy(S.clsSimilarities, withMatchBits)
    If withMatchBits Then
        Me.RuleInstanceValid = S.RuleInstanceValid
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                            ' cAllNameRules.AllRulesCopy

'---------------------------------------------------------------------------------------
' Method : AllRulesClone
' Author : Rolf-Günther Bercht
' Date   : 20211108@11_47
' Purpose: AllRulesClone Class (deep AllRulesCopy)
'---------------------------------------------------------------------------------------
Function AllRulesClone(dRuleType As String, thisObjDsc As cObjDsc, withMatchBits As Boolean) As cAllNameRules
Dim nCloneAllNameRules As cAllNameRules

    Set nCloneAllNameRules = New cAllNameRules     ' new Rules with sub classes
    aBugVer = Not thisObjDsc Is Nothing
    If Not DoVerify(aBugVer, "AllRulesClone for undefined thisObjDsc???") Then
        Me.ARName = thisObjDsc.objTypeName         ' cloning for thisObjDsc as default
    End If
    Call nCloneAllNameRules.AllRulesCopy(dRuleType, Me, withMatchBits)
    Set AllRulesClone = nCloneAllNameRules
    Set nCloneAllNameRules.RuleObjDsc = thisObjDsc ' parent link for Rules
    If dRuleType = "ClassRules" Then
        If Not thisObjDsc Is Nothing Then
            Set thisObjDsc.objClsRules = nCloneAllNameRules
        End If
    End If
    
FuncExit:
    Set nCloneAllNameRules = Nothing
End Function                                       ' cAllNameRules.AllRulesClone

