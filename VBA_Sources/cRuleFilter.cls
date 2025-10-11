VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRuleFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Rule Filter settings
Public RulesToFilter As String
Public RuleUbound As Long
Public RulePart As Variant
Public RulesIgnoreFront As Variant                 ' Boolean
Public RulesIgnoreTail As Variant                  ' Boolean
Public RuleLogic As Variant                        ' Boolean

'---------------------------------------------------------------------------------------
' Method : Function RuleFilter
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function RuleFilter(matchname As String) As Boolean ' matchname in/not in RulesToFilter
Dim zErr As cErr
Const zKey As String = "cRuleFilter.RuleFilter"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cRuleFilter")

Dim starter As Long
Dim aLogic As Boolean
Dim i As Long
    If LenB(RulesToFilter) = 0 Then
        RuleFilter = False                         ' nothing is filtered out
    Else
        For i = 0 To RuleUbound                    ' stay in loop until RuleFilter turns true
            starter = InStr(1, RulePart(i), matchname, vbTextCompare)
            aLogic = RuleLogic(i)
            If starter = 0 And i <= RuleUbound Then
                If aLogic Then                     ' we can't determine as true yet
                    GoTo nextPart
                End If
                RuleFilter = aLogic                ' rule not/is filtered out
                If aLogic Then
                    GoTo ProcReturn                ' this saves us a lot of else parts
                End If
            End If
            If RulesIgnoreFront(i) Then
                RuleFilter = aLogic
            Else
                If starter = 1 Then
                    RuleFilter = aLogic
                End If
            End If
            If RulesIgnoreTail(i) Then
                RuleFilter = aLogic                ' it occurs so drop/don,,3,3,3 drop it
            Else
                If Len(matchname) - starter + 1 = Len(RulePart(i)) Then
                    RuleFilter = aLogic
                Else
                    RuleFilter = Not aLogic
                End If
            End If
            If RuleFilter Then
                GoTo ProcReturn                    ' this saves us a lot of else parts
            End If
nextPart:
        Next i
    End If

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                       ' cRuleFilter.RuleFilter

'---------------------------------------------------------------------------------------
' Method : Sub setRuleFilter
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub setRuleFilter(Filter As String)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "cRuleFilter.setRuleFilter"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="cRuleFilter")

Dim i As Long
    RulesToFilter = Filter
    If LenB(Filter) = 0 Then
        RulesToFilter = vbNullString
        RulePart = Array(vbNullString)
        RulesIgnoreFront = True
        RulesIgnoreTail = True
        RuleUbound = UBound(RulePart)
        GoTo ProcReturn
    Else
        RulePart = split(Filter, "|")
        RuleUbound = UBound(RulePart)
    End If
    ReDim RulesIgnoreFront(0 To RuleUbound) As Boolean
    ReDim RulesIgnoreTail(0 To RuleUbound) As Boolean
    ReDim RuleLogic(0 To RuleUbound) As Boolean
    For i = 0 To RuleUbound
        If Left(RulePart(i), 1) = "*" Then
            RulesIgnoreFront(i) = True
            RulePart(i) = Mid(RulePart(i), 2)
        Else
            RulesIgnoreFront(i) = False
        End If
        If Right(RulePart(i), 1) = "*" Then
            RulesIgnoreTail(i) = True
            RulePart(i) = Left(RulePart(i), Len(RulePart(i)) - 1)
        Else
            RulesIgnoreTail(i) = False
        End If
        If Left(RulePart(i), 1) = "^" Then
            RuleLogic(i) = False
        Else
            RuleLogic(i) = True
        End If
    Next i

ProcReturn:
    Call ProcExit(zErr)

End Sub                                            ' cRuleFilter.setRuleFilter

