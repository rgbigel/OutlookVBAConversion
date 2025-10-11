VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTraceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TKey As String
Attribute TKey.VB_VarUserMemId = 0
Attribute TKey.VB_VarDescription = "Display Key of traced Err"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute TKey.VB_VarUserMemId = 0
' Attribute TKey.VB_VarDescription = "Display Key of traced Err"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public TErr As cErr
Public Tinx As Long
Public TPre As Long
Public TSuc As Long

Public TSrc As String
Public TLne As Long
Public TDet As String
Public TLD As String                               ' atLastInDate   : trace<>non-instance vals in cErr
Public TLS As Double                               ' atLastInSec
Public TPS As Double                               ' atPrevEntrySec
Public TES As Double                               ' atThisEntrySec
Public TLog As String                              ' LogProgress Line
Public TRL As Long                                 ' recursion Level (instance)
Public TCal As String                              ' Stack call situation: A if top of stack and active
' P if recursive and not yet inactive (one other instance=A)
' E exited because not (A or P)

'---------------------------------------------------------------------------------------
' Method : Function TCallState
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Function TCallState() As String
    TCallState = TCal
    Select Case TErr.atCallState
        Case eCExited
            If TCal <> "E" Then                    ' debug.assert not working??
                Debug.Print Tinx, TCal, TLog
                DoVerify False
                TCal = "E"
            End If
        Case Else
            Select Case TCal
                Case "A"
                    DoVerify TErr.atCallState <> eCUndef
                Case "E"
                    DoVerify TErr.atCallState <> eCpaused
                    DoVerify TErr.atCallState <> eCUndef
                Case "P"
                    DoVerify TErr.atCallState = eCpaused
                Case "u"
                    DoVerify TErr.atCallState = eCUndef
                Case Else
                    DoVerify False
            End Select
    End Select
    
    aCallState = TCal
    DoVerify aCallState = TCallState

End Function                                       ' cTraceEntry.TCallState

Property Get TraceSucc(Optional TT As Long) As cTraceEntry
    
Dim TP As Long
Dim ps As Long
Dim aPredTE As cTraceEntry
     
    If TT = 0 Then
        TP = TraceTop
    Else
        TP = TT
    End If
    TP = TP Mod ErrLifeTime
    
    Set aPredTE = C_CallTrace(TP)
    ps = aPredTE.TSuc Mod ErrLifeTime
    
    If ps > 0 And TErr.atCallDepth > 0 Then
        Set TraceSucc = C_CallTrace(ps)
        DoVerify TraceSucc.TErr.atKey = TErr.atKey, "** check integrity"
        If DebugMode Then
            DoVerify aPredTE.TErr.atErrPrev Is TraceSucc.TErr, "check integrity"
        End If
    End If

End Property                                       ' cTraceEntry.TraceSucc Get

Property Get TracePred(Optional TT As Long) As cTraceEntry
    
Dim TP As Long
Dim ps As Long
Dim aPredTE As cTraceEntry
     
    If TT = 0 Then
        TP = TraceTop
    Else
        TP = TT
    End If
    If C_CallTrace.Count >= ErrLifeTime - 1 Then
        TP = TP Mod ErrLifeTime
        If TP < 3 Then                             ' skipping 200..202 (modulus 200)
            TP = TP + 3
        End If
    End If
    
    Set aPredTE = C_CallTrace(TP)
    ps = aPredTE.TPre Mod ErrLifeTime
    
    If ps > 0 Then
        Set TracePred = C_CallTrace(ps)
    End If
    
    aCallState = TCal

End Property                                       ' cTraceEntry.TracePred Get

'---------------------------------------------------------------------------------------
' Method : TraceAdd
' Author : Rolf G. Bercht
' Date   : 20211108@11_47
' Purpose: Put Trace Entry onto Stack setting Tinx
'---------------------------------------------------------------------------------------
Sub TraceAdd(Optional ExplainS As String)
Dim TP As Long
Dim aPredTE As cTraceEntry

    Tinx = TraceTop + 1
    TKey = TErr.atKey
    
    If TraceTop = 0 Or C_CallTrace.Count < ErrLifeTime - 1 Then
        C_CallTrace.Add Me                         ' creating element up to ErrLifeTime-1
        TErr.atTraceStackPos = Tinx
    Else
        TP = Tinx
        TP = TP Mod ErrLifeTime
        If TP < 3 Then                             ' skipping 200..202 (modulus 200)
            TP = TP + 3
            Tinx = Tinx + 3
        End If
        Set aPredTE = C_CallTrace.Item(TP)
        If DebugMode Then
            DoVerify aPredTE.TErr.atCallState = eCExited, _
                     "entry #" & Tinx & b & aPredTE.TErr.atKey _
                     & " is " & CStateNames(aPredTE.TErr.atCallState) _
                     & " ??? It will be kicked from CallTrace!"
        End If
        C_CallTrace.Remove TP                      ' "replace" position Tp
        If TP > C_CallTrace.Count Then
            C_CallTrace.Add Me                     ' creating element up to ErrLifeTime-1
        Else
            C_CallTrace.Add Me, Before:=TP
        End If
        TErr.atTraceStackPos = Tinx
    End If
    
    If Not TErr.atCalledBy Is Nothing Then
        TPre = TErr.atCalledBy.atTraceStackPos
    End If
     
    If TPre > 0 Then                               ' check if TracePreDec is still ok
        TP = TPre Mod ErrLifeTime
        Set aPredTE = TracePred(Tinx)              ' who is the TracePreDec
        If TP <= Tinx Then                         ' probably has become invalid
            If aPredTE Is Nothing Then
                TPre = 0
                GoTo noPre
            Else
                If aPredTE.TErr.atCallState = eCExited Then
                    aPredTE.TPre = 0               ' caller exited ==> not in use, can overwrite
                    GoTo noPre
                ElseIf P_Active.CallMode <> eQnoDef And aPredTE.TErr.atCallState = eCActive Then
                    DoVerify aPredTE.TErr Is E_Active, _
                             "Caller " & aPredTE.TErr.atKey _
                             & " is active (and not Extern.Caller)"
                End If                             ' Paused is OK
            End If
        Else
            DoVerify 1 = 0, "TPre > 0 and alive"
        End If
        aPredTE.TSuc = Tinx
    End If
noPre:
    If TSuc > 1 Then                               ' check if TraceSucc is still ok, ignore Extern.Caller (TSuc=1)
        TP = TPre Mod ErrLifeTime
        Set aPredTE = TraceSucc(Tinx)              ' who is the TraceSucc
        If TP <= aPredTE.TErr.atTraceStackPos Mod ErrLifeTime Then ' probably has become invalid
            If aPredTE Is Nothing Then
                TSuc = 0
            Else
                If aPredTE.TErr.atCallState = eCExited Then ' this is ok, item not in use, can overwrite
                    TSuc = 0
                Else                               ' killing active TraceSucc??
                    DoVerify Not DebugMode, "killing active TraceSucc??"
                    TSuc = 0
                End If
            End If
        Else
            ' TSuc > 0 and alive
        End If
    End If
    If CallLogging And LogZProcs Then
        Call N_ShowProgress(CallNr, TErr.atDsc, "+T", vbNullString, ExplainS)
    End If

FuncExit:
    TraceTop = Tinx
    DoVerify aCallState = TCal, "design check on CallState ???"
End Sub                                            ' cTraceEntry.TraceAdd


