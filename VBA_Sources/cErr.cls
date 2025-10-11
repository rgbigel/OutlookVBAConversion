VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public atKey As String                             ' this is the string value of the cErr
Attribute atKey.VB_VarUserMemId = 0
Attribute atKey.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute atKey.VB_VarUserMemId = 0
' Attribute atKey.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public atFastMode As Boolean                       ' the real value at Entry to the Proc, used for exit
Public atCallState As eCState                      ' if <= eCPaused, the atDsc did not Exit
                                            
Public atCalledBy As cErr                          ' immediate Caller
Public atDsc As cProcItem                          ' cProcDsc attached to cErr, unique for all instances
Public at_Live As Dictionary                       ' DoCall defined Items: cCallEnv from LiveCallStack
Public atCatDict As String                         ' list of Procs from at_Live

Public atLastInDate As String                      ' date function
Public atLastInSec As Double                       ' timer [sec]
Public atPrevEntrySec As Double                    ' always >0, time of last entry
Public atThisEntrySec As Double                    ' may be 0 if called another entry, activated
                                                   ' again if called entry exits.
                                                   ' Time consumed is accumulated in atDsc.TotalProcTime

Public atMessage As String
Public atProcIndex As Long

Public atErrPrev As cErr                           ' back chain of same proc ???
Public atTraceStackPos As Long                     ' Position of the ProcErr in C_TraceStack

Public atRecursionLvl As Long                      ' incremented on every DoCall until DoExit
Public atCallDepth As Long                         ' position on the active call stack
Public atRecursionOK As Boolean
Public atShowStack As String

Public atFuncResult As String

Public atLiveLevel As Long                         ' Live Stack level of call

' Error data used for controlling the ErrrorHandlerModule

Public DebugState As Boolean
Public EventBlock As Boolean                       ' Events are blocked During current environment

Public NrMT As String                              ' short info for log purposes
Public ErrSnoCatch As Boolean                      ' No ErrHandler Recursion;  NO N_PublishBugState
Public ErrNoRec As Boolean                         ' No ErrHandler Recursion; use N_PublishBugState
Public errNumber As Long
Public Description As String
Public Source As String

Public FoundBadErrorNr As Long

Public Permitted As Variant
Public Explanations As String
Public Reasoning As String

Property Let Permit(allow As Variant)
    Permitted = allow
    T_DC.DCAllowedMatch = allow
    If isEmpty(allow) Then                         ' end of handled "Try"
        MayChangeErr = True
        Permitted = Empty
        errNumber = 0
        FoundBadErrorNr = 0
        Description = vbNullString
        Explanations = vbNullString
        Reasoning = vbNullString
        If E_Active Is Me Then
            ErrSnoCatch = False
            Z§ErrSnoCatch = False
        Else
            ' ErrSnoCatch   remains
            ' Z§ErrSnoCatch remains
        End If
        ErrorCaught = 0                            ' raw error last set by N_OnError
            
        If Not E_Active Is E_AppErr Then
            With E_AppErr
                .Permitted = Empty
                .errNumber = 0
                .FoundBadErrorNr = 0
                .Description = vbNullString
                .Explanations = vbNullString
            End With                               ' E_AppErr
        End If
        If ErrStatusFormUsable Then
            frmErrStatus.fAcceptableErrors = "No acceptable Errors"
        End If
    ElseIf ErrStatusFormUsable Then
        frmErrStatus.fAcceptableErrors = "Acceptable Errors: " & allow
    End If
    If MayChangeErr Then
        Err.Clear
    End If
End Property                                       ' cErr Permit Let

'---------------------------------------------------------------------------------------
' Method : Sub cPrint
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Sub cPrint(Optional otherErr As cErr, Optional All As Boolean, Optional DoDebug As Boolean = True)

Dim SbS As Boolean                                 ' Side-by-Side used if compare to otherErr
Dim oLine As String
Dim hLine As String
Dim lLine As Long

    If Not otherErr Is Nothing Then
        If Not otherErr Is Me Then
            SbS = True                             ' Side-by-Side
        End If
    End If
    If DoDebug Then
        GoTo doPrint
    End If
    If DebugMode Or DebugLogging Then
doPrint:
        lLine = lKeyM + 20
        oLine = String(20, "-") & LString(" Error Data ", lKeyM)
        If SbS Then
            hLine = LString(oLine, lLine) & " | " & String(lLine, "-")
        Else
            hLine = oLine
        End If
        Debug.Print hLine
        
        oLine = LString("Key", 20) & atKey
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atKey
        End If
        Debug.Print oLine
        
        oLine = LString("CallDepth", 20) & atCallDepth
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atCallDepth
        End If
        Debug.Print oLine
        
        oLine = LString("LiveLevel", 20) & atLiveLevel
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atLiveLevel
        End If
        Debug.Print oLine
        
        oLine = LString("EventBlock", 20) & EventBlock
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.EventBlock
        End If
        Debug.Print oLine
        
        oLine = LString("LiveLevel", 20) & atLiveLevel
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atLiveLevel
        End If
        Debug.Print oLine
        
        oLine = LString("CallState", 20) & atCallState & " = " & CStateNames(atCallState)
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atCallState & " = " & CStateNames(otherErr.atCallState)
        End If
        Debug.Print oLine
        
        oLine = LString("CalledBy", 20) & atCalledBy.atKey
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atCalledBy.atKey
        End If
        Debug.Print oLine
        
        oLine = LString("TraceStackPos", 20) & atTraceStackPos
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atTraceStackPos
        End If
        Debug.Print oLine
        
        oLine = LString("RecursionLvl", 20) & atRecursionLvl
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atRecursionLvl
        End If
        Debug.Print oLine
        
        oLine = LString("ProcIndex", 20) & atProcIndex
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atProcIndex
        End If
        Debug.Print oLine
        
        If atProcIndex >= 0 Then
            If D_ErrInterface.Items(atProcIndex).Key = atDsc.Key Then
                oLine = vbNullString
            Else
                oLine = LString("Proc/Index mismatch", 20) _
                      & D_ErrInterface.Items(atProcIndex).Key
            End If
        ElseIf atProcIndex < -1 Then
            If D_ErrInterface.Items(2 - atProcIndex).Key = atDsc.Key Then
                oLine = vbNullString
            Else
                oLine = LString("Proc/Index mismatch", 20) _
                      & D_ErrInterface.Items(2 - atProcIndex).Key
            End If
        End If
        
        If SbS Then
            If otherErr.atProcIndex > inv Then
                If D_ErrInterface.Items(otherErr.atProcIndex).Key = otherErr.atDsc.Key Then
                    If LenB(oLine) > 0 Then
                        oLine = LString(oLine, lLine) _
                              & " | " & LString("Proc/Index match", 20) _
                              & D_ErrInterface.Items(otherErr.atProcIndex).Key
                    End If
                Else
                    If LenB(oLine) = 0 Then
                        oLine = LString("Proc/Index match", 20) _
                              & D_ErrInterface.Items(atProcIndex).Key
                    End If
                    oLine = LString(oLine, lLine) _
                          & " | " & D_ErrInterface.Items(otherErr.atProcIndex).Key
                End If
            End If
            If otherErr.atProcIndex < -1 Then
                If D_ErrInterface.Items(2 - otherErr.atProcIndex).Key = otherErr.atDsc.Key Then
                    If LenB(oLine) > 0 Then
                        oLine = LString(oLine, lLine) _
                              & " | " & LString("Proc/Index match", 20) _
                              & D_ErrInterface.Items(2 - otherErr.atProcIndex).Key
                    End If
                Else
                    If LenB(oLine) = 0 Then
                        oLine = LString("Proc/Index match", 20) _
        & D_ErrInterface.Items(2 - atProcIndex).Key
                    End If
                    oLine = LString(oLine, lLine) _
                          & " | " & D_ErrInterface.Items(2 - otherErr.atProcIndex).Key
                End If
            End If
        End If
        If LenB(oLine) > 0 Then
            Debug.Print oLine
        End If
        
        oLine = LString("Parent Mode", 20) _
        & atDsc.CallMode & " = " & atDsc.ModeName
        If SbS Then
            oLine = LString(oLine, lLine) _
                  & " | " & otherErr.atDsc.CallMode _
                  & " = " & otherErr.atDsc.ModeName
        End If
        Debug.Print oLine
        
        oLine = LString("NrMT", 20) & NrMT
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.NrMT
        End If
        Debug.Print oLine
                
        oLine = LString("Last In Date", 20) & atLastInDate
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atLastInDate
        End If
        Debug.Print oLine
        
        oLine = LString("Last In Sec", 20) & FormatRight(atLastInSec, 14, 8)
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atLastInSec
        End If
        Debug.Print oLine
        
        oLine = LString("Time Prev Entry", 20) & FormatRight(atPrevEntrySec, 14, 8)
        If SbS Then
            oLine = LString(oLine, lLine) _
                  & " | " & FormatRight(otherErr.atPrevEntrySec, 14, 8)
        End If
        Debug.Print oLine
        
        oLine = LString("Time This Entry", 20) & FormatRight(atThisEntrySec, 14, 8)
        If SbS Then
            oLine = LString(oLine, lLine) _
                  & " | " & FormatRight(otherErr.atThisEntrySec, 14, 8)
        End If
        Debug.Print oLine
        
        oLine = LString("Live Dictionary", 20) & LString(atCatDict, 30)
        If SbS Then
            oLine = LString(oLine, lLine) _
                  & " | " & Left(otherErr.atCatDict, 30)
        End If
        Debug.Print oLine
                
        oLine = LString("Prev Instance", 20)
        If atErrPrev Is Nothing Then
            oLine = oLine & "is Nothing"
        Else
            oLine = oLine & " | " & atErrPrev.atKey
        End If
        If SbS Then
            If otherErr.atErrPrev Is Nothing Then
                oLine = LString(oLine, lLine) & " | " & "is Nothing"
            Else
                oLine = LString(oLine, lLine) & " | " & otherErr.atErrPrev.atKey
            End If
        End If
        Debug.Print oLine
        
        oLine = LString("FuncResult", 20)
        If LenB(atFuncResult) > 0 Then
            oLine = oLine & atFuncResult
        End If
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atFuncResult
        End If
        Debug.Print oLine
                
        If atDsc.Key <> atKey Or otherErr.atDsc.Key <> otherErr.atKey Then
            oLine = LString("ProcKey mismatch", 20) & "atDsc = " & atDsc.Key
            If SbS Then
                oLine = LString(oLine, lLine) & " | " & otherErr.atDsc.Key
            End If
            Debug.Print oLine
        End If
        
        If All Or DebugMode Or errNumber <> 0 Or otherErr.errNumber <> 0 _
        Or FoundBadErrorNr <> 0 Or otherErr.FoundBadErrorNr <> 0 Then
            oLine = LString("errNumber", 20) & errNumber
            If SbS Then
                oLine = LString(oLine, lLine) & " | " & otherErr.errNumber
            End If
            Debug.Print oLine
            
            oLine = LString("Description", 20) & Description
            If SbS Then
                oLine = LString(oLine, lLine) & " | " & otherErr.Description
            End If
            Debug.Print oLine
            
            oLine = LString("Bad Error", 20) & FoundBadErrorNr
            If SbS Then
                oLine = LString(oLine, lLine) & " | " & otherErr.FoundBadErrorNr
            End If
            Debug.Print oLine
        End If
        
        oLine = LString("Permit", 20) & Permitted
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.Permitted
        End If
        Debug.Print oLine
        
        oLine = LString("ErrSnoCatch", 20) & ErrSnoCatch
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.ErrSnoCatch
        End If
        Debug.Print oLine
                
        oLine = LString("DebugState", 20) & DebugState
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.DebugState
        End If
        Debug.Print oLine
        
        oLine = LString("Explanations", 20) & Explanations
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.Explanations
        End If
        Debug.Print oLine
        
        oLine = LString("Reasoning", 20) & Reasoning
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.Reasoning
        End If
        Debug.Print oLine
        
        oLine = LString("atShowStack", 20) & atShowStack
        If SbS Then
            oLine = LString(oLine, lLine) & " | " & otherErr.atShowStack
        End If
        Debug.Print oLine
        
        Debug.Print hLine & " End"
        
    End If

End Sub                                            ' cErr.cPrint

Property Get TraceEntry() As cTraceEntry
Dim TP As Long
     
    TP = atTraceStackPos Mod ErrLifeTime

    If TP > 0 Then
        Set TraceEntry = C_CallTrace(TP)
    End If
End Property                                       ' cErr.TraceEntry Get

Property Let TraceSucc(TraceSucc As Long)
Dim TP As Long
     
    TP = atTraceStackPos Mod ErrLifeTime
    
    If TP > 0 Then
        C_CallTrace.Item(TP).TSuc = TraceSucc
    End If
End Property                                       ' cErr.TraceSucc Let

Property Get TraceSucc() As Long
Dim TP As Long
     
    TP = atTraceStackPos Mod ErrLifeTime

    If TP > 0 Then
        TraceSucc = C_CallTrace.Item(TP).TSuc
    End If
End Property                                       ' cErr.TraceSucc Get

Property Let TracePreDec(Pre As Long)
Dim TP As Long
     
    TP = atTraceStackPos Mod ErrLifeTime
     
    If TP > 0 Then
        C_CallTrace.Item(TP).TSuc = Pre
    End If
End Property                                       ' cErr.TracePreDec Let

Property Get TracePreDec() As Long
Dim TP As Long
     
    TP = atTraceStackPos Mod ErrLifeTime
    
    If TP > 0 Then
        TracePreDec = C_CallTrace.Item(TP).TPre
    End If
End Property                                       ' cErr.TracePreDec Get

'---------------------------------------------------------------------------------------
' Method : Clone
' Author : rgbig
' Date   : 20211108@11_47
' Purpose: Replicates a cErr (for Recursion, Application start etc.)
'           Used until return from stack.
' Note   : the optional False Parm Exact=True will copy everything without exception.)
'---------------------------------------------------------------------------------------
Function Clone(Optional Exact As Boolean) As cErr

    Set Clone = New cErr

    Clone.atFastMode = atFastMode
    Clone.atCallState = atCallState
    Clone.atMessage = atMessage
    Set Clone.atDsc = atDsc
    Clone.atKey = atKey
    Clone.atProcIndex = atProcIndex
    Clone.atRecursionLvl = atRecursionLvl
    Clone.atCallDepth = atCallDepth
    Clone.atRecursionOK = atRecursionOK
    Clone.atShowStack = atShowStack
    Clone.DebugState = DebugState
    Clone.EventBlock = EventBlock
    Clone.NrMT = NrMT
    Clone.errNumber = errNumber
    Clone.Description = Description
    Clone.Source = Source
    Clone.FoundBadErrorNr = FoundBadErrorNr
    Clone.Permitted = Permitted
    Clone.Explanations = Explanations
    Clone.Reasoning = Reasoning
    Set Clone.atErrPrev = atErrPrev                ' links Clone to current always
    If Exact Then
        Clone.atTraceStackPos = atTraceStackPos
        Set Clone.atCalledBy = atCalledBy
        Set Clone.at_Live = at_Live
        Clone.atCatDict = atCatDict
        Clone.atLastInDate = atLastInDate
        Clone.atLastInSec = atLastInSec
        Clone.atPrevEntrySec = atPrevEntrySec
        Clone.atThisEntrySec = atThisEntrySec
        Clone.atFuncResult = atFuncResult
        Clone.ErrSnoCatch = ErrSnoCatch
        Clone.ErrNoRec = ErrNoRec
    Else                                           ' Entry to Recusion case:
        Clone.ErrSnoCatch = Z§ErrSnoCatch          ' No ErrHandler Recursion;  NO N_PublishBugState
        Clone.ErrNoRec = Z§ErrNoRec                ' No ErrHandler Recursion; use N_PublishBugState
        'Clone.atCatDict = vbNullString               '
        'Set Clone.at_Live = nothing                  ' no LiveStack exists as yet for this instance
        'Clone.atCatDict = vbNullString               '
        'Clone.atPrevEntrySec, Last, This, Date = 0   ' instance can not share time values
        'Set Clone.atCalledBy = Nothing               '    "     and no caller yet
        'Clone.atFuncResult = vbNullString            '    "     and no result yet
    End If
    CloneCounter = CloneCounter + 1

End Function                                       ' cErr.Clone


