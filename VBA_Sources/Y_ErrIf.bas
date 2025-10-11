Attribute VB_Name = "Y_ErrIf"
Option Explicit

' ---------- Generated Code, See ZZZErrInterfaceGenerator ---------------
Public Const UBndErrinterfaces As Long = 96
Public Const UBndErrinterfaceNames As Long = 108
'            13.12.2017 15:19:20
' The following is a list of all quasi-static Objects in this Module
' it is not necessary that it is complete
Public Z你ppStart As cProcItem                   '  | K:BugHelp.Z_StartUp R:Z_StartUp
Public Z判rocStart As cProcItem                  '  | K:BugHelp.ProcCall R:ProcCall
Public Z刨howErrorStatus As cProcItem            '  | K:BugHelp.ShowErrorStatus R:ShowErrorStatus
Public Z助sedThisCall As cProcItem               '  | K:BugHelp.Z_UsedThisCall R:Z_UsedThisCall

' Pre-Define Calls
Sub N_PreDefine()
Dim TEntry As cTraceEntry

    Set D_ErrInterface = New Dictionary
    
    dontIncrementCallDepth = True               ' no real calls in PreDefine
    
    ' define the External Caller (Dummy) as first Entry on all Stacks
    Call DoCall("Extern.Caller", tSub, eQnoDef, P_Active) ' also defines DoCall
    Call DoCall("Y_ErrIf.CurrEntry", tSub, eQnoDef, P_CurrEntry)
    Call DoCall("Y_ErrIf.LastEP", tSub, eQnoDef, P_LastEP)
    Call DoCall("Y_ErrIf.ProcCall", tSub, eQnoDef, Z判rocStart)
    Call DoCall("Y_ErrIf.ShowErrorStatus", tSub, eQnoDef, Z刨howErrorStatus)
    Call DoCall("Y_ErrIf.Z_UsedThisCall", tSub, eQnoDef, Z助sedThisCall)
    Call DoCall("ThisOutlookSession.ApplicationStartup", tSub, eQnoDef, Z刨tartApp)
    
    dontIncrementCallDepth = False
    
    With ExternCaller
        .ErrActive.atCallState = eCpaused        ' during this, the first inits are done
        .ErrActive.atCallDepth = 0
        Set TEntry = New cTraceEntry
        Set TEntry.TErr = .ErrActive             ' generate trace Entry for ExternCaller
        Call N_TraceEntry(TEntry, "not removable Trace " & .ErrActive.atKey)
    End With                                     ' ExternCaller
    
    Set E_AppErr = ExternCaller.ErrActive
    
    Set N低all.ErrActive.atCalledBy = ExternCaller.ErrActive

FuncExit:
    
    Set TEntry = Nothing

End Sub                                          ' Z_ErrIf.N_PreDefine

