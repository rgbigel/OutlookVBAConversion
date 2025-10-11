VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCallEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ModProc As String                           ' derived from CallStackProcKind
Attribute ModProc.VB_VarUserMemId = 0
Attribute ModProc.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"

' **************************************************************************************
' to insert default attribute, first export the <self>.cls                          ****
' lines below must be placed into <self>.cls by an editor after the declaration     ****
' Attribute ModProc.VB_VarUserMemId = 0
' Attribute ModProc.VB_VarDescription = "Display Key = corresponding ProcDsc.Key"
' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
' **************************************************************************************

Public ModuleName As String
Public ProcedureName As String
Public LineCode As String                          ' Code of Line on stack
Public LineNumber As Long                          ' Line in Procedure where call occurred
Public DscKind As String                           ' derived from CallStackProcKind

Public StackDepth As Long                          ' depth in Stack
Public CallerErr As cErr                           ' if Err/Proc is in D_ErrInterface and LookUpDsc
Public CallerInfo As String                        ' Info derived outside of LiveStack
