# Converted from cTermination.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cTermination"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public DCAllowedMatch As Variant
# Public DCUseErrExOn As String
# Public DCAppl As String
# Public DCerrNum As Long
# Public DCerrMsg As String
# Public DCerrSource As String

# Public TermRQ As Boolean
# Public LogFileLen As Long
# Public LogIsOpen As Boolean
# Public LogName As String
# Public LogNameNext As String
# Public LogNamePrev As String

# '---------------------------------------------------------------------------------------
# ' Method : Function Terminate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function Terminate(Optional ByVal msg As String) As Boolean
# Const zKey As String = "cTermination.Terminate"

# Dim ErrNoAtCall As Long

# ErrNoAtCall = E_AppErr.errNumber
if LenB(DCerrMsg) = 0 Then:
# msg = "Termination requested " & msg
else:
# msg = "Termination: " & DCerrMsg

print(Debug.Print String(80, "*") _)
# & vbCrLf & msg & ", Error Number at time of Termination: " & ErrNoAtCall _
# & vbCrLf & String(80, "*")

# TermRQ = True

# QuitStarted = True
# Call Application.ActiveExplorer.Close          ' may cause an event!

# End                                            ' there is no return from this procedure, but: ThisOutlookSession.Quit is run


# '---------------------------------------------------------------------------------------
# ' Method : Sub N_ClearTermination
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub N_ClearTermination()
# Const zKey As String = "cTermination.N_ClearTermination"

# '------------------- gated Entry -------------------------------------------------------
# Static Recursive As Boolean

if Recursive Then:
# ' choose Ignored or Forbidden and dependence on StackDebug
if StackDebug > 8 Then:
print(Debug.Print String(OffCal, b) & "Ignored recursion from " _)
# & P_Active.DbgId & " => " & zKey
# GoTo ProcRet
# Recursive = True
# '---------- End ---- gated Entry -------------------------------------------------------

# TermRQ = False
# ' DCUseErrExOn = vbNullString             T_DC.DCUseErrExOn and DCAppAct not changing
# ' DCAppAct = vbNullString
# DCerrNum = 0
# DCerrMsg = vbNullString
# DCerrSource = vbNullString
# DCAllowedMatch = Empty
# Recursive = False

# ProcRet:

