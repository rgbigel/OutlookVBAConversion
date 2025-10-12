# Converted from cBugTimer.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cBugTimer"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public BugStateTicks As Long
# Public BugTimerId As Long
# Public BugStateLast As Double
# Public BugStateElapsed As Double
# Public BugStateTrigCount As Long

# Private BugState_Save As Boolean
# Private BugState_ReCheck As Boolean
# Private sEventBlock As Boolean

# Property Let BugStateReCheck(setit As Boolean)
# BugState_ReCheck = setit
# BugState_Save = setit
# End Property                                       ' cBugTimer.BugStateReCheck Let

# Property Get BugStateReCheck() As Boolean
# BugStateReCheck = BugState_ReCheck
# End Property                                       ' Get cBugTimer.BugStateReCheck Get

# '---------------------------------------------------------------------------------------
# ' Method : Function BugState_SetPause
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bugstate_setpause():
    if BugTimerId <> 0 Then:
    # BugState_SetPause = BugState_ReCheck       ' return the old state
    # BugState_Save = BugState_ReCheck
    # BugState_ReCheck = False                       ' allways in pause
    # sEventBlock = BlockEvents
    # BlockEvents = True

# '---------------------------------------------------------------------------------------
# ' Method : Function BugState_UnPause
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
def bugstate_unpause():
    if BugTimerId <> 0 Then:
    # BugState_UnPause = BugState_Save           ' return the old state
    # BugState_ReCheck = True                    ' allways pause ended
    # BugState_Save = True
    # BlockEvents = sEventBlock
