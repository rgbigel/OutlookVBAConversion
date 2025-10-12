# Converted from cTriState.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cTriState"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Dim StateIs As Long

# Public Property Get Value() As Long
# Value = StateIs
# End Property                                       ' cTriState.Value Get

# Public Property Let Value(nv As Long)
match nv:
    case TristateTrue:
# StateIs = TristateTrue
    case TriStateFalse:
# StateIs = TriStateFalse
    case TriStateMixed:
# StateIs = TriStateMixed
    case TriStatsUndefined:
# StateIs = TriStatsUndefined
    case _:
if nv <> TriStatsUndefined Then:
if rsp = vbCancel Then:
# Err.Raise Hell
# StateIs = TriStatsUndefined
# End Property                                       ' cTriState.Value Let
