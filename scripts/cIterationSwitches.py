# Converted from cIterationSwitches.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cIterationSwitches"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public ReprocessLOGGEDItems As Boolean
# Public CategoryConfirmation As Boolean
# Public ReProcessDontAsk As Boolean
# Public ResetCategories As Boolean
# Public SaveItemRequested As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Function Assign
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function Assign(Source As cIterationSwitches) As cIterationSwitches
# Dim zErr As cErr
# Const zKey As String = "cIterationSwitches.Assign"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction, ExplainS:="cIterationSwitches")

# Assign.ReprocessLOGGEDItems = Source.ReprocessLOGGEDItems
# Assign.CategoryConfirmation = Source.CategoryConfirmation
# Assign.ReProcessDontAsk = Source.ReProcessDontAsk
# Assign.SaveItemRequested = Source.SaveItemRequested

# ProcReturn:
# Call ProcExit(zErr)
