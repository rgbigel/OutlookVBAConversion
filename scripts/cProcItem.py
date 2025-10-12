# Converted from cProcItem.py

# VERSION 1.0 CLASS
# BEGIN
# MultiUse = -1  'True
# END
# Attribute VB_Name = "cProcItem"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = False
# Attribute VB_Exposed = False
# Option Explicit

# Public Key As String                               ' atDsc Identification (Long)
# Attribute Key.VB_VarUserMemId = 0
# Attribute Key.VB_VarDescription = "Display Class Instance ID"

# ' **************************************************************************************
# ' to insert default attribute, first export the <self>.cls                          ****
# ' lines below must be placed into <self>.cls by an editor after the declaration     ****
# ' Attribute Key.VB_VarUserMemId = 0
# ' Attribute Key.VB_VarDescription = "Display Key of ProcDsc"
# ' when changes done (without copying the ' Chars), remove + reimport <self>.cls     ****
# ' **************************************************************************************

# Public ProcIndex As Long                           ' Position in D_ErrInterface if >=0,
# '           in D_ErrInterface if < -1

# Public Module As String                            ' Module Name
# Public CallType As String                          ' Sub, Function, ...

# Public DbgId As String                             ' atDsc Ident (short)
# Public CallCounter As Long
# Public pCallMode As eQMode                         ' Proc call Convention. Usage is Private to Class (Friend)
# ' CallMode is the True Public (Get/Let)
# Public ModeName As String                          ' long version of pCallMode
# Public ModeLetter As String                        ' one-letter form of pCallMode

# Public MaxRecursions As Long                       ' deepest recursion level reached
# Public TotalProcTime As Double                     ' not counting Time of called Procs on Stack
# Public TotalRunTime As Double                      ' total Time spent in this running instance

# Public ErrLevel As eLogLevel                       ' Log Calls depending on this, using only 1..4 (eLall..eLcritical)

# Public ErrActive As cErr                           ' this Error Instance

# '---------------------------------------------------------------------------------------
# ' Method : Function ShowDiff
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Function ShowDiff(OtherProcItm As cProcItem, Optional ShowIt As Boolean) As Boolean
# Const zKey As String = "cProcItem.ShowDiff"

# Dim Delta As String

if Me Is OtherProcItm Then:
# ShowDiff = True
# Delta = "cProcItem" & "s are identical"
if ShowIt Then:
print(Debug.Print Delta)
# GoTo ProcRet

if Key <> OtherProcItm.Key Then Delta = Delta & LString(Key, lDbgM) _:
# & Key & " <> " & OtherProcItm.Key & vbCrLf ' String

if ProcIndex <> OtherProcItm.ProcIndex Then _:
# Delta = Delta & LString(ProcIndex, lDbgM) _
# & CStr(ProcIndex) & " <> " & CStr(OtherProcItm.ProcIndex) & vbCrLf ' Long

if Module <> OtherProcItm.Module Then Delta = Delta & LString(Module, lDbgM) _:
# & Module & " <> " & OtherProcItm.Module & vbCrLf ' String

if CallType <> OtherProcItm.CallType Then Delta = Delta & LString(CallType, lDbgM) _:
# & CallType & " <> " & OtherProcItm.CallType & vbCrLf ' String

if DbgId <> OtherProcItm.DbgId Then _:
# Delta = Delta & Format(DbgId, lDbgM) _
# & DbgId & " <> " & OtherProcItm.DbgId & vbCrLf ' String

if CallCounter <> OtherProcItm.CallCounter Then _:
# Delta = Delta & Format(CallCounter, lDbgM) _
# & CStr(CallCounter) & " <> " & CStr(OtherProcItm.CallCounter) & vbCrLf ' Long

if pCallMode <> OtherProcItm.pCallMode Then _:
# Delta = Delta & Format(pCallMode, lDbgM) _
# & CStr(pCallMode) & " <> " & CStr(OtherProcItm.pCallMode) & vbCrLf ' Long

if MaxRecursions <> OtherProcItm.MaxRecursions Then _:
# Delta = Delta & Format(MaxRecursions, lDbgM) _
# & CStr(MaxRecursions) & " <> " & CStr(OtherProcItm.MaxRecursions) & vbCrLf ' Long

if TotalProcTime <> OtherProcItm.TotalProcTime Then _:
# Delta = Delta & Format(TotalProcTime, lDbgM) _
# & CStr(TotalProcTime) & " <> " & CStr(OtherProcItm.TotalProcTime) & vbCrLf ' Double

if TotalRunTime <> OtherProcItm.TotalRunTime Then _:
# Delta = Delta & Format(TotalRunTime, lDbgM) _
# & CStr(TotalRunTime) & " <> " & CStr(OtherProcItm.TotalRunTime) & vbCrLf ' Double

if ErrLevel <> OtherProcItm.ErrLevel Then _:
# Delta = Delta & Format(ErrLevel, lDbgM) _
# & CStr(ErrLevel) & " <> " & CStr(OtherProcItm.ErrLevel) & vbCrLf ' Long


if ShowIt Then:
if LenB(Delta) = 0 Then:
print(Debug.Print "cProcItem a and OtherProcItm have the identical property values")
if Not ErrActive Is OtherProcItm.ErrActive Then _:
# Delta = Delta & "ProcItem " & Format(ErrActive, lDbgM) & VarPtr(ErrActive) _
# & " <> " & VarPtr(OtherProcItm.ErrActive) & vbCrLf ' VarPtr(cErr)
print(Debug.Print Delta)
if Not ShowDiff Then:
# ShowDiff = (LenB(Delta) = 0)

# ProcRet:

# Property Get CallMode() As eQMode
# CallMode = pCallMode
# End Property                                       ' cProcItem.CallMode Get

# Property Let CallMode(Qmode As eQMode)
if Qmode <> pCallMode Then:
# pCallMode = Qmode
# ModeName = QModeNames(Qmode)
# ModeLetter = UCase(Left(ModeName, 1))
# End Property                                       ' cProcItem.CallMode Let

# '---------------------------------------------------------------------------------------
# ' Method : DscCopy
# ' Author : Rolf-Gnther Bercht
# ' Date   : 20211108@11_47
# ' Purpose: Copy Me to a New cProcDsc
# ' Note   : does NOT instanciate new Object.ErrActive, just copies the original reference
# '---------------------------------------------------------------------------------------
def dsccopy():
    # Const zKey As String = "cProcItem.DscCopy"

    # Set DscCopy = New cProcItem
    # With DscCopy
    # .Key = Key                                 ' atDsc Identification (Long)
    # .ProcIndex = ProcIndex                     ' Position in D_ErrInterface
    # .Module = Module                           ' Module Name
    # .CallType = CallType                       ' Sub, Function, ...
    # .DbgId = DbgId                             ' atDsc Ident (short)
    # .CallCounter = CallCounter                 ' Number of calls for the proc
    # .CallMode = CallMode                       ' Quiet, OnStack, NotDef etc.
    # .MaxRecursions = MaxRecursions             ' deepest recursion level reached
    # .TotalProcTime = TotalProcTime             ' not counting Time of called Procs on Stack
    # .TotalRunTime = TotalRunTime               ' total Time spent in this running instance
    # .ErrLevel = ErrLevel                       ' Log Calls depending on this, using only 1..4 (eLall..eLcritical)
    # Set .ErrActive = ErrActive                 ' this Error Instance
    # End With                                       ' DscCopy


