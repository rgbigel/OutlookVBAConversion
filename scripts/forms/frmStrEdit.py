# Converted from frmStrEdit.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStrEdit
# Caption         =   "StringModifier"
# ClientHeight    =   9405.001
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   10635
# OleObjectBlob   =   "frmStrEdit.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmStrEdit"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit


# '---------------------------------------------------------------------------------------
# ' Method : Sub EditRules_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub EditRules_Click()
# StringModifierRsp = vbYes                    ' go to rule editing
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub StringModifierCan_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub StringModifierCan_Click()
# StringToConfirm = StringModifierCancelValue.Text
# StringModifierRsp = vbCancel                 ' finish processing without further changes
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub StringModifierOK_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub StringModifierOK_Click()
# StringModifierRsp = vbOK
# Me.Hide

# Private Sub StringToConfirm_Initialize()
# Dim zErr As cErr
# Const zKey As String = "frmStrEdit.StringToConfirm_Initialize"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

# StringToConfirm = StringModifierCancelValue.Text

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# Dim zErr As cErr
# Const zKey As String = "frmStrEdit.UserForm_Activate"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

# With CurIterationSwitches
# ReprocessLOGGEDItems = .ReprocessLOGGEDItems
# CategoryConfirmation = .CategoryConfirmation
# IgnoriereBestehendeKategorien = .ResetCategories
# chSaveItemRequested = .SaveItemRequested
# End With                                     ' CurIterationSwitches
# Call ChkCatLogic
# StringModifierRsp = vbRetry

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Deactivate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Deactivate()
if StringModifierRsp = vbRetry Then:
# StringModifierCan_Click                  ' same as cancel, but:
# StringModifierRsp = 0                    ' ignore this item, but continue loop
else:
# StringToConfirm_Initialize

# Private Sub UserForm_Terminate()
# UserForm_Deactivate

# '---------------------------------------------------------------------------------------
# ' Method : Sub ReprocessLOGGEDItems_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ReprocessLOGGEDItems_AfterUpdate()
# CurIterationSwitches.ReprocessLOGGEDItems = ReprocessLOGGEDItems
# Call ChkCatLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub IgnoriereBestehendeKategorien_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub IgnoriereBestehendeKategorien_AfterUpdate()

# Const zKey As String = "frmStrEdit.IgnoriereBestehendeKategorien_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmStrEdit")

# CurIterationSwitches.ResetCategories = IgnoriereBestehendeKategorien

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ReProcessDontAsk_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ReProcessDontAsk_AfterUpdate()

# Const zKey As String = "frmStrEdit.ReProcessDontAsk_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmStrEdit")

# CurIterationSwitches.ReProcessDontAsk = ReProcessDontAsk
# Call ChkCatLogic

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub CategoryConfirmation_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub CategoryConfirmation_AfterUpdate()

# Const zKey As String = "frmStrEdit.CategoryConfirmation_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmStrEdit")

# CurIterationSwitches.CategoryConfirmation = CategoryConfirmation

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub chSaveItemRequested_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub chSaveItemRequested_AfterUpdate()

# Const zKey As String = "frmStrEdit.chSaveItemRequested_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmStrEdit")

# StringModifierRsp = vbNo
# CurIterationSwitches.SaveItemRequested = chSaveItemRequested

# ProcReturn:
# Call ProcExit(zErr)

