# Converted from frmMessage.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMessage
# Caption         =   "Nachricht"
# ClientHeight    =   3120
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   4710
# OleObjectBlob   =   "frmMessage.frx":0000
# ShowModal       =   0   'False
# End
# Attribute VB_Name = "frmMessage"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# Private MessageResponse As Long
# Const B1mr As Long = 1
# Const B2mr As Long = 2
# Public ResponseWaits As Long

# '---------------------------------------------------------------------------------------
# ' Method : Sub B1_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub B1_Click()
# MessageResponse = B1mr
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub B2_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub B2_Click()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmMessage.B2_Click"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmMessage")

# MessageResponse = B2mr
# Me.Hide

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# ResponseWaits = ResponseWaits + 1

# '---------------------------------------------------------------------------------------
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# MessageResponse = msgResponse

# ProcReturn:
# Call ProcExit(zErr)


# Private Sub UserForm_Initialize()
# ResponseWaits = 0
