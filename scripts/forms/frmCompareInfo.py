# Converted from frmCompareInfo.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompareInfo
# Caption         =   "Vergleichsinformationen"
# ClientHeight    =   11760
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   13695
# OleObjectBlob   =   "frmCompareInfo.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmCompareInfo"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub DetailsToFile_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub DetailsToFile_Click()
# Call DetailsToPrintFile(iPfad)
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub OK_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OK_Click()
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# Dim zErr As cErr
# Const zKey As String = "frmCompareInfo.UserForm_Activate"

# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

if rID > 0 Then:
# LPInfo.Text = ListContent(rID).MatchData
# LPDiff = ListContent(rID).DiffsRecognized
# UserDecisionRequest = False

# ProcReturn:
# Call ProcExit(zErr)

# pExit:

# '---------------------------------------------------------------------------------------
# ' Method : Sub VollVergleich_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub VollVergleich_Click()
# UserDecisionRequest = True
# Me.Hide
