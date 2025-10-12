# Converted from frmDeltaList.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDeltaList
# Caption         =   "Ordner Vergleich"
# ClientHeight    =   8835.001
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   13560
# OleObjectBlob   =   "frmDeltaList.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmDeltaList"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub AnzeigenDetails_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub AnzeigenDetails_Click()
# Call ShowDetails

# '---------------------------------------------------------------------------------------
# ' Method : Sub DruckenAlle_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub DruckenAlle_Click()
# Call DetailsToPrintFile(iPfad)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ItemList_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ItemList_Click()
# rID = ItemList.ListIndex
# frmCompareInfo.Show

# Private Sub UserForm_Initialize()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmDeltaList.UserForm_Initialize"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

# rID = 0
# Me.ItemList.Clear
# Me.ItemList.addItem "no."
# ItemList.List(rID, 1) = "indx1"
# ItemList.List(rID, 2) = MainObjectIdentification
# ItemList.List(rID, 3) = "comp."
# ItemList.List(rID, 4) = "indx2"
# ' ItemList.List(rId, 5) = "Matchinfo"
# ItemList.List(rID, 5) = "Diffinfo"
# Call ErrReset(4)

# ''''On Error Res'ume Next    ' if assignment fails, use previous values

# Me.ItemList.addItem rID
# ItemList.List(rID, 1) = ListContent(rID).Index1
# ItemList.List(rID, 2) = "***too long ***"
# ItemList.List(rID, 2) = ListContent(rID).MainId
# ItemList.List(rID, 3) = ListContent(rID).Compares
# ItemList.List(rID, 4) = ListContent(rID).Index2
# ItemList.List(rID, 5) = "***too many differrences***"
# ItemList.List(rID, 5) = ListContent(rID).DiffsRecognized

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)

