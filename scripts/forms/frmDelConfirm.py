# Converted from frmDelConfirm.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelConfirm
# Caption         =   "Entscheidung ber das Lschen treffen"
# ClientHeight    =   9450.001
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   10425
# OleObjectBlob   =   "frmDelConfirm.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmDelConfirm"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub Cancel_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Cancel_Click()
# rsp = vbCancel
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub DisplayItem
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub DisplayItem(px As Long)
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmDelConfirm.DisplayItem"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# Dim aktItem As Object

# Set aktItem = aID(px).idObjItem
# ' Wenn Outlook-Element existiert, dann...
if Not aktItem Is Nothing Then:
# Me.Hide
# ' ... Outlook-Element anzeigen
# aktItem.Display True
if DebugControlsUsable Then:
# Me.Show
# Call ModItem(aktItem, _
# "FirstName", _
# "LastName", _
# "*,")

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub Go_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Go_Click()
# rsp = vbYes
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub Label3_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Label3_Click()                       ' "Lschliste bisher" ganz unten
# DeleteNow = Not DeleteNow
if DeleteNow Then:
# Label3.BackColor = 0
else:
# Label3.BackColor = -2147483633

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPAskEveryFolder_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPAskEveryFolder_Click()
# AskEveryFolder = LPAskEveryFolder
if AskEveryFolder Then:
# LPWantConfirmationThisFolder.Caption = _
# "Lschen nur nach Einzelbesttigung," _
# & vbCrLf & "auch wenn identisch"
else:
# LPWantConfirmationThisFolder.Caption = "ohne Rckfrage:" _
# & vbCrLf & killType

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPMaxMisMatchesForCandidates_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPMaxMisMatchesForCandidates_AfterUpdate()
if IsNumeric(LPMaxMisMatchesForCandidates) Then:
if LPMaxMisMatchesForCandidates < 0 Then:
# LPMaxMisMatchesForCandidates = 0
# AcceptCloseMatches = LPMaxMisMatchesForCandidates > 0
# MaxMisMatchesForCandidates = LPMaxMisMatchesForCandidates
else:
# LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPWantConfirmation_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPWantConfirmation_Click()
# WantConfirmation = LPWantConfirmation

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPWantConfirmationThisFolder_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPWantConfirmationThisFolder_Click()
# WantConfirmationThisFolder = LPWantConfirmationThisFolder

# '---------------------------------------------------------------------------------------
# ' Method : Sub MoreParams_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub MoreParams_Click()
# askforParams = Not askforParams
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub No_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub No_Click()
# rsp = vbNo
# Me.Hide

# '---------------------------------------------------------------------------------------
# ' Method : Sub S1_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub S1_Click()
# S1.Enabled = False
# S2.Enabled = True
# S2.Value = False
# Call DisplayItem(1)
# S1.Enabled = True
# S1.Value = False

# '---------------------------------------------------------------------------------------
# ' Method : Sub S2_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub S2_Click()
# S2.Enabled = False
# S1.Enabled = True
# S1.Value = False
# Call DisplayItem(2)
# S2.Enabled = True
# S2.Value = False

# '---------------------------------------------------------------------------------------
# ' Method : Sub SucheBeenden_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub SucheBeenden_Click()
# StopLoop = Not StopLoop
if StopLoop Then:
# SucheBeenden.Caption = "fortfahren"
else:
# SucheBeenden.Caption = "Suche beenden"

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmDelConfirm.UserForm_Activate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

# Go.Caption = b1text
# Go.Visible = (b1text <> vbNullString)
# No.Caption = b2text
# No.Visible = (b2text <> vbNullString)
# Cancel.Caption = b3text
# LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
# LPMaxMisMatchesForCandidates.Enabled = AcceptCloseMatches
# LPWantConfirmation = WantConfirmation
# LPWantConfirmationThisFolder = WantConfirmationThisFolder
if AskEveryFolder Then:
# LPWantConfirmationThisFolder.Caption = _
# "Lschen nur nach Einzelbesttigung, auch wenn identisch:" _
# & vbCrLf & killType
else:
# LPWantConfirmationThisFolder.Caption = "ohne Rckfrage:" _
# & vbCrLf & killType
# LPAskEveryFolder = AskEveryFolder
# LPDiffs = Diffs
# LPmsg.Text = Message
if LenB(LoeschbesttigungCaption) > 0 Then:
# Me.Caption = LoeschbesttigungCaption
# Me.LBlliste = Mid(LListe, 2)
# Me.Go.Default = True
# Me.No.Default = False
# Me.Cancel.Default = False

# Me.Controls(bDefaultButton).Default = True   ' funktioniert nicht wenn label falsch
if askforParams Then:
# MoreParams.Caption = "Params OK"
# askforParams = False
else:
# MoreParams.Caption = "Weitere Parameter"
# Me.SucheBeenden = False
# S1.Enabled = True
# S1.Value = False                             ' wre schner, wenn wir nicht modal schaffen knnten
# S2.Enabled = True
# S2.Value = False                             ' dann knnten wir sehen, was schon aktiv ist....
if WindowSetForeground(LBF.Caption, FWP_LBF_Hdl) Then:
print(Debug.Print Quote1(Trim(LBF.Caption)) _)
# & "  wird als oberstes Fenster angezeigt."
# GoTo FuncExit
else:
print(Debug.Print Quote1(Trim(LBF.Caption)) & "  wird nicht angezeigt")

# FuncExit:
# Call ErrReset(0)

# ProcReturn:
# Call ProcExit(zErr)


# Private Sub UserForm_Initialize()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmDelConfirm.UserForm_Initialize"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmDelConfirm")

# S1.Enabled = True
# S2.Enabled = False
# S1.Value = False
# S2.Value = False
if LpLogLevel.ListCount = 0 Then:
# Call LPlogLevel_define(Me)
# LpLogLevel.Text = LpLogLevel.List(MinimalLogging + 1)
if LpLogLevel.ListIndex <= 0 Then            ' no choice so far, default one:
# LpLogLevel.Text = LpLogLevel.List(eLmin + 1)
# MinimalLogging = LpLogLevel.ListIndex
# Me.LpLogLevel = LogSelection

if xlApp Is Nothing Then:
if displayInExcel Then:
# Call XlgetApp
# xShowAllExcel.Visible = True
else:
# xShowAllExcel.Visible = False
else:
if O Is Nothing Or Not displayInExcel Then:
# xShowAllExcel.Visible = True
else:
# xShowAllExcel.Visible = False
# rsp = vbIgnore

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub xShowAllExcel_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub xShowAllExcel_Click()
# Me.Hide
# rsp = vbRetry

# '---------------------------------------------------------------------------------------
# ' Method : Sub LpLogLevel_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LpLogLevel_AfterUpdate()

# Const zKey As String = "frmDelConfirm.LpLogLevel_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmDelConfirm")

# Call LogLevelChecks(Me)

# ProcReturn:
# Call ProcExit(zErr)


