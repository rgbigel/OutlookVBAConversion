# Converted from frmDelParms.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelParms
# Caption         =   "Parameter fr das Vergleichen/Lschen von Doubletten"
# ClientHeight    =   6345
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   8295.001
# OleObjectBlob   =   "frmDelParms.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmDelParms"
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
# Call ShowOrHideForm(Me, ShowIt:=False)

# '---------------------------------------------------------------------------------------
# ' Method : Sub DatumsBedingung_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub DatumsBedingung_AfterUpdate()

# Const zKey As String = "frmDelParms.DatumsBedingung_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmDelParms")

# Dim toDay As Date
# CutOffDate = CDate("00:00:00")
# LPmsg.Caption = vbNullString
# toDay = CDate(Left(Now(), 10))
match Datumsbedingung.Text:
    case "keine Datumsbeschrnkung":
    case "heute":
# CutOffDate = DateAdd("d", -0, toDay)
    case "ab gestern":
# CutOffDate = DateAdd("d", -1, toDay)
    case "letzte Woche":
# CutOffDate = DateAdd("d", -7, toDay)
    case "letzte 30 Tage":
# CutOffDate = DateAdd("d", -30, toDay)
    case _:
if IsDate(Datumsbedingung) Then:
# CutOffDate = CDate(Datumsbedingung)
else:
# LPmsg.Caption = "Unzulssige Datumsbedingung: " _
# & Datumsbedingung

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub bDebugStop_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub bDebugStop_Click()
# rsp = vbNo
# Call ShowOrHideForm(Me, ShowIt:=False)
if StopLoop Or bDebugStop.Caption = "DoVerify" Then:
# DoVerify False, " You pressed the debug stop button! F5 or F8 to continue"
# b3text = vbNullString

# '---------------------------------------------------------------------------------------
# ' Method : Sub Go_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Go_Click()
# rsp = vbYes
# Call ShowOrHideForm(Me, ShowIt:=False)

# '---------------------------------------------------------------------------------------
# ' Method : Sub LpLogLevel_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LpLogLevel_AfterUpdate()
# Call LogLevelChecks(Me)

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
# "Lschen nur nach Einzelbesttigung, auch wenn identisch:" _
# & vbCrLf & killType
else:
# LPWantConfirmationThisFolder.Caption = "ohne Rckfrage:" _
# & vbCrLf & killType

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPErgebnisseAlsListe_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPErgebnisseAlsListe_Click()
# ErgebnisseAlsListe = LPErgebnisseAlsListe

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPMaxMisMatchesForCandidates_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPMaxMisMatchesForCandidates_AfterUpdate()
if IsNumeric(LPMaxMisMatchesForCandidates) Then:
# MaxMisMatchesForCandidates = LPMaxMisMatchesForCandidates
else:
# LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPOFast_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPOFast_Click()
# quickChecksOnly = LPOFast

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPOFolderChoice_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPOFolderChoice_Click()
# PickTopFolder = LPOFolderChoice

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPOInformative_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPOInformative_Click()
# quickChecksOnly = Not LPOInformative

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPOWalk_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPOWalk_Click()
# PickTopFolder = Not LPOWalk

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
# ' Method : Sub LPSaveAttachments_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPSaveAttachments_Click()
# SaveAttachments = LPSaveAttachments

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPNoShowEmptyAttrs_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPNoShowEmptyAttrs_Click()
# ShowEmptyAttributes = Not LPNoShowEmptyAttrs

# '---------------------------------------------------------------------------------------
# ' Method : Sub SelektionModifizieren_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub SelektionModifizieren_Click()
# frmSelParms.Show
# Cancel = frmSelParms.Cancel
# bDebugStop = frmSelParms.bDebugStop
if Not sRules Is Nothing Then:
# LPmsg = "Auswahlkriterien gendert, sind nun " _
# & sRules.clsObligMatches.aRuleString
# Go.SetFocus

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPSkipDontCompare_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPSkipDontCompare_Click()
# SkipDontCompare = LPSkipDontCompare

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
# bDebugStop = True
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
# Const zKey As String = "frmDelParms.UserForm_Activate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

# Go.Caption = b1text
# Go.Visible = (b1text <> vbNullString)
# bDebugStop.Caption = b2text
# bDebugStop.Visible = (b2text <> vbNullString)
if LenB(b3text) = 0 Then:
# b3text = "Cancel"
# Cancel.Caption = b3text
# LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
# LPErgebnisseAlsListe = ErgebnisseAlsListe
# AcceptCloseMatches = LPMaxMisMatchesForCandidates > 0
# LpLogLevel.Text = LpLogLevel.List(MinimalLogging)
if LpLogLevel.ListIndex <= 0 Then            ' no choice so far, default one:
# LpLogLevel.Text = LpLogLevel.List(eLmin)
# MinimalLogging = LpLogLevel.ListIndex
# LPWantConfirmation = WantConfirmation
# LPWantConfirmationThisFolder = WantConfirmationThisFolder

# LPOInformative = Not quickChecksOnly
# LPOFast = quickChecksOnly

# LPOWalk = Not PickTopFolder
# LPOFolderChoice = PickTopFolder
# Call Combo_Define_DatumsBedingungen(Datumsbedingung)

if AskEveryFolder Then:
# LPWantConfirmationThisFolder.Caption = _
# "Lschen nur nach Einzelbesttigung, auch wenn identisch:" _
# & vbCrLf & killType
else:
# LPWantConfirmationThisFolder.Caption = "ohne Rckfrage:" _
# & vbCrLf & killType
# LPSaveAttachments = SaveAttachments
# LPAskEveryFolder = AskEveryFolder

# xlShow = xUseExcel
# XlDeferred = xDeferExcel
if ActionID <> 0 And (xlShow Or XlDeferred) Then:
# LPOFast = False                          ' excel requires all attributes
# quickChecksOnly = False
if Not sRules Is Nothing Then:
# Me.LPmsg = Message & vbCrLf & "aktuelle Auswahl lautet: " & sRules.clsObligMatches.aRuleString
# LPmsg.Caption = Message
# bDebugStop.Default = False
# Controls(bDefaultButton).Default = True

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)


# Private Sub UserForm_Initialize()
# Call LPlogLevel_define(Me)

# '---------------------------------------------------------------------------------------
# ' Method : Sub XlDeferred_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub XlDeferred_Click()
# xDeferExcel = XlDeferred

# '---------------------------------------------------------------------------------------
# ' Method : Sub xlShow_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub xlShow_Click()
# xUseExcel = xlShow

