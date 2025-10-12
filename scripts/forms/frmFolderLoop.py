# Converted from frmFolderLoop.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFolderLoop
# Caption         =   "Ordner-Iteration"
# ClientHeight    =   7305
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   10515
# OleObjectBlob   =   "frmFolderLoop.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmFolderLoop"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# Private Initing As Boolean

# '---------------------------------------------------------------------------------------
# ' Method : Sub Cancel_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Cancel_Click()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmFolderLoop.Cancel_Click"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# End

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub CategoryConfirmation_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub CategoryConfirmation_AfterUpdate()

# Const zKey As String = "frmFolderLoop.CategoryConfirmation_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# Call CheckLogic

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub CheckLogic
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub CheckLogic()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmFolderLoop.checklogic"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# Static lastSelectorMode As Long

if Initing Then:
# With CurIterationSwitches
# ReprocessLOGGEDItems = .ReprocessLOGGEDItems
# CategoryConfirmation = .CategoryConfirmation
# ReProcessDontAsk = .ReProcessDontAsk
# OverrideCategories = .ResetCategories
# chSaveItemRequested = .SaveItemRequested
# End With                                 ' CurIterationSwitches
# GoTo ProcReturn
if Not UI_Show_Sel Then                      ' frmSelParms Sichtbarkeit verloren::
# UI_SelParameter.Value = True             ' dann Standard
else:
if UI_SelParameter.Value Then            ' on Standard: dont change visibility:
# UI_Show_Sel = UI_Show_Sel
elif OptionButton8.Value Then:
# UI_Show_Sel = True                   ' allow user to switch back
# UI_SelParameter.Visible = UI_Show_Sel

if Not UI_Show_Del Then                      ' frmDelParms Sichtbarkeit verloren::
# UI_DelOption.Value = True                ' dann Standard
else:
if UI_DelOption.Value Then               ' on Standard: dont change visibility:
# UI_Show_Del = UI_Show_Del
elif OptionButton8.Value Then:
# UI_Show_Del = True                   ' allow user to switch back
# UI_DelOption.Visible = UI_Show_Del

if lastSelectorMode <> SelectorMode Then:
# ' changed selector mode: compute all Bits again
# setMode:
# AllPublic.eActFolderChoice = False       ' SelectorMode 1
# AllPublic.eAllFoldersOfType = False      ' SelectorMode 2
# AllPublic.eOnlySelectedItems = False     ' Selectormode 3
match SelectorMode:
    case 1:
# AllPublic.eActFolderChoice = True
# SelectorButton1.Value = True
    case 2:
# AllPublic.eAllFoldersOfType = True
# SelectorButton2.Value = True
    case 3:
# AllPublic.eOnlySelectedItems = True
# SelectorButton3.Value = True
    case 99:
# ' all ExplainS remain False, they do not make sense at this time
    case Else                                ' default:
# SelectorMode = 1
# GoTo setMode
# lastSelectorMode = SelectorMode

# Call ChkCatLogic

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub chSaveItemRequested_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub chSaveItemRequested_AfterUpdate()

# Const zKey As String = "frmFolderLoop.chSaveItemRequested_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# CurIterationSwitches.SaveItemRequested = chSaveItemRequested

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub DatumsBedingung_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub DatumsBedingung_AfterUpdate()

# Const zKey As String = "frmFolderLoop.DatumsbedingunG_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# Dim toDay As Date
# CutOffDate = CDate("00:00:00")
# ShowError.Text = vbNullString
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
# ShowError.Text = "Unzulssige Datumsbedingung: " _
# & Datumsbedingung

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub InitActions
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Public Sub InitActions()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "frmFolderLoop.InitActions"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

# Dim i As Long
# Dim actionCounter As Long
# Dim actTitle As String
# Dim haveMod As Boolean
# Dim aControl As Control
# Dim ActionItemsCount As Long
# Dim aControlN As String

# Call Try(testAll)                               ' Try anything, autocatch
# Initing = True                               ' do not process buttons yet
if LenB(ActionTitle(UBound(ActionTitle))) = 0 Then:
# Call SetStaticActionTitles
# ActionItemsCount = UBound(ActionTitle)
if LpLogLevel.ListCount = 0 Then:
# Call LPlogLevel_define(Me)
# LpLogLevel.Text = LpLogLevel.List(MinimalLogging + 1)
if LpLogLevel.ListIndex <= 0 Then            ' no choice so far, default one:
# LpLogLevel.Text = LpLogLevel.List(eLmin + 1)
# MinimalLogging = LpLogLevel.ListIndex
# UI_DelOption.Value = UI_DontUseDel           ' request override of frmDelParms
# UI_DelOption.Visible = UI_Show_Del           ' allow user to choose override
# UI_SelParameter.Value = UI_DontUse_Sel       ' request override of SelecionParameter
# UI_SelParameter.Visible = UI_Show_Sel        ' allow user to choose override
# CurIterationSwitches.ResetCategories = False
if MaintenanceAction = 2 _:
# Or ActionTitle(ActionItemsCount) = vbNullString _
# Or Left(ActionTitle(0), 5) = IgnoredHeader Then
# actionCounter = 0
for acontrol in me:
# aControlN = aControl.Name
# aBugTxt = "set up Control " & aControlN
# Call Try(testAll)
# aBugTxt = "set up Control " & aControlN
# ' needed, because some controls have no name
if Len(aControlN) > 12 Then:
if Left(aControlN, 12) = "OptionButton" Then:
# actionCounter = actionCounter + 1
# actTitle = Me.Controls.Item("OptionButton" & actionCounter).Caption
if ActionItemsCount >= actionCounter Then:
if actTitle <> ActionTitle(actionCounter) Then:
# haveMod = True
# ActionTitle(actionCounter) = actTitle
else:
# haveMod = True
if Catch Then:
# GoTo FuncExit
# Call ErrReset(4)

if actionCounter = 0 Then:
print(Debug.Print "No option buttons in " & Me.Name)
# GoTo InitFinished

if Not haveMod And actionCounter = ActionItemsCount Then:
if Not ShutUpMode Then:
print(Debug.Print "Action titles are OK in " & Me.Name)
# GoTo InitFinished

print(Debug.Print "Paste this into AllPublic if changed or not present")
print(Debug.Print IgnoredHeader & " Start of generated code")
print(Debug.Print "Public ActionTitle(0 to " & actionCounter; ") As String")
# actTitle = ActionTitle(i)
print(Debug.Print "Public Const at" _)
# & RemoveChars(actTitle, """-/*:") _
# & " As Long = " & i

print(Debug.Print vbCrLf & "Sub SetStaticActionTitles")

# actTitle = ActionTitle(i)
print(Debug.Print vbTab & "ActionTitle" & Quote(i, Bracket) = vbNullString; "" _)
# & Replace(actTitle, Q, " & quote(") & Q
print(Debug.Print "En" & "d Su" & "b      ' SetStaticActionTitles" & vbCrLf)
print(Debug.Print IgnoredHeader & " End of generated code")
if haveMod Then:
# actTitle = IgnoredHeader & " Updating of InitActions/AllPublic required " _
# & "because the Array 'ActionTitle' is too short"
# ActionTitle(0) = actTitle
# DoVerify False, " copy and paste stuff from debug log"
else:
# ActionTitle(0) = "unspecified"
elif LenB(ActionTitle(ActionItemsCount)) = 0 Then:
# ActionTitle(i) = Me.Controls.Item("OptionButton" & i).Caption
# setMode:
match SelectorMode:
    case 1:
# SelectorButton1.Value = True
    case 2:
# SelectorButton2.Value = True
    case 3:
# SelectorButton3.Value = True
    case 99:
# ' all ExplainS remain False, they do not make sense at this time
# SelectorButton1.Value = False
# SelectorButton2.Value = False
# SelectorButton3.Value = False
    case Else                                    ' default:
# SelectorMode = 1
# GoTo setMode
match ActionID:
    case 0                                       ' use default:
# ActionID = atPostEingangsbearbeitungdurchfhren
# ' after FindingallDeferred...
if LF_UsrRqAtionId = atFindealleDeferredSuchordner Then:
# OptionButton7 = True
else:
# OptionButton3 = True
# LF_UsrRqAtionId = atPostEingangsbearbeitungdurchfhren
    case 1:
# OptionButton1 = True
    case 2:
# OptionButton2 = True
    case 3:
# OptionButton3 = True
    case 4:
# OptionButton4 = True
    case 5:
# OptionButton5 = True
    case 6:
# OptionButton6 = True
    case 7:
# OptionButton7 = True
    case 8:
# OptionButton8 = True
    case _:
# DoVerify False, "Undefined ActionID"
# OptionButton1 = False
# OptionButton2 = False
# OptionButton3 = False
# OptionButton4 = False
# OptionButton5 = False
# OptionButton6 = False
# OptionButton7 = False
# OptionButton8 = False
# Datumsbedingung = vbNullString
# ShowError.Text = vbNullString
# chSaveItemRequested = CurIterationSwitches.SaveItemRequested
# xlShow = xUseExcel
# eOnlySelectedItems = Not AllPublic.eActFolderChoice
# OverrideCategories = False

# Call Combo_Define_DatumsBedingungen(Datumsbedingung)
if XlDeferred <> xDeferExcel Then:
if Not (displayInExcel Or xUseExcel Or xDeferExcel) Then:
# XlDeferred = True
else:
# XlDeferred = False
# xDeferExcel = XlDeferred
# InitFinished:
# Initing = False
# Call CheckLogic
# GoTo FuncExit

# FuncExit:
# Call ErrReset(0)

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub OK_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OK_Click()
if LenB(ShowError.Text) = 0 Then:
# CurIterationSwitches.ReProcessDontAsk = ReProcessDontAsk
# Call CheckLogic
# Hide
# UI_Show_Del = False                      ' dont allow user to choose override
# UI_Show_Sel = False                      ' dont allow user to choose override
else:
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton1_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton1_Click()
# LF_UsrRqAtionId = atDefaultAktion
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = False
# PickTopFolder = False
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow  user choice of Selection Parameters
# Call CheckLogic



# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton2_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton2_Click()
# LF_UsrRqAtionId = atKategoriederMailbestimmen
# SelectorMode = 3                             ' eOnlySelectedItems = True
# CategoryProcessing.Visible = True
# CurIterationSwitches.CategoryConfirmation = True
# CurIterationSwitches.ReprocessLOGGEDItems = True
# ReprocessLOGGEDItems = True
# CategoryConfirmation = True
# OverrideCategories = True
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow  user choice of Selection Parameters
# Call CheckLogic



# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton3_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton3_Click()
# LF_UsrRqAtionId = atPostEingangsbearbeitungdurchfhren
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = True
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow  user choice of Selection Parameters
# Call CheckLogic


# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton4_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton4_Click()
# LF_UsrRqAtionId = atDoppelteItemslschen
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = False
# AcceptCloseMatches = True
# quickChecksOnly = Not AcceptCloseMatches
# AskEveryFolder = True
# WantConfirmation = True
# MatchMin = 1000
# IsComparemode = False
# StopLoop = False
# PickTopFolder = True
# UI_Show_Del = True                           ' allow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow automatic user choice of Selection Parameters
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton5_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton5_Click()
# LF_UsrRqAtionId = atNormalreprsentationerzwingen
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = False
# xDeferExcel = True
# XlDeferred = xDeferExcel
# chSaveItemRequested = True
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow  user choice of Selection Parameters


# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton6_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton6_Click()
# LF_UsrRqAtionId = atOrdnerinhalteZusammenfhren
# PickTopFolder = False
# SelectorMode = 1                             ' Es macht nur Sinn ganze Ordner zu whlen
# CategoryProcessing.Visible = False
# xDeferExcel = True
# XlDeferred = xDeferExcel
# xUseExcel = False
# displayInExcel = False                       ' wait until we need it
# xlShow = xUseExcel
# quickChecksOnly = True
# UI_Show_Del = True                           ' allow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow automatic user choice of Selection Parameters
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton7_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton7_Click()
# LF_UsrRqAtionId = atFindealleDeferredSuchordner
# CategoryProcessing.Visible = False
# AllPublic.eOnlySelectedItems = False
# AllPublic.eActFolderChoice = False
# AllPublic.eAllFoldersOfType = False
# PickTopFolder = False
# SelectorMode = 99
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow automatic user choice of Selection Parameters
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton8_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton8_Click()
# LF_UsrRqAtionId = atBearbeiteAllebereinstimmungenzueinerSuche
# PickTopFolder = True
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = False
# SelectOnlyOne = False                        ' loop operation!
# SelectMulti = True
# xDeferExcel = False
# XlDeferred = xDeferExcel
# xUseExcel = False
# displayInExcel = False                       ' wait until we need it
# xlShow = xUseExcel
# quickChecksOnly = True
# UI_Show_Del = False                          ' disallow user choice of Deletion Parameters
# UI_Show_Sel = True                           ' allow automatic user choice of Selection Parameters

# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub OptionButton9_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OptionButton9_Click()
# LF_UsrRqAtionId = atContactFixer
if SelectorMode = 99 Then:
# SelectorMode = 1
# CategoryProcessing.Visible = False
# PickTopFolder = False
# UI_Show_Del = False                          ' disallow user choice of Match Parameters
# UI_Show_Sel = False                          ' disallow  user choice of Selection Parameters
# Call CheckLogic


# '---------------------------------------------------------------------------------------
# ' Method : Sub UI_SelParameter_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UI_SelParameter_Change()

# Const zKey As String = "frmFolderLoop.UI_SelParameter_Change"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# UI_DontUse_Sel = UI_SelParameter.Value

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub OverrideCategories_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub OverrideCategories_AfterUpdate()

# Const zKey As String = "frmFolderLoop.OverrideCategories_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# CurIterationSwitches.ResetCategories = OverrideCategories
# Call CheckLogic

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ReProcessDontAsk_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ReProcessDontAsk_AfterUpdate()

# Const zKey As String = "frmFolderLoop.ReProcessDontAsk_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# CurIterationSwitches.ReProcessDontAsk = ReProcessDontAsk
if CurIterationSwitches.ReProcessDontAsk Then:
# ReprocessLOGGEDItems = True
# Call CheckLogic

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub ReprocessLOGGEDItems_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ReprocessLOGGEDItems_AfterUpdate()

# Const zKey As String = "frmFolderLoop.ReprocessLOGGEDItems_AfterUpdate"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmFolderLoop")

# CurIterationSwitches.ReprocessLOGGEDItems = ReprocessLOGGEDItems
# Call CheckLogic

# ProcReturn:
# Call ProcExit(zErr)


# '---------------------------------------------------------------------------------------
# ' Method : Sub SelectorButton1_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub SelectorButton1_Click()
# SelectorMode = 1
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub SelectorButton2_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub SelectorButton2_Click()
# SelectorMode = 2
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub SelectorButton3_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub SelectorButton3_Click()
# SelectorMode = 3
# Call CheckLogic

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# Dim zErr As cErr
# Const zKey As String = "frmFolderLoop.UserForm_Activate"
# Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

# LpLogLevel.ListIndex = MinimalLogging

# ProcReturn:
# Call ProcExit(zErr)

# Private Sub UserForm_Initialize()
# ' all Items in Form are/must be ErrHdlInited !
# Call InitActions
if MaintenanceAction = 2 Then:
# MaintenanceAction = 0

# '---------------------------------------------------------------------------------------
# ' Method : Sub XlDeferred_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub XlDeferred_Click()
# xDeferExcel = Not xDeferExcel
if xDeferExcel Then:
# xUseExcel = False                        ' probable, but not absolutely necessary
# displayInExcel = False
# xlShow = xUseExcel                       ' next init value = false
elif Not xUseExcel Then:
# displayInExcel = False

# '---------------------------------------------------------------------------------------
# ' Method : Sub xlShow_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub xlShow_Click()
# xUseExcel = xlShow

# '---------------------------------------------------------------------------------------
# ' Method : Sub LpLogLevel_AfterUpdate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LpLogLevel_AfterUpdate()
# Call LogLevelChecks(Me)

