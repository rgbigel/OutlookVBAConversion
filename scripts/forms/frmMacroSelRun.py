# Converted from frmMacroSelRun.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMacroSelRun
# Caption         =   "Liste verfgbarer Macros"
# ClientHeight    =   7500
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   9900.001
# OleObjectBlob   =   "frmMacroSelRun.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmMacroSelRun"
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
# '''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
# Const zKey As String = "frmMacroSelRun.Cancel_Click"

# rsp = vbNo
# Me.Hide


# '---------------------------------------------------------------------------------------
# ' Method : Sub Run_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub Run_Click()
# rsp = vbYes

if Not fKeepOpen Then:
# Me.Hide
match LPMacroListe.Value:
    case "C1SI":                Call C1SI:
    case "C2SI":                Call C2SI:
    case "ChangeBirthdayOf":    Call ChangeBirthdaySubject("#"):
    case "CopyAllBackupCats":   Call CopyAllBackupCats:
    case "CopyAllHotmailCats":  Call CopyAllHotmailCats:
    case "CreateRules":         Call CreateRules:
    case "ExcelShowItem":       Call ExcelShowItem:
    case "LoopFoldersDialog":   Call LoopFoldersDialog:
    case "LoopToDoItems":       Call LoopToDoItems:
    case "MPEmap":              Call MPEmap:
    case "NoDupes":             Call NoDupes:
    case "RunMissedRules":      Call RunMissedRules:
    case "ShowErrStack":        Call ShowErrStack:
    case "ShowCallTrace":       Call ShowCallTrace:
    case "ShowDbgStatus":       Call ShowDbgStatus:
    case "ShowDefProcs":        Call ShowDefProcs:
    case "ShowErr":             Call ShowErr:
    case "ShowErrInterface":    Call ShowErrInterface:
    case "ShowErrorStatus":     Call ShowErrorStatus:
    case "ShowLiveStack":       Call ShowLiveStack:
    case "ShowLog":             Call ShowLog:
    case "ShowStacks":          Call ShowStacks:
    case "StartUp":             Call StartUp:
    case "WasEmailProcessed":   Call WasEmailProcessed:
    case "ContactFixer":        Call ContactFixer:
    case _:
# DoVerify False, "invalid selection " & LPMacroListe.Value

# '---------------------------------------------------------------------------------------
# ' Method : Sub UserForm_Activate
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub UserForm_Activate()
# Me.Run.Default = True
# Me.Cancel.Default = False

# Private Sub UserForm_Initialize()
# Dim DftPos As Long
# Dim i As Long

if LPMacroListe.ListCount = 0 Then:
# LPMacroListe.Clear
if isEmpty(MacroArray) Then:
# MacroArray = Array("C1SI", "C2SI", "ChangeBirthdayOf", "CopyAllBackupCats", _
# "CopyAllHotmailCats", "CreateRules", "ExcelShowItem", _
# "LoopFoldersDialog", "LoopToDoItems", "MPEmap", "MacroSelRun", _
# "NoDupes", "RunMissedRules", "ShowErrStack", "ShowAppStack", _
# "ShowCallTrace", "ShowDbgStatus", "ShowDefProcs", "ShowErr", _
# "ShowErrInterface", "ShowErrorStatus", "ShowLiveStack", "ShowLog", _
# "ShowStacks", "StartUp", "WasEmailProcessed", "ContactFixer")
# LPMacroListe.addItem MacroArray(i), i

if LPMacroListe.ListIndex <= 0 Then          ' no choice so far, default one:
if MacroArray(i) = "ExcelShowItem" Then:
# DftPos = i
# Exit For
# LPMacroListe.ListIndex = DftPos
# rsp = vbIgnore

# '---------------------------------------------------------------------------------------
# ' Method : Sub LPMacroListe_Change
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub LPMacroListe_Change()
print(Debug.Print "Aktuelle Wahl zur Ausfhrung: " _)
# & LPMacroListe.Value & " Index=" & LPMacroListe.ListIndex

