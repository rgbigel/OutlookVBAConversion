VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMacroSelRun 
   Caption         =   "Liste verfügbarer Macros"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900.001
   OleObjectBlob   =   "frmMacroSelRun.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMacroSelRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Sub Cancel_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Cancel_Click()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "frmMacroSelRun.Cancel_Click"

    rsp = vbNo
    Me.Hide
    
End Sub                                          ' frmMacroSelRun.Cancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub Run_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Run_Click()
    rsp = vbYes
    
    If Not fKeepOpen Then
        Me.Hide
    End If
    Select Case LPMacroListe.Value
    Case "C1SI":                Call C1SI
    Case "C2SI":                Call C2SI
    Case "ChangeBirthdayOf":    Call ChangeBirthdaySubject("#")
    Case "CopyAllBackupCats":   Call CopyAllBackupCats
    Case "CopyAllHotmailCats":  Call CopyAllHotmailCats
    Case "CreateRules":         Call CreateRules
    Case "ExcelShowItem":       Call ExcelShowItem
    Case "LoopFoldersDialog":   Call LoopFoldersDialog
    Case "LoopToDoItems":       Call LoopToDoItems
    Case "MPEmap":              Call MPEmap
    Case "NoDupes":             Call NoDupes
    Case "RunMissedRules":      Call RunMissedRules
    Case "ShowErrStack":        Call ShowErrStack
    Case "ShowCallTrace":       Call ShowCallTrace
    Case "ShowDbgStatus":       Call ShowDbgStatus
    Case "ShowDefProcs":        Call ShowDefProcs
    Case "ShowErr":             Call ShowErr
    Case "ShowErrInterface":    Call ShowErrInterface
    Case "ShowErrorStatus":     Call ShowErrorStatus
    Case "ShowLiveStack":       Call ShowLiveStack
    Case "ShowLog":             Call ShowLog
    Case "ShowStacks":          Call ShowStacks
    Case "StartUp":             Call StartUp
    Case "WasEmailProcessed":   Call WasEmailProcessed
    Case "ContactFixer":        Call ContactFixer
    Case Else
        DoVerify False, "invalid selection " & LPMacroListe.Value
    End Select
End Sub                                          ' frmMacroSelRun.Run_Click

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    Me.Run.Default = True
    Me.Cancel.Default = False
End Sub                                          ' frmMacroSelRun.UserForm_Activate

Private Sub UserForm_Initialize()
Dim DftPos As Long
Dim i As Long

    If LPMacroListe.ListCount = 0 Then
        LPMacroListe.Clear
        If isEmpty(MacroArray) Then
            MacroArray = Array("C1SI", "C2SI", "ChangeBirthdayOf", "CopyAllBackupCats", _
                               "CopyAllHotmailCats", "CreateRules", "ExcelShowItem", _
                               "LoopFoldersDialog", "LoopToDoItems", "MPEmap", "MacroSelRun", _
                               "NoDupes", "RunMissedRules", "ShowErrStack", "ShowAppStack", _
                               "ShowCallTrace", "ShowDbgStatus", "ShowDefProcs", "ShowErr", _
                               "ShowErrInterface", "ShowErrorStatus", "ShowLiveStack", "ShowLog", _
                               "ShowStacks", "StartUp", "WasEmailProcessed", "ContactFixer")
        End If
        For i = 0 To UBound(MacroArray)
            LPMacroListe.addItem MacroArray(i), i
        Next i
    End If
    
    If LPMacroListe.ListIndex <= 0 Then          ' no choice so far, default one
        For i = 0 To UBound(MacroArray)
            If MacroArray(i) = "ExcelShowItem" Then
                DftPos = i
                Exit For
            End If
        Next i
        LPMacroListe.ListIndex = DftPos
    End If
    rsp = vbIgnore
End Sub                                          ' frmMacroSelRun.UserForm_Initialize

'---------------------------------------------------------------------------------------
' Method : Sub LPMacroListe_Change
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPMacroListe_Change()
    Debug.Print "Aktuelle Wahl zur Ausführung: " _
              & LPMacroListe.Value & " Index=" & LPMacroListe.ListIndex
End Sub                                          ' frmMacroSelRun.LPMacroListe_Change


