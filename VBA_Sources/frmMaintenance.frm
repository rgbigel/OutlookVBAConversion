VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMaintenance 
   Caption         =   "Wartung"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9780.001
   OleObjectBlob   =   "frmMaintenance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public someAction As Long

'---------------------------------------------------------------------------------------
' Method : Sub ActtabMod_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub ActtabMod_Click()
    someAction = 2
End Sub                                          ' frmMaintenance.ActtabMod_Click

'---------------------------------------------------------------------------------------
' Method : Sub Cancel_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Cancel_Click()
    someAction = 0
    Call UserForm_Terminate
End Sub                                          ' frmMaintenance.Cancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub OK_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub OK_Click()
    Select Case someAction
    Case 0
        If MsgBox("Eine Aktion muss ausgewählt werden.", vbOKCancel, _
                  "Auswahl einer Aktion steht aus") _
        = vbCancel Then
            Call UserForm_Terminate
        End If
    Case 1
        MaintenanceAction = someAction
        Call EditRulesTable
    Case 2
        MaintenanceAction = someAction
        frmFolderLoop.InitActions
    Case Else
        Call MsgBox("Diese Aktion wurde noch nicht definiert.", vbOKCancel, _
                    "Definition einer Aktion fehlt")
        someAction = 0
        MaintenanceAction = someAction
    End Select                                   ' someAction
    Call UserForm_Terminate
End Sub                                          ' frmMaintenance.OK_Click

'---------------------------------------------------------------------------------------
' Method : Sub RuleTableMod_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub RuleTableMod_Click()
    someAction = 1
End Sub                                          ' frmMaintenance.RuleTableMod_Click

Private Sub UserForm_Initialize()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmMaintenance.UserForm_Initialize"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    RuleTableMod.Value = True

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmMaintenance.UserForm_Initialize

Private Sub UserForm_Terminate()
'''' Proc Must ONLY CALL Z_Type PROCS                         ' trivial proc
Const zKey As String = "frmMaintenance.UserForm_Terminate"
    
    MaintenanceAction = 0
    Me.Hide
    
End Sub                                          ' frmMaintenance.UserForm_Terminate


