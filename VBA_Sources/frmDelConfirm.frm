VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelConfirm 
   Caption         =   "Entscheidung über das Löschen treffen"
   ClientHeight    =   9450.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   OleObjectBlob   =   "frmDelConfirm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDelConfirm"
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
    rsp = vbCancel
    Me.Hide
End Sub                                          ' frmDelConfirm.Cancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub DisplayItem
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub DisplayItem(px As Long)
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmDelConfirm.DisplayItem"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="")

Dim aktItem As Object

    Set aktItem = aID(px).idObjItem
    ' Wenn Outlook-Element existiert, dann...
    If Not aktItem Is Nothing Then
        Me.Hide
        ' ... Outlook-Element anzeigen
        aktItem.Display True
        If DebugControlsUsable Then
            Me.Show
        End If
        Call ModItem(aktItem, _
                     "FirstName", _
                     "LastName", _
                     "*,")
    End If

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelConfirm.DisplayItem

'---------------------------------------------------------------------------------------
' Method : Sub Go_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Go_Click()
    rsp = vbYes
    Me.Hide
End Sub                                          ' frmDelConfirm.Go_Click

'---------------------------------------------------------------------------------------
' Method : Sub Label3_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Label3_Click()                       ' "Löschliste bisher" ganz unten
    DeleteNow = Not DeleteNow
    If DeleteNow Then
        Label3.BackColor = 0
    Else
        Label3.BackColor = -2147483633
    End If
End Sub                                          ' frmDelConfirm.Label3_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPAskEveryFolder_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPAskEveryFolder_Click()
    AskEveryFolder = LPAskEveryFolder
    If AskEveryFolder Then
        LPWantConfirmationThisFolder.Caption = _
                                             "Löschen nur nach Einzelbestätigung," _
                                           & vbCrLf & "auch wenn identisch"
    Else
        LPWantConfirmationThisFolder.Caption = "ohne Rückfrage:" _
                                             & vbCrLf & killType
    End If
End Sub                                          ' frmDelConfirm.LPAskEveryFolder_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPMaxMisMatchesForCandidates_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPMaxMisMatchesForCandidates_AfterUpdate()
    If IsNumeric(LPMaxMisMatchesForCandidates) Then
        If LPMaxMisMatchesForCandidates < 0 Then
            LPMaxMisMatchesForCandidates = 0
        End If
        AcceptCloseMatches = LPMaxMisMatchesForCandidates > 0
        MaxMisMatchesForCandidates = LPMaxMisMatchesForCandidates
    Else
        LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
    End If
End Sub                                          ' frmDelConfirm.LPMaxMisMatchesForCandidates_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub LPWantConfirmation_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPWantConfirmation_Click()
    WantConfirmation = LPWantConfirmation
End Sub                                          ' frmDelConfirm.LPWantConfirmation_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPWantConfirmationThisFolder_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPWantConfirmationThisFolder_Click()
    WantConfirmationThisFolder = LPWantConfirmationThisFolder
End Sub                                          ' frmDelConfirm.LPWantConfirmationThisFolder_Click

'---------------------------------------------------------------------------------------
' Method : Sub MoreParams_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub MoreParams_Click()
    askforParams = Not askforParams
    Me.Hide
End Sub                                          ' frmDelConfirm.MoreParams_Click

'---------------------------------------------------------------------------------------
' Method : Sub No_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub No_Click()
    rsp = vbNo
    Me.Hide
End Sub                                          ' frmDelConfirm.No_Click

'---------------------------------------------------------------------------------------
' Method : Sub S1_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub S1_Click()
    S1.Enabled = False
    S2.Enabled = True
    S2.Value = False
    Call DisplayItem(1)
    S1.Enabled = True
    S1.Value = False
End Sub                                          ' frmDelConfirm.S1_Click

'---------------------------------------------------------------------------------------
' Method : Sub S2_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub S2_Click()
    S2.Enabled = False
    S1.Enabled = True
    S1.Value = False
    Call DisplayItem(2)
    S2.Enabled = True
    S2.Value = False
End Sub                                          ' frmDelConfirm.S2_Click

'---------------------------------------------------------------------------------------
' Method : Sub SucheBeenden_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub SucheBeenden_Click()
    StopLoop = Not StopLoop
    If StopLoop Then
        SucheBeenden.Caption = "fortfahren"
    Else
        SucheBeenden.Caption = "Suche beenden"
    End If
End Sub                                          ' frmDelConfirm.SucheBeenden_Click

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmDelConfirm.UserForm_Activate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)
    
    Go.Caption = b1text
    Go.Visible = (b1text <> vbNullString)
    No.Caption = b2text
    No.Visible = (b2text <> vbNullString)
    Cancel.Caption = b3text
    LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
    LPMaxMisMatchesForCandidates.Enabled = AcceptCloseMatches
    LPWantConfirmation = WantConfirmation
    LPWantConfirmationThisFolder = WantConfirmationThisFolder
    If AskEveryFolder Then
        LPWantConfirmationThisFolder.Caption = _
                                             "Löschen nur nach Einzelbestätigung, auch wenn identisch:" _
                                           & vbCrLf & killType
    Else
        LPWantConfirmationThisFolder.Caption = "ohne Rückfrage:" _
                                             & vbCrLf & killType
    End If
    LPAskEveryFolder = AskEveryFolder
    LPDiffs = Diffs
    LPmsg.Text = Message
    If LenB(LoeschbestätigungCaption) > 0 Then
        Me.Caption = LoeschbestätigungCaption
    End If
    Me.LBlöliste = Mid(LöListe, 2)
    Me.Go.Default = True
    Me.No.Default = False
    Me.Cancel.Default = False
    
    Me.Controls(bDefaultButton).Default = True   ' funktioniert nicht wenn label falsch
    If askforParams Then
        MoreParams.Caption = "Params OK"
        askforParams = False
    Else
        MoreParams.Caption = "Weitere Parameter"
    End If
    Me.SucheBeenden = False
    S1.Enabled = True
    S1.Value = False                             ' wäre schöner, wenn wir nicht modal schaffen könnten
    S2.Enabled = True
    S2.Value = False                             ' dann könnten wir sehen, was schon aktiv ist....
    If WindowSetForeground(LBF.Caption, FWP_LBF_Hdl) Then
        Debug.Print Quote1(Trim(LBF.Caption)) _
      & "  wird als oberstes Fenster angezeigt."
        GoTo FuncExit
    Else
        Debug.Print Quote1(Trim(LBF.Caption)) & "  wird nicht angezeigt"
    End If

FuncExit:
    Call ErrReset(0)

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelConfirm.UserForm_Activate

Private Sub UserForm_Initialize()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmDelConfirm.UserForm_Initialize"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="frmDelConfirm")

    S1.Enabled = True
    S2.Enabled = False
    S1.Value = False
    S2.Value = False
    If LpLogLevel.ListCount = 0 Then
        Call LPlogLevel_define(Me)
    End If
    LpLogLevel.Text = LpLogLevel.List(MinimalLogging + 1)
    If LpLogLevel.ListIndex <= 0 Then            ' no choice so far, default one
        LpLogLevel.Text = LpLogLevel.List(eLmin + 1)
        MinimalLogging = LpLogLevel.ListIndex
    End If
    Me.LpLogLevel = LogSelection
    
    If xlApp Is Nothing Then
        If displayInExcel Then
            Call XlgetApp
            xShowAllExcel.Visible = True
        Else
            xShowAllExcel.Visible = False
        End If
    Else
        If O Is Nothing Or Not displayInExcel Then
            xShowAllExcel.Visible = True
        Else
            xShowAllExcel.Visible = False
        End If
    End If
    rsp = vbIgnore

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelConfirm.UserForm_Initialize

'---------------------------------------------------------------------------------------
' Method : Sub xShowAllExcel_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub xShowAllExcel_Click()
    Me.Hide
    rsp = vbRetry
End Sub                                          ' frmDelConfirm.xShowAllExcel_Click

'---------------------------------------------------------------------------------------
' Method : Sub LpLogLevel_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LpLogLevel_AfterUpdate()

Const zKey As String = "frmDelConfirm.LpLogLevel_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmDelConfirm")

    Call LogLevelChecks(Me)

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelConfirm.LpLogLevel_AfterUpdate


