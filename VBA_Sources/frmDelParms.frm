VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelParms 
   Caption         =   "Parameter für das Vergleichen/Löschen von Doubletten"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295.001
   OleObjectBlob   =   "frmDelParms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDelParms"
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
    Call ShowOrHideForm(Me, ShowIt:=False)
End Sub                                          ' frmDelParms.Cancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub DatumsBedingung_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub DatumsBedingung_AfterUpdate()

Const zKey As String = "frmDelParms.DatumsBedingung_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmDelParms")

Dim toDay As Date
    CutOffDate = CDate("00:00:00")
    LPmsg.Caption = vbNullString
    toDay = CDate(Left(Now(), 10))
    Select Case Datumsbedingung.Text
    Case "keine Datumsbeschränkung"
    Case "heute"
        CutOffDate = DateAdd("d", -0, toDay)
    Case "ab gestern"
        CutOffDate = DateAdd("d", -1, toDay)
    Case "letzte Woche"
        CutOffDate = DateAdd("d", -7, toDay)
    Case "letzte 30 Tage"
        CutOffDate = DateAdd("d", -30, toDay)
    Case Else
        If IsDate(Datumsbedingung) Then
            CutOffDate = CDate(Datumsbedingung)
        Else
            LPmsg.Caption = "Unzulässige Datumsbedingung: " _
                          & Datumsbedingung
        End If
    End Select

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelParms.DatumsBedingung_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub bDebugStop_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub bDebugStop_Click()
    rsp = vbNo
    Call ShowOrHideForm(Me, ShowIt:=False)
    If StopLoop Or bDebugStop.Caption = "DoVerify" Then
        DoVerify False, " You pressed the debug stop button! F5 or F8 to continue"
        b3text = vbNullString
    End If
End Sub                                          ' frmDelParms.bDebugStop_Click

'---------------------------------------------------------------------------------------
' Method : Sub Go_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Go_Click()
    rsp = vbYes
    Call ShowOrHideForm(Me, ShowIt:=False)
End Sub                                          ' frmDelParms.Go_Click

'---------------------------------------------------------------------------------------
' Method : Sub LpLogLevel_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LpLogLevel_AfterUpdate()
    Call LogLevelChecks(Me)
End Sub                                          ' frmDelParms.LpLogLevel_AfterUpdate

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
                                             "Löschen nur nach Einzelbestätigung, auch wenn identisch:" _
                                           & vbCrLf & killType
    Else
        LPWantConfirmationThisFolder.Caption = "ohne Rückfrage:" _
                                             & vbCrLf & killType
    End If
End Sub                                          ' frmDelParms.LPAskEveryFolder_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPErgebnisseAlsListe_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPErgebnisseAlsListe_Click()
    ErgebnisseAlsListe = LPErgebnisseAlsListe
End Sub                                          ' frmDelParms.LPErgebnisseAlsListe_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPMaxMisMatchesForCandidates_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPMaxMisMatchesForCandidates_AfterUpdate()
    If IsNumeric(LPMaxMisMatchesForCandidates) Then
        MaxMisMatchesForCandidates = LPMaxMisMatchesForCandidates
    Else
        LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
    End If
End Sub                                          ' frmDelParms.LPMaxMisMatchesForCandidates_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub LPOFast_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPOFast_Click()
    quickChecksOnly = LPOFast
End Sub                                          ' frmDelParms.LPOFast_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPOFolderChoice_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPOFolderChoice_Click()
    PickTopFolder = LPOFolderChoice
End Sub                                          ' frmDelParms.LPOFolderChoice_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPOInformative_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPOInformative_Click()
    quickChecksOnly = Not LPOInformative
End Sub                                          ' frmDelParms.LPOInformative_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPOWalk_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPOWalk_Click()
    PickTopFolder = Not LPOWalk
End Sub                                          ' frmDelParms.LPOWalk_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPWantConfirmation_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPWantConfirmation_Click()
    WantConfirmation = LPWantConfirmation
End Sub                                          ' frmDelParms.LPWantConfirmation_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPWantConfirmationThisFolder_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPWantConfirmationThisFolder_Click()
    WantConfirmationThisFolder = LPWantConfirmationThisFolder
End Sub                                          ' frmDelParms.LPWantConfirmationThisFolder_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPSaveAttachments_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPSaveAttachments_Click()
    SaveAttachments = LPSaveAttachments
End Sub                                          ' frmDelParms.LPSaveAttachments_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPNoShowEmptyAttrs_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPNoShowEmptyAttrs_Click()
    ShowEmptyAttributes = Not LPNoShowEmptyAttrs
End Sub                                          ' frmDelParms.LPNoShowEmptyAttrs_Click

'---------------------------------------------------------------------------------------
' Method : Sub SelektionModifizieren_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub SelektionModifizieren_Click()
    frmSelParms.Show
    Cancel = frmSelParms.Cancel
    bDebugStop = frmSelParms.bDebugStop
    If Not sRules Is Nothing Then
        LPmsg = "Auswahlkriterien geändert, sind nun " _
              & sRules.clsObligMatches.aRuleString
    End If
    Go.SetFocus
End Sub                                          ' frmDelParms.SelektionModifizieren_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPSkipDontCompare_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPSkipDontCompare_Click()
    SkipDontCompare = LPSkipDontCompare
End Sub                                          ' frmDelParms.LPSkipDontCompare_Click

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
        bDebugStop = True
    Else
        SucheBeenden.Caption = "Suche beenden"
    End If
End Sub                                          ' frmDelParms.SucheBeenden_Click

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmDelParms.UserForm_Activate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    Go.Caption = b1text
    Go.Visible = (b1text <> vbNullString)
    bDebugStop.Caption = b2text
    bDebugStop.Visible = (b2text <> vbNullString)
    If LenB(b3text) = 0 Then
        b3text = "Cancel"
    End If
    Cancel.Caption = b3text
    LPMaxMisMatchesForCandidates = MaxMisMatchesForCandidates
    LPErgebnisseAlsListe = ErgebnisseAlsListe
    AcceptCloseMatches = LPMaxMisMatchesForCandidates > 0
    LpLogLevel.Text = LpLogLevel.List(MinimalLogging)
    If LpLogLevel.ListIndex <= 0 Then            ' no choice so far, default one
        LpLogLevel.Text = LpLogLevel.List(eLmin)
        MinimalLogging = LpLogLevel.ListIndex
    End If
    LPWantConfirmation = WantConfirmation
    LPWantConfirmationThisFolder = WantConfirmationThisFolder
    
    LPOInformative = Not quickChecksOnly
    LPOFast = quickChecksOnly
    
    LPOWalk = Not PickTopFolder
    LPOFolderChoice = PickTopFolder
    Call Combo_Define_DatumsBedingungen(Datumsbedingung)
    
    If AskEveryFolder Then
        LPWantConfirmationThisFolder.Caption = _
                                             "Löschen nur nach Einzelbestätigung, auch wenn identisch:" _
                                           & vbCrLf & killType
    Else
        LPWantConfirmationThisFolder.Caption = "ohne Rückfrage:" _
                                             & vbCrLf & killType
    End If
    LPSaveAttachments = SaveAttachments
    LPAskEveryFolder = AskEveryFolder
    
    xlShow = xUseExcel
    XlDeferred = xDeferExcel
    If ActionID <> 0 And (xlShow Or XlDeferred) Then
        LPOFast = False                          ' excel requires all attributes
        quickChecksOnly = False
    End If
    If Not sRules Is Nothing Then
        Me.LPmsg = Message & vbCrLf & "aktuelle Auswahl lautet: " & sRules.clsObligMatches.aRuleString
    End If
    LPmsg.Caption = Message
    bDebugStop.Default = False
    Controls(bDefaultButton).Default = True

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmDelParms.UserForm_Activate

Private Sub UserForm_Initialize()
    Call LPlogLevel_define(Me)
End Sub                                          ' frmDelParms.UserForm_Initialize

'---------------------------------------------------------------------------------------
' Method : Sub XlDeferred_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub XlDeferred_Click()
    xDeferExcel = XlDeferred
End Sub                                          ' frmDelParms.XlDeferred_Click

'---------------------------------------------------------------------------------------
' Method : Sub xlShow_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub xlShow_Click()
    xUseExcel = xlShow
End Sub                                          ' frmDelParms.xlShow_Click


