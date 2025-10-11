VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelParms 
   Caption         =   "Sortier- und Auswahlkriterien"
   ClientHeight    =   9510.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295.001
   OleObjectBlob   =   "frmSelParms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelParms"
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
End Sub                                          ' frmSelParms.Cancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub bDebugStop_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub bDebugStop_Click()
    rsp = vbNo
    Me.Hide
    If bDebugStop.Caption = "DoVerify" Then
        BugHelp.DoVerify False, "You pressed the debug button! F5 or F8 to continue"
        b3text = vbNullString
    End If
End Sub                                          ' frmSelParms.bDebugStop_Click

'---------------------------------------------------------------------------------------
' Method : Sub Go_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub Go_Click()
    rsp = vbYes
    Me.Hide
End Sub                                          ' frmSelParms.Go_Click

'---------------------------------------------------------------------------------------
' Method : Sub LPDontCompareList_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPDontCompareList_AfterUpdate()

Const zKey As String = "frmSelParms.LPDontCompareList_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmSelParms")

    UserRule.clsNeverCompare.ChangeTo = LPDontCompareList

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmSelParms.LPDontCompareList_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub LPMandatoryMatches_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPMandatoryMatches_AfterUpdate()

Const zKey As String = "frmSelParms.LPMandatoryMatches_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmSelParms")

    If InStr(LCase(Trim(LPMandatoryMatches)), "none") > 0 Then
        LPMandatoryMatches = vbNullString
        LPSimilarities = "none"
        UserRule.clsSimilarities.ChangeTo = vbNullString
    End If
    UserRule.RuleInstanceValid = False
    Call SplitMandatories(LPMandatoryMatches)
    Set UserRule = sRules

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmSelParms.LPMandatoryMatches_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub LPNotDecodablePropertiesList_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPNotDecodablePropertiesList_AfterUpdate()

Const zKey As String = "frmSelParms.LPNotDecodablePropertiesList_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmSelParms")

    UserRule.clsNotDecodable.ChangeTo = LPNotDecodablePropertiesList

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmSelParms.LPNotDecodablePropertiesList_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub LPSimilarities_AfterUpdate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub LPSimilarities_AfterUpdate()

Const zKey As String = "frmSelParms.LPSimilarities_AfterUpdate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQyMode, CallType:=tSub, ExplainS:="frmSelParms")

    UserRule.clsSimilarities.ChangeTo = LPSimilarities

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmSelParms.LPSimilarities_AfterUpdate

'---------------------------------------------------------------------------------------
' Method : Sub UserForm_Activate
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "frmSelParms.UserForm_Activate"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub)

    Go.Caption = b1text
    Go.Visible = (b1text <> vbNullString)
    LPMandatoryMatches.SetFocus
    bDebugStop.Caption = b2text
    bDebugStop.Visible = (b2text <> vbNullString)
    If LenB(b3text) = 0 Then
        b3text = "Cancel"
    End If
    Cancel.Caption = b3text
    
    Set UserRule = BestRule()
    If UserRule Is Nothing Then
        If DebugMode Then
            BugHelp.DoVerify False
        End If
    Else
        With UserRule
            LPDontCompareList = DontCompareListDefault ' never init with old user data
            ' init with confirmed or defaulted user data
            LPNotDecodablePropertiesList = Trim(.clsNotDecodable.aRuleString)
            LPSimilarities = Trim(.clsSimilarities.aRuleString)
            LPMandatoryMatches = Trim(.clsObligMatches.aRuleString)
        End With                                 ' UserRule
    End If
    LPmsg.Caption = Message
    bDebugStop.Default = False
    Me.Controls(bDefaultButton).Default = True

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmSelParms.UserForm_Activate

